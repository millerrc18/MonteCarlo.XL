/// <reference types="office-js" />
import { MonteCarloBridgeClient } from './bridge';
import {
  BRIDGE_VERSION,
  CUSTOM_XML_NAMESPACE,
  HIDDEN_CONFIG_SHEET,
  type AnalyzeSimulationResponse,
  type BenchmarkResponse,
  type FormulaDefinitionDto,
  type OfficeRunSettings,
  type OutputAnalysisDto,
  type ParseFormulaResponse,
  type SavedInput,
  type SavedOutput,
  type SimulationProfile,
  type UserSettings,
  type ValidateProfileResponse,
  type WorkbookEnvelope,
  type WorkbookUserSettingsOverrides,
} from './types';

const USER_SETTINGS_KEY = 'montecarlo-xl.office.user-settings';

export const DEFAULT_USER_SETTINGS: UserSettings = {
  createNewWorksheetForExports: true,
  defaultIterationCount: 5000,
  seedMode: 'Random',
  fixedRandomSeed: 42,
  samplingMethod: 'LatinHypercube',
  autoStopOnConvergence: false,
  pauseOnPreflightWarnings: true,
  defaultPercentiles: '1,5,10,25,50,75,90,95,99',
  theme: 'system',
};

type WorkbookStoredValue = {
  formula: string | null;
  value: unknown;
};

type RunExecutionResult = {
  generatedInputMatrix: number[][];
  outputMatrix: number[][];
  elapsedMilliseconds: number;
};

export function loadUserSettings(): UserSettings {
  try {
    const raw = window.localStorage.getItem(USER_SETTINGS_KEY);
    if (!raw) {
      return DEFAULT_USER_SETTINGS;
    }

    return {
      ...DEFAULT_USER_SETTINGS,
      ...(JSON.parse(raw) as Partial<UserSettings>),
    };
  } catch {
    return DEFAULT_USER_SETTINGS;
  }
}

export function saveUserSettings(settings: UserSettings): void {
  window.localStorage.setItem(USER_SETTINGS_KEY, JSON.stringify(settings));
}

export function resolveEffectiveSettings(
  globalSettings: UserSettings,
  workbookOverrides: WorkbookUserSettingsOverrides | null,
): UserSettings {
  return {
    ...globalSettings,
    ...(workbookOverrides ?? {}),
  };
}

export function createEmptyProfile(iterationCount: number): SimulationProfile {
  return {
    name: 'Default',
    iterationCount,
    randomSeed: null,
    inputs: [],
    outputs: [],
    correlationMatrixFlat: null,
    correlationMatrixSize: 0,
  };
}

export async function loadWorkbookEnvelope(): Promise<WorkbookEnvelope> {
  const fromCustomXml = await loadFromCustomXml();
  if (fromCustomXml.profile || fromCustomXml.workbookSettingsOverrides) {
    return fromCustomXml;
  }

  const hiddenProfile = await loadProfileFromHiddenSheet();
  return {
    profile: hiddenProfile,
    workbookSettingsOverrides: null,
  };
}

export async function saveWorkbookEnvelope(envelope: WorkbookEnvelope): Promise<void> {
  const xml = wrapInXml(envelope.profile, envelope.workbookSettingsOverrides);

  try {
    await Excel.run(async (context) => {
      const parts = context.workbook.customXmlParts.getByNamespace(CUSTOM_XML_NAMESPACE);
      parts.load('items/id');
      await context.sync();

      for (const part of parts.items) {
        part.delete();
      }

      if (envelope.profile || envelope.workbookSettingsOverrides) {
        context.workbook.customXmlParts.add(xml);
      }

      await context.sync();
    });
  } catch {
    // Custom XML is preferred, but some hosts are finicky. The hidden sheet
    // keeps basic profile compatibility with the Excel-DNA add-in.
  }

  await saveProfileToHiddenSheet(envelope.profile);
}

export async function scanMcInputs(): Promise<SavedInput[]> {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = worksheet.getUsedRangeOrNullObject();
    usedRange.load(['rowCount', 'columnCount', 'formulas']);
    await context.sync();

    if (usedRange.isNullObject) {
      return [];
    }

    const formulas = usedRange.formulas as (string | number | boolean)[][];
    const inputs: SavedInput[] = [];

    for (let row = 0; row < formulas.length; row++) {
      for (let column = 0; column < formulas[row].length; column++) {
        const value = formulas[row][column];
        if (typeof value !== 'string' || !value.startsWith('=MC.')) {
          continue;
        }

        const parsed = await MonteCarloBridgeClient.parseFormula(value);
        if (!parsed.isMatch || !parsed.rawArguments || !parsed.parameterNames || !parsed.distributionName) {
          continue;
        }

        const address = toA1Address(row + 1, column + 1);
        const parameters: Record<string, number> = {};
        let resolved = true;

        for (let index = 0; index < parsed.parameterNames.length; index++) {
          const argument = parsed.rawArguments[index];
          const numeric = Number(argument);
          if (Number.isFinite(numeric)) {
            parameters[parsed.parameterNames[index]] = numeric;
            continue;
          }

          try {
            const reference = worksheet.getRange(argument);
            reference.load('values');
            await context.sync();
            parameters[parsed.parameterNames[index]] = Number(reference.values[0][0]);
          } catch {
            resolved = false;
            break;
          }
        }

        if (!resolved) {
          continue;
        }

        inputs.push({
          sheetName: worksheet.name,
          cellAddress: address,
          label: address,
          distributionName: parsed.distributionName,
          parameters,
        });
      }
    }

    return inputs;
  });
}

export async function getSelectedCell(): Promise<{ sheetName: string; cellAddress: string; formula: string | null; value: unknown }> {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    const range = context.workbook.getSelectedRange();
    range.load(['address', 'formulas', 'values']);
    await context.sync();

    return {
      sheetName: worksheet.name,
      cellAddress: range.address.split('!').pop()?.replace(/\$/g, '') ?? range.address,
      formula: typeof range.formulas?.[0]?.[0] === 'string' ? (range.formulas[0][0] as string) : null,
      value: range.values?.[0]?.[0],
    };
  });
}

export async function addSelectedAsInput(existingInputs: SavedInput[]): Promise<SavedInput[]> {
  const selected = await getSelectedCell();
  let distributionName = 'Normal';
  let parameters: Record<string, number> = {
    mean: Number(selected.value ?? 0),
    stdDev: Math.max(1, Math.abs(Number(selected.value ?? 0)) * 0.1 || 1),
  };

  if (selected.formula?.startsWith('=MC.')) {
    const parsed = await MonteCarloBridgeClient.parseFormula(selected.formula);
    if (parsed.isMatch && parsed.parameterNames && parsed.rawArguments && parsed.distributionName) {
      distributionName = parsed.distributionName;
      parameters = {};
      for (let index = 0; index < parsed.parameterNames.length; index++) {
        parameters[parsed.parameterNames[index]] = Number(parsed.rawArguments[index]);
      }
    }
  }

  const withoutDuplicate = existingInputs.filter(
    (input) => !(input.sheetName === selected.sheetName && input.cellAddress === selected.cellAddress),
  );

  return [
    ...withoutDuplicate,
    {
      sheetName: selected.sheetName,
      cellAddress: selected.cellAddress,
      label: selected.cellAddress,
      distributionName,
      parameters,
    },
  ];
}

export async function addSelectedAsOutput(existingOutputs: SavedOutput[]): Promise<SavedOutput[]> {
  const selected = await getSelectedCell();
  const withoutDuplicate = existingOutputs.filter(
    (output) => !(output.sheetName === selected.sheetName && output.cellAddress === selected.cellAddress),
  );

  return [
    ...withoutDuplicate,
    {
      sheetName: selected.sheetName,
      cellAddress: selected.cellAddress,
      label: selected.cellAddress,
    },
  ];
}

export async function validateProfile(profile: SimulationProfile): Promise<ValidateProfileResponse> {
  return MonteCarloBridgeClient.validateProfile({
    version: BRIDGE_VERSION,
    profile,
  });
}

export async function executeSimulation(
  profile: SimulationProfile,
  settings: OfficeRunSettings,
  onProgress?: (completed: number, total: number) => void,
): Promise<RunExecutionResult> {
  const generation = await MonteCarloBridgeClient.generateInputMatrix({
    version: BRIDGE_VERSION,
    profile,
    settings,
  });

  const outputMatrix: number[][] = [];
  const startedAt = performance.now();

  await Excel.run(async (context) => {
    const workbook = context.workbook;
    const inputRanges = profile.inputs.map((input) =>
      workbook.worksheets.getItem(input.sheetName).getRange(input.cellAddress),
    );
    const outputRanges = profile.outputs.map((output) =>
      workbook.worksheets.getItem(output.sheetName).getRange(output.cellAddress),
    );

    inputRanges.forEach((range) => range.load(['formulas', 'values']));
    outputRanges.forEach((range) => range.load('values'));
    await context.sync();

    const originalInputValues: WorkbookStoredValue[] = inputRanges.map((range) => ({
      formula: typeof range.formulas?.[0]?.[0] === 'string' ? (range.formulas[0][0] as string) : null,
      value: range.values?.[0]?.[0],
    }));

    try {
      for (let iteration = 0; iteration < generation.inputMatrix.length; iteration++) {
        for (let inputIndex = 0; inputIndex < inputRanges.length; inputIndex++) {
          inputRanges[inputIndex].values = [[generation.inputMatrix[iteration][inputIndex]]];
        }

        context.workbook.application.calculate(Excel.CalculationType.full);
        outputRanges.forEach((range) => range.load('values'));
        await context.sync();

        outputMatrix.push(
          outputRanges.map((range) => Number(range.values?.[0]?.[0] ?? 0)),
        );

        if (onProgress && ((iteration + 1) % 25 === 0 || iteration === generation.inputMatrix.length - 1)) {
          onProgress(iteration + 1, generation.inputMatrix.length);
        }
      }
    } finally {
      for (let index = 0; index < inputRanges.length; index++) {
        const stored = originalInputValues[index];
        if (stored.formula && stored.formula.startsWith('=')) {
          inputRanges[index].formulas = [[stored.formula]];
        } else {
          inputRanges[index].values = [[stored.value]];
        }
      }

      context.workbook.application.calculate(Excel.CalculationType.full);
      await context.sync();
    }
  });

  return {
    generatedInputMatrix: generation.inputMatrix,
    outputMatrix,
    elapsedMilliseconds: performance.now() - startedAt,
  };
}

export async function analyzeSimulation(
  profile: SimulationProfile,
  settings: OfficeRunSettings,
  run: RunExecutionResult,
): Promise<AnalyzeSimulationResponse> {
  return MonteCarloBridgeClient.analyzeSimulation({
    version: BRIDGE_VERSION,
    profile,
    settings,
    inputMatrix: run.generatedInputMatrix,
    outputMatrix: run.outputMatrix,
    elapsedMilliseconds: run.elapsedMilliseconds,
  });
}

export async function runBenchmark(inputCount: number, iterationCount: number): Promise<BenchmarkResponse> {
  return MonteCarloBridgeClient.runBenchmark({
    version: BRIDGE_VERSION,
    inputCount,
    iterationCount,
  });
}

export async function exportSummarySheet(
  profile: SimulationProfile,
  results: AnalyzeSimulationResponse,
  createNewSheet: boolean,
): Promise<void> {
  if (results.outputs.length === 0) {
    throw new Error('Run a simulation before exporting results.');
  }

  const primaryOutput = results.outputs[0];

  await Excel.run(async (context) => {
    const workbook = context.workbook;
    const worksheetName = createNewSheet
      ? uniqueSheetName('MC Summary')
      : 'MC Summary';

    let worksheet: Excel.Worksheet;
    if (createNewSheet) {
      worksheet = workbook.worksheets.add(worksheetName);
    } else {
      worksheet = workbook.worksheets.getItemOrNullObject(worksheetName);
      await context.sync();
      worksheet = worksheet.isNullObject ? workbook.worksheets.add(worksheetName) : worksheet;
      const used = worksheet.getUsedRangeOrNullObject();
      used.load('address');
      await context.sync();
      if (!used.isNullObject) {
        used.clear();
      }
    }

    const metadata = [
      ['MonteCarlo.XL Office Host', 'ARM / Office.js foundation'],
      ['Profile', profile.name],
      ['Iterations', results.iterationCount],
      ['Elapsed ms', results.elapsedMilliseconds],
      ['Primary output', primaryOutput.outputLabel],
    ];
    worksheet.getRange('A1:B5').values = metadata;

    const statsRows = [
      ['Metric', 'Value'],
      ['Mean', primaryOutput.summary.mean],
      ['Median', primaryOutput.summary.median],
      ['Std Dev', primaryOutput.summary.stdDev],
      ['Min', primaryOutput.summary.min],
      ['Max', primaryOutput.summary.max],
      ['P10', primaryOutput.summary.p10],
      ['P50', primaryOutput.summary.p50],
      ['P90', primaryOutput.summary.p90],
      ['P95', primaryOutput.summary.p95],
      ['P99', primaryOutput.summary.p99],
    ];
    worksheet.getRange(`A7:B${statsRows.length + 6}`).values = statsRows;

    const histogramRows = [
      ['Bin', 'Frequency', 'Relative Frequency'],
      ...primaryOutput.histogram.binCenters.map((center, index) => [
        center,
        primaryOutput.histogram.frequencies[index],
        primaryOutput.histogram.relativeFrequencies[index],
      ]),
    ];
    const histogramEndRow = histogramRows.length + 6;
    worksheet.getRange(`D7:F${histogramEndRow}`).values = histogramRows;

    const cdfRows = [
      ['Value', 'CDF'],
      ...primaryOutput.cdf.x.map((value, index) => [value, primaryOutput.cdf.y[index]]),
    ];
    const cdfEndRow = cdfRows.length + 6;
    worksheet.getRange(`H7:I${cdfEndRow}`).values = cdfRows;

    const sensitivityRows = [
      ['Input', 'Rank Correlation'],
      ...primaryOutput.sensitivity.map((item) => [item.inputLabel, item.rankCorrelation]),
    ];
    const sensitivityEndRow = sensitivityRows.length + 6;
    worksheet.getRange(`K7:L${sensitivityEndRow}`).values = sensitivityRows;

    const histogramChart = worksheet.charts.add(
      Excel.ChartType.columnClustered,
      worksheet.getRange(`D7:E${histogramEndRow}`),
      Excel.ChartSeriesBy.columns,
    );
    histogramChart.title.text = `${primaryOutput.outputLabel} Histogram`;
    histogramChart.setPosition('D22', 'J38');

    const cdfChart = worksheet.charts.add(
      Excel.ChartType.line,
      worksheet.getRange(`H7:I${cdfEndRow}`),
      Excel.ChartSeriesBy.columns,
    );
    cdfChart.title.text = `${primaryOutput.outputLabel} CDF`;
    cdfChart.setPosition('K22', 'P38');

    const tornadoChart = worksheet.charts.add(
      Excel.ChartType.barClustered,
      worksheet.getRange(`K7:L${sensitivityEndRow}`),
      Excel.ChartSeriesBy.columns,
    );
    tornadoChart.title.text = `${primaryOutput.outputLabel} Sensitivity`;
    tornadoChart.setPosition('Q22', 'W38');

    worksheet.activate();
    await context.sync();
  });
}

function uniqueSheetName(prefix: string): string {
  return `${prefix} ${new Date().toISOString().slice(11, 19).replace(/:/g, '-')}`;
}

async function loadFromCustomXml(): Promise<WorkbookEnvelope> {
  try {
    return await Excel.run(async (context) => {
      const parts = context.workbook.customXmlParts.getByNamespace(CUSTOM_XML_NAMESPACE);
      parts.load('items/id');
      await context.sync();

      if (parts.items.length === 0) {
        return { profile: null, workbookSettingsOverrides: null };
      }

      const xml = parts.items[0].getXml();
      await context.sync();
      return parseEnvelope(xml.value);
    });
  } catch {
    return { profile: null, workbookSettingsOverrides: null };
  }
}

async function loadProfileFromHiddenSheet(): Promise<SimulationProfile | null> {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(HIDDEN_CONFIG_SHEET);
      sheet.load('name');
      await context.sync();
      if (sheet.isNullObject) {
        return null;
      }

      const range = sheet.getRange('A1');
      range.load('values');
      await context.sync();
      const raw = String(range.values?.[0]?.[0] ?? '');
      return raw ? (JSON.parse(raw) as SimulationProfile) : null;
    });
  } catch {
    return null;
  }
}

async function saveProfileToHiddenSheet(profile: SimulationProfile | null): Promise<void> {
  await Excel.run(async (context) => {
    const workbook = context.workbook;
    let sheet = workbook.worksheets.getItemOrNullObject(HIDDEN_CONFIG_SHEET);
    sheet.load('name');
    await context.sync();

    if (profile === null) {
      if (!sheet.isNullObject) {
        sheet.delete();
        await context.sync();
      }
      return;
    }

    if (sheet.isNullObject) {
      sheet = workbook.worksheets.add(HIDDEN_CONFIG_SHEET);
      sheet.visibility = Excel.SheetVisibility.hidden;
    }

    sheet.getRange('A1').values = [[JSON.stringify(profile)]];
    await context.sync();
  });
}

function wrapInXml(
  profile: SimulationProfile | null,
  workbookSettingsOverrides: WorkbookUserSettingsOverrides | null,
): string {
  const parts: string[] = [
    '<?xml version="1.0" encoding="utf-8"?>',
    `<MonteCarloConfig xmlns="${CUSTOM_XML_NAMESPACE}">`,
  ];

  if (profile) {
    parts.push(
      `  <Profile name="${escapeXml(profile.name)}"><![CDATA[${escapeCData(JSON.stringify(profile))}]]></Profile>`,
    );
  }

  if (workbookSettingsOverrides && Object.keys(workbookSettingsOverrides).length > 0) {
    parts.push(
      `  <WorkbookSettings><![CDATA[${escapeCData(JSON.stringify(workbookSettingsOverrides))}]]></WorkbookSettings>`,
    );
  }

  parts.push('</MonteCarloConfig>');
  return parts.join('\n');
}

function parseEnvelope(xml: string): WorkbookEnvelope {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const profileNode = doc.getElementsByTagName('Profile')[0];
  const settingsNode = doc.getElementsByTagName('WorkbookSettings')[0];

  return {
    profile: profileNode?.textContent ? (JSON.parse(profileNode.textContent) as SimulationProfile) : null,
    workbookSettingsOverrides: settingsNode?.textContent
      ? (JSON.parse(settingsNode.textContent) as WorkbookUserSettingsOverrides)
      : null,
  };
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

function escapeCData(value: string): string {
  return value.replaceAll(']]>', ']]]]><![CDATA[>');
}

function toA1Address(row: number, column: number): string {
  let dividend = column;
  let columnName = '';

  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }

  return `${columnName}${row}`;
}

export const distributionUseCases: Record<string, string> = {
  Normal: 'Use for mature processes with symmetric noise, like monthly demand variance around a forecast.',
  Triangular: 'Use when you know a low, most-likely, and high estimate but have limited history.',
  PERT: 'Use for project estimates where you want a smoother shape around the most-likely case.',
  Lognormal: 'Use for right-skewed positive values such as sales, cost overruns, or cycle times.',
  Uniform: 'Use when only a hard min and max are known and every value is equally plausible.',
  Beta: 'Use for ratios and probabilities that must stay between 0 and 1.',
  Weibull: 'Use for reliability, failure timing, and maintenance modeling.',
  Exponential: 'Use for time between random arrivals or failure events.',
  Poisson: 'Use for discrete event counts like claims, tickets, or defects per period.',
  Gamma: 'Use for positive skewed quantities such as repair cost, rainfall, or duration.',
  Logistic: 'Use for outcomes with heavier tails than Normal around a central point.',
  GEV: 'Use for extreme-value modeling like worst-case demand spikes or annual max losses.',
  Binomial: 'Use for success counts out of a fixed number of trials.',
  Geometric: 'Use for trials-until-success questions such as calls until a sale closes.',
};

export function buildRunSettings(settings: UserSettings, outputTargetText: string): OfficeRunSettings {
  const trimmedTarget = outputTargetText.trim();
  const outputTarget = trimmedTarget.length > 0 ? Number(trimmedTarget) : null;

  return {
    iterationCount: settings.defaultIterationCount,
    randomSeed: settings.seedMode === 'Fixed' ? settings.fixedRandomSeed : null,
    samplingMethod: settings.samplingMethod,
    autoStopOnConvergence: settings.autoStopOnConvergence,
    outputTarget: Number.isFinite(outputTarget) ? outputTarget : null,
  };
}

export async function syncAutoDetectedInputs(profile: SimulationProfile): Promise<SimulationProfile> {
  const detectedInputs = await scanMcInputs();
  return {
    ...profile,
    inputs: detectedInputs,
  };
}

export function sortFormulaCatalog(formulas: FormulaDefinitionDto[]): FormulaDefinitionDto[] {
  return [...formulas].sort((left, right) => left.functionName.localeCompare(right.functionName));
}

