export const BRIDGE_VERSION = 1;
export const CUSTOM_XML_NAMESPACE = 'urn:montecarlo-xl:config:v1';
export const HIDDEN_CONFIG_SHEET = '__MC_Config';

export type SamplingMethod = 'MonteCarlo' | 'LatinHypercube';
export type SeedMode = 'Random' | 'Fixed';
export type GoalSeekMetric =
  | 'Mean'
  | 'ProbabilityAboveTarget'
  | 'ProbabilityAtOrBelowTarget'
  | 'Percentile';

export interface SavedInput {
  sheetName: string;
  cellAddress: string;
  label: string;
  distributionName: string;
  parameters: Record<string, number>;
}

export interface SavedOutput {
  sheetName: string;
  cellAddress: string;
  label: string;
}

export interface SimulationProfile {
  name: string;
  iterationCount: number;
  randomSeed?: number | null;
  inputs: SavedInput[];
  outputs: SavedOutput[];
  correlationMatrixFlat?: number[] | null;
  correlationMatrixSize: number;
}

export interface WorkbookUserSettingsOverrides {
  createNewWorksheetForExports?: boolean;
  defaultIterationCount?: number;
  seedMode?: SeedMode;
  fixedRandomSeed?: number;
  samplingMethod?: SamplingMethod;
  autoStopOnConvergence?: boolean;
  pauseOnPreflightWarnings?: boolean;
  defaultPercentiles?: string;
}

export interface UserSettings {
  createNewWorksheetForExports: boolean;
  defaultIterationCount: number;
  seedMode: SeedMode;
  fixedRandomSeed: number;
  samplingMethod: SamplingMethod;
  autoStopOnConvergence: boolean;
  pauseOnPreflightWarnings: boolean;
  defaultPercentiles: string;
  theme: 'system' | 'light' | 'dark';
}

export interface OfficeRunSettings {
  iterationCount: number;
  randomSeed: number | null;
  samplingMethod: SamplingMethod;
  autoStopOnConvergence: boolean;
  outputTarget?: number | null;
}

export interface FormulaArgumentDto {
  name: string;
  displayName: string;
  description: string;
}

export interface FormulaDefinitionDto {
  functionName: string;
  excelName: string;
  distributionName: string;
  description: string;
  arguments: FormulaArgumentDto[];
}

export interface FormulaCatalogResponse {
  version: number;
  functions: FormulaDefinitionDto[];
}

export interface ParseFormulaResponse {
  version: number;
  isMatch: boolean;
  functionName?: string | null;
  distributionName?: string | null;
  rawArguments?: string[] | null;
  parameterNames?: string[] | null;
}

export interface EvaluateFormulaResponse {
  version: number;
  isValid: boolean;
  value?: number | null;
}

export interface PreflightIssueDto {
  severity: string;
  code: string;
  title: string;
  message: string;
  suggestedAction: string;
}

export interface ValidateProfileResponse {
  version: number;
  issues: PreflightIssueDto[];
}

export interface GenerateInputMatrixResponse {
  version: number;
  inputIds: string[];
  outputIds: string[];
  inputMatrix: number[][];
}

export interface SummaryStatisticsDto {
  count: number;
  mean: number;
  median: number;
  mode: number;
  stdDev: number;
  variance: number;
  min: number;
  max: number;
  p1: number;
  p5: number;
  p10: number;
  p25: number;
  p50: number;
  p75: number;
  p90: number;
  p95: number;
  p99: number;
  probabilityAboveTarget: number;
  probabilityAtOrBelowTarget: number;
}

export interface HistogramDto {
  binEdges: number[];
  binCenters: number[];
  frequencies: number[];
  relativeFrequencies: number[];
  kdeX: number[];
  kdeY: number[];
}

export interface CdfDto {
  x: number[];
  y: number[];
}

export interface SensitivityItemDto {
  inputId: string;
  inputLabel: string;
  rankCorrelation: number;
  standardizedRegression: number;
}

export interface ScenarioInputSummaryDto {
  inputId: string;
  inputLabel: string;
  overallMean: number;
  scenarioMean: number;
  delta: number;
  deltaPercent?: number | null;
  overallP10: number;
  overallP50: number;
  overallP90: number;
  scenarioP10: number;
  scenarioP50: number;
  scenarioP90: number;
}

export interface ScenarioSummaryDto {
  description: string;
  threshold: number;
  matchedIterations: number;
  matchedFraction: number;
  inputs: ScenarioInputSummaryDto[];
}

export interface TargetAnalysisDto {
  target: number;
  probabilityAbove: number;
  probabilityAtOrBelow: number;
}

export interface OutputAnalysisDto {
  outputId: string;
  outputLabel: string;
  summary: SummaryStatisticsDto;
  histogram: HistogramDto;
  cdf: CdfDto;
  sensitivity: SensitivityItemDto[];
  worstCase: ScenarioSummaryDto;
  bestCase: ScenarioSummaryDto;
  target?: TargetAnalysisDto | null;
}

export interface AnalyzeSimulationResponse {
  version: number;
  iterationCount: number;
  elapsedMilliseconds: number;
  outputs: OutputAnalysisDto[];
}

export interface BenchmarkResponse {
  version: number;
  inputCount: number;
  iterationCount: number;
  iterationsPerSecond: number;
  microsecondsPerIteration: number;
  totalMilliseconds: number;
}

export interface WorkbookEnvelope {
  profile: SimulationProfile | null;
  workbookSettingsOverrides: WorkbookUserSettingsOverrides | null;
}
