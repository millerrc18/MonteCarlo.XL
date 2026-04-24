import type {
  AnalyzeSimulationResponse,
  BenchmarkResponse,
  EvaluateFormulaResponse,
  FormulaCatalogResponse,
  GenerateInputMatrixResponse,
  ParseFormulaResponse,
  ValidateProfileResponse,
} from './types';

declare global {
  interface Window {
    MonteCarloBridge?: {
      ensureReady: () => Promise<unknown>;
      getFormulaCatalog: () => Promise<string>;
      parseFormula: (formula: string) => Promise<string>;
      evaluateFormula: (payload: string) => Promise<string>;
      validateProfile: (payload: string) => Promise<string>;
      generateInputMatrix: (payload: string) => Promise<string>;
      analyzeSimulation: (payload: string) => Promise<string>;
      evaluateGoalSeekMetric: (payload: string) => Promise<string>;
      runBenchmark: (payload: string) => Promise<string>;
    };
  }
}

let readyPromise: Promise<void> | undefined;

async function ensureBridgeLoaded(): Promise<NonNullable<Window['MonteCarloBridge']>> {
  if (!readyPromise) {
    readyPromise = (async () => {
      const bridgeModulePath = '/wasm/main.js';
      await import(/* @vite-ignore */ bridgeModulePath);
      if (!window.MonteCarloBridge) {
        throw new Error('MonteCarlo WebAssembly bridge failed to initialize.');
      }

      await window.MonteCarloBridge.ensureReady();
    })();
  }

  await readyPromise;
  return window.MonteCarloBridge!;
}

async function invokeJson<TResponse>(method: keyof NonNullable<Window['MonteCarloBridge']>, payload?: unknown): Promise<TResponse> {
  const bridge = await ensureBridgeLoaded();
  const raw = payload === undefined
    ? await (bridge[method] as () => Promise<string>)()
    : await (bridge[method] as (json: string) => Promise<string>)(JSON.stringify(payload));
  return JSON.parse(raw) as TResponse;
}

async function invokeRaw<TResponse>(method: keyof NonNullable<Window['MonteCarloBridge']>, payload: string): Promise<TResponse> {
  const bridge = await ensureBridgeLoaded();
  const raw = await (bridge[method] as (value: string) => Promise<string>)(payload);
  return JSON.parse(raw) as TResponse;
}

export const MonteCarloBridgeClient = {
  ensureReady: () => ensureBridgeLoaded().then(() => undefined),
  getFormulaCatalog: () => invokeJson<FormulaCatalogResponse>('getFormulaCatalog'),
  parseFormula: (formula: string) => invokeRaw<ParseFormulaResponse>('parseFormula', formula),
  evaluateFormula: (payload: unknown) => invokeJson<EvaluateFormulaResponse>('evaluateFormula', payload),
  validateProfile: (payload: unknown) => invokeJson<ValidateProfileResponse>('validateProfile', payload),
  generateInputMatrix: (payload: unknown) => invokeJson<GenerateInputMatrixResponse>('generateInputMatrix', payload),
  analyzeSimulation: (payload: unknown) => invokeJson<AnalyzeSimulationResponse>('analyzeSimulation', payload),
  runBenchmark: (payload: unknown) => invokeJson<BenchmarkResponse>('runBenchmark', payload),
};
