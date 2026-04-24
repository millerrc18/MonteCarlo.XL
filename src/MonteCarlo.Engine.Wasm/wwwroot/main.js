import { dotnet } from './_framework/dotnet.js';

let exportsPromise;

async function getBridgeExports() {
  if (!exportsPromise) {
    exportsPromise = (async () => {
      const runtime = await dotnet.create();
      const config = runtime.getConfig();
      const exports = await runtime.getAssemblyExports(config.mainAssemblyName);
      await runtime.runMain(config.mainAssemblyName, []);
      return exports.MonteCarlo.Engine.Wasm.BridgeExports;
    })();
  }

  return exportsPromise;
}

async function invokeJson(methodName, payload) {
  const bridge = await getBridgeExports();
  const method = bridge[methodName];

  if (typeof method !== 'function') {
    throw new Error(`MonteCarlo bridge method '${methodName}' is not available.`);
  }

  return method(payload);
}

async function invokeNoArgs(methodName) {
  const bridge = await getBridgeExports();
  const method = bridge[methodName];

  if (typeof method !== 'function') {
    throw new Error(`MonteCarlo bridge method '${methodName}' is not available.`);
  }

  return method();
}

globalThis.MonteCarloBridge = {
  ensureReady: () => getBridgeExports(),
  getFormulaCatalog: () => invokeNoArgs('GetFormulaCatalogJson'),
  parseFormula: (formula) => invokeJson('ParseFormulaJson', formula),
  evaluateFormula: (payload) => invokeJson('EvaluateFormulaJson', payload),
  validateProfile: (payload) => invokeJson('ValidateProfileJson', payload),
  generateInputMatrix: (payload) => invokeJson('GenerateInputMatrixJson', payload),
  analyzeSimulation: (payload) => invokeJson('AnalyzeSimulationJson', payload),
  evaluateGoalSeekMetric: (payload) => invokeJson('EvaluateGoalSeekMetricJson', payload),
  runBenchmark: (payload) => invokeJson('RunBenchmarkJson', payload),
};
