import React, { useEffect, useMemo, useState } from 'react';
import Plot from 'react-plotly.js';
import {
  Button,
  Dropdown,
  Field,
  FluentProvider,
  Input,
  Option,
  Spinner,
  Switch,
  webDarkTheme,
  webLightTheme,
} from '@fluentui/react-components';
import { MonteCarloBridgeClient } from './bridge';
import {
  analyzeSimulation,
  buildRunSettings,
  createEmptyProfile,
  DEFAULT_USER_SETTINGS,
  distributionUseCases,
  executeSimulation,
  exportSummarySheet,
  loadUserSettings,
  loadWorkbookEnvelope,
  resolveEffectiveSettings,
  runBenchmark,
  saveUserSettings,
  saveWorkbookEnvelope,
  sortFormulaCatalog,
  addSelectedAsInput,
  addSelectedAsOutput,
  syncAutoDetectedInputs,
  validateProfile,
} from './excel';
import type {
  AnalyzeSimulationResponse,
  BenchmarkResponse,
  FormulaDefinitionDto,
  OutputAnalysisDto,
  SimulationProfile,
  UserSettings,
  WorkbookUserSettingsOverrides,
} from './types';

function App() {
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState('Loading Office host...');
  const [formulaCatalog, setFormulaCatalog] = useState<FormulaDefinitionDto[]>([]);
  const [profile, setProfile] = useState<SimulationProfile>(createEmptyProfile(DEFAULT_USER_SETTINGS.defaultIterationCount));
  const [globalSettings, setGlobalSettings] = useState<UserSettings>(loadUserSettings());
  const [workbookOverrides, setWorkbookOverrides] = useState<WorkbookUserSettingsOverrides | null>(null);
  const [results, setResults] = useState<AnalyzeSimulationResponse | null>(null);
  const [benchmark, setBenchmark] = useState<BenchmarkResponse | null>(null);
  const [selectedOutputId, setSelectedOutputId] = useState<string>('');
  const [outputTargetText, setOutputTargetText] = useState('');
  const [progress, setProgress] = useState<{ completed: number; total: number } | null>(null);

  const effectiveSettings = useMemo(
    () => resolveEffectiveSettings(globalSettings, workbookOverrides),
    [globalSettings, workbookOverrides],
  );

  const theme = effectiveSettings.theme === 'dark'
    ? webDarkTheme
    : effectiveSettings.theme === 'light'
      ? webLightTheme
      : window.matchMedia('(prefers-color-scheme: dark)').matches
        ? webDarkTheme
        : webLightTheme;

  const selectedOutput = results?.outputs.find((output) => output.outputId === selectedOutputId) ?? results?.outputs[0] ?? null;

  useEffect(() => {
    saveUserSettings(globalSettings);
  }, [globalSettings]);

  useEffect(() => {
    void initialize();
  }, []);

  async function initialize() {
    setLoading(true);
    try {
      await MonteCarloBridgeClient.ensureReady();
      const catalog = await MonteCarloBridgeClient.getFormulaCatalog();
      const workbookEnvelope = await loadWorkbookEnvelope();
      const nextProfile = workbookEnvelope.profile ?? createEmptyProfile(globalSettings.defaultIterationCount);

      setFormulaCatalog(sortFormulaCatalog(catalog.functions));
      setProfile(nextProfile);
      setWorkbookOverrides(workbookEnvelope.workbookSettingsOverrides);
      setSelectedOutputId(nextProfile.outputs[0] ? `'${nextProfile.outputs[0].sheetName}'!${nextProfile.outputs[0].cellAddress}` : '');
      setStatus('Office host ready.');
    } catch (error) {
      setStatus(`Initialization failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setLoading(false);
    }
  }

  async function saveWorkbookState() {
    setBusy(true);
    try {
      await saveWorkbookEnvelope({
        profile,
        workbookSettingsOverrides: workbookOverrides,
      });
      setStatus('Workbook configuration saved.');
    } catch (error) {
      setStatus(`Save failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setBusy(false);
    }
  }

  async function detectInputs() {
    setBusy(true);
    try {
      const nextProfile = await syncAutoDetectedInputs(profile);
      setProfile(nextProfile);
      setStatus(`Detected ${nextProfile.inputs.length} MC.* input(s) on the active sheet.`);
    } catch (error) {
      setStatus(`Scan failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setBusy(false);
    }
  }

  async function addInput() {
    setBusy(true);
    try {
      const inputs = await addSelectedAsInput(profile.inputs);
      setProfile({ ...profile, inputs });
      setStatus('Selected cell added as an input.');
    } catch (error) {
      setStatus(`Add input failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setBusy(false);
    }
  }

  async function addOutput() {
    setBusy(true);
    try {
      const outputs = await addSelectedAsOutput(profile.outputs);
      const nextProfile = { ...profile, outputs };
      setProfile(nextProfile);
      if (!selectedOutputId && outputs[0]) {
        setSelectedOutputId(`'${outputs[0].sheetName}'!${outputs[0].cellAddress}`);
      }
      setStatus('Selected cell added as an output.');
    } catch (error) {
      setStatus(`Add output failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setBusy(false);
    }
  }

  async function runModel() {
    setBusy(true);
    setBenchmark(null);
    setResults(null);
    setProgress(null);

    try {
      const runSettings = buildRunSettings(effectiveSettings, outputTargetText);
      const profileToRun = {
        ...profile,
        iterationCount: runSettings.iterationCount,
        randomSeed: runSettings.randomSeed,
      };

      const validation = await validateProfile(profileToRun);
      const blockingIssues = validation.issues.filter((issue) => issue.severity === 'Error');
      if (blockingIssues.length > 0) {
        setStatus(`Model check found ${blockingIssues.length} blocking issue(s). Fix them before running.`);
        return;
      }

      const execution = await executeSimulation(profileToRun, runSettings, (completed, total) => {
        setProgress({ completed, total });
      });
      const analysis = await analyzeSimulation(profileToRun, runSettings, execution);
      setResults(analysis);
      setSelectedOutputId(analysis.outputs[0]?.outputId ?? '');
      setStatus(`Simulation finished in ${analysis.elapsedMilliseconds.toFixed(0)} ms across ${analysis.iterationCount.toLocaleString()} iterations.`);
    } catch (error) {
      setStatus(`Simulation failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setBusy(false);
      setProgress(null);
    }
  }

  async function exportSummary() {
    if (!results) {
      setStatus('Run a simulation before exporting results.');
      return;
    }

    setBusy(true);
    try {
      await exportSummarySheet(profile, results, effectiveSettings.createNewWorksheetForExports);
      setStatus('Summary sheet exported to the workbook.');
    } catch (error) {
      setStatus(`Export failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setBusy(false);
    }
  }

  async function benchmarkHost() {
    setBusy(true);
    try {
      const result = await runBenchmark(Math.max(5, profile.inputs.length || 5), effectiveSettings.defaultIterationCount);
      setBenchmark(result);
      setStatus(`Benchmark finished at ${result.iterationsPerSecond.toLocaleString(undefined, { maximumFractionDigits: 0 })} iterations/sec.`);
    } catch (error) {
      setStatus(`Benchmark failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setBusy(false);
    }
  }

  if (loading) {
    return (
      <div className="app-loading">
        <Spinner label={status} />
      </div>
    );
  }

  return (
    <FluentProvider theme={theme}>
      <div className="app-shell">
        <header className="hero">
          <div>
            <div className="eyebrow">MonteCarlo.XL Office Host</div>
            <h1>Native ARM path for Excel</h1>
            <p>{status}</p>
          </div>
          <div className="hero-actions">
            <Switch
              label="Dark mode"
              checked={effectiveSettings.theme === 'dark'}
              onChange={(_, data) => setGlobalSettings({
                ...globalSettings,
                theme: data.checked ? 'dark' : 'light',
              })}
            />
            <Button appearance="primary" onClick={() => void saveWorkbookState()} disabled={busy}>
              Save Workbook Config
            </Button>
          </div>
        </header>

        <section className="toolbar">
          <Button appearance="secondary" onClick={() => void initialize()} disabled={busy}>
            Refresh
          </Button>
          <Button appearance="secondary" onClick={() => void detectInputs()} disabled={busy}>
            Scan MC Formulas
          </Button>
          <Button appearance="secondary" onClick={() => void addInput()} disabled={busy}>
            Add Selected Input
          </Button>
          <Button appearance="secondary" onClick={() => void addOutput()} disabled={busy}>
            Add Selected Output
          </Button>
          <Button appearance="primary" onClick={() => void runModel()} disabled={busy}>
            Run Simulation
          </Button>
          <Button appearance="secondary" onClick={() => void exportSummary()} disabled={busy || !results}>
            Export Summary
          </Button>
          <Button appearance="secondary" onClick={() => void benchmarkHost()} disabled={busy}>
            Benchmark
          </Button>
        </section>

        {busy && progress && (
          <section className="panel">
            <div className="progress-header">
              <strong>Simulation progress</strong>
              <span>{progress.completed.toLocaleString()} / {progress.total.toLocaleString()}</span>
            </div>
            <progress max={progress.total} value={progress.completed} />
          </section>
        )}

        <section className="grid">
          <div className="panel">
            <h2>Run Settings</h2>
            <div className="form-grid">
              <Field label="Iterations">
                <Input
                  type="number"
                  value={String(globalSettings.defaultIterationCount)}
                  onChange={(_, data) => setGlobalSettings({
                    ...globalSettings,
                    defaultIterationCount: Math.max(1, Number(data.value) || DEFAULT_USER_SETTINGS.defaultIterationCount),
                  })}
                />
              </Field>

              <Field label="Seed Mode">
                <Dropdown
                  value={globalSettings.seedMode}
                  selectedOptions={[globalSettings.seedMode]}
                  onOptionSelect={(_, data) => setGlobalSettings({
                    ...globalSettings,
                    seedMode: data.optionValue === 'Fixed' ? 'Fixed' : 'Random',
                  })}
                >
                  <Option value="Random">Random</Option>
                  <Option value="Fixed">Fixed</Option>
                </Dropdown>
              </Field>

              <Field label="Fixed Seed">
                <Input
                  type="number"
                  value={String(globalSettings.fixedRandomSeed)}
                  onChange={(_, data) => setGlobalSettings({
                    ...globalSettings,
                    fixedRandomSeed: Math.max(0, Number(data.value) || 0),
                  })}
                />
              </Field>

              <Field label="Sampling">
                <Dropdown
                  value={globalSettings.samplingMethod}
                  selectedOptions={[globalSettings.samplingMethod]}
                  onOptionSelect={(_, data) => setGlobalSettings({
                    ...globalSettings,
                    samplingMethod: data.optionValue === 'MonteCarlo' ? 'MonteCarlo' : 'LatinHypercube',
                  })}
                >
                  <Option value="LatinHypercube">Latin Hypercube</Option>
                  <Option value="MonteCarlo">Monte Carlo</Option>
                </Dropdown>
              </Field>

              <Field label="Target (optional)">
                <Input value={outputTargetText} onChange={(_, data) => setOutputTargetText(data.value)} />
              </Field>

              <Switch
                label="New worksheet for exports"
                checked={globalSettings.createNewWorksheetForExports}
                onChange={(_, data) => setGlobalSettings({
                  ...globalSettings,
                  createNewWorksheetForExports: data.checked,
                })}
              />
              <Switch
                label="Auto-stop on convergence"
                checked={globalSettings.autoStopOnConvergence}
                onChange={(_, data) => setGlobalSettings({
                  ...globalSettings,
                  autoStopOnConvergence: data.checked,
                })}
              />
            </div>
          </div>

          <div className="panel">
            <h2>Model Setup</h2>
            <div className="counts">
              <div>
                <strong>{profile.inputs.length}</strong>
                <span>Inputs</span>
              </div>
              <div>
                <strong>{profile.outputs.length}</strong>
                <span>Outputs</span>
              </div>
            </div>
            <div className="subpanel">
              <h3>Inputs</h3>
              <table>
                <thead>
                  <tr>
                    <th>Cell</th>
                    <th>Distribution</th>
                    <th>Label</th>
                  </tr>
                </thead>
                <tbody>
                  {profile.inputs.map((input) => (
                    <tr key={`${input.sheetName}:${input.cellAddress}`}>
                      <td>{input.sheetName}!{input.cellAddress}</td>
                      <td>{input.distributionName}</td>
                      <td>{input.label}</td>
                    </tr>
                  ))}
                  {profile.inputs.length === 0 && (
                    <tr>
                      <td colSpan={3}>Use Scan MC Formulas or Add Selected Input to start.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            <div className="subpanel">
              <h3>Outputs</h3>
              <table>
                <thead>
                  <tr>
                    <th>Cell</th>
                    <th>Label</th>
                  </tr>
                </thead>
                <tbody>
                  {profile.outputs.map((output) => (
                    <tr key={`${output.sheetName}:${output.cellAddress}`}>
                      <td>{output.sheetName}!{output.cellAddress}</td>
                      <td>{output.label}</td>
                    </tr>
                  ))}
                  {profile.outputs.length === 0 && (
                    <tr>
                      <td colSpan={2}>Add at least one output before running.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </section>

        <section className="grid">
          <div className="panel">
            <h2>Results</h2>
            {selectedOutput ? (
              <>
                <Field label="Output">
                  <Dropdown
                    value={selectedOutput.outputLabel}
                    selectedOptions={[selectedOutput.outputId]}
                    onOptionSelect={(_, data) => setSelectedOutputId(data.optionValue ?? '')}
                  >
                    {results?.outputs.map((output) => (
                      <Option key={output.outputId} value={output.outputId}>
                        {output.outputLabel}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>

                <div className="stats-grid">
                  <StatCard label="Mean" value={selectedOutput.summary.mean} />
                  <StatCard label="P50" value={selectedOutput.summary.p50} />
                  <StatCard label="P90" value={selectedOutput.summary.p90} />
                  <StatCard label="Std Dev" value={selectedOutput.summary.stdDev} />
                </div>

                <div className="chart-grid">
                  <ChartCard title="Histogram">
                    <Plot
                      data={[
                        {
                          x: selectedOutput.histogram.binCenters,
                          y: selectedOutput.histogram.relativeFrequencies,
                          type: 'bar',
                          marker: { color: '#0f6cbd' },
                          name: 'Histogram',
                        },
                        {
                          x: selectedOutput.histogram.kdeX,
                          y: selectedOutput.histogram.kdeY,
                          type: 'scatter',
                          mode: 'lines',
                          line: { color: '#ca5010', width: 2 },
                          name: 'KDE',
                        },
                      ]}
                      layout={plotLayout(selectedOutput.outputLabel)}
                      style={{ width: '100%', height: '100%' }}
                      config={{ displayModeBar: false, responsive: true }}
                    />
                  </ChartCard>
                  <ChartCard title="CDF">
                    <Plot
                      data={[
                        {
                          x: selectedOutput.cdf.x,
                          y: selectedOutput.cdf.y,
                          type: 'scatter',
                          mode: 'lines',
                          line: { color: '#107c10', width: 2 },
                          name: 'CDF',
                        },
                      ]}
                      layout={plotLayout('CDF')}
                      style={{ width: '100%', height: '100%' }}
                      config={{ displayModeBar: false, responsive: true }}
                    />
                  </ChartCard>
                  <ChartCard title="Sensitivity">
                    <Plot
                      data={[
                        {
                          x: selectedOutput.sensitivity.map((item) => item.rankCorrelation),
                          y: selectedOutput.sensitivity.map((item) => item.inputLabel),
                          type: 'bar',
                          orientation: 'h',
                          marker: { color: '#5c2e91' },
                        },
                      ]}
                      layout={{
                        margin: { l: 120, r: 20, t: 10, b: 30 },
                        paper_bgcolor: 'transparent',
                        plot_bgcolor: 'transparent',
                      }}
                      style={{ width: '100%', height: '100%' }}
                      config={{ displayModeBar: false, responsive: true }}
                    />
                  </ChartCard>
                </div>
              </>
            ) : (
              <p>Run a simulation to see charts and summary statistics here.</p>
            )}
          </div>

          <div className="panel">
            <h2>Distribution Guide</h2>
            <div className="guide-list">
              {formulaCatalog.map((formula) => (
                <details key={formula.functionName}>
                  <summary>{formula.excelName}</summary>
                  <p>{formula.description}</p>
                  <p>{distributionUseCases[formula.distributionName] ?? 'Use when this distribution best matches the business mechanism behind the uncertainty.'}</p>
                  <ul>
                    {formula.arguments.map((argument) => (
                      <li key={argument.name}>
                        <strong>{argument.displayName}:</strong> {argument.description}
                      </li>
                    ))}
                  </ul>
                </details>
              ))}
            </div>
          </div>
        </section>

        {benchmark && (
          <section className="panel">
            <h2>Benchmark</h2>
            <div className="stats-grid">
              <StatCard label="Iterations/sec" value={benchmark.iterationsPerSecond} />
              <StatCard label="Microseconds/iter" value={benchmark.microsecondsPerIteration} />
              <StatCard label="Inputs" value={benchmark.inputCount} />
              <StatCard label="Iterations" value={benchmark.iterationCount} />
            </div>
          </section>
        )}
      </div>
    </FluentProvider>
  );
}

function plotLayout(title: string) {
  return {
    title: { text: title, font: { size: 14 } },
    margin: { l: 50, r: 20, t: 36, b: 40 },
    paper_bgcolor: 'transparent',
    plot_bgcolor: 'transparent',
  };
}

function StatCard({ label, value }: { label: string; value: number }) {
  return (
    <div className="stat-card">
      <span>{label}</span>
      <strong>{Number.isFinite(value) ? value.toLocaleString(undefined, { maximumFractionDigits: 4 }) : 'n/a'}</strong>
    </div>
  );
}

function ChartCard({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div className="chart-card">
      <h3>{title}</h3>
      <div className="chart-body">{children}</div>
    </div>
  );
}

export default App;
