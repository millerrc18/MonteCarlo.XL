using Microsoft.Office.Interop.Excel;
using MonteCarlo.Addin.Services;

namespace MonteCarlo.Addin.Excel;

/// <summary>
/// Captures Excel application state and restores it when disposed.
/// </summary>
public sealed class ExcelStateScope : IDisposable
{
    private readonly Application _app;
    private readonly string _phase;
    private readonly bool _restoreSelection;
    private readonly bool? _screenUpdating;
    private readonly bool? _enableEvents;
    private readonly bool? _displayAlerts;
    private readonly XlCalculation? _calculation;
    private readonly object? _statusBar;
    private readonly object? _activeSheet;
    private readonly object? _selection;
    private bool _disposed;

    private ExcelStateScope(Application app, string phase, bool restoreSelection)
    {
        _app = app;
        _phase = string.IsNullOrWhiteSpace(phase) ? "Excel state" : phase;
        _restoreSelection = restoreSelection;

        _screenUpdating = TryCapture(() => _app.ScreenUpdating, nameof(_app.ScreenUpdating));
        _enableEvents = TryCapture(() => _app.EnableEvents, nameof(_app.EnableEvents));
        _displayAlerts = TryCapture(() => _app.DisplayAlerts, nameof(_app.DisplayAlerts));
        _calculation = TryCapture(() => _app.Calculation, nameof(_app.Calculation));
        _statusBar = TryCapture(() => _app.StatusBar, nameof(_app.StatusBar));

        if (restoreSelection)
        {
            _activeSheet = TryCapture(() => _app.ActiveSheet, nameof(_app.ActiveSheet));
            _selection = TryCapture(() => _app.Selection, nameof(_app.Selection));
        }
    }

    /// <summary>
    /// Capture the current Excel state.
    /// </summary>
    public static ExcelStateScope Capture(Application app, string phase, bool restoreSelection = false)
    {
        ArgumentNullException.ThrowIfNull(app);
        return new ExcelStateScope(app, phase, restoreSelection);
    }

    /// <summary>
    /// Restore Excel to interactive defaults after an external failure leaves global state disabled.
    /// </summary>
    public static void RestoreInteractiveDefaults(Application app, string phase = "Excel state recovery")
    {
        ArgumentNullException.ThrowIfNull(app);

        RestoreDefault(phase, nameof(app.Calculation), () => app.Calculation = XlCalculation.xlCalculationAutomatic);
        RestoreDefault(phase, nameof(app.DisplayAlerts), () => app.DisplayAlerts = true);
        RestoreDefault(phase, nameof(app.StatusBar), () => app.StatusBar = false);
        RestoreDefault(phase, nameof(app.EnableEvents), () => app.EnableEvents = true);
        RestoreDefault(phase, nameof(app.ScreenUpdating), () => app.ScreenUpdating = true);
    }

    /// <summary>
    /// Creates a compact diagnostic snapshot of the current Excel interactive state.
    /// </summary>
    public static string DescribeCurrentState(Application app)
    {
        ArgumentNullException.ThrowIfNull(app);

        var parts = new[]
        {
            $"Calculation={SafeRead(() => app.Calculation.ToString(), "unknown")}",
            $"ScreenUpdating={SafeRead(() => app.ScreenUpdating.ToString(), "unknown")}",
            $"EnableEvents={SafeRead(() => app.EnableEvents.ToString(), "unknown")}",
            $"DisplayAlerts={SafeRead(() => app.DisplayAlerts.ToString(), "unknown")}",
            $"StatusBar={FormatStatusBar(SafeRead<object?>(() => app.StatusBar, null))}",
            $"Workbook={SafeRead(() => app.ActiveWorkbook?.Name ?? "none", "unknown")}",
            $"Sheet={SafeRead(() => ((Worksheet)app.ActiveSheet).Name, "unknown")}",
            $"Selection={SafeRead(() => ((Range)app.Selection).Address[false, false], "unknown")}"
        };

        return string.Join(", ", parts);
    }

    /// <summary>
    /// Apply temporary Excel settings for a protected operation.
    /// </summary>
    public void Apply(
        bool? screenUpdating = null,
        bool? enableEvents = null,
        bool? displayAlerts = null,
        XlCalculation? calculation = null,
        object? statusBar = null)
    {
        if (screenUpdating.HasValue)
            _app.ScreenUpdating = screenUpdating.Value;
        if (enableEvents.HasValue)
            _app.EnableEvents = enableEvents.Value;
        if (displayAlerts.HasValue)
            _app.DisplayAlerts = displayAlerts.Value;
        if (calculation.HasValue)
            _app.Calculation = calculation.Value;
        if (statusBar != null)
            _app.StatusBar = statusBar;
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;

        Restore(nameof(_app.Calculation), () =>
        {
            if (_calculation.HasValue)
                _app.Calculation = _calculation.Value;
        });

        Restore(nameof(_app.DisplayAlerts), () =>
        {
            if (_displayAlerts.HasValue)
                _app.DisplayAlerts = _displayAlerts.Value;
        });

        Restore(nameof(_app.StatusBar), () => _app.StatusBar = _statusBar ?? false);

        if (_restoreSelection)
        {
            Restore("active sheet", () =>
            {
                if (_activeSheet != null)
                    ((dynamic)_activeSheet).Activate();
            });

            Restore("selection", () =>
            {
                if (_selection != null)
                    ((dynamic)_selection).Select();
            });
        }

        Restore(nameof(_app.EnableEvents), () =>
        {
            if (_enableEvents.HasValue)
                _app.EnableEvents = _enableEvents.Value;
        });

        Restore(nameof(_app.ScreenUpdating), () =>
        {
            if (_screenUpdating.HasValue)
                _app.ScreenUpdating = _screenUpdating.Value;
        });
    }

    private T? TryCapture<T>(Func<T> capture, string propertyName)
    {
        try
        {
            return capture();
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException($"{_phase}: failed to capture Excel {propertyName}.", ex);
            return default;
        }
    }

    private void Restore(string propertyName, System.Action restore)
    {
        try
        {
            restore();
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException($"{_phase}: failed to restore Excel {propertyName}.", ex);
        }
    }

    private static void RestoreDefault(string phase, string propertyName, System.Action restore)
    {
        try
        {
            restore();
        }
        catch (Exception ex)
        {
            StartupDiagnostics.LogException($"{phase}: failed to restore Excel {propertyName}.", ex);
        }
    }

    private static T SafeRead<T>(Func<T> read, T fallback)
    {
        try
        {
            return read();
        }
        catch
        {
            return fallback;
        }
    }

    private static string FormatStatusBar(object? statusBar) =>
        statusBar switch
        {
            null => "unknown",
            bool value => value ? "true" : "false",
            _ => statusBar.ToString() ?? "unknown"
        };
}
