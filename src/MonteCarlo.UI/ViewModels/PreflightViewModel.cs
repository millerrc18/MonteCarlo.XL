using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MonteCarlo.Engine.Validation;

namespace MonteCarlo.UI.ViewModels;

/// <summary>
/// View model for pre-run model validation results.
/// </summary>
public partial class PreflightViewModel : ObservableObject
{
    [ObservableProperty]
    private ObservableCollection<PreflightIssue> _issues = new();

    [ObservableProperty]
    private string _summary = "Run a model check to validate the current setup.";

    public int ErrorCount => Issues.Count(i => i.Severity == PreflightSeverity.Error);

    public int WarningCount => Issues.Count(i => i.Severity == PreflightSeverity.Warning);

    public int InfoCount => Issues.Count(i => i.Severity == PreflightSeverity.Info);

    public bool CanRunAnyway => ErrorCount == 0;

    public event Action? RunAnywayRequested;

    public event Action? BackToSetupRequested;

    public void Load(PreflightReport report)
    {
        Issues = new ObservableCollection<PreflightIssue>(report.Issues);
        Summary = report.HasErrors
            ? $"{report.ErrorCount} blocking issue(s) found. Fix these before running the simulation."
            : report.WarningCount > 0
                ? $"{report.WarningCount} warning(s) found. Review them before running a final simulation."
                : "No blocking issues found. The model is ready to simulate.";

        OnPropertyChanged(nameof(ErrorCount));
        OnPropertyChanged(nameof(WarningCount));
        OnPropertyChanged(nameof(InfoCount));
        OnPropertyChanged(nameof(CanRunAnyway));
    }

    [RelayCommand]
    private void RunAnyway()
    {
        if (CanRunAnyway)
            RunAnywayRequested?.Invoke();
    }

    [RelayCommand]
    private void BackToSetup()
    {
        BackToSetupRequested?.Invoke();
    }
}
