using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MonteCarlo.UI.Views;

namespace MonteCarlo.UI.ViewModels;

/// <summary>
/// Main view model for the task pane. Manages navigation between views.
/// </summary>
public partial class MainViewModel : ObservableObject
{
    [ObservableProperty]
    private string _title = "MonteCarlo.XL";

    [ObservableProperty]
    private object? _currentView;

    [ObservableProperty]
    private string _currentViewName = "Setup";

    public MainViewModel()
    {
        NavigateToSetup();
    }

    [RelayCommand]
    private void NavigateToSetup()
    {
        CurrentView = new SetupView();
        CurrentViewName = "Setup";
    }

    [RelayCommand]
    private void NavigateToResults()
    {
        CurrentView = new ResultsView();
        CurrentViewName = "Results";
    }

    [RelayCommand]
    private void NavigateToSettings()
    {
        CurrentView = new SettingsView();
        CurrentViewName = "Settings";
    }

    [RelayCommand]
    private void NavigateToRun()
    {
        CurrentView = new RunView();
        CurrentViewName = "Run";
    }
}
