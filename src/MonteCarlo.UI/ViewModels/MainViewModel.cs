using CommunityToolkit.Mvvm.ComponentModel;

namespace MonteCarlo.UI.ViewModels;

/// <summary>
/// Placeholder view model for the main task pane. Will be replaced with navigation logic in Phase 1.
/// </summary>
public partial class MainViewModel : ObservableObject
{
    [ObservableProperty]
    private string _title = "MonteCarlo.XL";
}
