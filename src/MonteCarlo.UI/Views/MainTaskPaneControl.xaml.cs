using System.ComponentModel;
using System.Windows.Controls;
using System.Windows.Threading;
using MonteCarlo.UI.Services;
using MonteCarlo.UI.ViewModels;

namespace MonteCarlo.UI.Views;

/// <summary>
/// Root task pane control that hosts view navigation and the content area.
/// </summary>
public partial class MainTaskPaneControl : UserControl
{
    private readonly ThemeManager _themeManager = new();
    private bool _isViewModelWired;

    public MainTaskPaneControl()
    {
        InitializeComponent();
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, System.Windows.RoutedEventArgs e)
    {
        ApplySavedTheme();

        if (!_isViewModelWired && DataContext is MainViewModel viewModel)
        {
            viewModel.PropertyChanged += OnViewModelPropertyChanged;
            _isViewModelWired = true;
        }
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName != nameof(MainViewModel.CurrentView))
            return;

        Dispatcher.BeginInvoke(ApplySavedTheme, DispatcherPriority.Loaded);
    }

    private void ApplySavedTheme()
    {
        _themeManager.ApplyTheme(_themeManager.LoadPreference(), this);
    }
}
