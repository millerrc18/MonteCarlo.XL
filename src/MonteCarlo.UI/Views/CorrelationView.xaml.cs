using System.Windows.Controls;
using MonteCarlo.UI.ViewModels;

namespace MonteCarlo.UI.Views;

/// <summary>
/// Correlation matrix editor view where users define pairwise correlations between inputs.
/// </summary>
public partial class CorrelationView : UserControl
{
    public CorrelationView()
    {
        InitializeComponent();
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, System.Windows.RoutedEventArgs e)
    {
        if (DataContext is CorrelationViewModel vm)
        {
            MatrixGridControl.CellValueChanged += vm.OnCellValueChanged;

            // Show scroll hint if many inputs
            if (vm.InputLabels.Count > 6)
                ScrollHint.Visibility = System.Windows.Visibility.Visible;
        }
    }
}
