using System.Windows.Controls;
using MonteCarlo.UI.ViewModels;

namespace MonteCarlo.UI.Views;

/// <summary>
/// Results dashboard — histogram, CDF, stats panel, and target analysis.
/// </summary>
public partial class ResultsView : UserControl
{
    public ResultsView()
    {
        InitializeComponent();
    }

    /// <summary>
    /// Gets the ResultsViewModel for external data loading.
    /// </summary>
    public ResultsViewModel ViewModel => (ResultsViewModel)DataContext;
}
