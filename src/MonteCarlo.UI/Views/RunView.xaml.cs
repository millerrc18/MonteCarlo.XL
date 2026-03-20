using System.Windows.Controls;
using MonteCarlo.UI.ViewModels;

namespace MonteCarlo.UI.Views;

/// <summary>
/// Run view — displays simulation progress, live stats, and convergence.
/// </summary>
public partial class RunView : UserControl
{
    public RunView()
    {
        InitializeComponent();
    }

    /// <summary>
    /// Gets the RunViewModel for external progress updates.
    /// </summary>
    public RunViewModel ViewModel => (RunViewModel)DataContext;
}
