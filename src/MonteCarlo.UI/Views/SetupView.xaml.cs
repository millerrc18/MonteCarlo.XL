using System.Windows.Controls;
using MonteCarlo.UI.ViewModels;

namespace MonteCarlo.UI.Views;

/// <summary>
/// Setup view — configure simulation inputs and outputs.
/// </summary>
public partial class SetupView : UserControl
{
    public SetupView()
    {
        InitializeComponent();
    }

    /// <summary>
    /// Gets the SetupViewModel for external wiring (e.g., cell selection, run events).
    /// </summary>
    public SetupViewModel ViewModel => (SetupViewModel)DataContext;
}
