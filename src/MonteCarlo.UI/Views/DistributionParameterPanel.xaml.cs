using System.Windows.Controls;
using MonteCarlo.UI.ViewModels;

namespace MonteCarlo.UI.Views;

/// <summary>
/// Dynamic parameter fields that change based on the selected distribution type.
/// </summary>
public partial class DistributionParameterPanel : UserControl
{
    public DistributionParameterPanel()
    {
        InitializeComponent();
        AddHandler(TextBox.TextChangedEvent, new TextChangedEventHandler(OnParameterTextChanged));
    }

    private void OnParameterTextChanged(object sender, TextChangedEventArgs e)
    {
        if (DataContext is SetupViewModel viewModel
            && viewModel.UpdateEditorPreviewCommand.CanExecute(null))
        {
            viewModel.UpdateEditorPreviewCommand.Execute(null);
        }
    }
}
