using CommunityToolkit.Mvvm.ComponentModel;

namespace MonteCarlo.UI.ViewModels;

/// <summary>
/// View model for a configured simulation output displayed as a card.
/// </summary>
public partial class OutputCardViewModel : ObservableObject
{
    /// <summary>Cell address in A1 notation (e.g., "D10").</summary>
    public string CellAddress { get; }

    /// <summary>Worksheet name.</summary>
    public string SheetName { get; }

    /// <summary>Human-readable label.</summary>
    public string Label { get; }

    /// <summary>Full cell reference (e.g., "Sheet1!D10").</summary>
    public string FullReference => $"{SheetName}!{CellAddress}";

    public OutputCardViewModel(string sheetName, string cellAddress, string label)
    {
        SheetName = sheetName;
        CellAddress = cellAddress;
        Label = label;
    }
}
