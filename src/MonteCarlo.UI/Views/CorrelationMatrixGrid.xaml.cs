using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace MonteCarlo.UI.Views;

/// <summary>
/// Dynamically generates an editable correlation matrix grid.
/// Diagonal cells are read-only (1.0), lower triangle mirrors upper triangle.
/// Cells are color-coded by correlation strength and sign.
/// </summary>
public partial class CorrelationMatrixGrid : UserControl
{
    private TextBox?[,]? _cellBoxes;

    public static readonly DependencyProperty InputLabelsProperty =
        DependencyProperty.Register(nameof(InputLabels), typeof(ObservableCollection<string>),
            typeof(CorrelationMatrixGrid), new PropertyMetadata(null, OnLabelsChanged));

    public static readonly DependencyProperty MatrixValuesProperty =
        DependencyProperty.Register(nameof(MatrixValues), typeof(double[,]),
            typeof(CorrelationMatrixGrid), new PropertyMetadata(null, OnMatrixChanged));

    public ObservableCollection<string>? InputLabels
    {
        get => (ObservableCollection<string>?)GetValue(InputLabelsProperty);
        set => SetValue(InputLabelsProperty, value);
    }

    public double[,]? MatrixValues
    {
        get => (double[,]?)GetValue(MatrixValuesProperty);
        set => SetValue(MatrixValuesProperty, value);
    }

    /// <summary>
    /// Raised when a cell value is edited by the user.
    /// Parameters: row, column, new value.
    /// </summary>
    public event Action<int, int, double>? CellValueChanged;

    public CorrelationMatrixGrid()
    {
        InitializeComponent();
    }

    private static void OnLabelsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        ((CorrelationMatrixGrid)d).RebuildGrid();
    }

    private static void OnMatrixChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        ((CorrelationMatrixGrid)d).UpdateCellValues();
    }

    private void RebuildGrid()
    {
        MatrixGrid.Children.Clear();
        MatrixGrid.RowDefinitions.Clear();
        MatrixGrid.ColumnDefinitions.Clear();
        _cellBoxes = null;

        var labels = InputLabels;
        var matrix = MatrixValues;
        if (labels == null || labels.Count < 2 || matrix == null)
            return;

        int n = labels.Count;
        _cellBoxes = new TextBox?[n, n];

        // Create grid: n+1 rows and n+1 columns (first row/col = headers)
        for (int i = 0; i <= n; i++)
        {
            MatrixGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            MatrixGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        }

        // Column headers (row 0, starting at col 1)
        for (int j = 0; j < n; j++)
        {
            var header = new TextBlock
            {
                Text = TruncateLabel(labels[j], 8),
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(4, 4, 4, 4),
                ToolTip = labels[j]
            };
            Grid.SetRow(header, 0);
            Grid.SetColumn(header, j + 1);
            MatrixGrid.Children.Add(header);
        }

        // Row headers (col 0, starting at row 1)
        for (int i = 0; i < n; i++)
        {
            var header = new TextBlock
            {
                Text = TruncateLabel(labels[i], 8),
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 4, 8, 4),
                ToolTip = labels[i]
            };
            Grid.SetRow(header, i + 1);
            Grid.SetColumn(header, 0);
            MatrixGrid.Children.Add(header);
        }

        // Matrix cells
        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j < n; j++)
            {
                if (i == j)
                {
                    // Diagonal: read-only, grayed out
                    var diag = new TextBlock
                    {
                        Text = "1.0",
                        FontSize = 12,
                        FontFamily = new FontFamily("Consolas"),
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center,
                        Foreground = new SolidColorBrush(Color.FromRgb(148, 163, 184)), // slate-400
                        Margin = new Thickness(2)
                    };
                    var diagBorder = new Border
                    {
                        Background = new SolidColorBrush(Color.FromRgb(241, 245, 249)), // slate-100
                        BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)), // slate-200
                        BorderThickness = new Thickness(1),
                        MinWidth = 52,
                        MinHeight = 32,
                        Child = diag
                    };
                    Grid.SetRow(diagBorder, i + 1);
                    Grid.SetColumn(diagBorder, j + 1);
                    MatrixGrid.Children.Add(diagBorder);
                }
                else if (j > i)
                {
                    // Upper triangle: editable
                    var tb = new TextBox
                    {
                        Text = matrix[i, j].ToString("F2"),
                        FontSize = 12,
                        FontFamily = new FontFamily("Consolas"),
                        HorizontalContentAlignment = HorizontalAlignment.Center,
                        VerticalContentAlignment = VerticalAlignment.Center,
                        MinWidth = 52,
                        MinHeight = 32,
                        Padding = new Thickness(4, 4, 4, 4),
                        Margin = new Thickness(1),
                        BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
                        BorderThickness = new Thickness(1)
                    };
                    int row = i, col = j; // Capture for closure
                    tb.LostFocus += (_, _) => OnCellLostFocus(row, col, tb);
                    tb.Background = GetCellBackground(matrix[i, j]);
                    _cellBoxes[i, j] = tb;

                    Grid.SetRow(tb, i + 1);
                    Grid.SetColumn(tb, j + 1);
                    MatrixGrid.Children.Add(tb);
                }
                else
                {
                    // Lower triangle: read-only mirror
                    var mirror = new TextBlock
                    {
                        Text = matrix[i, j].ToString("F2"),
                        FontSize = 12,
                        FontFamily = new FontFamily("Consolas"),
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center,
                        Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139)), // slate-500
                        Margin = new Thickness(2)
                    };
                    var mirrorBorder = new Border
                    {
                        Background = GetCellBackground(matrix[i, j]),
                        BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
                        BorderThickness = new Thickness(1),
                        MinWidth = 52,
                        MinHeight = 32,
                        Child = mirror
                    };
                    Grid.SetRow(mirrorBorder, i + 1);
                    Grid.SetColumn(mirrorBorder, j + 1);
                    MatrixGrid.Children.Add(mirrorBorder);
                }
            }
        }
    }

    private void OnCellLostFocus(int row, int col, TextBox tb)
    {
        if (!double.TryParse(tb.Text, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out double value))
        {
            tb.BorderBrush = new SolidColorBrush(Color.FromRgb(239, 68, 68)); // red-500
            return;
        }

        value = Math.Clamp(value, -1.0, 1.0);
        tb.Text = value.ToString("F2");
        tb.Background = GetCellBackground(value);
        tb.BorderBrush = new SolidColorBrush(Color.FromRgb(226, 232, 240));

        CellValueChanged?.Invoke(row, col, value);
    }

    private void UpdateCellValues()
    {
        var matrix = MatrixValues;
        if (_cellBoxes == null || matrix == null)
        {
            RebuildGrid();
            return;
        }

        int n = matrix.GetLength(0);
        if (_cellBoxes.GetLength(0) != n)
        {
            RebuildGrid();
            return;
        }

        // Update cell text and colors without rebuilding
        for (int i = 0; i < n; i++)
        {
            for (int j = i + 1; j < n; j++)
            {
                if (_cellBoxes[i, j] is TextBox tb)
                {
                    tb.Text = matrix[i, j].ToString("F2");
                    tb.Background = GetCellBackground(matrix[i, j]);
                }
            }
        }

        // Rebuild to update mirror cells
        RebuildGrid();
    }

    private static SolidColorBrush GetCellBackground(double corr)
    {
        if (Math.Abs(corr) < 0.01)
            return Brushes.Transparent;

        byte alpha = (byte)(Math.Min(Math.Abs(corr), 1.0) * 120);
        if (corr > 0)
            return new SolidColorBrush(Color.FromArgb(alpha, 59, 130, 246));
        return new SolidColorBrush(Color.FromArgb(alpha, 249, 115, 22));
    }

    private static string TruncateLabel(string label, int maxLen)
    {
        return label.Length <= maxLen ? label : label[..(maxLen - 1)] + "…";
    }
}
