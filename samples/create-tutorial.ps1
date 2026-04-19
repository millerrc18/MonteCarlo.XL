$savePath = "C:\Users\mille\OneDrive\04 - Projects\MonteCarlo.XL\samples\Product Launch Model.xlsx"
New-Item -ItemType Directory -Path (Split-Path $savePath) -Force | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Name = "Product Launch Model"

# TITLE
$ws.Range("A1").Value2 = "NEW PRODUCT LAUNCH - COST & REVENUE MODEL"
$ws.Range("A1").Font.Size = 16
$ws.Range("A1").Font.Bold = $true
$ws.Range("A1:F1").Merge()
$ws.Range("A2").Value2 = "MonteCarlo.XL Tutorial - Run a simulation to see the distribution of outcomes"
$ws.Range("A2").Font.Italic = $true
$ws.Range("A2").Font.Color = 0x808080
$ws.Range("A2:F2").Merge()

# ASSUMPTIONS HEADER
$ws.Range("A4").Value2 = "ASSUMPTIONS (Uncertain Inputs)"
$ws.Range("A4").Font.Bold = $true
$ws.Range("A4").Font.Size = 12
$ws.Range("A4:C4").Interior.Color = 0xFFF2CC
$ws.Range("A5").Value2 = "Variable"
$ws.Range("B5").Value2 = "Value (Expected)"
$ws.Range("C5").Value2 = "Distribution"
$ws.Range("A5:C5").Font.Bold = $true

# Inputs
$ws.Range("A6").Value2 = "Unit Mfg Cost ($)"
$ws.Range("B6").Formula = "=MC.Triangular(12, 15, 22)"
$ws.Range("C6").Value2 = "Triangular(12, 15, 22)"

$ws.Range("A7").Value2 = "Units Sold (Year 1)"
$ws.Range("B7").Formula = "=MC.Lognormal(9.2, 0.4)"
$ws.Range("C7").Value2 = "Lognormal(mu=9.2, sigma=0.4)"

$ws.Range("A8").Value2 = "Selling Price ($)"
$ws.Range("B8").Formula = "=MC.Normal(45, 5)"
$ws.Range("C8").Value2 = "Normal(45, 5)"

$ws.Range("A9").Value2 = "Marketing Spend ($)"
$ws.Range("B9").Formula = "=MC.PERT(50000, 80000, 150000)"
$ws.Range("C9").Value2 = "PERT(50k, 80k, 150k)"

$ws.Range("A10").Value2 = "Defect Rate"
$ws.Range("B10").Formula = "=MC.Beta(2, 50)"
$ws.Range("C10").Value2 = "Beta(2, 50) ~3.8% mean"

$ws.Range("A11").Value2 = "Dev Delay (months)"
$ws.Range("B11").Formula = "=MC.Poisson(2)"
$ws.Range("C11").Value2 = "Poisson(lambda=2)"

$ws.Range("A12").Value2 = "Competitor Enters? (0/1)"
$ws.Range("B12").Formula = "=MC.Binomial(1, 0.3)"
$ws.Range("C12").Value2 = "Binomial(1, 0.3) - 30% chance"

# Format inputs
$ws.Range("B6:B12").Font.Bold = $true
$ws.Range("B6:B12").Interior.Color = 0xE2EFDA
$ws.Range("B9").NumberFormat = "#,##0"
$ws.Range("B10").NumberFormat = "0.00%"

# FIXED ASSUMPTIONS
$ws.Range("A14").Value2 = "FIXED ASSUMPTIONS"
$ws.Range("A14").Font.Bold = $true
$ws.Range("A14").Font.Size = 12
$ws.Range("A14:C14").Interior.Color = 0xD9E2F3

$ws.Range("A15").Value2 = "Development Cost ($)"
$ws.Range("B15").Value2 = 200000
$ws.Range("B15").NumberFormat = "#,##0"

$ws.Range("A16").Value2 = "Delay Penalty ($/month)"
$ws.Range("B16").Value2 = 15000
$ws.Range("B16").NumberFormat = "#,##0"

$ws.Range("A17").Value2 = "Competitor Price Impact"
$ws.Range("B17").Value2 = 0.15
$ws.Range("B17").NumberFormat = "0%"

# MODEL CALCULATIONS
$ws.Range("A19").Value2 = "MODEL CALCULATIONS"
$ws.Range("A19").Font.Bold = $true
$ws.Range("A19").Font.Size = 12
$ws.Range("A19:C19").Interior.Color = 0xFCE4D6

$ws.Range("A20").Value2 = "Effective Price ($)"
$ws.Range("B20").Formula = "=B8 * (1 - B12 * B17)"
$ws.Range("B20").NumberFormat = "#,##0.00"

$ws.Range("A21").Value2 = "Gross Revenue ($)"
$ws.Range("B21").Formula = "=B20 * B7"
$ws.Range("B21").NumberFormat = "#,##0"

$ws.Range("A22").Value2 = "Cost of Goods Sold ($)"
$ws.Range("B22").Formula = "=B6 * B7"
$ws.Range("B22").NumberFormat = "#,##0"

$ws.Range("A23").Value2 = "Warranty Costs ($)"
$ws.Range("B23").Formula = "=B10 * B22 * 2"
$ws.Range("C23").Value2 = "Defects cost 2x manufacturing"
$ws.Range("C23").Font.Italic = $true
$ws.Range("C23").Font.Color = 0x808080
$ws.Range("B23").NumberFormat = "#,##0"

$ws.Range("A24").Value2 = "Delay Penalty ($)"
$ws.Range("B24").Formula = "=B11 * B16"
$ws.Range("B24").NumberFormat = "#,##0"

$ws.Range("A25").Value2 = "Total Costs ($)"
$ws.Range("B25").Formula = "=B22 + B9 + B15 + B23 + B24"
$ws.Range("B25").NumberFormat = "#,##0"
$ws.Range("A25").Font.Bold = $true

# OUTPUTS
$ws.Range("A27").Value2 = "KEY OUTPUTS (Simulation Targets)"
$ws.Range("A27").Font.Bold = $true
$ws.Range("A27").Font.Size = 12
$ws.Range("A27:C27").Interior.Color = 0xD5A6BD

$ws.Range("A28").Value2 = "Net Profit ($)"
$ws.Range("B28").Formula = "=B21 - B25"
$ws.Range("B28").NumberFormat = "#,##0"
$ws.Range("B28").Font.Bold = $true
$ws.Range("B28").Font.Size = 13
$ws.Range("B28").Interior.Color = 0xE2EFDA

$ws.Range("A29").Value2 = "ROI (%)"
$ws.Range("B29").Formula = "=B28 / B25"
$ws.Range("B29").NumberFormat = "0.0%"
$ws.Range("B29").Font.Bold = $true
$ws.Range("B29").Font.Size = 13
$ws.Range("B29").Interior.Color = 0xE2EFDA

$ws.Range("A30").Value2 = "Break-Even Units"
$ws.Range("B30").Formula = "=(B9 + B15 + B24) / (B20 - B6)"
$ws.Range("B30").NumberFormat = "#,##0"
$ws.Range("B30").Font.Bold = $true
$ws.Range("B30").Font.Size = 13
$ws.Range("B30").Interior.Color = 0xE2EFDA

# INSTRUCTIONS
$ws.Range("E4").Value2 = "HOW TO USE"
$ws.Range("E4").Font.Bold = $true
$ws.Range("E4").Font.Size = 12

$instructions = @(
    "1. Open the MonteCarlo.XL task pane (Ctrl+Shift+T)",
    "2. Click 'Add Input' and select each green cell (B6:B12)",
    "3. Click 'Add Output' and select each purple cell (B28:B30)",
    "4. Click Run Simulation (Ctrl+Shift+R)",
    "5. Watch the live histogram build during iteration",
    "6. Review results: histogram, CDF, tornado sensitivity",
    "7. Enter a target value (e.g. 0 for Net Profit) to see P(loss)",
    "8. Toggle PDF/CDF to see the probability density overlay",
    "9. Select inputs in the scatter dropdown to explore correlations",
    "10. Click Export to create a summary worksheet",
    "",
    "TIP: Try 10,000 iterations with LHS sampling for smooth",
    "results. This model has 7 uncertain inputs and 3 outputs.",
    "",
    "DISTRIBUTIONS USED:",
    "  Triangular - vendor cost quotes (min/mode/max)",
    "  Lognormal  - demand (right-skewed, always positive)",
    "  Normal     - price (symmetric around market research)",
    "  PERT       - budget (weighted toward most likely)",
    "  Beta       - defect rate (bounded 0-1, low mean)",
    "  Poisson    - schedule delays (discrete events)",
    "  Binomial   - competitor entry (yes/no event)"
)

for ($i = 0; $i -lt $instructions.Count; $i++) {
    $ws.Range("E$($i + 5)").Value2 = $instructions[$i]
    $ws.Range("E$($i + 5)").Font.Size = 10
}

# Column widths
$ws.Columns.Item("A").ColumnWidth = 28
$ws.Columns.Item("B").ColumnWidth = 18
$ws.Columns.Item("C").ColumnWidth = 32
$ws.Columns.Item("D").ColumnWidth = 3
$ws.Columns.Item("E").ColumnWidth = 58

# Save
$wb.SaveAs($savePath, 51)
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Output "Tutorial workbook saved to: $savePath"
