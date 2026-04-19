[CmdletBinding()]
param(
    [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $PSScriptRoot "MonteCarlo.XL.Smoke.xlsx"
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = "Smoke Model"

    $worksheet.Range("A1").Value2 = "MonteCarlo.XL Smoke Model"
    $worksheet.Range("A3").Value2 = "Input"
    $worksheet.Range("B3").Value2 = "Value"
    $worksheet.Range("C3").Value2 = "Distribution"

    $worksheet.Range("A4").Value2 = "Units"
    $worksheet.Range("B4").Formula = "=MC.Normal(100,10)"
    $worksheet.Range("C4").Value2 = "MC.Normal(100,10)"

    $worksheet.Range("A5").Value2 = "Price"
    $worksheet.Range("B5").Formula = "=MC.Triangular(900,1000,1200)"
    $worksheet.Range("C5").Value2 = "MC.Triangular(900,1000,1200)"

    $worksheet.Range("A6").Value2 = "Variable Cost"
    $worksheet.Range("B6").Formula = "=MC.PERT(450,600,800)"
    $worksheet.Range("C6").Value2 = "MC.PERT(450,600,800)"

    $worksheet.Range("A7").Value2 = "Fixed Costs"
    $worksheet.Range("B7").Value2 = 20000
    $worksheet.Range("C7").Value2 = "Fixed"

    $worksheet.Range("A9").Value2 = "Profit"
    $worksheet.Range("B9").Formula = "=B4*(B5-B6)-B7"
    $worksheet.Range("C9").Value2 = "Output cell to tag"

    $worksheet.Range("E3").Value2 = "Smoke steps"
    $worksheet.Range("E4").Value2 = "1. Load MonteCarlo.XL."
    $worksheet.Range("E5").Value2 = "2. Add B9 as an output."
    $worksheet.Range("E6").Value2 = "3. Run 1,000 iterations."
    $worksheet.Range("E7").Value2 = "4. Confirm results appear."

    $worksheet.Range("A1:C1").Merge() | Out-Null
    $worksheet.Range("A1").Font.Bold = $true
    $worksheet.Range("A3:C3").Font.Bold = $true
    $worksheet.Range("E3").Font.Bold = $true
    $worksheet.Columns.AutoFit() | Out-Null

    if (Test-Path $OutputPath) {
        Remove-Item $OutputPath -Force
    }

    $workbook.SaveAs($OutputPath, 51)
    Write-Host "Created $OutputPath"
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }

    if ($excel -ne $null) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
