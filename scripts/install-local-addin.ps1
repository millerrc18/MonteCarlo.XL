[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",

    [switch]$Build,

    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$installDir = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
$installedAddin = Join-Path $installDir "MonteCarlo.XL.xll"

if ($Uninstall) {
    if (Test-Path $installedAddin) {
        Remove-Item $installedAddin -Force
        Write-Host "Removed $installedAddin"
    }
    else {
        Write-Host "No local MonteCarlo.XL add-in found at $installedAddin"
    }
    return
}

if ($Build) {
    Push-Location $repoRoot
    try {
        dotnet build MonteCarlo.XL.sln
    }
    finally {
        Pop-Location
    }
}

$officeConfigPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
)

$platform = "x64"
foreach ($path in $officeConfigPaths) {
    if (Test-Path $path) {
        $config = Get-ItemProperty $path
        if ($config.Platform) {
            $platform = $config.Platform
            break
        }
    }
}

if ($platform -eq "x86") {
    $artifactName = "MonteCarlo.Addin-AddIn-packed.xll"
}
else {
    $artifactName = "MonteCarlo.Addin-AddIn64-packed.xll"
}

$artifact = Join-Path $repoRoot "src\MonteCarlo.Addin\bin\$Configuration\net8.0-windows\publish\$artifactName"
if (-not (Test-Path $artifact)) {
    throw "Packed add-in not found: $artifact. Run with -Build or build the solution first."
}

New-Item -ItemType Directory -Path $installDir -Force | Out-Null
try {
    Copy-Item $artifact $installedAddin -Force
}
catch {
    throw "Could not update $installedAddin. Close every Excel window and run this script again. Original error: $($_.Exception.Message)"
}
Unblock-File $installedAddin -ErrorAction SilentlyContinue

Write-Host "Installed $artifactName for $platform Excel:"
Write-Host "  $installedAddin"
Write-Host "Restart Excel to load MonteCarlo.XL."
