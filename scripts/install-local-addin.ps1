[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",

    [switch]$Build,

    [switch]$Uninstall,

    [switch]$AllowUnsupportedArm
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$installDir = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
$installedAddin = Join-Path $installDir "MonteCarlo.XL.xll"

function Get-SystemArchitecture {
    try {
        $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
        switch ($cpu.Architecture) {
            12 { return "ARM64" }
            9 { return "x64" }
            0 { return "x86" }
        }
    }
    catch {
    }

    if ($env:PROCESSOR_ARCHITECTURE -match "ARM64" -or $env:PROCESSOR_IDENTIFIER -match "ARM") {
        return "ARM64"
    }

    if ($env:PROCESSOR_ARCHITECTURE) {
        return $env:PROCESSOR_ARCHITECTURE
    }

    return "Unknown"
}

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

$systemArchitecture = Get-SystemArchitecture
if ($systemArchitecture -eq "ARM64" -and -not $AllowUnsupportedArm) {
    throw @"
MonteCarlo.XL does not currently support native ARM64 Excel installs.

This repository packages Excel-DNA XLL loaders for x86 and x64 Excel only. Microsoft 365 Excel on Windows ARM is typically ARM-native, so the add-in will not load there as-is.

Use an x64 Windows Excel machine today, or rerun this script with -AllowUnsupportedArm only if you have already verified an x64-compatible Excel environment and want to experiment.

See README.md and docs/ARM64_SUPPORT.md for the current status and the real upgrade paths.
"@
}

if ($systemArchitecture -eq "ARM64" -and $AllowUnsupportedArm) {
    Write-Warning "ARM64 system detected. Proceeding only because -AllowUnsupportedArm was specified."
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
