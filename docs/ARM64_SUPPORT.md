# ARM64 Support Status

As of April 22, 2026, MonteCarlo.XL does not ship a native ARM64 Excel add-in build.

## Current Status

- The add-in host is Excel-DNA.
- This repository currently packages `ExcelDna.xll` and `ExcelDna64.xll`, which cover 32-bit and 64-bit Excel on Windows.
- The packaged local artifact we install for modern Excel is `MonteCarlo.Addin-AddIn64-packed.xll`, which is an x64 XLL.
- Microsoft 365 on Windows ARM now uses ARM-native Office apps on Windows 11 ARM devices.

That combination means the current MonteCarlo.XL build should be treated as x86/x64-only.

## Why It Does Not Work Natively Today

There are two blockers:

1. The host loader is the issue, not the simulation engine.

   `MonteCarlo.Engine`, most of `MonteCarlo.UI`, and the charting code are ordinary .NET projects that could be made ARM-friendly. But the Excel entry point is an in-process XLL loaded by Excel, and that loader has to match the Excel process architecture.

2. The packaged Excel-DNA toolchain in this repo does not include an ARM64 XLL loader.

   Today the repo references `ExcelDna.AddIn` and publishes the x64 packed `.xll`. There is no native ARM64 Excel-DNA loader wired into this build.

There is also a smaller downstream blocker already visible in the project file: native chart dependencies are currently copied from `runtimes\win-x64\native`, so even after the host-layer problem is solved, packaging still needs ARM64-aware native dependency handling.

## What Will Work Today

- Windows desktop Excel on x64 hardware
- 64-bit Excel with the packed `MonteCarlo.Addin-AddIn64-packed.xll`

## What To Expect On A Surface Pro Or Other Windows ARM Device

- If Excel is running as a native ARM64 process, the current add-in should be considered unsupported.
- The install script now stops on ARM64 by default so you do not get a misleading "installed successfully" result.
- You can override that guard with `-AllowUnsupportedArm` if you have a verified x64-compatible Excel environment and want to experiment, but that is not a supported path.

## Real Paths To ARM Support

### Path 1: Keep Excel-DNA, Build A New ARM-Capable XLL Host

This is the smallest product change, but the riskiest technical path.

Needed work:

- Replace or fork the native Excel-DNA loader so Excel can load the add-in inside native ARM64 Excel.
- Rework packing/publish so ARM64 or Arm64X artifacts are produced intentionally.
- Audit every native dependency, including SkiaSharp/HarfBuzz packaging, for ARM64.
- Validate custom task pane hosting, ribbon callbacks, COM interop, and chart rendering on real Windows ARM hardware.

This is a deep host-layer project, not a small repo tweak.

### Path 2: Add An Office.js Host For ARM And Cross-Platform Reach

This is the cleaner product path if ARM support matters strategically.

Needed work:

- Build an Office Add-in host with task pane UI and custom functions.
- Reuse `MonteCarlo.Engine` as the calculation core behind the new host.
- Recreate the Excel integration layer that currently depends on Excel-DNA, COM interop, WPF task panes, and ribbon callbacks.
- Decide whether MonteCarlo.XL becomes a dual-host product: Excel-DNA for Windows power users, Office.js for ARM/web/Mac coverage.

This is a larger rewrite, but it is also the path that improves portability instead of only unblocking one architecture.

## Recommendation

If the immediate goal is to use MonteCarlo.XL at work, the fastest reliable answer is still an x64 Windows Excel machine.

If ARM support becomes a real product requirement, treat it as a host-platform project and choose deliberately between:

- a custom ARM-capable Excel-XLL host, or
- an Office.js host alongside or instead of Excel-DNA.
