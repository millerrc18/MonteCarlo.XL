# ARM64 Support Status

As of April 23, 2026, MonteCarlo.XL still does not ship a production-ready native ARM64 Excel add-in build, but the repository now contains the first working Office.js plus .NET WebAssembly host foundation for that path.

## Current Status

- The add-in host is Excel-DNA.
- This repository currently packages `ExcelDna.xll` and `ExcelDna64.xll`, which cover 32-bit and 64-bit Excel on Windows.
- The packaged local artifact we install for modern Excel is `MonteCarlo.Addin-AddIn64-packed.xll`, which is an x64 XLL.
- The repository now also contains:
  - `src/MonteCarlo.Shared` for host-neutral formula, parser, workbook-config, and simulation-analysis contracts
  - `src/MonteCarlo.Engine.Wasm` for browser-hosted `[JSExport]` bridge methods over the simulation engine
  - `src/MonteCarlo.OfficeAddin` for an Office.js task pane, custom functions, workbook scanning, simulation execution, and Excel-native summary-chart export
- Microsoft 365 on Windows ARM now uses ARM-native Office apps on Windows 11 ARM devices.

That means the current production add-in should still be treated as x86/x64-only, while the ARM path is now an active dual-host implementation track.

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
- Building the experimental Office.js host from this repo and bundling it successfully on the local development machine

## What Exists Now For The ARM Path

The new ARM-oriented host is not just a note in the roadmap anymore. The repo now has:

- a shared formula catalog and parser used by both the Excel-DNA add-in and the Office host
- shared deterministic `MC.*` normal-mode evaluation logic
- a shared JSON bridge contract for validation, sampling, simulation analysis, goal-seek metrics, and benchmarks
- a browser-hosted .NET engine bridge that publishes into the Office add-in's `public/wasm` folder
- an Office.js React task pane with:
  - scan `MC.*` formulas on the active worksheet
  - add selected input and output cells
  - run simulations by writing sampled inputs into Excel, recalculating, and reading outputs
  - render histogram, CDF, and sensitivity charts in-pane
  - export a summary worksheet with native Excel charts
  - sideload manifest and custom function metadata for `MC.*` worksheet functions

## What Is Still Missing Before I Would Trust It On A Surface Pro Tomorrow

- Manual acceptance on a real ARM Surface Pro running native desktop Excel
- Full ribbon-command parity with the x64 Excel-DNA add-in
- End-to-end parity for every task-pane workflow, especially settings scopes, richer goal seek, sharing utilities, and cancel/recovery UX
- Deployment packaging beyond the current localhost sideload path
- Real-world validation for custom XML persistence, hidden-sheet fallback, and workbook reopen behavior in the Office host
- Performance tuning and optional `wasm-tools` optimization; the current publish succeeds without it but explicitly says the optimized workload is still recommended

## What To Expect On A Surface Pro Or Other Windows ARM Device

- If Excel is running as a native ARM64 process, the current Excel-DNA `.xll` add-in should still be considered unsupported.
- The install script now stops on ARM64 by default so you do not get a misleading "installed successfully" result.
- You can override that guard with `-AllowUnsupportedArm` if you have a verified x64-compatible Excel environment and want to experiment, but that is not a supported path.
- The new Office.js host is the path that is meant to become ARM-native, but it is currently an experimental sideload workflow, not a finished installer.

## How To Try The New Office Host

From the repo root on the target Windows machine:

```powershell
cd src\MonteCarlo.OfficeAddin
npm install
npm run dev
```

Then sideload:

- manifest: `src\MonteCarlo.OfficeAddin\manifest.xml`
- task pane origin: `https://localhost:3000`

For a static production-style bundle instead of live development hosting:

```powershell
cd src\MonteCarlo.OfficeAddin
npm run build
```

That writes the bundled task pane into `src\MonteCarlo.OfficeAddin\dist` and republishes the `.NET` browser bridge into `src\MonteCarlo.OfficeAddin\public\wasm` during the build.

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

This is the cleaner product path if ARM support matters strategically, and it is the path now being implemented in this repository.

Remaining work:

- Broaden the current Office task pane into near-parity with the Excel-DNA host.
- Finish ARM hardware validation and deployment packaging.
- Decide when the Office.js host is trustworthy enough to call supported instead of experimental.

This is a larger rewrite, but it is also the path that improves portability instead of only unblocking one architecture.

## Recommendation

If the immediate goal is to use MonteCarlo.XL at work tomorrow, the fastest reliable answer is still an x64 Windows Excel machine.

If ARM support becomes a real product requirement, treat it as a host-platform project and choose deliberately between:

- a custom ARM-capable Excel-XLL host, or
- an Office.js host alongside or instead of Excel-DNA.
