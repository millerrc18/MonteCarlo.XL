# Local Excel Debug Path

Use this path when validating MonteCarlo.XL on a Windows desktop Excel install.

## Build

From the repository root:

```powershell
dotnet build MonteCarlo.XL.sln
```

From WSL on this machine:

```bash
"/mnt/c/Program Files/dotnet/dotnet.exe" build MonteCarlo.XL.sln --no-restore
```

The local Office install is 64-bit Microsoft 365 Excel, so use:

```text
src/MonteCarlo.Addin/bin/Debug/net8.0-windows/publish/MonteCarlo.Addin-AddIn64-packed.xll
```

Use `MonteCarlo.Addin-AddIn-packed.xll` only for 32-bit Excel.

## Install For Local Testing

PowerShell:

```powershell
.\scripts\install-local-addin.ps1 -Configuration Debug -Build
```

The script detects Office bitness, builds if requested, copies the correct packed `.xll` to Excel's user `XLSTART` folder, and unblocks the file. Restart Excel after installing.

To remove it:

```powershell
.\scripts\install-local-addin.ps1 -Uninstall
```

## Manual Load

1. Close all Excel instances.
2. Build the solution.
3. Open Excel.
4. Use File > Open > Browse and select the 64-bit packed `.xll`.
5. Enable the add-in if Excel asks.
6. Confirm the `MonteCarlo.XL` ribbon tab appears.

If Excel blocks the add-in, place the `.xll` in a trusted location or use the install script above.

## Smoke Workbook

Create the sample workbook:

```powershell
.\samples\create-smoke-workbook.ps1
```

Then open:

```text
samples/MonteCarlo.XL.Smoke.xlsx
```

Expected smoke flow:

1. Load the add-in.
2. Open the smoke workbook.
3. Click the `MonteCarlo.XL` ribbon tab.
4. Open the task pane.
5. Add output cell `B9` on the `Smoke Model` sheet.
6. Run 1,000 iterations.
7. Confirm results appear in the task pane.

The workbook uses `MC.Normal`, `MC.Triangular`, and `MC.PERT` formulas for uncertain inputs, so the add-in can auto-detect those inputs during a run.

## Diagnostics

Startup and runtime diagnostics are written to:

```text
%LOCALAPPDATA%\MonteCarlo.XL\Logs\startup.log
```

Check this log when Excel loads the `.xll` but the ribbon/task pane does not appear, or when a simulation fails before the UI can show a useful error.
