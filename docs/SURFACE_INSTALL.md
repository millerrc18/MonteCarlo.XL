# Surface Pro Install Guide

This guide is the best current path for trying `MonteCarlo.XL` on a **Surface Pro running Windows on ARM**.

Important:

- Do **not** use the Excel-DNA `.xll` installer as your main path on native ARM Excel.
- The current ARM path is the new **Office.js + .NET WebAssembly** host.
- This is still an **experimental sideload workflow**, not a finished end-user installer.

If you only need the add-in to work with the least risk tomorrow, an x64 Windows Excel machine is still the safer choice.

## What You Are Installing

On the Surface, you are installing:

- the Office.js task pane host from `src/MonteCarlo.OfficeAddin`
- the browser-hosted engine bridge from `src/MonteCarlo.Engine.Wasm`
- the Office add-in manifest at `src/MonteCarlo.OfficeAddin/manifest.xml`

You are **not** installing the packed `.xll`.

## Prerequisites

Before you start, make sure the Surface has:

1. Windows 11 on ARM
2. Desktop Excel for Windows
3. Microsoft Edge installed
4. Node.js and `npm`
5. .NET 8 SDK
6. A local checkout of the `MonteCarlo.XL` repo

Microsoft's current Office add-in requirements and sideload guidance:

- Office add-in requirements:
  https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/requirements-for-running-office-add-ins
- Sideload from a shared folder on Windows desktop Office:
  https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins

## Step 1: Start The Local Office Host

Open **Windows PowerShell** on the Surface and run:

```powershell
cd "C:\path\to\MonteCarlo.XL\src\MonteCarlo.OfficeAddin"
npm install
npm run dev
```

Leave that terminal window running.

Why: the current manifest points to `https://localhost:3000`, so Excel must be able to reach the task pane and custom-function files from the Surface itself.

## Step 2: Trust The Local HTTPS Endpoint

Open this address in **Microsoft Edge** on the Surface:

```text
https://localhost:3000/taskpane.html
```

If the browser warns about the certificate:

1. continue to the page
2. if needed, trust the local development certificate on the machine

Office add-ins should use HTTPS, and Microsoft says self-signed certificates are acceptable for development and testing as long as the certificate is trusted locally.

## Step 3: Fix The `localhost` Loopback Issue If Excel Blocks It

If Excel later shows:

```text
We can't open this add-in from localhost.
```

open an **Administrator Command Prompt** and run:

```cmd
CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"
```

Microsoft reference:
https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/office-suite-issues/cannot-open-add-in-from-localhost

## Step 4: Create A Shared Folder Catalog On The Surface

Excel desktop on Windows needs a trusted add-in catalog for this sideload flow.

Create a folder such as:

```text
C:\OfficeAddins
```

Copy this file into it:

```text
src\MonteCarlo.OfficeAddin\manifest.xml
```

Then share the folder in Windows:

1. In File Explorer, right-click `C:\OfficeAddins`
2. Open `Properties`
3. Open the `Sharing` tab
4. Click `Share`
5. Add your Windows user with read/write access
6. Finish sharing and note the **UNC path**

Example UNC path:

```text
\\SurfaceName\OfficeAddins
```

## Step 5: Add The Shared Folder As A Trusted Catalog In Excel

In desktop Excel on the Surface:

1. Open a blank workbook
2. Go to `File > Options`
3. Open `Trust Center`
4. Click `Trust Center Settings`
5. Open `Trusted Add-in Catalogs`
6. In `Catalog Url`, enter the UNC path from the previous step
7. Click `Add catalog`
8. Enable `Show in Menu`
9. Click `OK`
10. Close Excel completely
11. Reopen Excel

## Step 6: Sideload The Add-In

In Excel:

1. Go to `Home > Add-ins > Advanced`
2. Open the `Shared Folder` tab
3. Select `MonteCarlo.XL Office Host`
4. Click `Add`

Depending on your Excel build, the entry point may appear under a slightly older add-in UI path, but `Shared Folder` is the important target.

## Step 7: First-Run Verification

When the add-in opens, verify:

1. the task pane loads
2. it does not fail on `localhost`
3. it can read the active workbook
4. it can scan `MC.*` formulas on the active sheet
5. it can add a selected output cell

If you want a quick functional smoke test:

1. open `samples\MonteCarlo.XL.Smoke.xlsx`
2. open the Office host task pane
3. scan formulas
4. add output cell `B9`
5. run a small simulation
6. export a summary sheet

## What Works Today

On the current Office host foundation, the following pieces are implemented:

- task pane build and sideload path
- `MC.*` custom functions for the Office host
- shared formula parsing and deterministic normal-mode behavior
- workbook scan for `MC.*` formulas
- add selected input and output
- simulation execution by writing sampled input values into Excel and recalculating
- in-pane histogram, CDF, and sensitivity charts
- summary sheet export with native Excel charts
- benchmark call into the shared engine layer

## What Is Still Experimental

You should still treat the Surface path as a real prototype, not a finished deployment:

- no polished installer yet
- no production hosting story yet
- no full parity with the Excel-DNA ribbon/workflows yet
- richer settings, sharing, and goal-seek flows still need validation in the Office host
- persistence and reopen behavior still need more real-device testing

## Troubleshooting

### The Task Pane Does Not Open

Check:

- the `npm run dev` terminal is still running
- `https://localhost:3000/taskpane.html` loads in Edge
- the certificate is trusted locally
- Excel was restarted after adding the trusted catalog

### Excel Says It Cannot Open The Add-In From `localhost`

Run the loopback exemption command from **Step 3**, then restart Excel.

### The Add-In Is Missing From `Shared Folder`

Check:

- the manifest file is really in the shared folder
- the UNC path in `Trusted Add-in Catalogs` is correct
- `Show in Menu` is enabled
- Excel was fully closed and reopened

### Port `3000` Is Already In Use

The current manifest is hardcoded to `https://localhost:3000`.

If something else is using that port, either:

1. stop the other process, or
2. change both:
   - `src/MonteCarlo.OfficeAddin/vite.config.ts`
   - `src/MonteCarlo.OfficeAddin/manifest.xml`

Then restart the dev server and re-sideload.

### `npm run build` Worked, But Excel Still Does Not Install Anything

That is expected. `npm run build` creates a bundle, but it does **not** create a finished Surface installer by itself.

For the current repo state, the simplest working path is still:

- `npm run dev`
- local HTTPS on the Surface
- sideload through the trusted shared-folder catalog

## Related Docs

- [ARM64 support status](ARM64_SUPPORT.md)
- [Quickstart](QUICKSTART.md)
- [README](../README.md)
