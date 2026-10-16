# ZYVORA Tools — Excel Compatibility Build

ZYVORA Tools is an Excel productivity toolkit with seven sections: Formulas Master, Text Tools, Code Generator, Data Tools, AI Tools, Custom Formula Builder, and Settings & Help.

## No sign-up or login

There is **no ZYVORA Sign Up, Login, password, subscription account, or login screen**. Open the add-in and use the tools directly.

Google connection is separate and optional. It uses Google's OAuth flow only when the user chooses to connect it.

## Compatibility

ZYVORA now uses an **Excel API compatibility check** at startup.

- **Excel Online / Microsoft 365 / newer Excel:** full web-add-in experience where the supported Excel API is available.
- **Excel 2016/2019/2021/2024:** supported web-add-in experience, subject to the installed Office build and organization policy.
- **Excel 2013:** compatibility mode. Core formula insertion, text tools, code generation and compatible actions can work when the Office add-in platform is available. Advanced actions that require ExcelApi 1.2+ are disabled with a clear message instead of failing silently.
- **Excel 2010/2007:** use the legacy VBA compatibility module at `legacy/ZYVORA_Legacy.bas`. Import it into the VBA editor and save the workbook as `ZYVORA_Legacy.xlam`. This is a separate legacy add-in because those versions cannot run the current Office.js web add-in architecture.
- **Excel 2003 and earlier:** not supported by the modern Office.js build.

### Important formula compatibility

`XLOOKUP` and `UNIQUE` are modern worksheet functions and are marked as modern-only in the ZYVORA formula list. Older Excel users should use alternatives such as `VLOOKUP`, `INDEX/MATCH`, or helper-column techniques.

## Publish with GitHub Pages

1. Repository **Settings → Pages → Deploy from a branch**.
2. Select `main` and `/(root)` and save.
3. Verify `https://teamzyvorax.github.io/gen/taskpane.html` loads over HTTPS.
4. In supported Excel versions, use **Insert → My Add-ins / Upload My Add-in** to select `manifest.xml` when sideloading is available.

The current manifest loads the task pane from GitHub Pages and exposes the ZYVORA Tools ribbon tab where the Office build supports custom add-in commands.

## Legacy Excel 2007/2010

The file `legacy/ZYVORA_Legacy.bas` provides VBA equivalents for the main productivity actions. To make a legacy add-in:

1. Open Excel 2007/2010.
2. Press `Alt+F11` to open the VBA editor.
3. Use **File → Import File…** and select `legacy/ZYVORA_Legacy.bas`.
4. Save the workbook as **Excel Add-In (*.xlam)**, for example `ZYVORA_Legacy.xlam`.
5. Install it from **File → Options → Add-ins → Manage Excel Add-ins → Go → Browse**.
6. Run the ZYVORA legacy commands from **Developer → Macros** or assign them to a Quick Access Toolbar button.

This legacy route is intentionally separate from `manifest.xml`: the Browse dialog shown in older Excel is for `.xlam/.xla`-style add-ins, not Office web-add-in manifests.

## Optional Google services

Google is not required for the core Excel tools. Configure your own OAuth Web client ID in `taskpane.js` if you want the optional Drive/Sheets connection. Never put a Google client secret in browser code.

## AI

AI prompt helpers currently work locally. They do not send workbook data to an AI provider. For live AI, connect a secure backend/provider and keep API keys off the client.

## Repository

https://github.com/teamzyvorax/gen
