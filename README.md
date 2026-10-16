# ZYVORA Tools — Excel Add-in

ZYVORA Tools is a Windows Excel + Excel Online productivity add-in with seven sections:

1. **Formulas Master** — search and insert common Excel formulas.
2. **Text Tools** — TRIM, UPPER, LOWER, PROPER, CLEAN and LEN helpers.
3. **Code Generator** — generate IDs such as `=LEFT("LG",2)&L21&LEFT(SUBSTITUTE(M21,".",""),3)`.
4. **Data Tools** — auto-fit, duplicate removal, freeze top row, clear formats, sorting and fill-down.
5. **AI Tools** — generate AI-ready prompts locally without uploading workbook data.
6. **Custom Formula Builder** — build join/clean/code formulas from a simple UI.
7. **Settings & Help** — installation and optional Google services.

## No sign-up or login

There is **no ZYVORA Sign Up, Login, password, subscription account, or login screen** in this build. Open the add-in and use the tools directly.

Google connection is separate and **optional**. It authorizes access to Google Drive/Sheets through Google's OAuth flow only when the user chooses to connect it. ZYVORA does not receive or handle the Google password.

## Publish with GitHub Pages

1. Repository **Settings → Pages → Deploy from a branch**.
2. Select `main` and `/(root)` and save.
3. Verify `https://teamzyvorax.github.io/gen/taskpane.html` loads over HTTPS.
4. In Excel, use **Insert → My Add-ins / Upload My Add-in** to select `manifest.xml` when sideloading is available.

## Optional Google services

Google is not required for the core Excel tools.

1. Open Google Cloud Console and create/select a project.
2. Configure the OAuth consent screen.
3. Enable Google Drive API and Google Sheets API.
4. Create an OAuth Client ID of type **Web application**.
5. Add `https://teamzyvorax.github.io` as an authorized JavaScript origin.
6. Replace the placeholder client ID in `taskpane.js`.
7. Never put a Google client secret in browser code.

The optional integration lists Drive files and Google spreadsheets. The OAuth access token is kept in memory for the current task-pane session.

## AI

AI prompt helpers currently work locally. They do **not** send workbook data to an AI provider. To add live AI, connect a secure backend/provider and keep API keys off the client.

## Compatibility

XLOOKUP and UNIQUE require compatible modern Excel versions. Excel API support, sideloading, and custom add-in deployment can depend on the Microsoft 365 build and organization policy. Test data-changing actions on a workbook copy first.

## Repository

https://github.com/teamzyvorax/gen
