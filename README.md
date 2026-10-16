# ZYVORA Tools — Excel Add-in

Starter project for Windows Excel and Excel Online with seven sections: Formulas Master, Text Tools, Code Generator, Data Tools, AI Tools, Custom Formula Builder, and Settings & Help.

## Publish with GitHub Pages
1. Repository Settings → Pages → Deploy from a branch → `main` → `/(root)` → Save.
2. Verify `https://teamzyvorax.github.io/gen/taskpane.html` loads over HTTPS.
3. In Excel, use Insert → My Add-ins / Upload My Add-in to select `manifest.xml` (if your Microsoft 365 account permits sideloading).

## Configure Google
1. Open https://console.cloud.google.com/ and select/create a project.
2. Configure Google Auth Platform / OAuth consent screen.
3. Enable Google Drive API and Google Sheets API.
4. Create an OAuth Client ID of type Web application.
5. Add `https://teamzyvorax.github.io` as an Authorized JavaScript origin.
6. Replace the placeholder client ID in `taskpane.js`. Never place a client secret in browser code.
7. Deploy and test Connect with Google. Google consent/verification requirements may apply.

Google integration currently lists Drive files and Google spreadsheets; it does not import sheet contents into Excel. The token is held in memory only for the current task-pane session.

## Important limits
AI Tools is a UI placeholder and needs a secure backend/provider. XLOOKUP and UNIQUE require newer Excel versions. Excel API support and sideloading depend on the user's Office build and organization policies. Test data-changing tools on a workbook copy.
