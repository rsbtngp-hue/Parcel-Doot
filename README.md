# Parcel Doot — Priority Service

Static customer portal hosted on GitHub Pages, with Google Sheets storage handled by Google Apps Script.

## Files
- `index.html` — GitHub Pages website.
- `Code.gs` — Apps Script backend source. This file is **not executed by GitHub Pages**; install it in the bound Google Apps Script project.

## Required setup before the form can save data
1. Open the Google Spreadsheet that should store inquiries, then open **Extensions → Apps Script**.
2. Replace the project code with the contents of `Code.gs` from this repository and save.
3. Run `setup()` from the Apps Script editor and authorize access. This creates `Form_Config` and `Inquiries` sheets.
4. In `setAdminPassword()`, replace the placeholder with a strong unique password of at least 14 characters, save, and run that function once. Do not publish your password in this repository.
5. **Deploy → Manage deployments → Edit** the Web app deployment, select a new version, execute as **Me**, and set access to **Anyone** only if permitted by your organization's policies. Save the deployment.
6. If the deployment URL changes, update `API_URL` near the top of `index.html` with the new `/exec` URL. The current page is configured with the previously supplied URL.

The frontend sends requests using a hidden form/iframe and receives a `postMessage` acknowledgement from Apps Script. The backend must be redeployed with this version for that protocol to work.

## Enable GitHub Pages
1. Open **Settings → Pages** for this repository.
2. Under **Build and deployment**, choose **Deploy from a branch**.
3. Select branch `main` and folder `/ (root)`, then save.
4. Wait for the Pages deployment to finish. The intended URL is https://rsbtngp-hue.github.io/Parcel-Doot/.

## Notes
- The administrator panel lets an authenticated administrator edit the form-field JSON configuration stored in Google Sheets.
- Inquiry submissions are stored in the `Inquiries` sheet, including the full submitted field set as JSON.
- Do not collect passwords, payment-card data, or unnecessary sensitive personal information.
- The starter administrator password is only a setup fallback. Replace it immediately before public use.
