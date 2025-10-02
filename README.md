# Village Photo Uploader (React + Vite)

Mobile-friendly SPA that uploads photos to a SharePoint library and creates a list item via Microsoft Graph (delegated auth with MSAL popup). Auto-renders a Village dropdown when the list column is a Choice.

## Local Dev
1. Copy `.env.example` to `.env` and set your values.
2. Install & run:
   ```bash
   npm install
   npm run dev
   ```

## Build
```bash
npm run build
```
Output is in `dist/`.

## Deploy to **Azure App Service (Linux)** (Node runtime)

**Option A: Zip Deploy**
1. Build locally:
   ```bash
   npm install
   npm run build
   ```
2. Ensure your App Service has these **Application Settings** (Configuration → Application settings):
   - `WEBSITES_PORT` = `8080`
   - `SCM_DO_BUILD_DURING_DEPLOYMENT` = `true` (if you deploy source and want Kudu to run `npm install && npm run build`)
   - Your SPA env values as App Settings if you bake them at runtime (or keep them in the built files):
     - (If you use runtime env injection you’ll need a server; this template bakes at build time.)
3. Push the whole folder (including `package.json`) via Zip Deploy or from Git. The start command is:
   ```
   npm start
   ```
   which serves `dist` on port `8080`.

**Option B: Build on App Service**
- Set `SCM_DO_BUILD_DURING_DEPLOYMENT=true` so Kudu runs `npm install` and `npm run build`.
- Deploy the repo/folder. App will start with `npm start`.

### Required Entra App (SPA)
- Platform: **Single-page application (SPA)**
- Redirect URI: `https://<your-appservice-name>.azurewebsites.net/`
- Graph delegated permissions (scopes used by MSAL): `User.Read`, `Sites.ReadWrite.All`, `Files.ReadWrite.All`

### SharePoint
- Ensure the site, list, and library match your config. If the **Village** column is a **Choice** column, the dropdown will populate automatically.

---
_Note_: Browser SPAs cannot use certificate app-only auth. This app signs users in (delegated). For app-only with certs, put a tiny API in front (Azure Functions) and call that from this SPA.
