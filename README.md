# AUV Procurement Tracker

A live procurement tracker for the AUV (Al Maktoum Vertiport) fit-out — status, notes,
and costings, backed by a SharePoint list so Joseph can pull confirmed costs straight
into the business model.

## Current state: demo mode

Right now `config.js` has no `clientId`/`tenantId`, so the app runs entirely against
the bundled sample data in `data/seed.json` — nothing is saved anywhere. This is
intentional: it lets anyone review and click through the look and feel before the
SharePoint connection is wired up (see below).

To try it locally:

```powershell
./serve.ps1
```

Then open http://127.0.0.1:8844/.

## Going live: what IT/an Entra ID admin needs to do

The app talks to Microsoft Graph on behalf of whoever is signed in — no secrets are
stored anywhere. Someone with **Application Administrator** (or Global Admin) rights
in the Skyports Entra ID tenant needs to:

1. **Register an app** in [Entra admin center](https://entra.microsoft.com) → App
   registrations → New registration.
   - Name: `AUV Procurement Tracker` (or similar).
   - Supported account types: single tenant (Skyports only).
   - Platform: **Single-page application** — redirect URI = the URL this app is
     hosted at once deployed (e.g. `https://<your-swa-name>.azurestaticapps.net/`).
2. **Add API permissions**: Microsoft Graph → Delegated permissions →
   `Sites.Selected` and `User.Read`. Grant admin consent.
   - We deliberately use `Sites.Selected` instead of `Sites.ReadWrite.All` — it's the
     least-privileged option and means this app can be granted access to **only** the
     `Infra-Technology` SharePoint site, not the whole tenant.
3. **Grant the app access to the `Infra-Technology` site.** `Sites.Selected` requires
   a one-time Graph call to attach the app to that specific site with write access:
   ```http
   POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
   Content-Type: application/json

   {
     "roles": ["write"],
     "grantedToIdentities": [{
       "application": { "id": "{app-client-id}", "displayName": "AUV Procurement Tracker" }
     }]
   }
   ```
   (Run with an account that has `Sites.FullControl.All` — e.g. via Graph Explorer.)
4. **Hand Ryan**: the **Application (client) ID** and **Directory (tenant) ID** from
   step 1's app registration Overview page. Paste them into `config.js`:
   ```js
   clientId: "...",
   tenantId: "...",
   ```
5. If the app registration restricts sign-in to specific users/groups ("Assignment
   required" on the Enterprise Application), make sure everyone who needs to use the
   tracker — including Joseph — is in an assigned group. Right now only the
   **Infra Technology** group is assigned; Joseph will get an "unauthorized" error on
   sign-in unless he's added.

Once that's done, reload the app and click **Sign in with Microsoft**. The first
person to sign in will trigger creation of the `AUV Procurement Tracker` SharePoint
list (in `Infra-Technology`) and seed it from `data/seed.json` if it doesn't exist yet.
After that, everyone with access to the site can read/write through the app, and the
data is also visible/editable as an ordinary SharePoint list if needed.

## Hosting

Live at **https://witty-mushroom-0d743f603.7.azurestaticapps.net/**, deployed from
[github.com/ryanchoy-skyports/auv-procurement-tracker](https://github.com/ryanchoy-skyports/auv-procurement-tracker)
via the Azure Static Web Apps GitHub Action Azure generated automatically
(`.github/workflows/azure-static-web-apps-witty-mushroom-0d743f603.yml`) — push to
`main` to redeploy. Redirect URI for the Entra app registration should be set to the
URL above.

## Data schema (SharePoint list columns)

| Column | Type | Notes |
|---|---|---|
| Item | Text | |
| CoreFunction | Text | Ops / Tech / Launch / Electrification |
| Owner | Text | |
| Status | Choice | Not Started / Pending / On Hold / Done |
| Update | Text (multi-line) | Free-text status note |
| NextAction | Text (multi-line) | |
| EstimatedCost | Currency | |
| ConfirmedCost | Currency | This is the number Joseph needs for the business model |
| CostStatus | Choice | Estimate / Quote Received / PO Issued / Invoiced / Paid |
| Currency | Text | Defaults to AED |

`lastModifiedDateTime` / `lastModifiedBy` come from Graph automatically per item — no
manual "last updated" tracking needed.

Seed data in `data/seed.json` is the 26-item AUV Procurement Scope Matrix plus the
Ceilometer line from the 13 Jul procurement update email, with status/notes filled in
from that email where known. Worth a quick sanity check against the source
`AUV Procurement Scope Matrix.xlsx` before the first real sync, since a couple of the
notes (rows 17-20) were reconstructed from a slightly ambiguous table layout.
