# AUV Procurement Tracker — technical setup

This is the IT/admin/developer reference. For "how do I use this," see `README.md`.

## Going live: what an Entra ID admin needs to do

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
   a one-time Graph call to attach the app to that specific site — role must be
   `owner` (or `fullcontrol`), not `write`: creating a new list (which this app does
   on first run) needs `Sites.Selected+Owner` per
   [Microsoft's permissions table](https://learn.microsoft.com/graph/permissions-selected-overview#how-selected-scopes-work-with-sharepoint-and-onedrive-permissions);
   `write` only covers editing existing list items.
   ```http
   POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
   Content-Type: application/json

   {
     "roles": ["owner"],
     "grantedToIdentities": [{
       "application": { "id": "{app-client-id}", "displayName": "AUV Procurement Tracker" }
     }]
   }
   ```
   Run via [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer),
   signed in as a Global/SharePoint Admin who has also consented to `Sites.FullControl.All`
   (delegated) for Graph Explorer itself — that's a separate consent from this app's.
   No native SharePoint admin center UI for this exists in this tenant (checked
   Active sites → Infra Technology → no Permissions/API access tab) — Graph Explorer
   is the only path.
4. **Hand Ryan**: the **Application (client) ID** and **Directory (tenant) ID** from
   step 1's app registration Overview page. Paste them into `config.js`:
   ```js
   clientId: "...",
   tenantId: "...",
   ```
5. If the app registration restricts sign-in to specific users/groups ("Assignment
   required" on the Enterprise Application), make sure everyone who needs to use the
   tracker — including Joseph — is in an assigned group. Right now only the
   **Infra Technology** group is assigned; anyone outside it will get an "unauthorized"
   error on sign-in unless added.

Once that's done, reload the app and click **Sign in with Microsoft**. The first
person to sign in will trigger creation of the `AUV Procurement Tracker` SharePoint
list (in `Infra-Technology`) and seed it from `data/seed.json` if it doesn't exist yet.
`data/seed.json` is only ever read on that first creation — after that it's inert, all
reads/writes go straight to the live SharePoint list.

## Hosting

Live at **https://witty-mushroom-0d743f603.7.azurestaticapps.net/**, deployed from
[github.com/ryanchoy-skyports/auv-procurement-tracker](https://github.com/ryanchoy-skyports/auv-procurement-tracker)
via the Azure Static Web Apps GitHub Action Azure generated automatically
(`.github/workflows/azure-static-web-apps-witty-mushroom-0d743f603.yml`) — push to
`main` to redeploy. Redirect URI for the Entra app registration should be set to the
URL above.

## Local dev

```powershell
./serve.ps1
```
Then open http://127.0.0.1:8844/. Without `clientId`/`tenantId` set in `config.js`,
the app runs against the bundled sample data in `data/seed.json` (demo mode) —
nothing is saved anywhere, which is useful for reviewing UI changes before touching
the real SharePoint list.

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
manual "last updated" tracking needed. There's no way to rename an item's title from
inside the app itself — that needs a direct edit in the SharePoint list.

Seed data in `data/seed.json` is the 26-item AUV Procurement Scope Matrix (verified
directly against `AUV Procurement Scope Matrix.xlsx`) plus the Ceilometer line, with
status/notes for the items covered in the 13 Jul procurement update email filled in
from that email.

`AUV Fit Out Schedule.xlsx` (installation task plan with start/end dates, RAG status,
assignees) is a separate document — task-level install scheduling, not procurement
cost/status — and is intentionally not reflected in this tracker.
