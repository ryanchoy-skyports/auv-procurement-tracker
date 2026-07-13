// MSAL auth + Microsoft Graph calls for the AUV Procurement Tracker.
// Loaded after msal-browser (CDN) and config.js in index.html.

const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";

const COLUMNS = [
  { name: "Item", text: {} },
  { name: "CoreFunction", text: {} },
  { name: "Owner", text: {} },
  { name: "Status", choice: { choices: ["Not Started", "Pending", "On Hold", "Done"], displayAs: "dropDownMenu" } },
  { name: "Update", text: { allowMultipleLines: true } },
  { name: "NextAction", text: { allowMultipleLines: true } },
  { name: "EstimatedCost", currency: { locale: "en-AE" } },
  { name: "ConfirmedCost", currency: { locale: "en-AE" } },
  { name: "CostStatus", choice: { choices: ["Estimate", "Quote Received", "PO Issued", "Invoiced", "Paid"] } },
  { name: "Currency", text: {} },
];

let msalApp = null;

function isConfigured() {
  return Boolean(window.APP_CONFIG && window.APP_CONFIG.clientId && window.APP_CONFIG.tenantId);
}

function getMsalApp() {
  if (!msalApp) {
    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: window.APP_CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${window.APP_CONFIG.tenantId}`,
        redirectUri: window.APP_CONFIG.redirectUri,
      },
      cache: { cacheLocation: "sessionStorage" },
    });
  }
  return msalApp;
}

// Redirect flow (not popup) — corporate browsers/policies often block popups,
// so this must work without relying on one.
async function initAuth() {
  const app = getMsalApp();
  await app.initialize();
  const result = await app.handleRedirectPromise();
  if (result) {
    app.setActiveAccount(result.account);
    return result.account;
  }
  const accounts = app.getAllAccounts();
  if (accounts.length) {
    app.setActiveAccount(accounts[0]);
    return accounts[0];
  }
  return null;
}

function signIn() {
  const app = getMsalApp();
  return app.loginRedirect({ scopes: window.APP_CONFIG.graphScopes });
}

function signOut() {
  const app = getMsalApp();
  const account = app.getActiveAccount();
  return app.logoutRedirect({ account });
}

async function getToken() {
  const app = getMsalApp();
  const account = app.getActiveAccount();
  const request = { scopes: window.APP_CONFIG.graphScopes, account };
  try {
    const result = await app.acquireTokenSilent(request);
    return result.accessToken;
  } catch (err) {
    await app.acquireTokenRedirect(request);
    throw err; // acquireTokenRedirect navigates away; this line only runs if it didn't
  }
}

async function graphFetch(path, options = {}) {
  const token = await getToken();
  const res = await fetch(`${GRAPH_ROOT}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph API error ${res.status} on ${path}: ${body}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

let cachedSiteId = null;
async function getSiteId() {
  if (cachedSiteId) return cachedSiteId;
  const { siteHostname, sitePath } = window.APP_CONFIG;
  const site = await graphFetch(`/sites/${siteHostname}:${sitePath}`);
  cachedSiteId = site.id;
  return cachedSiteId;
}

let cachedListId = null;
async function ensureList() {
  if (cachedListId) return cachedListId;
  const siteId = await getSiteId();
  const { listName } = window.APP_CONFIG;

  const existing = await graphFetch(`/sites/${siteId}/lists?$filter=displayName eq '${listName}'`);
  if (existing.value && existing.value.length) {
    cachedListId = existing.value[0].id;
    return cachedListId;
  }

  const created = await graphFetch(`/sites/${siteId}/lists`, {
    method: "POST",
    body: JSON.stringify({
      displayName: listName,
      columns: COLUMNS,
      list: { template: "genericList" },
    }),
  });
  cachedListId = created.id;

  const seedRows = await fetch("data/seed.json").then((r) => r.json());
  for (const row of seedRows) {
    await createItem(row);
  }

  return cachedListId;
}

function toFields(row) {
  return {
    Item: row.item,
    CoreFunction: row.coreFunction,
    Owner: row.owner,
    Status: row.status,
    Update: row.update,
    NextAction: row.nextAction,
    EstimatedCost: row.estimatedCost,
    ConfirmedCost: row.confirmedCost,
    CostStatus: row.costStatus,
    Currency: row.currency,
  };
}

function fromFields(item) {
  const f = item.fields || {};
  return {
    id: item.id,
    item: f.Item || "",
    coreFunction: f.CoreFunction || "",
    owner: f.Owner || "",
    status: f.Status || "Not Started",
    update: f.Update || "",
    nextAction: f.NextAction || "",
    estimatedCost: f.EstimatedCost ?? null,
    confirmedCost: f.ConfirmedCost ?? null,
    costStatus: f.CostStatus || "Estimate",
    currency: f.Currency || "AED",
    lastModified: item.lastModifiedDateTime,
    lastModifiedBy: item.lastModifiedBy?.user?.displayName || "",
  };
}

async function listItems() {
  const siteId = await getSiteId();
  const listId = await ensureList();
  const res = await graphFetch(`/sites/${siteId}/lists/${listId}/items?expand=fields&$top=200`);
  return res.value.map(fromFields);
}

async function createItem(row) {
  const siteId = await getSiteId();
  const listId = await ensureList();
  const created = await graphFetch(`/sites/${siteId}/lists/${listId}/items`, {
    method: "POST",
    body: JSON.stringify({ fields: toFields(row) }),
  });
  return created.id;
}

async function updateItem(itemId, partialRow) {
  const siteId = await getSiteId();
  const listId = await ensureList();
  const fields = toFields(partialRow);
  await graphFetch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, {
    method: "PATCH",
    body: JSON.stringify(fields),
  });
}

window.SkyportsGraph = { isConfigured, initAuth, signIn, signOut, listItems, createItem, updateItem };
