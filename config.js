// Client ID / tenant ID are public identifiers for a SPA (PKCE, no client secret
// involved or needed) — safe to ship in static JS. Never add a client secret here.
window.APP_CONFIG = {
  clientId: "6baf804e-e13a-4e6b-b7ff-733a2b06cb4c",
  tenantId: "7c20608d-4a1b-45e8-b553-3ef51e6a1960",
  redirectUri: window.location.origin + window.location.pathname,

  // SharePoint site + list that back the tracker.
  siteHostname: "skyportsuk.sharepoint.com",
  sitePath: "/sites/Infra-Technology",
  listName: "AUV Procurement Tracker",

  // Least-privilege delegated Graph scope for this site only (see README).
  graphScopes: ["Sites.Selected", "User.Read"],
};
