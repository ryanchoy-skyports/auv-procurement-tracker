// Filled in once IT registers the Entra ID app (see README.md "IT / Admin setup").
// Until clientId is set, the app runs in demo mode against data/seed.json only.
window.APP_CONFIG = {
  clientId: "", // TODO: Application (client) ID from Entra ID app registration
  tenantId: "", // TODO: Directory (tenant) ID
  redirectUri: window.location.origin + window.location.pathname,

  // SharePoint site + list that back the tracker.
  siteHostname: "skyportsuk.sharepoint.com",
  sitePath: "/sites/Infra-Technology",
  listName: "AUV Procurement Tracker",

  // Least-privilege delegated Graph scope for this site only (see README).
  graphScopes: ["Sites.Selected", "User.Read"],
};
