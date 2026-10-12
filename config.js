// RightOwner settings.
window.RIGHTOWNER_CONFIG = {
  // Application (client) ID from the RightOwner app registration in Microsoft Entra.
  clientId: "c7bef32f-a7e5-44d7-a01e-ee97f9afe064",
  listName: "RightOwner Directory",
  supportEmail: "support@rubuild.com",
  version: "1.0.0",
  teamsAppId: "8e9cfc41-8ed3-40a8-9da2-8c192c816a7d",
  // Billing service (Cloudflare Worker). Leave empty to switch paid plans off (everything unlimited).
  billingUrl: "https://rightowner-billing.deepakdas.workers.dev",
  // RightOwner's page on Microsoft AppSource (used by the Upgrade button outside Teams).
  storeUrl: "",
  // Plan limits.
  freeItemLimit: 25,
  proMaxPeople: 250,
  checkupDays: 90
};
