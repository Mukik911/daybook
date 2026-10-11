// Daybook configuration.
// Fill in clientId after registering the app in Microsoft Entra ID (see SETUP.md).
// Both values can also be set at runtime from the app's Settings screen.
window.DAYBOOK_CONFIG = {
  // Application (client) ID from Entra ID → App registrations → Daybook → Overview
  clientId: "",
  // "organizations" = any work/school account (Betsson M365).
  // Use your tenant ID to restrict sign-in to one organization,
  // or "common" to also allow personal Outlook.com accounts.
  tenant: "organizations",
};
