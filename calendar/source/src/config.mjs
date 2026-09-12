// Public Application (client) ID from your Microsoft Entra SPA registration.
// Alternatively enter the ID in the calendar's Connect Outlook dialog.
export const microsoftConfig = {
  clientId: "942aaf74-fa78-45d7-a764-a1a693ca4070",
  authority: "https://login.microsoftonline.com/consumers",
  scopes: ["User.Read", "Calendars.ReadWrite"],
};
