Groceries Reminder - Trigger Setup

1. Update `CONFIG.SPREADSHEET_ID` in `groceries.gs` or set Script Properties with your spreadsheet ID.
2. Run `createGroceriesTriggers()` once from the Apps Script editor (or run via `clasp run`) to create the weekly reminder and installable onEdit trigger.
3. Note: Installable triggers require authorization and can send emails to external addresses.
