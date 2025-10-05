// Groceries reminder script
// Sends weekly reminder for unchecked groceries and immediate email on new item

var CONFIG = {
  // Spreadsheet ID for: https://docs.google.com/spreadsheets/d/1UQC8zrrjyJLRDO49UrE6BIsHaYfZr8-Ci1HJMpkh9XY
  SPREADSHEET_ID: '1UQC8zrrjyJLRDO49UrE6BIsHaYfZr8-Ci1HJMpkh9XY',
  SHEET_NAME: 'item',
  // Recipients for reminders and notifications
  EMAILS: ['yusufajarmoekti@gmail.com', 'tiarahediati@gmail.com'],
  // Column indexes (1-based)
  COL_ITEM: 1,
  COL_DONE: 2
};

function getSpreadsheet() {
  if (CONFIG.SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('SPREADSHEET_ID not configured');
  return ss;
}

/**
 * Send reminder for all unchecked items in column B
 */
function sendGroceriesReminder() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // no data

  var rows = sheet.getRange(2, CONFIG.COL_ITEM, lastRow - 1, CONFIG.COL_DONE).getValues();
  var unchecked = [];

  for (var i = 0; i < rows.length; i++) {
    var item = rows[i][0];
    var done = rows[i][1];

    // Treat checkbox true as done. Blank or false -> not done
    var isDone = (done === true || String(done).toLowerCase() === 'true');

    if (item && !isDone) unchecked.push(item);
  }

  if (unchecked.length === 0) return;

  var body = 'Reminder: the following groceries are not yet checked as done:\n\n' + unchecked.join('\n');
  MailApp.sendEmail(CONFIG.EMAILS.join(','), 'Groceries reminder', body);
}

/**
 * onEdit handler to send immediate email when new row added
 * Note: simple triggers cannot send email to external addresses unless installed; recommend installing as an installable trigger
 */
/**
 * Installable onEdit handler. Sends immediate email when a new item appears in column A.
 * To avoid duplicate notifications, we persist notified items in PropertiesService.
 */
function handleGroceriesEdit(e) {
  try {
    if (!e || !e.range) return;
    var range = e.range;
    var sheet = range.getSheet();
    if (!sheet || sheet.getName() !== CONFIG.SHEET_NAME) return;

    // Only act on edits in the Item column (A)
    if (range.getColumn() === CONFIG.COL_ITEM && range.getRow() >= 2) {
      var newValue = range.getValue();
      if (!newValue) return; // ignore clears

      var key = 'notified:' + range.getRow() + '|' + String(newValue).trim();
      var props = PropertiesService.getScriptProperties();
      if (props.getProperty(key)) {
        // already notified for this row+value
        return;
      }

      // send notification and mark as notified
      sendImmediateNewItemEmail(newValue, range.getRow());
      props.setProperty(key, new Date().toISOString());
    }
  } catch (err) {
    Logger.log('Error in handleGroceriesEdit: ' + err);
  }
}

function sendImmediateNewItemEmail(itemName, row) {
  var body = 'A new grocery item was added (row ' + row + '):\n\n' + itemName + '\n\nPlease check it when you shop.';
  MailApp.sendEmail(CONFIG.EMAILS.join(','), 'New grocery item added', body);
}

/**
 * Helper to clear notification cache (useful during testing)
 */
function clearNotifiedCache() {
  PropertiesService.getScriptProperties().deleteAllProperties();
}

function createGroceriesTriggers() {
  // create a weekly time-driven trigger (e.g., Saturday 18:00 Jakarta)
  ScriptApp.newTrigger('sendGroceriesReminder')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY)
    .atHour(18)
    .create();

  // create an installable onEdit trigger for immediate emails
  ScriptApp.newTrigger('handleGroceriesEdit')
    .forSpreadsheet(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID))
    .onEdit()
    .create();
}

/**
 * Debug helper: list installed triggers for this project (logs and returns an array)
 */
function listGroceriesTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var out = triggers.map(function(t) {
    try {
      return t.getHandlerFunction() + ' | ' + t.getEventType() + ' | ' + t.getTriggerSource() + ' | ' + (t.getTriggerSourceId ? t.getTriggerSourceId() : 'n/a');
    } catch (e) {
      return 'trigger-info-error: ' + e;
    }
  });
  Logger.log(out.join('\n'));
  return out;
}

/**
 * Send a small test email to the configured recipients so you can verify MailApp works
 */
function sendTestEmail() {
  var body = 'Test email from groceries-reminder script. If you receive this, MailApp is working.';
  MailApp.sendEmail(CONFIG.EMAILS.join(','), 'Groceries reminder - test email', body);
}

/**
 * Test helper: simulate adding a new item programmatically (sets the value then calls the handler)
 * Use from the Apps Script editor: testSimulateNewItem(3, 'Milk')
 */
function testSimulateNewItem(row, value) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + CONFIG.SHEET_NAME);
  var range = sheet.getRange(row, CONFIG.COL_ITEM);
  range.setValue(value);

  // build a minimal event object and call the handler
  var e = { range: range };
  handleGroceriesEdit(e);
}
