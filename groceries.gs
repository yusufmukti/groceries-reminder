/**
 * Installable onFormSubmit handler for Form Responses tab
 * Sends immediate email when a new item is submitted via Google Form
 */
function handleFormSubmit(e) {
  try {
    if (!e || !e.values) return;
    // e.values: [Timestamp, New Item to purchase?, Done]
    var itemValue = e.values[1];
    if (!itemValue) return;
    var row = e.range ? e.range.getRow() : 'unknown';
    var key = 'formsubmit_notified:' + row + '|' + String(itemValue).trim();
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty(key)) {
      // already notified for this row+value
      return;
    }
    sendImmediateNewItemEmail(itemValue, row);
    props.setProperty(key, new Date().toISOString());
  } catch (err) {
    Logger.log('Error in handleFormSubmit: ' + err);
  }
}
/**
 * Installable onEdit handler for Form Responses tab
 * Sends immediate email when a new item is submitted via Google Form
 */
function handleFormResponseEdit(e) {
  try {
    if (!e || !e.range) return;
    var range = e.range;
    var sheet = range.getSheet();
    if (!sheet || sheet.getName() !== 'Form Responses 1') return;

    // Only act on edits in the 'New Item to purchase?' column (column 2)
    if (range.getColumn() === 2 && range.getRow() >= 2) {
      var itemValue = range.getValue();
      if (!itemValue) return;
      var key = 'form_notified:' + range.getRow() + '|' + String(itemValue).trim();
      var props = PropertiesService.getScriptProperties();
      if (props.getProperty(key)) {
        // already notified for this row+value
        return;
      }
      sendImmediateNewItemEmail(itemValue, range.getRow());
      props.setProperty(key, new Date().toISOString());
    }
  } catch (err) {
    Logger.log('Error in handleFormResponseEdit: ' + err);
  }
}
// Groceries reminder script
// Sends Daily reminder for unchecked groceries and immediate email on new item

var CONFIG = {
  // Spreadsheet ID for: https://docs.google.com/spreadsheets/d/1UQC8zrrjyJLRDO49UrE6BIsHaYfZr8-Ci1HJMpkh9XY
  SPREADSHEET_ID: '1UQC8zrrjyJLRDO49UrE6BIsHaYfZr8-Ci1HJMpkh9XY',
  SHEET_NAME: 'Form Responses 1',
  // Recipients for reminders and notifications
  EMAILS: ['yusufajarmoekti@gmail.com', 'tiarahediati@gmail.com'],
  // Column indexes (1-based)
  COL_ITEM: 2, // 'New Item to purchase?'
  COL_DONE: 3  // 'Done'
  ,
  // Attach spreadsheet export (Excel .xlsx) to emails when true
  SEND_ATTACHMENT: true
};

/**
 * Return the spreadsheet as an attachment blob (Excel .xlsx) suitable for email.
 */
function getSpreadsheetAttachmentBlob() {
  try {
    var file = DriveApp.getFileById(CONFIG.SPREADSHEET_ID);
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    var baseName = file.getName() + '-' + now;
    // Export as Excel workbook
    var blob = file.getAs(MimeType.MICROSOFT_EXCEL).setName(baseName + '.xlsx');
    return blob;
  } catch (e) {
    Logger.log('Error creating spreadsheet attachment blob: ' + e);
    return null;
  }
}

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

  var rows = sheet.getRange(2, CONFIG.COL_ITEM, lastRow - 1, 2).getValues(); // COL_ITEM and COL_DONE
  var unchecked = [];

  for (var i = 0; i < rows.length; i++) {
    var item = rows[i][0];
    var done = rows[i][1];

    // Treat checkbox true as done. Blank or false -> not done
    var isDone = (done === true || String(done).toLowerCase() === 'true');

    if (item && !isDone) unchecked.push(item);
  }

  if (unchecked.length === 0) return;

  var body = [
    'Hello,',
    '',
    'This is your daily groceries reminder. The following items are still unchecked:',
    '',
    unchecked.map(function(item, idx) { return (idx+1) + '. ' + item; }).join('\n'),
    '',
    'Please review and check them off once purchased.',
    '',
    'Best regards,',
    'Groceries Reminder Bot'
  ].join('\n');
  var subject = 'Daily Groceries Reminder: Items Pending';
  var sheetLink = 'https://docs.google.com/spreadsheets/d/' + CONFIG.SPREADSHEET_ID;
  var htmlBody = body.replace(/\n/g, '<br>') + '\n<p><a href="' + sheetLink + '" style="display:inline-block;padding:10px 14px;background:#1a73e8;color:#fff;text-decoration:none;border-radius:4px">Open Spreadsheet</a></p>';
  var options = { htmlBody: htmlBody };
  if (CONFIG.SEND_ATTACHMENT) {
    var blob = getSpreadsheetAttachmentBlob();
    if (blob) options.attachments = [blob];
  }
  MailApp.sendEmail(CONFIG.EMAILS.join(','), subject, body, options);
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

    var props = PropertiesService.getScriptProperties();

    // If editing the Item column (Form Responses 1, column 2), send new item notification
    if (range.getSheet().getName() === CONFIG.SHEET_NAME && range.getColumn() === CONFIG.COL_ITEM && range.getRow() >= 2) {
      var newValue = range.getValue();
      if (!newValue) return; // ignore clears

      var key = 'notified:' + range.getRow() + '|' + String(newValue).trim();
      if (props.getProperty(key)) {
        // already notified for this row+value
        return;
      }

      // send notification and mark as notified
      sendImmediateNewItemEmail(newValue, range.getRow());
      props.setProperty(key, new Date().toISOString());
    }

    // If editing the Done column (Form Responses 1, column 3), send checked notification if checked
    if (range.getSheet().getName() === CONFIG.SHEET_NAME && range.getColumn() === CONFIG.COL_DONE && range.getRow() >= 2) {
      var checked = range.getValue();
      var itemValue = range.getSheet().getRange(range.getRow(), CONFIG.COL_ITEM).getValue();
      if (!itemValue) return;
      var checkedKey = 'checked:' + range.getRow() + '|' + String(itemValue).trim();
      if (checked === true || String(checked).toLowerCase() === 'true') {
        if (!props.getProperty(checkedKey)) {
          sendItemCheckedEmail(itemValue, range.getRow());
          props.setProperty(checkedKey, new Date().toISOString());
        }
      } else {
        // If unchecked, clear the checked notification cache for this item
        props.deleteProperty(checkedKey);
      }
    }
  } catch (err) {
    Logger.log('Error in handleGroceriesEdit: ' + err);
  }
}

function sendImmediateNewItemEmail(itemName, row) {
  var body = [
    'Hello,',
    '',
    'A new grocery item has been added to your list:',
    '',
    'Item: ' + itemName,
    'Row: ' + row,
    '',
    'Please check it off once purchased.',
    '',
    'Best regards,',
    'Groceries Reminder Bot'
  ].join('\n');
  var subject = 'New Grocery Item Added: ' + itemName;
  var sheetLink = 'https://docs.google.com/spreadsheets/d/' + CONFIG.SPREADSHEET_ID;
  var htmlBody = body.replace(/\n/g, '<br>') + '\n<p><a href="' + sheetLink + '" style="display:inline-block;padding:10px 14px;background:#1a73e8;color:#fff;text-decoration:none;border-radius:4px">Open Spreadsheet</a></p>';
  var options = { htmlBody: htmlBody };
  if (CONFIG.SEND_ATTACHMENT) {
    var blob = getSpreadsheetAttachmentBlob();
    if (blob) options.attachments = [blob];
  }
  MailApp.sendEmail(CONFIG.EMAILS.join(','), subject, body, options);
}

/**
 * Send an email when an item is checked as done
 */
function sendItemCheckedEmail(itemName, row) {
  var body = [
    'Hello,',
    '',
    'The following grocery item has been checked as done:',
    '',
    'Item: ' + itemName,
    'Row: ' + row,
    '',
    'You may remove it from your shopping list.',
    '',
    'Best regards,',
    'Groceries Reminder Bot'
  ].join('\n');
  var subject = 'Grocery Item Checked: ' + itemName;
  var sheetLink = 'https://docs.google.com/spreadsheets/d/' + CONFIG.SPREADSHEET_ID;
  var htmlBody = bodyLines.join('\n').replace(/\n/g, '<br>') + '\n<p><a href="' + sheetLink + '" style="display:inline-block;padding:10px 14px;background:#1a73e8;color:#fff;text-decoration:none;border-radius:4px">Open Spreadsheet</a></p>';
  var options = { htmlBody: htmlBody };
  if (CONFIG.SEND_ATTACHMENT) {
    var blob = getSpreadsheetAttachmentBlob();
    if (blob) options.attachments = [blob];
  }
  MailApp.sendEmail(CONFIG.EMAILS.join(','), subject, bodyLines.join('\n'), options);
}

/**
 * Polling function to detect items checked since last run.
 * Runs every 15 minutes. Sends a single email listing all items
 * that changed from unchecked -> checked since the last poll.
 */
function pollCheckedItems() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var range = sheet.getRange(2, CONFIG.COL_ITEM, lastRow - 1, 2); // item + done
  var values = range.getValues();

  var props = PropertiesService.getScriptProperties();
  var stateJson = props.getProperty('checkedState') || '{}';
  var prevState = {};
  try { prevState = JSON.parse(stateJson); } catch (e) { prevState = {}; }

  var newlyChecked = [];
  var newState = {};

  for (var i = 0; i < values.length; i++) {
    var rowIndex = i + 2; // sheet row
    var item = values[i][0];
    var done = values[i][1];
    var isDone = (done === true || String(done).toLowerCase() === 'true');

    // record new state
    newState[rowIndex] = !!isDone;

    // if previously not done (false/undefined) and now done -> notify
    if (isDone && !prevState[rowIndex] && item) {
      newlyChecked.push({ row: rowIndex, item: item });
    }
  }

  // persist new state
  props.setProperty('checkedState', JSON.stringify(newState));

  if (newlyChecked.length === 0) return; // nothing to do

  var bodyLines = ['Hello,', '', 'The following grocery items were checked in the last 15 minutes:', ''];
  newlyChecked.forEach(function(it, idx) {
    bodyLines.push((idx+1) + '. ' + it.item + ' (row ' + it.row + ')');
  });
  bodyLines.push('', 'Best regards,', 'Groceries Reminder Bot');

  var subject = 'Checked Items Notification — ' + newlyChecked.length + ' item(s)';
  MailApp.sendEmail(CONFIG.EMAILS.join(','), subject, bodyLines.join('\n'));
}

/**
 * Helper to clear notification cache (useful during testing)
 */
function clearNotifiedCache() {
  PropertiesService.getScriptProperties().deleteAllProperties();
}

function createGroceriesTriggers() {
  // create an installable onFormSubmit trigger for form responses tab
  ScriptApp.newTrigger('handleFormSubmit')
    .forSpreadsheet(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID))
    .onFormSubmit()
    .create();
  // create a daily time-driven trigger (18:00 Jakarta)
  ScriptApp.newTrigger('sendGroceriesReminder')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();

  // create an installable onEdit trigger for immediate emails (item tab)
  ScriptApp.newTrigger('handleGroceriesEdit')
    .forSpreadsheet(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID))
    .onEdit()
    .create();

  // create an installable onEdit trigger for form responses tab
  ScriptApp.newTrigger('handleFormResponseEdit')
    .forSpreadsheet(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID))
    .onEdit()
    .create();

  // create a 1-minute polling trigger to detect checked items from mobile edits
  ScriptApp.newTrigger('pollCheckedItems')
    .timeBased()
    .everyMinutes(1)
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

/**
 * Recreate correct triggers: remove any unexpected triggers and ensure
 * `sendGroceriesReminder` (time-driven) and `handleGroceriesEdit` (installable onEdit)
 * exist. Run this once from the Apps Script editor and authorize when prompted.
 */
function recreateGroceriesTriggers() {
  var expected = { sendGroceriesReminder: true, handleGroceriesEdit: true, handleFormResponseEdit: true, handleFormSubmit: true };
  var triggers = ScriptApp.getProjectTriggers();
  var removed = [];

  // Remove triggers that are not expected
  triggers.forEach(function(t) {
    try {
      var fn = t.getHandlerFunction();
      if (!expected[fn]) {
        ScriptApp.deleteTrigger(t);
        removed.push(fn || '(unknown)');
      }
    } catch (e) {
      // ignore
    }
  });

  // Refresh list and create missing expected triggers
  var present = {};
  ScriptApp.getProjectTriggers().forEach(function(t) {
    try { present[t.getHandlerFunction()] = true; } catch (e) {}
  });

  if (!present.handleFormSubmit) {
    ScriptApp.newTrigger('handleFormSubmit')
      .forSpreadsheet(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID))
      .onFormSubmit()
      .create();
    present.handleFormSubmit = true;
  }

  if (!present.sendGroceriesReminder) {
    ScriptApp.newTrigger('sendGroceriesReminder')
      .timeBased()
      .everyDays(1)
      .atHour(18)
      .create();
    present.sendGroceriesReminder = true;
  }

  if (!present.handleGroceriesEdit) {
    ScriptApp.newTrigger('handleGroceriesEdit')
      .forSpreadsheet(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID))
      .onEdit()
      .create();
    present.handleGroceriesEdit = true;
  }

  if (!present.handleFormResponseEdit) {
    ScriptApp.newTrigger('handleFormResponseEdit')
      .forSpreadsheet(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID))
      .onEdit()
      .create();
    present.handleFormResponseEdit = true;
  }

  if (!present.pollCheckedItems) {
    ScriptApp.newTrigger('pollCheckedItems')
      .timeBased()
      .everyMinutes(1)
      .create();
    present.pollCheckedItems = true;
  }

  var ensured = Object.keys(expected).filter(function(k){ return !!present[k]; });
  Logger.log('Removed triggers: ' + (removed.length ? removed.join(', ') : 'none'));
  Logger.log('Current/ensured triggers: ' + ensured.join(', '));
  return { removed: removed, ensured: ensured };
}
