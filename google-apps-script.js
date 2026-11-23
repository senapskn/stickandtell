// Stick & Tell – receives answers from form.html and writes them to a Google Sheet.
//
// SETUP
// 1. Paste this whole file into Apps Script (replace everything that is there) and press Save (💾).
// 2. Choose the function "testSetup" in the toolbar and press Run. Allow the permissions Google asks for.
//    The log at the bottom shows the link of the sheet the answers will go to.
// 3. Deploy → Manage deployments → pencil (Edit) → Version: "New version" → Deploy.
//    (Execute as: Me · Who has access: Anyone). The /exec address stays the same.
// 4. Open the /exec address in your browser: it should say "Stick & Tell receiver is running".
//
// Works both when the script was opened from a sheet (Extensions → Apps Script)
// and when it was created on its own at script.google.com (then it creates its own sheet).

const SHEET_NAME = 'Responses';

function getSpreadsheet_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) return ss;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  ss = SpreadsheetApp.create('Stick & Tell – responses');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function getSheet_() {
  const ss = getSpreadsheet_();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sh = getSheet_();
    if (sh.getLastRow() === 0) {
      sh.appendRow(data.header);
      sh.setFrozenRows(1);
    }
    // Ignore a submission that is already in the sheet (the form may send twice).
    const last = sh.getLastRow();
    if (last > 1) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues().flat();
      if (ids.indexOf(data.id) !== -1) return reply_({ ok: true, duplicate: true });
    }
    if (data.rows && data.rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, data.rows.length, data.rows[0].length).setValues(data.rows);
    }
    return reply_({ ok: true });
  } catch (err) {
    return reply_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  let sheet = '';
  try { sheet = getSpreadsheet_().getUrl(); } catch (e) { sheet = 'error: ' + e; }
  return reply_({ ok: true, message: 'Stick & Tell receiver is running.', sheet: sheet });
}

// Run this once from the editor: it asks for permissions and shows where the answers will go.
function testSetup() {
  const sh = getSheet_();
  Logger.log('Answers will be written to: ' + sh.getParent().getUrl() + '  (tab: ' + SHEET_NAME + ')');
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
