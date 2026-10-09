// Run setup() once from Extensions > Apps Script in your chosen Google Sheet.
function setup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', spreadsheet.getId());
  let sheet = spreadsheet.getSheetByName('KOSICT 2026');
  if (!sheet) sheet = spreadsheet.insertSheet('KOSICT 2026');
  if (!sheet.getLastRow()) sheet.appendRow(['Timestamp','First name','Last name','Email','Company or university','Message','Event follow-up','Future updates','Consent version','Submission ID']);
  sheet.setFrozenRows(1);
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.website || !d.followUp || !d.firstName || !d.lastName || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email || '') || !/^[a-f0-9-]{36}$/i.test(d.submissionId || '')) throw new Error('Invalid submission');
    const safe = (v, max) => { const s=String(v || '').trim().slice(0,max); return /^[=+\-@]/.test(s) ? "'"+s : s; };
    lock.waitLock(15000);
    const sheet=SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getSheetByName('KOSICT 2026');
    if(sheet.getLastRow()>1 && sheet.getRange(2,10,sheet.getLastRow()-1,1).createTextFinder(d.submissionId).matchEntireCell(true).findNext()) return json({ok:true});
    sheet.appendRow([new Date(),safe(d.firstName,80),safe(d.lastName,80),safe(d.email,254),safe(d.organisation,160),safe(d.message,2000),'Yes',d.updates === true ? 'Yes':'No','KOSICT2026-v1',d.submissionId]);
    return json({ok:true});
  } catch(error) { return json({ok:false}); }
  finally { if(lock.hasLock()) lock.releaseLock(); }
}
function json(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
