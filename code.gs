const SHEET_NAME = 'Website Leads';
const SHEET_ID = '1VDv2Czf0oXwSWF04jPrr8hBb5vwY2eVodXj3yFo8ZR8';
const NOTIFICATION_EMAIL = 'varunsures@gmail.com';

function doGet() {
  return json_({ ok: true, message: 'Lead endpoint is running' });
}

function doPost(e) {
  try {
    const p = (e && e.parameter) ? e.parameter : {};
    if (p.website) return json_({ ok: true, ignored: true });
    const name = clean_(p.name, 100);
    const phone = clean_(p.phone, 20);
    const email = clean_(p.email, 150);
    const plan = clean_(p.plan, 120);
    if (!name || !phone || !plan) return json_({ ok: false, error: 'Missing required fields' });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      let sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['Timestamp', 'Name', 'Mobile', 'Email', 'Interested Plan', 'Preferred Callback', 'Message', 'Source']);
        sheet.setFrozenRows(1);
      }
      sheet.appendRow([
        new Date(), safeCell_(name), safeCell_(phone), safeCell_(email), safeCell_(plan),
        safeCell_(clean_(p.callback, 80)), safeCell_(clean_(p.message, 500)), safeCell_(clean_(p.source, 500))
      ]);
    } finally { lock.releaseLock(); }

    if (NOTIFICATION_EMAIL) {
      MailApp.sendEmail({
        to: NOTIFICATION_EMAIL,
        subject: 'New website enquiry: ' + name,
        body: 'Name: ' + name + '\nMobile: ' + phone + '\nEmail: ' + email + '\nPlan: ' + plan +
              '\nCallback: ' + clean_(p.callback, 80) + '\nMessage: ' + clean_(p.message, 500)
      });
    }
    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  }
}

function clean_(value, max) {
  return String(value || '').trim().replace(/[\u0000-\u001F\u007F]/g, '').slice(0, max);
}
function safeCell_(value) {
  const s = String(value || '');
  return /^[=+@\-]/.test(s) ? "'" + s : s;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}