// ============================================================
// كود Google Apps Script — استقبال الليدز من الفورم + عرضهم
// ============================================================

// غيّر الكلمة دي لأي كلمة سر تختارها إنت (تستخدمها بس إنت في صفحة الليدز)
const SECRET_KEY = 'rosto2026';

// استقبال بيانات جديدة من الفورم (POST)
function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leads');
  const data = JSON.parse(e.postData.contents);
  sheet.appendRow([
    new Date(),
    data.name || '',
    data.phone || '',
    data.city || '',
    data.units || '',
    data.q || ''
  ]);
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// عرض كل الليدز (GET) — لازم المفتاح الصح عشان يشتغل
function doGet(e) {
  if (!e.parameter.key || e.parameter.key !== SECRET_KEY) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'unauthorized' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leads');
  const rows = sheet.getDataRange().getValues();
  const headers = rows.shift();
  const leads = rows.map(function (row) {
    const obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  }).reverse(); // الأحدث فوق
  return ContentService.createTextOutput(JSON.stringify(leads))
    .setMimeType(ContentService.MimeType.JSON);
}
