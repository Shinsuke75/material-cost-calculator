const SHEET_ID = 'YOUR_SPREADSHEET_ID'; // プレースホルダーのまま

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const now = new Date();

    if (payload.type === 'apply') {
      const applySheet = getOrCreateSheet_(ss, '申請記録');
      applySheet.appendRow([
        now,
        payload.studentId || '',
        payload.studentName || '',
        JSON.stringify(payload.materials || []),
        payload.totalAmount || 0,
        '',
        payload.notes || '',
        ''
      ]);
    } else if (payload.type === 'approve') {
      const approveSheet = getOrCreateSheet_(ss, '承認記録');
      approveSheet.appendRow([
        now,
        payload.studentId || '',
        payload.studentName || '',
        JSON.stringify(payload.materials || []),
        payload.totalAmount || 0,
        payload.payDate || '',
        payload.notes || '',
        payload.teacherName || ''
      ]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'getPrices') {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName('単価マスタ');
      if (!sheet || sheet.getLastRow() === 0) {
        return ContentService
          .createTextOutput(JSON.stringify({ wood: [], filament: [] }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const values = sheet.getDataRange().getValues();
      const prices = { wood: [], filament: [] };
      values.forEach(function(row) {
        const type = String(row[0] || '').trim();
        const name = String(row[1] || '').trim();
        const unitPrice = Number(row[2]);
        if (!name || Number.isNaN(unitPrice)) return;
        if (type === 'wood') {
          prices.wood.push({ name: name, pricePerM3: unitPrice });
        } else if (type === 'filament') {
          prices.filament.push({ name: name, pricePerG: unitPrice });
        }
      });

      return ContentService
        .createTextOutput(JSON.stringify(prices))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  return sheet || ss.insertSheet(sheetName);
}
