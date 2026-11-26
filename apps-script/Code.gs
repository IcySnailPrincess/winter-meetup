// Backend for the winter meetup availability page.
// Deploy: Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
const SHEET_ID = '1xjXAdBd90J1rFODyk3pdB8M5BWEF9a3peQ6510Ljn3g';

function sheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

function rows_() {
  const values = sheet_().getDataRange().getValues().slice(1);
  return values.filter(r => r[0]).map(r => {
    let slots = {};
    try { slots = JSON.parse(r[1] || '{}'); } catch (e) {}
    return { name: String(r[0]), slots: slots, updated: r[2] ? new Date(r[2]).toISOString() : null };
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_(rows_());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const name = String(body.name || '').trim().slice(0, 40);
    if (!name) return json_({ error: 'name required' });
    const sh = sheet_();
    const data = sh.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === name.toLowerCase()) { rowIdx = i + 1; break; }
    }
    if (body.delete) {
      if (rowIdx > 0) sh.deleteRow(rowIdx);
      return json_(rows_());
    }
    const slots = {};
    const allowed = { free: 1, maybe: 1, busy: 1 };
    Object.keys(body.slots || {}).forEach(k => {
      if (/^\d{4}-\d{2}-\d{2}-(AM|PM)$/.test(k) && allowed[body.slots[k]]) slots[k] = body.slots[k];
    });
    const row = [name, JSON.stringify(slots), new Date()];
    if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, 3).setValues([row]);
    else sh.appendRow(row);
    return json_(rows_());
  } finally {
    lock.releaseLock();
  }
}
