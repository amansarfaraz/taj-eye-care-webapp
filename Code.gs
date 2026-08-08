/**
 * Taj Eye Care and Opticals — Google Sheets sync backend
 * -------------------------------------------------------
 * Deploy this under the Google account you want tied to the shop
 * (e.g. tajstorecustomerhelp@gmail.com) as a Web App — see README.md
 * for the step-by-step. The app POSTs each eye-test / billing record
 * and each inventory change here, writing/updating one row per item
 * across two tabs: "Records" and "Inventory". doGet supports
 * ?action=listRecords and ?action=listInventory (with a JSONP
 * ?callback= parameter, since Apps Script doesn't send the CORS
 * headers a plain fetch() would need to read a cross-origin response).
 */

const RECORDS_SHEET_NAME = 'Records';
const INVENTORY_SHEET_NAME = 'Inventory';

const RECORD_HEADERS = [
  'ID', 'Status', 'Date', 'Name', 'Phone', 'Age',
  'Previous Power Used', 'Prev R SPH', 'Prev R CYL', 'Prev R AXIS', 'Prev L SPH', 'Prev L CYL', 'Prev L AXIS',
  'Lens Category', 'New R SPH', 'New R CYL', 'New R AXIS', 'New R ADD', 'New R PD', 'New L SPH', 'New L CYL', 'New L AXIS', 'New L ADD', 'New L PD',
  'Optometrist Note',
  'Lens Type', 'Frame Type', 'Delivery Date', 'Delivery Time',
  'Total Amount', 'Advance Amount', 'Balance Amount', 'Payment Status',
  'Created At', 'Updated At'
];

const INVENTORY_HEADERS = ['ID', 'Name', 'Category', 'Stock', 'Price', 'Updated At'];

function getSheetByName_(name, headers){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if(!sheet){
    sheet = ss.insertSheet(name);
  }
  if(sheet.getLastRow() === 0){
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getRecordsSheet_(){ return getSheetByName_(RECORDS_SHEET_NAME, RECORD_HEADERS); }
function getInventorySheet_(){ return getSheetByName_(INVENTORY_SHEET_NAME, INVENTORY_HEADERS); }

/* ---------- shared row helpers ---------- */

function upsertRowById_(sheet, id, row){
  const lastRow = sheet.getLastRow();
  const ids = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat() : [];
  const rowIndex = ids.indexOf(id); // -1 if not found
  if(rowIndex === -1){
    sheet.appendRow(row);
  } else {
    sheet.getRange(rowIndex + 2, 1, 1, row.length).setValues([row]);
  }
}

function deleteRowById_(sheet, id){
  const lastRow = sheet.getLastRow();
  if(lastRow < 2) return;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const rowIndex = ids.indexOf(id);
  if(rowIndex !== -1){
    sheet.deleteRow(rowIndex + 2);
  }
}

/* ---------- records ---------- */

function recordToRow_(r){
  const p = r.power || {};
  const pp = r.prevPower || {};
  const b = r.billing || {};
  return [
    r.id, r.status, r.date, r.name, r.phone, r.age,
    r.prevPowerUsed, pp.rSph, pp.rCyl, pp.rAxis, pp.lSph, pp.lCyl, pp.lAxis,
    r.lensCategory, p.rSph, p.rCyl, p.rAxis, p.rAdd, p.rPd, p.lSph, p.lCyl, p.lAxis, p.lAdd, p.lPd,
    r.notes || '',
    b.lensType || '', b.frameType || '', b.deliveryDate || '', b.deliveryTime || '',
    b.totalAmount || '', b.advanceAmount || '', b.balanceAmount || '', b.paymentStatus || '',
    r.createdAt || '', r.updatedAt || ''
  ];
}

function rowToRecord_(headers, row){
  const get = (name) => row[headers.indexOf(name)];
  return {
    id: get('ID'), status: get('Status'), date: fmtDate_(get('Date')), name: get('Name'), phone: String(get('Phone')), age: get('Age'),
    prevPowerUsed: get('Previous Power Used'),
    prevPower: { rSph:get('Prev R SPH'), rCyl:get('Prev R CYL'), rAxis:get('Prev R AXIS'), lSph:get('Prev L SPH'), lCyl:get('Prev L CYL'), lAxis:get('Prev L AXIS') },
    lensCategory: get('Lens Category'),
    power: { rSph:get('New R SPH'), rCyl:get('New R CYL'), rAxis:get('New R AXIS'), rAdd:get('New R ADD'), rPd:get('New R PD'), lSph:get('New L SPH'), lCyl:get('New L CYL'), lAxis:get('New L AXIS'), lAdd:get('New L ADD'), lPd:get('New L PD') },
    notes: get('Optometrist Note'),
    billing: get('Lens Type') || get('Total Amount') ? {
      lensType: get('Lens Type'), frameType: get('Frame Type'),
      deliveryDate: fmtDate_(get('Delivery Date')), deliveryTime: get('Delivery Time'),
      totalAmount: get('Total Amount'), advanceAmount: get('Advance Amount'), balanceAmount: get('Balance Amount'),
      paymentStatus: get('Payment Status')
    } : null,
    createdAt: get('Created At'), updatedAt: get('Updated At')
  };
}

/* ---------- inventory ---------- */

function inventoryToRow_(item){
  return [item.id, item.name, item.category, item.stock, item.price, item.updatedAt || ''];
}

function rowToInventory_(headers, row){
  const get = (name) => row[headers.indexOf(name)];
  return { id: get('ID'), name: get('Name'), category: get('Category'), stock: get('Stock'), price: get('Price'), updatedAt: get('Updated At') };
}

function fmtDate_(v){
  if(v instanceof Date){
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v;
}

/* ---------- doPost: saveRecord / saveInventoryItem / deleteInventoryItem ---------- */

function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents);

    if(body.action === 'saveRecord' && body.record){
      const sheet = getRecordsSheet_();
      upsertRowById_(sheet, body.record.id, recordToRow_(body.record));
      return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
    }

    if(body.action === 'saveInventoryItem' && body.item){
      const sheet = getInventorySheet_();
      upsertRowById_(sheet, body.item.id, inventoryToRow_(body.item));
      return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
    }

    if(body.action === 'deleteInventoryItem' && body.id){
      const sheet = getInventorySheet_();
      deleteRowById_(sheet, body.id);
      return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ok:false, error:'Unknown action'})).setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

/* ---------- doGet: listRecords / listInventory, JSONP-wrapped ---------- */

function doGet(e){
  const callback = e && e.parameter && e.parameter.callback;
  const action = e && e.parameter && e.parameter.action;
  let payload;

  if(action === 'listRecords'){
    const sheet = getRecordsSheet_();
    const lastRow = sheet.getLastRow();
    if(lastRow < 2){
      payload = {ok:true, records:[]};
    } else {
      const values = sheet.getRange(1, 1, lastRow, RECORD_HEADERS.length).getValues();
      const headers = values[0];
      const records = values.slice(1).filter(row => row[0]).map(row => rowToRecord_(headers, row));
      payload = {ok:true, records:records};
    }
  } else if(action === 'listInventory'){
    const sheet = getInventorySheet_();
    const lastRow = sheet.getLastRow();
    if(lastRow < 2){
      payload = {ok:true, items:[]};
    } else {
      const values = sheet.getRange(1, 1, lastRow, INVENTORY_HEADERS.length).getValues();
      const headers = values[0];
      const items = values.slice(1).filter(row => row[0]).map(row => rowToInventory_(headers, row));
      payload = {ok:true, items:items};
    }
  } else if(action === 'saveRecord'){
    // Writes go through GET (with the record as a URL parameter) rather
    // than POST, because some browsers/extensions silently block fetch()
    // POST requests to script.google.com while leaving simple <script>-tag
    // GET requests (what the read side already uses) alone. Routing writes
    // through the same proven channel sidesteps that entirely.
    try{
      const record = JSON.parse(e.parameter.data);
      const sheet = getRecordsSheet_();
      upsertRowById_(sheet, record.id, recordToRow_(record));
      payload = {ok:true};
    }catch(err){
      payload = {ok:false, error:String(err)};
    }
  } else if(action === 'saveInventoryItem'){
    try{
      const item = JSON.parse(e.parameter.data);
      const sheet = getInventorySheet_();
      upsertRowById_(sheet, item.id, inventoryToRow_(item));
      payload = {ok:true};
    }catch(err){
      payload = {ok:false, error:String(err)};
    }
  } else if(action === 'deleteInventoryItem'){
    try{
      const sheet = getInventorySheet_();
      deleteRowById_(sheet, e.parameter.id);
      payload = {ok:true};
    }catch(err){
      payload = {ok:false, error:String(err)};
    }
  } else {
    payload = {ok:true, message:'Taj Eye Care and Opticals — Sheets sync is running.'};
  }

  const json = JSON.stringify(payload);

  // Apps Script Web Apps don't send CORS headers, so a plain fetch() from a
  // different origin (like GitHub Pages) can't read the response even
  // though the request succeeds. JSONP (a <script> tag instead of fetch)
  // sidesteps that restriction — the app requests ?callback=NAME and we
  // wrap the JSON in a call to that function instead of returning raw JSON.
  if(callback){
    return ContentService.createTextOutput(callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
