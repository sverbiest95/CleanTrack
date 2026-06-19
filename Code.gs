// ╔══════════════════════════════════════════════════════════════╗
// ║  CleanTrack — Google Apps Script backend                     ║
// ║  Plak deze volledige code in je Apps Script project          ║
// ╚══════════════════════════════════════════════════════════════╝

// ── INSTELLINGEN ─────────────────────────────────────────────────
// Vervang dit door de ID van jouw Google Sheet (zie installatiegids stap 2)
var SHEET_ID = 'JOUW_SHEET_ID_HIER';

var CLIENTS_SHEET = 'Klanten';
var APPTS_SHEET   = 'Afspraken';

// ── HOOFDFUNCTIE ─────────────────────────────────────────────────
function doPost(e) {
  var result;
  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action;

    if      (action === 'getAll')        result = getAll();
    else if (action === 'saveClient')    result = saveClient(params.client);
    else if (action === 'deleteClient')  result = deleteClient(params.id);
    else if (action === 'importClients') result = importClients(params.clients);
    else if (action === 'saveAppt')      result = saveAppt(params.appt);
    else if (action === 'deleteAppt')    result = deleteAppt(params.id);
    else result = {error: 'Onbekende actie: ' + action};

  } catch(err) {
    result = {error: err.toString()};
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// Ook GET ondersteunen voor connectiviteitstest
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({status: 'ok', app: 'CleanTrack'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── SHEET HELPERS ────────────────────────────────────────────────
function getSheet(name) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    // Voeg headers toe
    if (name === CLIENTS_SHEET) {
      sheet.appendRow(['id','name','addr','city','phone','email','defaultPrice','notes']);
    } else if (name === APPTS_SHEET) {
      sheet.appendRow(['id','clientId','date','time','service','price','status','notes']);
    }
  }
  return sheet;
}

function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var results = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue; // sla lege rijen over
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j] !== undefined ? String(data[i][j]) : '';
    }
    results.push(obj);
  }
  return results;
}

function findRowById(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-indexed
  }
  return -1;
}

// ── ACTIES ───────────────────────────────────────────────────────
function getAll() {
  var clientSheet = getSheet(CLIENTS_SHEET);
  var apptSheet   = getSheet(APPTS_SHEET);
  return {
    clients: sheetToObjects(clientSheet),
    appts:   sheetToObjects(apptSheet)
  };
}

function saveClient(client) {
  var sheet = getSheet(CLIENTS_SHEET);
  var row = findRowById(sheet, client.id);
  var values = [
    client.id, client.name, client.addr||'', client.city||'',
    client.phone||'', client.email||'', client.defaultPrice||'', client.notes||''
  ];
  if (row > 0) {
    sheet.getRange(row, 1, 1, values.length).setValues([values]);
  } else {
    sheet.appendRow(values);
  }
  return {ok: true};
}

function deleteClient(id) {
  var sheet = getSheet(CLIENTS_SHEET);
  var row = findRowById(sheet, id);
  if (row > 0) sheet.deleteRow(row);
  return {ok: true};
}

function importClients(clientsArray) {
  var sheet = getSheet(CLIENTS_SHEET);
  for (var i = 0; i < clientsArray.length; i++) {
    var c = clientsArray[i];
    var row = findRowById(sheet, c.id);
    var values = [c.id, c.name, c.addr||'', c.city||'', c.phone||'', c.email||'', c.defaultPrice||'', c.notes||''];
    if (row > 0) {
      sheet.getRange(row, 1, 1, values.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }
  }
  return {ok: true, count: clientsArray.length};
}

function saveAppt(appt) {
  var sheet = getSheet(APPTS_SHEET);
  var row = findRowById(sheet, appt.id);
  var values = [
    appt.id, appt.clientId, appt.date, appt.time||'',
    appt.service||'', appt.price||'', appt.status||'gepland', appt.notes||''
  ];
  if (row > 0) {
    sheet.getRange(row, 1, 1, values.length).setValues([values]);
  } else {
    sheet.appendRow(values);
  }
  return {ok: true};
}

function deleteAppt(id) {
  var sheet = getSheet(APPTS_SHEET);
  var row = findRowById(sheet, id);
  if (row > 0) sheet.deleteRow(row);
  return {ok: true};
}
