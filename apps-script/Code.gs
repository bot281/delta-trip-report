// Google Apps Script — Delta Trip Report API v3.1 (no config cache)
// Deploy as Web App (Execute as: Me, Access: Anyone)

const SHEET_ID = '1yWzRFhSScCtNdMB8Dk82sY5v9RHYzQFYSerVQA3KKzE';
const SHEET_NAME = 'Cau tra loi bieu mau 1';
const CONFIG_SHEET = 'Cau hinh';
const CODE_CHARS = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

// ===== ENTRY POINT (GET with ?data=JSON) =====
function doGet(e) {
  try {
    const dataParam = e && e.parameter && e.parameter.data;
    if (!dataParam) {
      return jsonResponse({ status: 'ok', message: 'Delta Trip Report API v3' });
    }

    const data = JSON.parse(dataParam);

    if (data.action === 'getConfig') {
      return handleGetConfig();
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (data.action === 'arrive') {
      return handleArrive(sheet, data);
    } else if (data.action === 'depart') {
      return handleDepart(sheet, data);
    } else if (data.action === 'getPending') {
      return handleGetPending(sheet, data);
    } else if (data.action === 'findByCode') {
      return handleFindByCode(sheet, data);
    }

    return jsonResponse({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// Keep doPost for backward compat
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (data.action === 'arrive') return handleArrive(sheet, data);
    if (data.action === 'depart') return handleDepart(sheet, data);
    if (data.action === 'getPending') return handleGetPending(sheet, data);
    if (data.action === 'findByCode') return handleFindByCode(sheet, data);

    return jsonResponse({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ===== CONFIG (live — no cache) =====
function handleGetConfig() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const configSheet = ss.getSheetByName(CONFIG_SHEET);
  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ config: {} });

  const data = configSheet.getRange(2, 1, lastRow - 1, 7).getValues();

  const vehicles = [], trailers = [], drivers = [], locations = [], lots = [], tasksArrive = [], tasksDepart = [];

  data.forEach(row => {
    if (row[0]) vehicles.push(String(row[0]).trim());
    if (row[1]) trailers.push(String(row[1]).trim());
    if (row[2]) drivers.push(String(row[2]).trim());
    if (row[3]) locations.push(String(row[3]).trim());
    if (row[4]) lots.push(String(row[4]).trim());
    if (row[5]) tasksArrive.push(String(row[5]).trim());
    if (row[6]) tasksDepart.push(String(row[6]).trim());
  });

  const unique = arr => [...new Set(arr)];
  const config = {
    vehicles: unique(vehicles),
    trailers: unique(trailers),
    drivers: unique(drivers),
    locations: unique(locations),
    lots: unique(lots),
    tasksArrive: unique(tasksArrive),
    tasksDepart: unique(tasksDepart),
  };

  return jsonResponse({ config: config });
}

// ===== TRIP CODE (optimized: only scan recent rows) =====
function generateTripCode(sheet) {
  const lastRow = sheet.getLastRow();
  const existingCodes = new Set();

  if (lastRow >= 2) {
    // Only check last 300 rows for active codes (pending trips)
    const scanRows = Math.min(lastRow - 1, 300);
    const startRow = lastRow - scanRows + 1;
    const codes = sheet.getRange(startRow, 21, scanRows, 1).getValues();
    const departures = sheet.getRange(startRow, 12, scanRows, 1).getValues();
    for (let i = 0; i < codes.length; i++) {
      if (!departures[i][0] && codes[i][0]) {
        existingCodes.add(String(codes[i][0]).toUpperCase());
      }
    }
  }

  for (let attempt = 0; attempt < 10; attempt++) {
    let code = '';
    for (let i = 0; i < 4; i++) {
      code += CODE_CHARS.charAt(Math.floor(Math.random() * CODE_CHARS.length));
    }
    if (!existingCodes.has(code)) {
      return code;
    }
  }

  return CODE_CHARS.charAt(Math.floor(Math.random() * CODE_CHARS.length)) +
         String(Date.now()).slice(-3);
}

// ===== ARRIVE =====
function handleArrive(sheet, data) {
  const timeStr = data.datetime || data.time || '';

  const row = [
    timeStr,                    // A: Thoi gian den noi
    'Den noi',                  // B: Den noi / Roi di
    data.vehicle || '',         // C: Ma so xe
    data.trailer || '',         // D: Bien so romooc
    data.driver || '',          // E: Ten lai xe
    data.location || '',        // F: Dia diem
    data.lot || '',             // G: So lo
    data.km || '',              // H: So km
    data.battery || '',         // I: % pin
    data.task || '',            // J: Nghiep vu
    data.note || '',            // K: Ghi chu
  ];

  const lastRow = sheet.getLastRow();
  
  // Only check last 50 rows for empty slots (instead of all rows)
  const scanStart = Math.max(4, lastRow - 50);
  const scanCount = lastRow - scanStart + 2; // +2 for next row
  const checkRange = sheet.getRange(scanStart, 1, scanCount, 12).getValues();
  let insertRow = lastRow + 1;
  for (let i = 0; i < checkRange.length; i++) {
    if (!checkRange[i][0] && !checkRange[i][11]) {
      insertRow = scanStart + i;
      break;
    }
  }
  sheet.getRange(insertRow, 1, 1, row.length).setValues([row]);

  const tripCode = generateTripCode(sheet);
  sheet.getRange(insertRow, 21).setValue(tripCode);

  return jsonResponse({
    success: true,
    rowId: insertRow,
    tripCode: tripCode,
    message: 'Da ghi nhan den noi'
  });
}

// ===== DEPART =====
function handleDepart(sheet, data) {
  const rowId = data.rowId;
  if (!rowId) return jsonResponse({ error: 'Missing rowId' });

  const now = new Date();
  const timestamp = Utilities.formatDate(now, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');

  // Batch update: write all departure columns at once (L-S = 12-19)
  const depRow = [
    timestamp,                   // L: Dau thoi gian
    'Roi di',                    // M: Den noi / Roi di
    data.trailer || '',          // N: Bien so romooc
    data.km || '',               // O: So km
    data.battery || '',          // P: % pin
    data.note || '',             // Q: Ghi chu
    data.lot || '',              // R: So lo
    data.task || '',             // S: Nghiep vu
  ];
  sheet.getRange(rowId, 12, 1, depRow.length).setValues([depRow]);

  // Driver & vehicle in V (22), W (23)
  if (data.driver || data.vehicle) {
    sheet.getRange(rowId, 22, 1, 2).setValues([[data.driver || '', data.vehicle || '']]);
  }

  return jsonResponse({
    success: true,
    message: 'Da ghi nhan roi di'
  });
}

// ===== GET PENDING (optimized: scan last 200 rows only) =====
function handleGetPending(sheet, data) {
  const driver = data.driver;
  if (!driver) return jsonResponse({ error: 'Missing driver' });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ pending: [] });

  // Pending trips are always recent — scan last 200 rows only
  const scanRows = Math.min(lastRow - 1, 200);
  const startRow = lastRow - scanRows + 1;
  const allData = sheet.getRange(startRow, 1, scanRows, 21).getValues();
  const pending = [];

  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const rowDriver = row[4];   // E: Ten lai xe
    const departTime = row[11]; // L: Dau thoi gian (departure)

    if (rowDriver === driver && !departTime) {
      pending.push({
        rowId: startRow + i,
        time: row[0],
        vehicle: row[2],
        trailer: row[3],
        driver: row[4],
        location: row[5],
        lot: row[6],
        km: row[7],
        battery: row[8],
        task: row[9],
        note: row[10],
        tripCode: row[20] || '',
      });
    }
  }

  return jsonResponse({ pending: pending });
}

// ===== FIND BY CODE (optimized: TextFinder) =====
function handleFindByCode(sheet, data) {
  const code = (data.code || '').toUpperCase().trim();
  if (!code || code.length !== 4) {
    return jsonResponse({ error: 'Ma chuyen phải có 4 ký tự' });
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ error: 'Khong tim thay chuyen' });

  // Use TextFinder on column U (21) — much faster than reading all data
  const codeColumn = sheet.getRange(2, 21, lastRow - 1, 1);
  const finder = codeColumn.createTextFinder(code).matchEntireCell(true).matchCase(false);
  const matches = finder.findAll();

  for (let m = matches.length - 1; m >= 0; m--) {
    const matchRow = matches[m].getRow();
    const departTime = sheet.getRange(matchRow, 12).getValue();
    
    if (!departTime) {
      // Found pending trip — read row data
      const row = sheet.getRange(matchRow, 1, 1, 21).getValues()[0];
      return jsonResponse({
        success: true,
        trip: {
          rowId: matchRow,
          time: row[0],
          vehicle: row[2],
          trailer: row[3],
          driver: row[4],
          location: row[5],
          lot: row[6],
          km: row[7],
          battery: row[8],
          task: row[9],
          note: row[10],
          tripCode: code,
        }
      });
    }
  }

  return jsonResponse({ error: 'Khong tim thay chuyen với mã: ' + code });
}

// ===== UTILS =====
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
