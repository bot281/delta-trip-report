// Google Apps Script — Delta Trip Report API v4 (smart cache + onEdit invalidation)
// Deploy as Web App (Execute as: Me, Access: Anyone)

const SHEET_ID = '1yWzRFhSScCtNdMB8Dk82sY5v9RHYzQFYSerVQA3KKzE';
const SHEET_NAME = 'Cau tra loi bieu mau 1';
const CONFIG_SHEET = 'Cau hinh';
const EXPENSE_SHEET = 'Chi Phi';
const CODE_CHARS = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
const EXPENSE_TYPES_FALLBACK = ['LOLO', 'Vé cổng', 'Sửa chữa', 'Khác']; // fallback if col L empty
const VAT_OPTIONS = [
  { label: '10%', rate: 0.10 },
  { label: '8%', rate: 0.08 },
  { label: '0%', rate: 0 },
  { label: 'Không hóa đơn', rate: 0 },
];
const CONFIG_CACHE_KEY = 'config_v2';
const CONFIG_CACHE_TTL = 300; // 5 minutes
const CONFIG_VERSION_KEY = 'config_version';

// ===== ENTRY POINT (GET with ?data=JSON) =====
function doGet(e) {
  try {
    const dataParam = e && e.parameter && e.parameter.data;
    if (!dataParam) {
      return jsonResponse({ status: 'ok', message: 'Delta Trip Report API v4' });
    }

    const data = JSON.parse(dataParam);

    if (data.action === 'getConfig') {
      return handleGetConfig(data.version);
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (data.action === 'arrive') {
      return handleArrive(ss, sheet, data);
    } else if (data.action === 'depart') {
      return handleDepart(ss, sheet, data);
    } else if (data.action === 'getPending') {
      return handleGetPending(sheet, data);
    } else if (data.action === 'findByCode') {
      return handleFindByCode(ss, sheet, data);
    } else if (data.action === 'saveExpenses') {
      return handleSaveExpenses(ss, sheet, data);
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

    if (data.action === 'arrive') return handleArrive(ss, sheet, data);
    if (data.action === 'depart') return handleDepart(ss, sheet, data);
    if (data.action === 'getPending') return handleGetPending(sheet, data);
    if (data.action === 'findByCode') return handleFindByCode(ss, sheet, data);
    if (data.action === 'saveExpenses') return handleSaveExpenses(ss, sheet, data);

    return jsonResponse({ error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ===== CONFIG (smart cache: 5min TTL + version check + onEdit invalidation) =====
function handleGetConfig(clientVersion) {
  const cache = CacheService.getScriptCache();
  const props = PropertiesService.getScriptProperties();
  const currentVersion = props.getProperty(CONFIG_VERSION_KEY) || '0';

  // If client has same version, return 304-like response (no data transfer)
  if (clientVersion && clientVersion === currentVersion) {
    return jsonResponse({ unchanged: true, version: currentVersion });
  }

  // Try cache
  const cached = cache.get(CONFIG_CACHE_KEY);
  if (cached) {
    return jsonResponse({ config: JSON.parse(cached), version: currentVersion, cached: true });
  }

  // Read fresh from sheet
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const configSheet = ss.getSheetByName(CONFIG_SHEET);
  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ config: {}, version: currentVersion });

  const data = configSheet.getRange(2, 1, lastRow - 1, 12).getValues();

  const vehicles = [], trailers = [], drivers = [], locations = [], lots = [], tasksArrive = [], tasksDepart = [];
  const tollStations = [], expenseTypes = [];

  data.forEach(row => {
    if (row[0]) vehicles.push(String(row[0]).trim());
    if (row[1]) trailers.push(String(row[1]).trim());
    if (row[2]) drivers.push(String(row[2]).trim());
    if (row[3]) locations.push(String(row[3]).trim());
    if (row[4]) lots.push(String(row[4]).trim());
    if (row[5]) tasksArrive.push(String(row[5]).trim());
    if (row[6]) tasksDepart.push(String(row[6]).trim());
    if (row[10]) tollStations.push(String(row[10]).trim());  // col K
    if (row[11]) expenseTypes.push(String(row[11]).trim());   // col L
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
    expenseTypes: unique(expenseTypes).length > 0 ? unique(expenseTypes) : EXPENSE_TYPES_FALLBACK,
    tollStations: unique(tollStations),
    vatOptions: VAT_OPTIONS.map(v => v.label),
  };

  // Cache for 5 minutes
  cache.put(CONFIG_CACHE_KEY, JSON.stringify(config), CONFIG_CACHE_TTL);

  return jsonResponse({ config: config, version: currentVersion, cached: false });
}

// ===== onEdit TRIGGER — invalidate cache when Cau hinh is edited =====
function onEdit(e) {
  try {
    const sheetName = e.source.getActiveSheet().getName();
    if (sheetName === CONFIG_SHEET) {
      // Bump version → clients with old version will re-fetch
      const props = PropertiesService.getScriptProperties();
      props.setProperty(CONFIG_VERSION_KEY, String(Date.now()));
      // Clear cache
      CacheService.getScriptCache().remove(CONFIG_CACHE_KEY);
    }
  } catch (err) {
    // Silent fail — don't block the edit
  }
}

// ===== TRIP CODE (optimized: only scan recent rows) =====
function generateTripCode(sheet) {
  const lastRow = sheet.getLastRow();
  const existingCodes = new Set();

  if (lastRow >= 2) {
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
    if (!existingCodes.has(code)) return code;
  }

  return CODE_CHARS.charAt(Math.floor(Math.random() * CODE_CHARS.length)) +
         String(Date.now()).slice(-3);
}

// ===== ARRIVE =====
function handleArrive(ss, sheet, data) {
  const timeStr = data.datetime || data.time || '';

  const row = [
    timeStr,                    // A
    'Den noi',                  // B
    data.vehicle || '',         // C
    data.trailer || '',         // D
    data.driver || '',          // E
    data.location || '',        // F
    data.lot || '',             // G
    data.km || '',              // H
    data.battery || '',         // I
    data.task || '',            // J
    data.note || '',            // K
  ];

  const lastRow = sheet.getLastRow();
  const scanStart = Math.max(4, lastRow - 50);
  const scanCount = lastRow - scanStart + 2;
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

  // Save expenses if provided
  let savedExpenses = 0;
  if (data.expenses && data.expenses.length > 0) {
    savedExpenses = saveExpenseRows(ss, sheet, {
      rowId: insertRow,
      tripCode: tripCode,
      phase: 'arrival',
      expenses: data.expenses,
      createdBy: data.driver || '',
    });
  }

  return jsonResponse({
    success: true,
    rowId: insertRow,
    tripCode: tripCode,
    savedExpenses: savedExpenses,
    message: 'Da ghi nhan den noi'
  });
}

// ===== DEPART =====
function handleDepart(ss, sheet, data) {
  const rowId = data.rowId;
  if (!rowId) return jsonResponse({ error: 'Missing rowId' });

  const now = new Date();
  const timestamp = Utilities.formatDate(now, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');

  const depRow = [
    timestamp,                   // L
    'Roi di',                    // M
    data.trailer || '',          // N
    data.km || '',               // O
    data.battery || '',          // P
    data.note || '',             // Q
    data.lot || '',              // R
    data.task || '',             // S
  ];
  sheet.getRange(rowId, 12, 1, depRow.length).setValues([depRow]);

  if (data.driver || data.vehicle) {
    sheet.getRange(rowId, 22, 1, 2).setValues([[data.driver || '', data.vehicle || '']]);
  }

  // Save expenses if provided
  let savedExpenses = 0;
  if (data.expenses && data.expenses.length > 0) {
    const tripCode = sheet.getRange(rowId, 21).getValue();
    savedExpenses = saveExpenseRows(ss, sheet, {
      rowId: rowId,
      tripCode: tripCode,
      phase: 'departure',
      expenses: data.expenses,
      createdBy: data.driver || '',
    });
  }

  return jsonResponse({
    success: true,
    savedExpenses: savedExpenses,
    message: 'Da ghi nhan roi di'
  });
}

// ===== SAVE EXPENSES =====
function handleSaveExpenses(ss, sheet, data) {
  if (!data.tripCode && !data.rowId) return jsonResponse({ error: 'Missing tripCode or rowId' });
  if (!data.expenses || !data.expenses.length) return jsonResponse({ error: 'No expenses' });

  let rowId = data.rowId;
  if (!rowId && data.tripCode) {
    // Find row by trip code
    const lastRow = sheet.getLastRow();
    const codeColumn = sheet.getRange(2, 21, lastRow - 1, 1);
    const finder = codeColumn.createTextFinder(data.tripCode).matchEntireCell(true).matchCase(false);
    const match = finder.findNext();
    if (!match) return jsonResponse({ error: 'Trip not found' });
    rowId = match.getRow();
  }

  const saved = saveExpenseRows(ss, sheet, {
    rowId: rowId,
    tripCode: data.tripCode || '',
    phase: data.phase || 'arrival',
    expenses: data.expenses,
    createdBy: data.createdBy || '',
  });

  return jsonResponse({ success: true, savedExpenses: saved });
}

function saveExpenseRows(ss, sheet, options) {
  const expenseSheet = ss.getSheetByName(EXPENSE_SHEET);
  if (!expenseSheet) return 0;

  const now = new Date();
  const timestamp = Utilities.formatDate(now, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');

  // Read context from the trip row
  const rowData = sheet.getRange(options.rowId, 1, 1, 23).getValues()[0];
  const ctx = {};
  if (options.phase === 'departure') {
    ctx.driver = rowData[21] || rowData[4] || '';
    ctx.vehicle = rowData[22] || rowData[2] || '';
    ctx.trailer = rowData[13] || rowData[3] || '';
    ctx.location = rowData[5] || '';
    ctx.lot = rowData[17] || '';
    ctx.task = rowData[18] || '';
  } else {
    ctx.driver = rowData[4] || '';
    ctx.vehicle = rowData[2] || '';
    ctx.trailer = rowData[3] || '';
    ctx.location = rowData[5] || '';
    ctx.lot = rowData[6] || '';
    ctx.task = rowData[9] || '';
  }

  const rows = [];
  for (const item of options.expenses) {
    rows.push([
      timestamp,
      options.tripCode,
      options.rowId,
      options.phase,
      ctx.driver,
      ctx.vehicle,
      ctx.trailer,
      ctx.location,
      ctx.lot,
      ctx.task,
      item.type,
      item.expenseName || '',
      item.amount,
      item.vatType,
      '',
      item.note || '',
      options.createdBy,
      item.tollStation || '',   // col R: Trạm thu phí
    ]);
  }

  if (rows.length > 0) {
    const lastRow = expenseSheet.getLastRow();
    expenseSheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  return rows.length;
}

// ===== GET PENDING (optimized: scan last 200 rows only) =====
function handleGetPending(sheet, data) {
  const driver = data.driver;
  if (!driver) return jsonResponse({ error: 'Missing driver' });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ pending: [] });

  const scanRows = Math.min(lastRow - 1, 200);
  const startRow = lastRow - scanRows + 1;
  const allData = sheet.getRange(startRow, 1, scanRows, 21).getValues();
  const pending = [];

  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const rowDriver = row[4];
    const departTime = row[11];

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
function handleFindByCode(ss, sheet, data) {
  const code = (data.code || '').toUpperCase().trim();
  if (!code || code.length !== 4) {
    return jsonResponse({ error: 'Ma chuyen phải có 4 ký tự' });
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ error: 'Khong tim thay chuyen' });

  const codeColumn = sheet.getRange(2, 21, lastRow - 1, 1);
  const finder = codeColumn.createTextFinder(code).matchEntireCell(true).matchCase(false);
  const matches = finder.findAll();

  for (let m = matches.length - 1; m >= 0; m--) {
    const matchRow = matches[m].getRow();
    const rowData = sheet.getRange(matchRow, 1, 1, 23).getValues()[0];
    const departTime = rowData[11];

    // Read expenses for this trip
    const expenseSheet = ss.getSheetByName(EXPENSE_SHEET);
    let arrivalExpenses = [], departureExpenses = [];
    if (expenseSheet) {
      const expLastRow = expenseSheet.getLastRow();
      if (expLastRow >= 2) {
        const expData = expenseSheet.getRange(2, 1, expLastRow - 1, 18).getValues();
        for (const er of expData) {
          if (String(er[1]).toUpperCase() === code || Number(er[2]) === matchRow) {
            const expense = { type: er[10], expenseName: er[11], amount: er[12], vatType: er[13], note: er[15], tollStation: er[17] || '' };
            if (er[3] === 'departure') departureExpenses.push(expense);
            else arrivalExpenses.push(expense);
          }
        }
      }
    }

    const trip = {
      rowId: matchRow,
      time: rowData[0],
      vehicle: rowData[2],
      trailer: rowData[3],
      driver: rowData[4],
      location: rowData[5],
      lot: rowData[6],
      km: rowData[7],
      battery: rowData[8],
      task: rowData[9],
      note: rowData[10],
      tripCode: code,
      status: departTime ? 'completed' : 'pending',
      expenses: { arrival: arrivalExpenses, departure: departureExpenses },
    };

    if (departTime) {
      trip.departTime = rowData[11];
      trip.departTrailer = rowData[13];
      trip.departKm = rowData[14];
      trip.departBattery = rowData[15];
      trip.departNote = rowData[16];
      trip.departLot = rowData[17];
      trip.departTask = rowData[18];
      trip.departDriver = rowData[21];
      trip.departVehicle = rowData[22];
    }

    return jsonResponse({ success: true, trip: trip });
  }

  return jsonResponse({ error: 'Khong tim thay chuyen với mã: ' + code });
}

// ===== UTILS =====
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
