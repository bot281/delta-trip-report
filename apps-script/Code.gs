// Google Apps Script — Delta Trip Report API v2
// Deploy as Web App (Execute as: Me, Access: Anyone)

const SHEET_ID = '1yWzRFhSScCtNdMB8Dk82sY5v9RHYzQFYSerVQA3KKzE';
const SHEET_NAME = 'Cau tra loi bieu mau 1';
const CONFIG_SHEET = 'Cau hinh';
const EXPENSE_SHEET = 'Chi Phi';
const CODE_CHARS = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
const EXPENSE_TYPES = ['LOLO', 'Vé cổng', 'Sửa chữa', 'Khác'];
const VAT_OPTIONS = [
  { label: '10%', rate: 0.10 },
  { label: '8%', rate: 0.08 },
  { label: '0%', rate: 0.0 },
  { label: 'Không hóa đơn', rate: '' },
];

// ===== ENTRY POINT (GET with ?data=JSON) =====
function doGet(e) {
  try {
    const dataParam = e && e.parameter && e.parameter.data;
    if (!dataParam) {
      return jsonResponse({ status: 'ok', message: 'Delta Trip Report API v2' });
    }

    const data = JSON.parse(dataParam);
    return dispatchRequest(data);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// Keep doPost for backward compat
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    return dispatchRequest(data);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function dispatchRequest(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  if (data.action === 'getConfig') return handleGetConfig(ss);

  const sheet = ss.getSheetByName(SHEET_NAME);

  if (data.action === 'arrive') return handleArrive(ss, sheet, data);
  if (data.action === 'depart') return handleDepart(ss, sheet, data);
  if (data.action === 'getPending') return handleGetPending(sheet, data);
  if (data.action === 'findByCode') return handleFindByCode(ss, sheet, data);
  if (data.action === 'saveExpenses') return handleSaveExpenses(ss, sheet, data);

  return jsonResponse({ error: 'Unknown action' });
}

// ===== CONFIG =====
function handleGetConfig(ss) {
  const configSheet = ss.getSheetByName(CONFIG_SHEET);
  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) {
    return jsonResponse({ config: { expenseTypes: EXPENSE_TYPES, vatOptions: VAT_OPTIONS.map(v => v.label) } });
  }

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

  return jsonResponse({
    config: {
      vehicles: unique(vehicles),
      trailers: unique(trailers),
      drivers: unique(drivers),
      locations: unique(locations),
      lots: unique(lots),
      tasksArrive: unique(tasksArrive),
      tasksDepart: unique(tasksDepart),
      expenseTypes: EXPENSE_TYPES,
      vatOptions: VAT_OPTIONS.map(v => v.label),
    }
  });
}

// ===== TRIP CODE =====
function generateTripCode(sheet) {
  const lastRow = sheet.getLastRow();
  const existingCodes = new Set();

  if (lastRow >= 2) {
    const codes = sheet.getRange(2, 21, lastRow - 1, 1).getValues();
    const departures = sheet.getRange(2, 12, lastRow - 1, 1).getValues();
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

  let fallback = '';
  for (let i = 0; i < 4; i++) {
    fallback += CODE_CHARS.charAt(Math.floor(Math.random() * CODE_CHARS.length));
  }
  return fallback;
}

function parseWholeNumber(value) {
  const cleaned = String(value || '').trim().replace(/[^0-9]/g, '');
  if (!cleaned) return NaN;
  return Number(cleaned);
}

function validateKmValue(value) {
  const km = parseWholeNumber(value);
  if (!Number.isFinite(km) || km < 0) return 'Số km phải là số nguyên không âm';
  return null;
}

function validateBatteryValue(value) {
  const battery = parseWholeNumber(value);
  if (!Number.isFinite(battery)) return 'Pin phải là số từ 0 đến 100';
  if (battery < 0 || battery > 100) return 'Pin phải nằm trong khoảng 0-100%';
  return null;
}

function getVatRate(vatType) {
  for (let i = 0; i < VAT_OPTIONS.length; i++) {
    if (VAT_OPTIONS[i].label === vatType) return VAT_OPTIONS[i].rate;
  }
  return '';
}

function normalizeExpenseRows(rows) {
  const result = [];
  (rows || []).forEach(raw => {
    const type = String(raw.type || '').trim();
    const amountRaw = String(raw.amount || '').trim();
    const vatType = String(raw.vatType || '').trim();
    const expenseName = String(raw.expenseName || raw.customName || '').trim();
    const note = String(raw.note || '').trim();
    const amount = parseWholeNumber(amountRaw);

    const isBlank = !type && !amountRaw && !vatType && !expenseName && !note;
    if (isBlank) return;

    if (!type || EXPENSE_TYPES.indexOf(type) === -1) {
      throw new Error('Loại chi phí không hợp lệ');
    }
    if (!Number.isFinite(amount) || amount <= 0) {
      throw new Error('Số tiền chi phí phải lớn hơn 0');
    }
    if (!vatType || VAT_OPTIONS.map(v => v.label).indexOf(vatType) === -1) {
      throw new Error('VAT chi phí không hợp lệ');
    }
    if (type === 'Khác' && !expenseName) {
      throw new Error('Vui lòng nhập tên phí cho mục Khác');
    }

    result.push({
      type: type,
      amount: amount,
      vatType: vatType,
      vatRate: getVatRate(vatType),
      expenseName: type === 'Khác' ? expenseName : expenseName,
      note: note,
    });
  });
  return result;
}

function ensureExpenseSheet(ss) {
  let sheet = ss.getSheetByName(EXPENSE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(EXPENSE_SHEET);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'timestamp',
      'tripCode',
      'rowId',
      'phase',
      'driver',
      'vehicle',
      'trailer',
      'location',
      'Số lô',
      'task',
      'expenseType',
      'expenseName',
      'amount',
      'vatType',
      'vatRate',
      'note',
      'createdBy',
    ]);
  }
  return sheet;
}

function findRowByTripCode(sheet, tripCode) {
  const code = String(tripCode || '').trim().toUpperCase();
  if (!code) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const codes = sheet.getRange(2, 21, lastRow - 1, 1).getValues();
  for (let i = 0; i < codes.length; i++) {
    if (String(codes[i][0] || '').trim().toUpperCase() === code) return i + 2;
  }
  return null;
}

function getPhaseContext(row, phase) {
  if (phase === 'departure') {
    return {
      driver: row[21] || row[4] || '',
      vehicle: row[22] || row[2] || '',
      trailer: row[13] || row[3] || '',
      location: row[5] || '',
      lot: row[17] || '',
      task: row[18] || '',
    };
  }
  return {
    driver: row[4] || '',
    vehicle: row[2] || '',
    trailer: row[3] || '',
    location: row[5] || '',
    lot: row[6] || '',
    task: row[9] || '',
  };
}

function saveExpenseRows(ss, sheet, options) {
  const expenses = normalizeExpenseRows(options.expenses || []);
  if (!expenses.length) return 0;

  const phase = String(options.phase || '').trim();
  if (['arrival', 'departure'].indexOf(phase) === -1) {
    throw new Error('Phase chi phí không hợp lệ');
  }

  let rowId = Number(options.rowId || 0);
  let tripCode = String(options.tripCode || '').trim().toUpperCase();

  if (!rowId && tripCode) {
    rowId = findRowByTripCode(sheet, tripCode);
  }
  if (!rowId) throw new Error('Không xác định được chuyến để lưu chi phí');

  const row = sheet.getRange(rowId, 1, 1, 23).getValues()[0];
  tripCode = tripCode || String(row[20] || '').trim().toUpperCase();
  if (!tripCode) throw new Error('Chuyến chưa có mã chuyến');

  if (phase === 'departure' && !row[11]) {
    // Allow departure expenses after arrival if team wants to prefill later? keep permissive.
  }

  const ctx = getPhaseContext(row, phase);
  const expenseSheet = ensureExpenseSheet(ss);
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
  const createdBy = String(options.createdBy || ctx.driver || '').trim();

  const rows = expenses.map(item => [
    timestamp,
    tripCode,
    rowId,
    phase,
    ctx.driver,
    ctx.vehicle,
    ctx.trailer,
    ctx.location,
    ctx.lot,
    ctx.task,
    item.type,
    item.expenseName,
    item.amount,
    item.vatType,
    item.vatRate,
    item.note,
    createdBy,
  ]);

  expenseSheet.getRange(expenseSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return rows.length;
}

function getExpensesByTripCode(ss, tripCode) {
  const expenseSheet = ss.getSheetByName(EXPENSE_SHEET);
  const result = { arrival: [], departure: [] };
  if (!expenseSheet || expenseSheet.getLastRow() < 2) return result;

  const allRows = expenseSheet.getRange(2, 1, expenseSheet.getLastRow() - 1, 17).getValues();
  const code = String(tripCode || '').trim().toUpperCase();

  allRows.forEach(row => {
    if (String(row[1] || '').trim().toUpperCase() !== code) return;
    const phase = String(row[3] || '').trim();
    const item = {
      timestamp: row[0],
      type: row[10],
      expenseName: row[11],
      amount: row[12],
      vatType: row[13],
      vatRate: row[14],
      note: row[15],
    };
    if (phase === 'departure') result.departure.push(item);
    else result.arrival.push(item);
  });

  return result;
}

// ===== ARRIVE =====
function handleArrive(ss, sheet, data) {
  const timeStr = data.datetime || data.time || '';

  const kmError = validateKmValue(data.km);
  if (kmError) return jsonResponse({ error: kmError });

  const batteryError = validateBatteryValue(data.battery);
  if (batteryError) return jsonResponse({ error: batteryError });

  const row = [
    timeStr,
    'Đến nơi',
    data.vehicle || '',
    data.trailer || '',
    data.driver || '',
    data.location || '',
    data.lot || '',
    data.km || '',
    data.battery || '',
    data.task || '',
    data.note || '',
  ];

  sheet.appendRow(row);
  const lastRow = sheet.getLastRow();

  const tripCode = generateTripCode(sheet);
  sheet.getRange(lastRow, 21).setValue(tripCode);

  let savedExpenses = 0;
  if (data.expenses && data.expenses.length) {
    savedExpenses = saveExpenseRows(ss, sheet, {
      rowId: lastRow,
      tripCode: tripCode,
      phase: 'arrival',
      expenses: data.expenses,
      createdBy: data.driver || '',
    });
  }

  return jsonResponse({
    success: true,
    rowId: lastRow,
    tripCode: tripCode,
    savedExpenses: savedExpenses,
    message: 'Đã ghi nhận đến nơi'
  });
}

// ===== DEPART =====
function handleDepart(ss, sheet, data) {
  const rowId = data.rowId;
  if (!rowId) return jsonResponse({ error: 'Missing rowId' });

  const kmError = validateKmValue(data.km);
  if (kmError) return jsonResponse({ error: kmError });

  const batteryError = validateBatteryValue(data.battery);
  if (batteryError) return jsonResponse({ error: batteryError });

  const arriveKm = parseWholeNumber(sheet.getRange(rowId, 8).getValue());
  const departKm = parseWholeNumber(data.km);
  if (Number.isFinite(arriveKm) && Number.isFinite(departKm) && departKm < arriveKm) {
    return jsonResponse({ error: 'Km rời đi không được nhỏ hơn km đến nơi' });
  }

  const now = new Date();
  const timestamp = Utilities.formatDate(now, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');

  sheet.getRange(rowId, 12).setValue(timestamp);
  sheet.getRange(rowId, 13).setValue('Rời đi');
  sheet.getRange(rowId, 14).setValue(data.trailer || '');
  sheet.getRange(rowId, 15).setValue(data.km || '');
  sheet.getRange(rowId, 16).setValue(data.battery || '');
  sheet.getRange(rowId, 17).setValue(data.note || '');
  sheet.getRange(rowId, 18).setValue(data.lot || '');
  sheet.getRange(rowId, 19).setValue(data.task || '');

  if (data.driver) sheet.getRange(rowId, 22).setValue(data.driver);
  if (data.vehicle) sheet.getRange(rowId, 23).setValue(data.vehicle);

  const tripCode = String(sheet.getRange(rowId, 21).getValue() || '').trim().toUpperCase();
  let savedExpenses = 0;
  if (data.expenses && data.expenses.length) {
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
    message: 'Đã ghi nhận rời đi'
  });
}

function handleSaveExpenses(ss, sheet, data) {
  const phase = String(data.phase || '').trim();
  if (['arrival', 'departure'].indexOf(phase) === -1) {
    return jsonResponse({ error: 'Phase không hợp lệ' });
  }

  const saved = saveExpenseRows(ss, sheet, {
    rowId: data.rowId,
    tripCode: data.tripCode,
    phase: phase,
    expenses: data.expenses || [],
    createdBy: data.createdBy || data.driver || '',
  });

  return jsonResponse({ success: true, savedExpenses: saved, message: 'Đã lưu chi phí' });
}

// ===== GET PENDING (by driver name) =====
function handleGetPending(sheet, data) {
  const driver = data.driver;
  if (!driver) return jsonResponse({ error: 'Missing driver' });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ pending: [] });

  const allData = sheet.getRange(2, 1, lastRow - 1, 21).getValues();
  const pending = [];

  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const rowDriver = row[4];
    const departTime = row[11];

    if (rowDriver === driver && !departTime) {
      pending.push({
        rowId: i + 2,
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

// ===== FIND BY CODE =====
function handleFindByCode(ss, sheet, data) {
  const code = (data.code || '').toUpperCase().trim();
  if (!code || code.length !== 4) {
    return jsonResponse({ error: 'Mã chuyến phải có 4 ký tự' });
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ error: 'Không tìm thấy chuyến' });

  const allData = sheet.getRange(2, 1, lastRow - 1, 23).getValues();

  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const tripCode = String(row[20] || '').toUpperCase().trim();
    const departTime = row[11];

    if (tripCode === code) {
      const trip = {
        rowId: i + 2,
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
        tripCode: tripCode,
        status: departTime ? 'completed' : 'pending',
        expenses: getExpensesByTripCode(ss, tripCode),
      };

      if (departTime) {
        trip.departTime = departTime;
        trip.departTrailer = row[13];
        trip.departKm = row[14];
        trip.departBattery = row[15];
        trip.departNote = row[16];
        trip.departLot = row[17];
        trip.departTask = row[18];
        trip.departDriver = row[21] || row[4];
        trip.departVehicle = row[22] || row[2];
      }

      return jsonResponse({ success: true, trip: trip });
    }
  }

  return jsonResponse({ error: 'Không tìm thấy chuyến với mã: ' + code });
}

// ===== UTILS =====
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
