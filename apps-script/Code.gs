// Google Apps Script — Delta Trip Report API v2
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
      return jsonResponse({ status: 'ok', message: 'Delta Trip Report API v2' });
    }

    const data = JSON.parse(dataParam);
    const ss = SpreadsheetApp.openById(SHEET_ID);

    if (data.action === 'getConfig') {
      return handleGetConfig(ss);
    }

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

// ===== CONFIG =====
function handleGetConfig(ss) {
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

  // Deduplicate
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
    }
  });
}

// ===== TRIP CODE =====
function generateTripCode(sheet) {
  const lastRow = sheet.getLastRow();
  const existingCodes = new Set();

  if (lastRow >= 2) {
    // Column U = 21
    const codes = sheet.getRange(2, 21, lastRow - 1, 1).getValues();
    const departures = sheet.getRange(2, 12, lastRow - 1, 1).getValues();
    for (let i = 0; i < codes.length; i++) {
      // Only active (no departure) codes matter for uniqueness
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

  // Fallback: all random from CODE_CHARS
  let fallback = '';
  for (let i = 0; i < 4; i++) {
    fallback += CODE_CHARS.charAt(Math.floor(Math.random() * CODE_CHARS.length));
  }
  return fallback;
}

// ===== ARRIVE =====
function handleArrive(sheet, data) {
  // v2: datetime includes date (DD/MM/YYYY HH:mm:ss)
  const timeStr = data.datetime || data.time || '';

  // Columns A-K
  const row = [
    timeStr,                    // A: Thời gian đến nơi (now DD/MM/YYYY HH:mm)
    'Đến nơi',                  // B: Đến nơi / Rời đi
    data.vehicle || '',         // C: Mã số xe / Biển số xe
    data.trailer || '',         // D: Biển số romooc
    data.driver || '',          // E: Tên lái xe
    data.location || '',        // F: Địa điểm
    data.lot || '',             // G: Số lô
    data.km || '',              // H: Số km trên đồng hồ
    data.battery || '',         // I: % pin trên đồng hồ
    data.task || '',            // J: Nghiệp vụ
    data.note || '',            // K: Ghi chú
  ];

  sheet.appendRow(row);
  const lastRow = sheet.getLastRow();

  // Generate and store trip code in column U (21)
  const tripCode = generateTripCode(sheet);
  sheet.getRange(lastRow, 21).setValue(tripCode);

  return jsonResponse({
    success: true,
    rowId: lastRow,
    tripCode: tripCode,
    message: 'Đã ghi nhận đến nơi'
  });
}

// ===== DEPART =====
function handleDepart(sheet, data) {
  const rowId = data.rowId;
  if (!rowId) return jsonResponse({ error: 'Missing rowId' });

  const now = new Date();
  const timestamp = Utilities.formatDate(now, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');

  // Columns L-S (12-19)
  sheet.getRange(rowId, 12).setValue(timestamp);                    // L: Dấu thời gian
  sheet.getRange(rowId, 13).setValue('Rời đi');                     // M: Đến nơi / Rời đi (Chỉnh sửa)
  sheet.getRange(rowId, 14).setValue(data.trailer || '');            // N: Biển số romooc (Chỉnh sửa)
  sheet.getRange(rowId, 15).setValue(data.km || '');                 // O: Số km trên đồng hồ (Chỉnh sửa)
  sheet.getRange(rowId, 16).setValue(data.battery || '');            // P: % pin trên đồng hồ (Chỉnh sửa)
  sheet.getRange(rowId, 17).setValue(data.note || '');               // Q: Ghi chú (Chỉnh sửa)
  sheet.getRange(rowId, 18).setValue(data.lot || '');                // R: Số lô (khi rời đi)
  sheet.getRange(rowId, 19).setValue(data.task || '');               // S: Nghiệp vụ (khi rời đi)

  // v2: Departure driver & vehicle in columns V (22) and W (23)
  if (data.driver) sheet.getRange(rowId, 22).setValue(data.driver); // V: Tên lái xe (rời đi)
  if (data.vehicle) sheet.getRange(rowId, 23).setValue(data.vehicle); // W: Biển số xe (rời đi)

  return jsonResponse({
    success: true,
    message: 'Đã ghi nhận rời đi'
  });
}

// ===== GET PENDING (by driver name) =====
function handleGetPending(sheet, data) {
  const driver = data.driver;
  if (!driver) return jsonResponse({ error: 'Missing driver' });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ pending: [] });

  // Read columns A-U (1-21)
  const allData = sheet.getRange(2, 1, lastRow - 1, 21).getValues();
  const pending = [];

  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const rowDriver = row[4];   // E: Tên lái xe
    const departTime = row[11]; // L: Dấu thời gian (departure)

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
        tripCode: row[20] || '', // U: Mã chuyến
      });
    }
  }

  return jsonResponse({ pending: pending });
}

// ===== FIND BY CODE =====
function handleFindByCode(sheet, data) {
  const code = (data.code || '').toUpperCase().trim();
  if (!code || code.length !== 4) {
    return jsonResponse({ error: 'Mã chuyến phải có 4 ký tự' });
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ error: 'Không tìm thấy chuyến' });

  // Read columns A-W (1-23)
  const allData = sheet.getRange(2, 1, lastRow - 1, 23).getValues();

  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const tripCode = String(row[20] || '').toUpperCase().trim();
    const departTime = row[11]; // L: Dấu thời gian

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
      };

      // Include departure info if completed
      if (departTime) {
        trip.departTime = departTime;
        trip.departTrailer = row[13];
        trip.departKm = row[14];
        trip.departBattery = row[15];
        trip.departNote = row[16];
        trip.departLot = row[17];
        trip.departTask = row[18];
        trip.departDriver = row[21] || row[4]; // V or fallback to arrival driver
        trip.departVehicle = row[22] || row[2]; // W or fallback to arrival vehicle
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
