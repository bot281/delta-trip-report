const REPORT_SHEETS = {
  normalized: 'Chuan hoa du lieu',
  vehicle: 'BC Xe',
  driver: 'BC Lai Xe',
  other: 'BC KHAC',
  lot: 'BC Lo Hang',
  trailer: 'BC Mooc',
};

const REPORT_TASK_COLUMNS = [
  'Lấy rỗng',
  'Hạ cont đầy',
  'Sạc pin',
  'Sửa chữa',
  'Chạy không',
  'Đổi ca',
  'Đổi xe',
  'Về lấy mooc',
  'Vệ sinh cont',
];

const CHARGE_REGEX = /sạc/i;
const TASK_SPLIT_REGEX = /\s*,\s*|\s*;\s*|\s*\n\s*/;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Delta Reports')
    .addItem('🔄 Refresh toàn bộ báo cáo', 'refreshAllReports')
    .addToUi();
}

function refreshAllReports() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sourceSheet = ss.getSheetByName(SHEET_NAME);
    if (!sourceSheet) throw new Error('Không tìm thấy sheet nguồn: ' + SHEET_NAME);

    const normalizedRows = buildNormalizedRows_(sourceSheet);
    writeNormalizedSheet_(ss, normalizedRows);
    writeVehicleReport_(ss, normalizedRows);
    writeDriverReport_(ss, normalizedRows);
    writeOtherReport_(ss, normalizedRows);
    writeLotReport_(ss, normalizedRows);
    writeTrailerReport_(ss, normalizedRows);

    SpreadsheetApp.getActive().toast('Đã refresh toàn bộ báo cáo', 'Delta Reports', 5);
  } finally {
    lock.releaseLock();
  }
}

function buildNormalizedRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return [];
  const rawRows = sheet.getRange(4, 1, lastRow - 3, 23).getValues();
  const rows = [];
  const prevByVehicle = {};

  rawRows.forEach((row, idx) => {
    const rowId = idx + 4;
    const arriveTime = normalizeDateTime_(row[0]);
    if (!arriveTime) return;

    const arriveVehicle = cleanString_(row[2]);
    const departVehicle = cleanString_(row[22]);
    const effectiveVehicle = departVehicle || arriveVehicle;
    const arriveTrailer = cleanString_(row[3]);
    const departTrailer = cleanString_(row[13]);
    const effectiveDepartTrailer = departTrailer || '';
    const arriveDriver = cleanString_(row[4]);
    const departDriver = cleanString_(row[21]);
    const effectiveDriver = departDriver || arriveDriver;
    const location = cleanString_(row[5]);
    const arriveLot = cleanString_(row[6]);
    const departLot = cleanString_(row[17]);
    const arriveTask = cleanString_(row[9]);
    const departTask = cleanString_(row[18]);
    const arriveKm = parseNumberOrBlank_(row[7]);
    const departKm = parseNumberOrBlank_(row[14]);
    const tripKm = isFiniteNumber_(arriveKm) && isFiniteNumber_(departKm) ? Math.max(0, departKm - arriveKm) : '';
    const arrivePin = parseNumberOrBlank_(row[8]);
    const departPin = parseNumberOrBlank_(row[15]);
    const arriveNote = cleanString_(row[10]);
    const departNote = cleanString_(row[16]);
    const hasDeparture = !!normalizeDateTime_(row[11]);
    const departTime = normalizeDateTime_(row[11]);
    const tripCode = cleanString_(row[20]);
    const month = arriveTime ? Utilities.formatDate(arriveTime, 'Asia/Ho_Chi_Minh', 'MM/yyyy') : '';
    const arriveDate = arriveTime ? Utilities.formatDate(arriveTime, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy') : '';

    const pinDeltaRaw = isFiniteNumber_(arrivePin) && isFiniteNumber_(departPin) ? departPin - arrivePin : '';
    const chargeContext = [arriveTask, departTask, arriveNote, departNote, location].join(' | ');
    const pinChargeAdded = isFiniteNumber_(arrivePin) && isFiniteNumber_(departPin) && departPin > arrivePin && CHARGE_REGEX.test(chargeContext)
      ? departPin - arrivePin
      : '';

    let pinUsedAdjusted = '';
    const prev = effectiveVehicle ? prevByVehicle[effectiveVehicle] : null;
    const transitPinUsed = prev && isFiniteNumber_(prev.departPin) && isFiniteNumber_(arrivePin) && arrivePin > 0
      ? Math.max(0, prev.departPin - arrivePin)
      : 0;
    const parkingPinUsed = isFiniteNumber_(arrivePin) && isFiniteNumber_(departPin) && departPin > 0
      ? Math.max(0, arrivePin - departPin)
      : 0;

    const canCountArrival = isFiniteNumber_(arrivePin) && arrivePin > 0;
    const canCountDeparture = isFiniteNumber_(departPin) && departPin > 0 && hasDeparture;
    if (canCountArrival || canCountDeparture) {
      pinUsedAdjusted = transitPinUsed + parkingPinUsed;
    }

    rows.push({
      rowId,
      tripCode,
      arriveTime,
      departTime,
      month,
      arriveDate,
      arriveVehicle,
      departVehicle,
      effectiveVehicle,
      arriveTrailer,
      departTrailer,
      effectiveDepartTrailer,
      arriveDriver,
      departDriver,
      effectiveDriver,
      location,
      arriveLot,
      departLot,
      arriveTask,
      departTask,
      arriveKm,
      departKm,
      tripKm,
      arrivePin,
      departPin,
      pinUsed: isFiniteNumber_(arrivePin) && isFiniteNumber_(departPin) ? Math.max(0, arrivePin - departPin) : '',
      arriveNote,
      departNote,
      isArriveKhac: arriveLot === 'KHAC' ? 1 : 0,
      isDepartKhac: departLot === 'KHAC' ? 1 : 0,
      khacTaskArrive: arriveLot === 'KHAC' ? arriveTask : '',
      khacTaskDepart: departLot === 'KHAC' ? departTask : '',
      hasDeparture: hasDeparture ? 1 : 0,
      pinDeltaRaw,
      pinChargeAdded,
      pinUsedAdjusted,
      vehicleChanged: departVehicle && arriveVehicle && departVehicle !== arriveVehicle ? 1 : 0,
      trailerChanged: departTrailer && arriveTrailer && departTrailer !== arriveTrailer ? 1 : 0,
    });

    if (effectiveVehicle && canCountDeparture) {
      prevByVehicle[effectiveVehicle] = {
        departPin,
        departTime,
        rowId,
      };
    }
  });

  return rows;
}

function writeNormalizedSheet_(ss, rows) {
  const headers = [
    'row_id','trip_code','arrive_time','depart_time','month','arrive_date',
    'arrive_vehicle','depart_vehicle','effective_vehicle','arrive_trailer','depart_trailer','effective_depart_trailer',
    'arrive_driver','depart_driver','effective_driver','location','arrive_lot','depart_lot',
    'arrive_task','depart_task','arrive_km','depart_km','trip_km','arrive_pin','depart_pin','pin_used',
    'arrive_note','depart_note','is_arrive_khac','is_depart_khac','khac_task_arrive','khac_task_depart',
    'has_departure','pin_delta_raw','pin_charge_added','pin_used_adjusted','vehicle_changed','trailer_changed'
  ];
  const values = rows.map(r => [
    r.rowId, r.tripCode, formatDateTime_(r.arriveTime), formatDateTime_(r.departTime), r.month, r.arriveDate,
    r.arriveVehicle, r.departVehicle, r.effectiveVehicle, r.arriveTrailer, r.departTrailer, r.effectiveDepartTrailer,
    r.arriveDriver, r.departDriver, r.effectiveDriver, r.location, r.arriveLot, r.departLot,
    r.arriveTask, r.departTask, blankIfNaN_(r.arriveKm), blankIfNaN_(r.departKm), blankIfNaN_(r.tripKm), blankIfNaN_(r.arrivePin), blankIfNaN_(r.departPin), blankIfNaN_(r.pinUsed),
    r.arriveNote, r.departNote, r.isArriveKhac, r.isDepartKhac, r.khacTaskArrive, r.khacTaskDepart,
    r.hasDeparture, blankIfNaN_(r.pinDeltaRaw), blankIfNaN_(r.pinChargeAdded), blankIfNaN_(r.pinUsedAdjusted), r.vehicleChanged, r.trailerChanged
  ]);
  const sheet = getOrCreateSheet_(ss, REPORT_SHEETS.normalized);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (values.length) sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

function writeVehicleReport_(ss, rows) {
  const grouped = {};
  rows.forEach(r => {
    const month = r.month;
    const vehicle = r.effectiveVehicle;
    if (!month || !vehicle) return;
    const key = month + '||' + vehicle;
    if (!grouped[key]) {
      grouped[key] = initTaskBucket_({ month: month, vehicle: vehicle, totalKm: 0, totalPin: 0 });
    }
    addTaskCounts_(grouped[key], r.arriveTask);
    addTaskCounts_(grouped[key], r.departTask);
    grouped[key].totalKm += numberOrZero_(r.tripKm);
    grouped[key].totalPin += numberOrZero_(r.pinUsedAdjusted);
  });

  const data = Object.keys(grouped).sort().map(key => {
    const g = grouped[key];
    return [
      g.month, g.vehicle,
      g.tasks['Lấy rỗng'], g.tasks['Hạ cont đầy'], g.tasks['Sạc pin'], g.tasks['Sửa chữa'],
      g.tasks['Chạy không'], g.tasks['Đổi ca'], g.tasks['Đổi xe'], g.tasks['Về lấy mooc'], g.tasks['Vệ sinh cont'],
      g.totalTasks, round1_(g.totalKm), round1_(g.totalPin)
    ];
  });

  writeSimpleReport_(ss, REPORT_SHEETS.vehicle,
    'BÁO CÁO THEO XE (theo tháng)',
    ['Tháng','Biển số xe','Lấy rỗng','Hạ cont đầy','Sạc pin','Sửa chữa','Chạy không','Đổi ca','Đổi xe','Về lấy mooc','Vệ sinh cont','Tổng NV','Tổng km','Pin tiêu hao (%)'],
    data
  );
}

function writeDriverReport_(ss, rows) {
  const grouped = {};
  rows.forEach(r => {
    const month = r.month;
    const driver = r.effectiveDriver;
    if (!month || !driver) return;
    const key = month + '||' + driver;
    if (!grouped[key]) {
      grouped[key] = initTaskBucket_({ month: month, driver: driver, totalLots: new Set(), totalHours: 0 });
    }
    addTaskCounts_(grouped[key], r.arriveTask);
    addTaskCounts_(grouped[key], r.departTask);
    if (r.arriveLot && r.arriveLot !== 'KHAC') grouped[key].totalLots.add(r.arriveLot);
    if (r.departLot && r.departLot !== 'KHAC') grouped[key].totalLots.add(r.departLot);
    if (r.arriveTime && r.departTime) {
      grouped[key].totalHours += Math.max(0, (r.departTime.getTime() - r.arriveTime.getTime()) / 3600000);
    }
  });

  const data = Object.keys(grouped).sort().map(key => {
    const g = grouped[key];
    return [
      g.month, g.driver,
      g.tasks['Lấy rỗng'], g.tasks['Hạ cont đầy'], g.tasks['Sạc pin'], g.tasks['Sửa chữa'],
      g.tasks['Chạy không'], g.tasks['Đổi ca'], g.tasks['Đổi xe'], g.tasks['Về lấy mooc'], g.tasks['Vệ sinh cont'],
      g.totalTasks, g.totalLots.size, round1_(g.totalHours)
    ];
  });

  writeSimpleReport_(ss, REPORT_SHEETS.driver,
    'BÁO CÁO THEO LÁI XE (theo tháng)',
    ['Tháng','Tài xế','Lấy rỗng','Hạ cont đầy','Sạc pin','Sửa chữa','Chạy không','Đổi ca','Đổi xe','Về lấy mooc','Vệ sinh cont','Tổng NV','Tổng lô','Giờ lái xe (h)'],
    data
  );
}

function writeOtherReport_(ss, rows) {
  const grouped = {};
  rows.forEach(r => {
    const day = r.arriveDate;
    const vehicle = r.effectiveVehicle;
    if (!day || !vehicle) return;
    const tasks = [];
    if (r.arriveLot === 'KHAC') tasks.push.apply(tasks, splitTasks_(r.arriveTask));
    if (r.departLot === 'KHAC') tasks.push.apply(tasks, splitTasks_(r.departTask));
    if (!tasks.length) return;
    const key = r.month + '||' + day + '||' + vehicle;
    if (!grouped[key]) {
      grouped[key] = { month: r.month, day: day, vehicle: vehicle, taskCounts: {}, km: 0, pin: 0, lots: new Set() };
    }
    tasks.forEach(task => {
      grouped[key].taskCounts[task] = (grouped[key].taskCounts[task] || 0) + 1;
    });
    grouped[key].km += numberOrZero_(r.tripKm);
    grouped[key].pin += numberOrZero_(r.pinUsedAdjusted);
    if (r.arriveLot && r.arriveLot !== 'KHAC') grouped[key].lots.add(r.arriveLot);
    if (r.departLot && r.departLot !== 'KHAC') grouped[key].lots.add(r.departLot);
  });

  const data = Object.keys(grouped).sort().map(key => {
    const g = grouped[key];
    const lotCount = g.lots.size || 0;
    return [
      g.month,
      g.day,
      g.vehicle,
      formatTaskCounts_(g.taskCounts),
      round1_(g.km),
      round1_(g.pin),
      lotCount,
      lotCount ? round2_(g.km / lotCount) : '',
      lotCount ? round2_(g.pin / lotCount) : ''
    ];
  });

  writeSimpleReport_(ss, REPORT_SHEETS.other,
    'BÁO CÁO CHI PHÍ VẬN HÀNH (KHAC) THEO NGÀY',
    ['Tháng','Ngày','Xe','Nghiệp vụ KHAC','Km KHAC','Pin KHAC (%)','Số lô trong ngày','Phân bổ/lô (km)','Phân bổ/lô (pin)'],
    data
  );
}

function writeLotReport_(ss, rows) {
  const grouped = {};
  rows.forEach(r => {
    const day = r.arriveDate;
    const month = r.month;
    const vehicle = r.effectiveVehicle;
    if (!day || !month || !vehicle) return;
    const lots = [];
    if (r.arriveLot && r.arriveLot !== 'KHAC') lots.push(r.arriveLot);
    if (r.departLot && r.departLot !== 'KHAC' && lots.indexOf(r.departLot) === -1) lots.push(r.departLot);
    lots.forEach(lot => {
      const key = [month, lot, day, vehicle].join('||');
      if (!grouped[key]) grouped[key] = { month: month, lot: lot, day: day, vehicle: vehicle, km: 0, pin: 0, trips: 0 };
      grouped[key].km += numberOrZero_(r.tripKm);
      grouped[key].pin += numberOrZero_(r.pinUsedAdjusted);
      grouped[key].trips += 1;
    });
  });

  const totals = {};
  Object.keys(grouped).forEach(key => {
    const g = grouped[key];
    const totalKey = [g.month, g.lot, g.day].join('||');
    if (!totals[totalKey]) totals[totalKey] = { km: 0, pin: 0, trips: 0 };
    totals[totalKey].km += g.km;
    totals[totalKey].pin += g.pin;
    totals[totalKey].trips += g.trips;
  });

  const khacAlloc = {};
  rows.forEach(r => {
    if (!r.arriveDate || !r.month || !(r.arriveLot === 'KHAC' || r.departLot === 'KHAC')) return;
    const vehicleLots = [];
    if (r.arriveLot && r.arriveLot !== 'KHAC') vehicleLots.push(r.arriveLot);
    if (r.departLot && r.departLot !== 'KHAC' && vehicleLots.indexOf(r.departLot) === -1) vehicleLots.push(r.departLot);
    if (!vehicleLots.length) return;
    vehicleLots.forEach(lot => {
      const key = [r.month, lot, r.arriveDate, r.effectiveVehicle].join('||');
      if (!khacAlloc[key]) khacAlloc[key] = { km: 0, pin: 0 };
      khacAlloc[key].km += numberOrZero_(r.tripKm);
      khacAlloc[key].pin += numberOrZero_(r.pinUsedAdjusted);
    });
  });

  const data = Object.keys(grouped).sort().map(key => {
    const g = grouped[key];
    const alloc = khacAlloc[key] || { km: 0, pin: 0 };
    const total = totals[[g.month, g.lot, g.day].join('||')] || { km: 0, pin: 0, trips: 0 };
    return [
      g.month, g.lot, g.day, g.vehicle,
      round1_(g.km), round1_(g.pin), g.trips,
      round1_(total.km), round1_(total.pin), total.trips,
      round2_(alloc.km), round2_(alloc.pin),
      round2_(g.km + alloc.km), round2_(g.pin + alloc.pin)
    ];
  });

  writeSimpleReport_(ss, REPORT_SHEETS.lot,
    'BÁO CÁO THEO LÔ HÀNG (chi tiết theo xe + phân bổ KHAC)',
    ['Tháng','Số lô','Ngày lô','Xe','Km (xe)','Pin (xe)','Chuyến (xe)','Tổng km','Tổng pin (%)','Tổng chuyến','Km phân bổ KHAC','Pin phân bổ KHAC','Tổng km + KHAC','Tổng pin + KHAC'],
    data
  );
}

function writeTrailerReport_(ss, rows) {
  const month = getLatestMonth_(rows);
  const monthRows = rows.filter(r => r.month === month);
  const daySet = {};
  const trailerMap = {};

  monthRows.forEach(r => {
    const day = dayOfMonth_(r.arriveDate);
    if (!day) return;
    daySet[day] = true;
    [r.arriveTrailer, r.departTrailer].forEach(trailer => {
      trailer = cleanString_(trailer);
      if (!trailer || /không mooc/i.test(trailer)) return;
      trailerMap[trailer] = trailerMap[trailer] || {};
      trailerMap[trailer][day] = '✔';
    });
  });

  const days = Object.keys(daySet).map(Number).sort((a,b) => a-b);
  const headers = ['Biển số mooc'].concat(days.map(String));
  const data = Object.keys(trailerMap).sort().map(trailer => {
    return [trailer].concat(days.map(day => trailerMap[trailer][day] || ''));
  });

  writeSimpleReport_(ss, REPORT_SHEETS.trailer,
    month ? ('THÁNG ' + month) : 'THÁNG',
    headers,
    data
  );
}

function writeSimpleReport_(ss, sheetName, title, headers, rows) {
  const sheet = getOrCreateSheet_(ss, sheetName);
  sheet.clearContents();
  sheet.getRange(1, 1).setValue(title);
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);
}

function initTaskBucket_(base) {
  const bucket = Object.assign({ tasks: {}, totalTasks: 0 }, base);
  REPORT_TASK_COLUMNS.forEach(name => bucket.tasks[name] = 0);
  return bucket;
}

function addTaskCounts_(bucket, taskString) {
  splitTasks_(taskString).forEach(task => {
    if (bucket.tasks.hasOwnProperty(task)) {
      bucket.tasks[task] += 1;
      bucket.totalTasks += 1;
    }
  });
}

function splitTasks_(taskString) {
  return cleanString_(taskString)
    .split(TASK_SPLIT_REGEX)
    .map(s => cleanString_(s))
    .filter(Boolean);
}

function formatTaskCounts_(taskCounts) {
  return Object.keys(taskCounts).sort().map(name => name + '(' + taskCounts[name] + ')').join(', ');
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function cleanString_(value) {
  return String(value || '').trim();
}

function parseNumberOrBlank_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const n = Number(String(value).replace(/,/g, '.').replace(/[^0-9.\-]/g, ''));
  return Number.isFinite(n) ? n : '';
}

function isFiniteNumber_(value) {
  return typeof value === 'number' && Number.isFinite(value);
}

function blankIfNaN_(value) {
  return isFiniteNumber_(value) ? value : (value || '');
}

function numberOrZero_(value) {
  return isFiniteNumber_(value) ? value : 0;
}

function round1_(value) {
  return Math.round(numberOrZero_(value) * 10) / 10;
}

function round2_(value) {
  return Math.round(numberOrZero_(value) * 100) / 100;
}

function normalizeDateTime_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return value;
  const s = cleanString_(value);
  if (!s) return null;
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));
}

function formatDateTime_(date) {
  return date ? Utilities.formatDate(date, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy H:mm:ss') : '';
}

function dayOfMonth_(dateString) {
  const m = cleanString_(dateString).match(/^(\d{1,2})\//);
  return m ? Number(m[1]) : null;
}

function getLatestMonth_(rows) {
  const months = rows.map(r => r.month).filter(Boolean).sort();
  return months.length ? months[months.length - 1] : '';
}
