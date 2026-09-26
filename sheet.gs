/**
 * ============================================================================
 *  DATABASE PROJECT – GOOGLE APPS SCRIPT BACKEND
 * ============================================================================
 *  File Excel "master.xlsx" adalah TEMPLATE database untuk satu project,
 *  berisi 8 sheet: PO_Tracking, Milestone_Tracking, FAT_Schedule,
 *  Shipment_Tracking, Action_Log, Risk_Register, S_Curve_Data, Master_Data
 *  — persis seperti struktur "Kutai.xlsx" yang sudah berjalan.
 *
 *  Cara memakai untuk PROJECT BARU:
 *  1. Duplikat master.xlsx (mis. jadi "ProjectBaru.xlsx") lalu upload/import
 *     ke Google Sheets (atau File > Save as Google Sheets kalau sudah di Drive).
 *  2. Buka Google Sheet hasil import > Extensions > Apps Script, tempel
 *     SELURUH isi file Code.gs ini (satu Code.gs dipakai untuk semua project,
 *     tidak perlu diubah).
 *  3. Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 *  4. Salin URL /exec, tambahkan sebagai entry baru pada array PROJECTS di
 *     index.html (field apiUrl) untuk project tsb.
 *  Setiap project punya Google Sheet + deployment sendiri, tapi memakai
 *  Code.gs dan struktur kolom yang sama persis — jadi tinggal copy-paste.
 *
 *  ARSITEKTUR: Excel / Google Sheets di sini HANYA berperan sebagai
 *  DATABASE PASIF — sekadar menyimpan & mengembalikan data apa adanya.
 *  Tidak ada formula, tidak ada logika perhitungan di dalam sheet maupun
 *  di Code.gs ini. SEMUA instruksi/logika bisnis (termasuk menghitung
 *  status selesai/belum tiap milestone dan Cummulative Progress PO)
 *  dilakukan di index.html (JavaScript sisi browser) SEBELUM data dikirim
 *  ke sini lewat addRow/updateRow. Code.gs hanya menerjemahkan nama field
 *  camelCase <-> header kolom Excel, lalu menulis apa adanya (CRUD murni).
 * ============================================================================
 */

var SHEET_NAME_PO = 'PO_Tracking';
var CUMULATIVE_PROGRESS_HEADER = 'Cummulative Progress (Actual Progress)';
var PO_ID_HEADER = 'PO Tracking ID';

// Sheet lain dalam database per-project (mengikuti struktur Kutai.xlsx / master.xlsx)
var OTHER_SHEETS = {
  Milestone_Tracking: { keyColumn: 'Milestone_ID' },
  FAT_Schedule: { keyColumn: 'FAT_ID' },
  Shipment_Tracking: { keyColumn: 'Shipment_ID' },
  Action_Log: { keyColumn: 'Action_ID' },
  Risk_Register: { keyColumn: 'Risk_ID' },
  S_Curve_Data: { keyColumn: 'Period' },
  Master_Data: { keyColumn: 'Key' }
};

/**
 * PETA FIELD PO_Tracking: key camelCase (dipakai index.html, sesuai
 * KEY_COLUMNS.PO_Tracking = "poTrackingId" dan id-id pada SHEET_FORMS.PO_Tracking)
 * <-> nama kolom asli di sheet PO_Tracking (header baris 1 di master.xlsx).
 * Frontend HANYA bicara dalam camelCase; semua konversi ke/dari header asli
 * Excel terjadi di sini supaya struktur sheet & tampilan tetap bisa
 * berkembang independen.
 */
var PO_FIELD_MAP = {
  poTrackingId: 'PO Tracking ID',
  purchaseOrderId: 'Purchase Order ID',
  buyer: 'Buyer',
  expeditor: 'Expeditor',
  category: 'Category',
  itemDescription: 'Item Description',
  poDate: 'PO Date (Issued to Vendor)',
  revision: 'Revision',
  discipline: 'Discipline',
  qty: 'Qty',
  uom: 'UoM',
  deliveryIncoterm: 'Delivery (Incoterm)',
  deliveryPlan: 'Delivery (as per PO)',
  deliveryForecast: 'Delivery (Forecast)',
  criticalRating: 'Critical Rating',
  rosDate: 'ROS Date (Latest Baseline)',
  supplier: 'Supplier / Manufacturer',
  countryOfOrigin: 'Country Of Origin',

  kickOffPlan: 'Kick Off Meeting (Plan)',
  kickOffActual: 'Kick Off Meeting (Actual)',
  kickOffWeight: 'Kick Off Meeting (Weight Factor Percentage 12.5%)',

  topPlan: 'Term of Payment (Plan)',
  topActual: 'Term of Payment (Actual)',
  topType: 'Term of Payment (Type)',
  topPercentage: 'Term of Payment (Percentage)',
  topStatus: 'Term of Payment (Status)',

  keyDocSubmissionPlan: 'Key Document Submission (Plan)',
  keyDocSubmissionForecast: 'Key Document Submission (Forecast)',
  keyDocSubmissionActual: 'Key Document Submission (Actual)',

  keyDocApprovalPlan: 'Key Document Approval (Plan)',
  keyDocApprovalForecast: 'Key Document Approval (Forecast)',
  keyDocApprovalActual: 'Key Document Approval (Actual)',
  keyDocApprovalWeight: 'Key Document Approval (Weight Factor Percentage 12.5%)',

  materialOrderedPlan: 'Material Ordered (Plan)',
  materialOrderedForecast: 'Material Ordered (Forecast)',
  materialOrderedActual: 'Material Ordered (Actual)',
  materialOrderedWeight: 'Material Ordered (Weight Factor Percentage 12.5%)',

  materialReceiptVendorPlan: 'Material Receipt by Vendor (Plan)',
  materialReceiptVendorForecast: 'Material Receipt by Vendor (Forecast)',
  materialReceiptVendorActual: 'Material Receipt by Vendor (Actual)',
  materialReceiptVendorWeight: 'Material Receipt by Vendor (Weight Factor Percentage 12.5%)',

  preInspectionPlan: 'Pre-Inspection Meeting (Plan)',
  preInspectionForecast: 'Pre-Inspection Meeting (Forecast)',
  preInspectionActual: 'Pre-Inspection Meeting (Actual)',

  fabricationStartPlan: 'Fabrication Start (Plan)',
  fabricationStartForecast: 'Fabrication Start (Forecast)',
  fabricationStartActual: 'Fabrication Start (Actual)',
  fabricationStartWeight: 'Fabrication Start (Weight Factor Percentage 12.5%)',

  fabricationCompletionPlan: 'Fabrication Completion (Plan)',
  fabricationCompletionForecast: 'Fabrication Completion (Forecast)',
  fabricationCompletionActual: 'Fabrication Completion (Actual)',
  fabricationCompletionWeight: 'Fabrication Completion (Weight Factor Percentage 12.5%)',

  finalInspectionPlan: 'Final Inspection (Plan)',
  finalInspectionForecast: 'Final Inspection (Forecast)',
  finalInspectionActual: 'Final Inspection (Actual)',

  inspectionReleaseNotePlan: 'Inspection Release Note (Plan)',
  inspectionReleaseNoteForecast: 'Inspection Release Note (Forecast)',
  inspectionReleaseNoteActual: 'Inspection Release Note (Actual)',

  packingDispatchPlan: 'Packing and Dispatch (Plan)',
  packingDispatchForecast: 'Packing and Dispatch (Forecast)',
  packingDispatchActual: 'Packing and Dispatch (Actual)',
  packingDispatchWeight: 'Packing and Dispatch (Weight Factor Percentage 12.5%)',

  materialReceivedSitePlan: 'Material Received at Site (Plan)',
  materialReceivedSiteForecast: 'Material Received at Site (Forecast)',
  materialReceivedSiteActual: 'Material Received at Site (Actual)',
  materialReceivedSiteWeight: 'Material Received at Site (Weight Factor Percentage 12.5%)',

  cumulativeProgress: 'Cummulative Progress (Actual Progress)',
  remark: 'Remark / Status',
  areaOfConcern: 'Area of Concern & Mitigation Plan',
  unpricedPO: 'Unpriced PO (PDF)',
  orderStatus: 'Order Status (Open / Close)'
};

/** camelCase field id -> header asli. Fallback: kembalikan apa adanya (identity). */
function poHeaderForKey_(key) {
  return PO_FIELD_MAP[key] || key;
}

/** Ubah objek {camelKey: value} (dari frontend) -> {headerAsli: value} (untuk ditulis ke sheet) */
function poCamelToHeaderData_(data) {
  var out = {};
  if (!data) return out;
  Object.keys(data).forEach(function (k) {
    out[poHeaderForKey_(k)] = data[k];
  });
  return out;
}

// ----------------------------------------------------------------------------
// ENTRY POINTS
// ----------------------------------------------------------------------------

function doGet(e) {
  return handleRequest_(e, 'GET');
}

function doPost(e) {
  return handleRequest_(e, 'POST');
}

function handleRequest_(e, method) {
  var action, payload;
  try {
    if (method === 'GET') {
      action = e && e.parameter ? e.parameter.action : '';
      payload = e && e.parameter ? e.parameter : {};
    } else {
      var body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
      action = body.action;
      payload = body.payload || {};
    }
    var result = routeAction_(action, payload);
    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ success: false, error: err.message });
  }
}

function routeAction_(action, payload) {
  switch (action) {
    case 'getAllPO':
      return { success: true, data: getAllPO_() };
    case 'getAllMilestones':
      return { success: true, data: getAllGeneric_('Milestone_Tracking') };
    case 'getAllFAT':
      return { success: true, data: getAllGeneric_('FAT_Schedule') };
    case 'getAllShipments':
      return { success: true, data: getAllGeneric_('Shipment_Tracking') };
    case 'getAllActions':
      return { success: true, data: getAllGeneric_('Action_Log') };
    case 'getAllSCurve':
      return { success: true, data: getAllGeneric_('S_Curve_Data') };
    case 'getDashboardData':
      return getDashboardData_(payload && payload.upcomingDays);
    case 'getCurrencyRates':
      return { success: true, data: { USD: 1, IDR: 16300, JPY: 157 } };
    case 'diagnosePO':
      return diagnosePOTracking_();
    case 'closeAction':
      return updateRow_('Action_Log', 'Action_ID', payload.actionId || payload.id, {
        Status: 'CLOSED',
        Closed_Date: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM-dd')
      });
    case 'addRow':
      return addRow_(payload.sheet, payload.rowData);
    case 'updateRow':
      return updateRow_(payload.sheet, payload.keyColumn, payload.keyValue, payload.newData);
    case 'deleteRow':
      return deleteRow_(payload.sheet, payload.keyColumn, payload.keyValue);
    case 'uploadFileToDrive':
      return uploadFileToDrive_(payload.base64Data, payload.fileName, payload.mimeType);
    default:
      return { success: false, error: 'Aksi tidak dikenal: ' + action };
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------------------
// SHEET / HEADER HELPERS
// ----------------------------------------------------------------------------

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" tidak ditemukan');
  return sh;
}

function getHeaders_(sheet) {
  var lastCol = sheet.getLastColumn();
  var raw = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return raw.map(function (h) {
    return String(h).replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
  });
}

function formatCell_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM-dd');
  }
  return v;
}

function rowToObject_(headers, rowArr) {
  var obj = {};
  headers.forEach(function (h, i) {
    obj[h] = formatCell_(rowArr[i]);
  });
  return obj;
}

function isRowEmpty_(rowArr) {
  return rowArr.every(function (v) {
    return v === '' || v === null || typeof v === 'undefined';
  });
}

function columnToLetter_(col) {
  var letter = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// ----------------------------------------------------------------------------
// WEIGHT FACTOR / CUMMULATIVE PROGRESS — HANYA UNTUK KEBUTUHAN DIAGNOSTIK &
// MIGRASI DATA LAMA. Perhitungan nilai weight/progress yang SEBENARNYA untuk
// data baru dilakukan di index.html sebelum dikirim ke sini; sheet & Code.gs
// ini tidak lagi menulis formula apa pun.
// ----------------------------------------------------------------------------

/**
 * Menemukan semua kolom header yang mengandung "Weight Factor Percentage".
 * Dipakai HANYA oleh diagnosePOTracking_() dan rebuildLegacyPOData_() untuk
 * membaca struktur sheet — tidak dipakai untuk menulis formula.
 */
function getWeightFactorColumns_(headers) {
  var cols = [];
  headers.forEach(function (h, idx) {
    var m = h.match(/Weight Factor Percentage\s*(\d+(?:\.\d+)?)\s*%/i);
    if (m) {
      cols.push({
        index: idx, // 0-based
        header: h,
        weight: parseFloat(m[1]) / 100,
        milestone: h.split('(')[0].trim()
      });
    }
  });
  return cols;
}

function getCumulativeProgressColIndex_(headers) {
  var idx = headers.indexOf(CUMULATIVE_PROGRESS_HEADER);
  if (idx === -1) {
    idx = headers.findIndex(function (h) {
      return /Cummulative Progress/i.test(h);
    });
  }
  return idx; // 0-based, -1 jika tidak ditemukan
}

/**
 * DIAGNOSTIK: jalankan via browser dengan membuka
 *   <apiUrl>?action=diagnosePO
 * atau lewat editor Apps Script (pilih fungsi ini > Run > View > Logs).
 * Berguna untuk mengetahui persis struktur sheet PO_Tracking: apakah
 * ke-8 kolom Weight Factor terdeteksi, dan apa isi nilai aktual di baris contoh.
 */
function diagnosePOTracking_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet;
  try {
    sheet = getSheet_(SHEET_NAME_PO);
  } catch (e) {
    return {
      success: false,
      error: 'Sheet "' + SHEET_NAME_PO + '" tidak ditemukan di spreadsheet ini.',
      availableSheetNames: ss.getSheets().map(function (s) { return s.getName(); })
    };
  }

  var headers = getHeaders_(sheet);
  var weightCols = getWeightFactorColumns_(headers);
  var cumIdx = getCumulativeProgressColIndex_(headers);
  var lastRow = sheet.getLastRow();

  var sampleRow = null;
  if (lastRow >= 2 && cumIdx !== -1) {
    var cumLetter = columnToLetter_(cumIdx + 1);
    sampleRow = {
      row: 2,
      cumulativeProgressColumn: cumLetter,
      value: sheet.getRange(2, cumIdx + 1).getValue(),
      weightCellValues: weightCols.map(function (c) {
        return {
          milestone: c.milestone,
          column: columnToLetter_(c.index + 1),
          value: sheet.getRange(2, c.index + 1).getValue()
        };
      })
    };
  }

  return {
    success: true,
    spreadsheetName: ss.getName(),
    sheetName: sheet.getName(),
    totalColumns: headers.length,
    headers: headers,
    weightFactorColumnsFound: weightCols.length,
    weightFactorColumns: weightCols.map(function (c) {
      return { header: c.header, milestone: c.milestone, column: columnToLetter_(c.index + 1), weight: c.weight };
    }),
    cumulativeProgressHeaderFound: cumIdx !== -1,
    cumulativeProgressColumn: cumIdx !== -1 ? columnToLetter_(cumIdx + 1) : null,
    lastRow: lastRow,
    sampleRow2: sampleRow,
    note: weightCols.length === 0
      ? 'Tidak ada kolom header yang cocok dengan pola "Weight Factor Percentage". Cek daftar "headers" di atas.'
      : (cumIdx === -1
        ? 'Header "Cummulative Progress (Actual Progress)" tidak ditemukan di sheet.'
        : 'Struktur sheet terbaca normal. Nilai weight/progress sekarang murni data statis yang dikirim index.html, bukan formula.')
  };
}

/**
 * UTILITY MIGRASI (manual, dijalankan sekali dari editor Apps Script):
 * memperbaiki baris data LAMA yang kolom Weight/Cummulative Progress-nya
 * masih berupa formula rusak atau tidak sinkron peninggalan versi sebelumnya.
 * Menghitung ulang berdasarkan kolom "(Actual)" tiap milestone, lalu menulis
 * NILAI STATIS (bukan formula) — meniru logika yang sama dengan
 * computePOWeightsAndProgress() di index.html.
 * Setelah dijalankan, buka View > Logs untuk lihat ringkasannya.
 */
function rebuildLegacyPOData_() {
  var sheet = getSheet_(SHEET_NAME_PO);
  var headers = getHeaders_(sheet);
  var weightCols = getWeightFactorColumns_(headers);
  var cumIdx = getCumulativeProgressColIndex_(headers);
  var lastRow = sheet.getLastRow();

  if (weightCols.length === 0) {
    Logger.log('BERHENTI: tidak ada kolom "Weight Factor Percentage" yang terdeteksi di header.');
    return;
  }
  if (cumIdx === -1) {
    Logger.log('BERHENTI: header "Cummulative Progress (Actual Progress)" tidak ditemukan.');
    return;
  }
  if (lastRow < 2) {
    Logger.log('Tidak ada baris data. Tidak ada yang diproses.');
    return;
  }

  var range = sheet.getRange(2, 1, lastRow - 1, headers.length);
  var values = range.getValues();
  var processed = 0;

  values.forEach(function (rowArr) {
    if (isRowEmpty_(rowArr)) return;
    var completed = 0;
    weightCols.forEach(function (c) {
      var actualIdx = headers.indexOf(c.milestone + ' (Actual)');
      var done = actualIdx !== -1 && rowArr[actualIdx] !== '' && rowArr[actualIdx] !== null && typeof rowArr[actualIdx] !== 'undefined';
      rowArr[c.index] = done ? 1 : 0;
      if (done) completed++;
    });
    rowArr[cumIdx] = completed / weightCols.length;
    processed++;
  });

  range.setValues(values);
  Logger.log('SELESAI: ' + processed + ' baris ditulis ulang dengan nilai statis (bukan formula).');
}

// ----------------------------------------------------------------------------
// AUTO-SYNC SAAT EDIT LANGSUNG DI GOOGLE SHEETS (bukan lewat web app)
// ----------------------------------------------------------------------------
// index.html menghitung Weight Factor (0/1) & Cummulative Progress sendiri
// SEBELUM mengirim data lewat addRow/updateRow. Tapi kalau user mengedit
// langsung di Google Sheets (bukan lewat aplikasi web), tidak ada JS
// index.html yang berjalan untuk menghitung ulang. onEdit(e) di bawah ini
// adalah SIMPLE TRIGGER bawaan Apps Script (otomatis aktif tanpa perlu
// instalasi manual) yang mendeteksi kalau kolom "(Actual)" salah satu dari
// 8 milestone diubah langsung di sheet PO_Tracking, lalu menulis ULANG
// nilai statis (BUKAN formula) ke kolom Weight & Cummulative Progress baris
// tsb - meniru persis logika computePOWeightsAndProgress() di index.html.
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAME_PO) return;

    var headers = getHeaders_(sheet);
    var weightCols = getWeightFactorColumns_(headers);
    var cumIdx = getCumulativeProgressColIndex_(headers);
    if (!weightCols.length || cumIdx === -1) return;

    // kolom "(Actual)" -> weightCols entry yang terkait
    var relevantCols = {}; // 0-based col index -> true, kalau kolom ini relevan (Actual atau Weight)
    weightCols.forEach(function (c) {
      var actualIdx = headers.indexOf(c.milestone + ' (Actual)');
      if (actualIdx !== -1) relevantCols[actualIdx] = true;
      relevantCols[c.index] = true; // jaga-jaga kalau weight diedit manual juga
    });

    var startRow = e.range.getRow();
    var numRows = e.range.getNumRows();
    var startCol = e.range.getColumn(); // 1-based
    var numCols = e.range.getNumColumns();

    var touchedRelevant = false;
    for (var c = startCol; c < startCol + numCols; c++) {
      if (relevantCols[c - 1]) { touchedRelevant = true; break; }
    }
    if (!touchedRelevant) return;

    for (var r = startRow; r < startRow + numRows; r++) {
      if (r === 1) continue; // baris header
      recomputePORow_(sheet, headers, weightCols, cumIdx, r);
    }
  } catch (err) {
    // Simple trigger: kegagalan di sini tidak boleh mengganggu proses edit user di UI,
    // jadi cukup diabaikan (tidak ada cara menampilkan error ke user dari simple trigger).
  }
}

/** Hitung ulang & tulis nilai statis Weight (0/1) + Cummulative Progress untuk satu baris */
function recomputePORow_(sheet, headers, weightCols, cumIdx, row) {
  var completed = 0;
  weightCols.forEach(function (c) {
    var actualIdx = headers.indexOf(c.milestone + ' (Actual)');
    var actualVal = actualIdx !== -1 ? sheet.getRange(row, actualIdx + 1).getValue() : '';
    var done = actualVal !== '' && actualVal !== null && typeof actualVal !== 'undefined';
    sheet.getRange(row, c.index + 1).setValue(done ? 1 : 0);
    if (done) completed++;
  });
  sheet.getRange(row, cumIdx + 1).setValue(completed / weightCols.length);
  sheet.getRange(row, cumIdx + 1).setNumberFormat('0%');
}

// ----------------------------------------------------------------------------
// PO TRACKING ID GENERATOR
// ----------------------------------------------------------------------------

function generateNextPOId_(sheet, headers, idIdx) {
  var lastRow = sheet.getLastRow();
  var max = 0;
  if (lastRow >= 2) {
    var ids = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues();
    ids.forEach(function (r) {
      var m = String(r[0]).match(/(\d+)\s*$/);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  var next = max + 1;
  var padded = next < 1000 ? ('000' + next).slice(-3) : String(next);
  return 'PID' + padded;
}

// ----------------------------------------------------------------------------
// GET ALL PO
// ----------------------------------------------------------------------------

function getAllPO_() {
  var sheet = getSheet_(SHEET_NAME_PO);
  var headers = getHeaders_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  var headerIdx = {};
  headers.forEach(function (h, i) { headerIdx[h] = i; });

  var data = [];
  values.forEach(function (row) {
    if (isRowEmpty_(row)) return;
    var obj = {};
    Object.keys(PO_FIELD_MAP).forEach(function (camelKey) {
      var idx = headerIdx[PO_FIELD_MAP[camelKey]];
      obj[camelKey] = idx !== undefined ? formatCell_(row[idx]) : '';
    });
    // cumulativeProgress dikembalikan sebagai pecahan 0..1 (mis. 0.55 = 55%),
    // sesuai yang diharapkan index.html (Math.round(100 * cumulativeProgress)).
    obj.cumulativeProgress = Number(obj.cumulativeProgress) || 0;

    // --- Alias field gaya-lama, untuk kompatibilitas halaman "Reports" (export
    // PDF/Excel) di index.html yang masih memakai nama field skema PO lama.
    // Tidak memengaruhi field camelCase di atas; ini murni tambahan.
    // Delay_Days = hari ini - Delivery Date PO, TAPI berhenti bertambah begitu
    // Order Status jadi Close - dibekukan berdasarkan tanggal actual PO tsb
    // selesai (Material Received at Site), bukan terus jalan mengikuti hari ini.
    var planDate = parseDateSafe_(obj.deliveryPlan);
    var today0 = getTodayLocal_();
    var legacyDelayDays;
    if (obj.orderStatus === 'Close') {
      var actualDate = parseDateSafe_(obj.materialReceivedSiteActual);
      legacyDelayDays = (planDate && actualDate) ? Math.max(0, diffDays_(actualDate, planDate)) : 0;
    } else {
      legacyDelayDays = planDate ? Math.max(0, diffDays_(today0, planDate)) : 0;
    }
    var legacyStatus = obj.category === 'On Schedule' ? 'ON TIME' : obj.category === 'At Risk' ? 'AT RISK' : obj.category === 'Delay' ? 'DELAY' : '';
    obj.PO_ID = obj.poTrackingId;
    obj.Package = obj.discipline;
    obj.Item_Description = obj.itemDescription;
    obj.Vendor = obj.supplier;
    obj.PO_Date = obj.poDate;
    obj.RDD = obj.deliveryPlan;
    obj.FDD = obj.deliveryForecast;
    obj.Actual_Delivery = obj.materialReceivedSiteActual;
    obj.Status = legacyStatus;
    obj.Delay_Days = legacyDelayDays;
    obj.Notes = obj.remark;

    data.push(obj);
  });
  return data;
}

/** Baca seluruh baris dari sheet lain (Milestone_Tracking, FAT_Schedule, dst) apa adanya */
function getAllGeneric_(sheetName) {
  var sheet;
  try {
    sheet = getSheet_(sheetName);
  } catch (e) {
    return [];
  }
  var headers = getHeaders_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var data = [];
  values.forEach(function (row) {
    if (isRowEmpty_(row)) return;
    data.push(rowToObject_(headers, row));
  });
  return data;
}

// ----------------------------------------------------------------------------
// DASHBOARD
// Bentuk objek balikan ini WAJIB mengikuti apa yang dibaca renderAll() di
// index.html: { success, kpi:{...}, progress:{...}, sCurve:[...],
// criticalItems:[...], risks:[...], warnings:[...], fatSchedule:[...],
// actions:[...], shipments:[...] }
// ----------------------------------------------------------------------------

function parseDateSafe_(v) {
  if (!v && v !== 0) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return v;
  var d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * "Hari ini" yang benar berdasarkan timezone project (Session.getScriptTimeZone(),
 * biasanya timezone spreadsheet, mis. Asia/Jakarta), BUKAN dari `new Date()` polos.
 *
 * AKAR MASALAH yang diperbaiki fungsi ini: runtime V8 di Apps Script SELALU
 * berjalan di UTC untuk semua method Date bawaan (getHours, setHours, dst) -
 * tidak peduli timezone project di-set apa. Kalau `new Date()` dipakai langsung
 * saat waktu Jakarta sudah lewat tengah malam tapi UTC masih hari sebelumnya
 * (Jakarta 00:00-06:59 = UTC masih hari kemarin, karena Jakarta = UTC+7), maka
 * "hari ini" versi script jadi mundur 1 hari dari kalender Jakarta yang
 * sebenarnya - inilah penyebab selisih 1 hari pada kolom Delay.
 * Fungsi ini mengambil tanggal kalender yang benar via Utilities.formatDate
 * (yang MEMANG timezone-aware), lalu membungkusnya jadi Date UTC-tengah-malam
 * yang konsisten dipakai oleh diffDays_().
 */
function getTodayLocal_() {
  var tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
  var parts = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd').split('-');
  return new Date(Date.UTC(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10)));
}

/**
 * Selisih hari kalender antara dua tanggal (a - b), dibulatkan berbasis
 * TANGGAL saja (tahun/bulan/tanggal), BUKAN dari selisih milidetik mentah.
 * Ini sengaja menghindari getTime()/24*60*60*1000 langsung, karena kalau
 * salah satu Date punya jam yang tidak persis tengah malam yang sama
 * (mis. akibat perbedaan timezone server Apps Script vs data di sheet),
 * hasilnya bisa meleset 1 hari akibat pembulatan. Dengan menormalkan kedua
 * tanggal ke UTC tengah malam berdasarkan komponen kalendernya, hasilnya
 * selalu bilangan hari yang bulat & akurat, apa pun timezone-nya.
 */
function diffDays_(a, b) {
  if (!a || !b) return 0;
  var MS = 24 * 60 * 60 * 1000;
  var utcA = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
  var utcB = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
  return Math.round((utcA - utcB) / MS);
}

/** Baca seluruh baris mentah sebuah sheet sebagai array of {header: value}, dengan tanggal apa adanya (Date object, bukan string) */
function getRawRows_(sheetName) {
  var sheet;
  try {
    sheet = getSheet_(sheetName);
  } catch (e) {
    return [];
  }
  var headers = getHeaders_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var out = [];
  values.forEach(function (row) {
    if (isRowEmpty_(row)) return;
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    out.push(obj);
  });
  return out;
}

function getDashboardData_(upcomingDays) {
  var windowDays = parseInt(upcomingDays, 10);
  if (!windowDays || windowDays <= 0) windowDays = 30;
  var today = getTodayLocal_();
  var windowEnd = new Date(today.getTime() + windowDays * 24 * 60 * 60 * 1000);

  var poRows = getRawRows_(SHEET_NAME_PO);
  var fatRows = getRawRows_('FAT_Schedule');
  var actionRows = getRawRows_('Action_Log');
  var riskRows = getRawRows_('Risk_Register');
  var shipmentRows = getRawRows_('Shipment_Tracking');
  var sCurveRows = getRawRows_('S_Curve_Data');

  // ---------- PO summary ----------
  var onSchedule = 0, atRisk = 0, delay = 0, overdueActive = 0, progressSum = 0;
  var criticalItems = [];
  // Kelompokkan per Purchase Order ID unik untuk KPI Total PO (Plan/Actual) di
  // dashboard: satu Purchase Order ID bisa punya beberapa baris/items. Selama
  // MASIH ADA item dengan Order Status = Open di bawah Purchase Order ID yang
  // sama, PO tsb dihitung sebagai "Plan". Hanya dihitung "Actual" kalau SEMUA
  // item di bawah Purchase Order ID itu sudah Close.
  var poIdGroups = {}; // purchaseOrderId -> { anyOpen: bool }

  poRows.forEach(function (r) {
    var status = r['Order Status (Open / Close)'];
    var poNo = r['Purchase Order ID'] || r['PO Tracking ID'] || '';
    if (poNo) {
      if (!poIdGroups[poNo]) poIdGroups[poNo] = { anyOpen: false };
      if (status === 'Open') poIdGroups[poNo].anyOpen = true;
    }

    var cat = r['Category'];
    if (cat === 'On Schedule') onSchedule++;
    else if (cat === 'At Risk') atRisk++;
    else if (cat === 'Delay') delay++;

    progressSum += Number(r['Cummulative Progress (Actual Progress)']) || 0;

    if (cat === 'Delay' && status === 'Open') overdueActive++;

    // Critical Items HANYA untuk PO yang masih Open. Begitu Order Status jadi
    // Close, item tsb dianggap selesai/tidak lagi "kritis" - delay berhenti
    // dihitung/ditampilkan di sini (tidak terus bertambah mengikuti hari ini).
    if ((cat === 'At Risk' || cat === 'Delay') && status === 'Open') {
      var plan = parseDateSafe_(r['Delivery (as per PO)']);
      // Delay_Days = hari ini - Delivery Date PO (bukan lagi forecast - plan)
      var delayDays = plan ? Math.max(0, diffDays_(today, plan)) : 0;
      criticalItems.push({
        Item_Description: r['Item Description'] || '',
        Vendor: r['Supplier / Manufacturer'] || '',
        RDD: formatCell_(r['Delivery (as per PO)']),
        Delay_Days: delayDays,
        Status: cat // "At Risk" / "Delay" -> cocok dengan statusBadge() & filter toUpperCase() di index.html
      });
    }
  });
  criticalItems.sort(function (a, b) { return (b.Delay_Days || 0) - (a.Delay_Days || 0); });
  criticalItems = criticalItems.slice(0, 25);
  var avgProgress = poRows.length ? progressSum / poRows.length : 0;

  // ---------- FAT ----------
  var fatUpcoming = 0, fatPlanCount = 0, fatActualCount = 0;
  var fatSchedule = [];
  fatRows.forEach(function (r) {
    var d = parseDateSafe_(r['FAT_Date']);
    var inWindow = d && d.getTime() >= today.getTime() && d.getTime() <= windowEnd.getTime();
    var st = r['Status'];
    if (inWindow) {
      if (st !== 'Complete') fatUpcoming++;
      if (st === 'Tentative' || st === 'Scheduled') fatPlanCount++;
      else if (st === 'Confirmed' || st === 'Complete') fatActualCount++;
      fatSchedule.push({
        Item: r['Item'] || '',
        Vendor: r['Vendor'] || '',
        FAT_Date: formatCell_(r['FAT_Date']),
        Location: r['Location'] || '',
        Status: st || ''
      });
    }
  });
  fatSchedule.sort(function (a, b) { return String(a.FAT_Date).localeCompare(String(b.FAT_Date)); });
  fatSchedule = fatSchedule.slice(0, 15);

  // ---------- Actions ----------
  var openActions = 0, closedActions = 0;
  var openActionRows = [];
  actionRows.forEach(function (r) {
    var st = r['Status'];
    if (st === 'CLOSED') closedActions++;
    else {
      openActions++;
      openActionRows.push({
        Action_ID: r['Action_ID'] || '',
        Date: formatCell_(r['Date']),
        Item: r['Item'] || '',
        Issue_Action: r['Issue_Action'] || '',
        PIC_Vendor: r['PIC_Vendor'] || '',
        Due_Date: formatCell_(r['Due_Date']),
        Status: st || ''
      });
    }
  });
  openActionRows.sort(function (a, b) { return String(a.Due_Date).localeCompare(String(b.Due_Date)); });
  openActionRows = openActionRows.slice(0, 15);

  // ---------- Shipments (aktif / belum arrived) ----------
  var shipments = shipmentRows
    .filter(function (r) { return r['Status'] !== 'Arrived'; })
    .slice(0, 15)
    .map(function (r) {
      return {
        Item: r['Item'] || '',
        Vessel_Mode: r['Vessel_Mode'] || '',
        ETD: formatCell_(r['ETD']),
        ETA: formatCell_(r['ETA']),
        POL: r['POL'] || '',
        POD: r['POD'] || '',
        Current_Location: r['Current_Location'] || '',
        Status: r['Status'] || ''
      };
    });

  // ---------- Risks (dipakai widget Risk di dashboard) ----------
  var risks = riskRows.slice(0, 15).map(function (r) {
    return {
      Vendor: r['Vendor'] || '',
      Schedule_Risk: r['Schedule_Risk'] || '',
      Quality_Risk: r['Quality_Risk'] || '',
      Overall_Risk: r['Overall_Risk'] || ''
    };
  });

  // ---------- S-Curve ----------
  // Procurement Overview murni input manual dari sheet S_Curve_Data - TIDAK ada
  // fallback otomatis dari rata-rata Cummulative Progress PO_Tracking.
  var sCurve = sCurveRows.map(function (r) {
    return {
      Month: formatCell_(r['Month']),
      Planned_Pct: r['Planned_Pct'] === '' || r['Planned_Pct'] === null || typeof r['Planned_Pct'] === 'undefined' ? '' : Number(r['Planned_Pct']),
      Actual_Pct: r['Actual_Pct'] === '' || r['Actual_Pct'] === null || typeof r['Actual_Pct'] === 'undefined' ? '' : Number(r['Actual_Pct'])
    };
  });
  var lastActualRow = null;
  for (var i = sCurve.length - 1; i >= 0; i--) {
    if (sCurve[i].Actual_Pct !== '' && sCurve[i].Actual_Pct !== null) { lastActualRow = sCurve[i]; break; }
  }
  var progressActual = lastActualRow ? Number(lastActualRow.Actual_Pct) : 0;
  var progressPlanned = lastActualRow ? Number(lastActualRow.Planned_Pct) : 0;
  var progressVariance = Math.round((progressActual - progressPlanned) * 10) / 10;

  // ---------- Warnings (disintesis dari PO Delay & At Risk, serta Action lewat due date) ----------
  var warnings = [];
  poRows.forEach(function (r) {
    if (r['Category'] === 'Delay') {
      warnings.push({
        type: 'danger',
        message: 'PO ' + (r['PO Tracking ID'] || '') + ' (' + (r['Item Description'] || '') + ') mengalami keterlambatan',
        date: formatCell_(r['Delivery (Forecast)']) || ''
      });
    } else if (r['Category'] === 'At Risk') {
      warnings.push({
        type: 'warning',
        message: 'PO ' + (r['PO Tracking ID'] || '') + ' (' + (r['Item Description'] || '') + ') berisiko terlambat',
        date: formatCell_(r['Delivery (Forecast)']) || ''
      });
    }
  });
  actionRows.forEach(function (r) {
    var due = parseDateSafe_(r['Due_Date']);
    if (due && due.getTime() < today.getTime() && r['Status'] !== 'CLOSED') {
      warnings.push({
        type: 'warning',
        message: 'Action ' + (r['Action_ID'] || '') + ': ' + (r['Issue_Action'] || '') + ' telah lewat due date',
        date: formatCell_(r['Due_Date']) || ''
      });
    }
  });
  warnings = warnings.slice(0, 15);

  // Total PO (Plan/Actual) dashboard = dari Purchase Order ID unik, bukan jumlah baris/items.
  var poUniquePlan = 0, poUniqueActual = 0;
  Object.keys(poIdGroups).forEach(function (poNo) {
    if (poIdGroups[poNo].anyOpen) poUniquePlan++;
    else poUniqueActual++;
  });

  return {
    success: true,
    kpi: {
      poPlan: poUniquePlan,
      poActual: poUniqueActual,
      criticalPlan: atRisk,
      criticalActual: delay,
      criticalItems: atRisk + delay,
      overdueItems: overdueActive,
      openActions: openActions,
      closedActions: closedActions,
      upcomingFAT: fatUpcoming,
      fatPlan: fatPlanCount,
      fatActual: fatActualCount
    },
    progress: {
      actual: progressActual,
      variance: progressVariance
    },
    sCurve: sCurve,
    criticalItems: criticalItems,
    risks: risks,
    warnings: warnings,
    fatSchedule: fatSchedule,
    actions: openActionRows,
    shipments: shipments
  };
}

// ----------------------------------------------------------------------------
// CRUD GENERIK (dipakai oleh menu PO Tracking; sheet lain bisa memakai
// helper yang sama selama nama & header sheet-nya sudah dibuat)
// ----------------------------------------------------------------------------

function addRow_(sheetName, rowData) {
  if (!sheetName) return { success: false, error: 'Sheet tidak ditentukan' };
  var sheet = getSheet_(sheetName);
  var headers = getHeaders_(sheet);
  var isPO = sheetName === SHEET_NAME_PO;
  var dataByHeader = isPO ? poCamelToHeaderData_(rowData) : (rowData || {});

  var rowArr = headers.map(function (h) {
    return dataByHeader && Object.prototype.hasOwnProperty.call(dataByHeader, h) ? dataByHeader[h] : '';
  });

  if (isPO) {
    var idIdx = headers.indexOf(PO_ID_HEADER);
    if (idIdx !== -1 && !rowArr[idIdx]) {
      rowArr[idIdx] = generateNextPOId_(sheet, headers, idIdx);
    }
  }

  sheet.appendRow(rowArr);

  return { success: true, message: 'Data berhasil ditambahkan' };
}

function updateRow_(sheetName, keyColumn, keyValue, newData) {
  if (!sheetName || !keyColumn) return { success: false, error: 'Parameter tidak lengkap' };
  var sheet = getSheet_(sheetName);
  var headers = getHeaders_(sheet);
  var isPO = sheetName === SHEET_NAME_PO;
  var actualKeyColumn = isPO ? poHeaderForKey_(keyColumn) : keyColumn;
  var dataByHeader = isPO ? poCamelToHeaderData_(newData) : (newData || {});

  var keyIdx = headers.indexOf(actualKeyColumn);
  if (keyIdx === -1) return { success: false, error: 'Kolom kunci "' + actualKeyColumn + '" tidak ditemukan' };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, error: 'Data tidak ditemukan' };
  var range = sheet.getRange(2, 1, lastRow - 1, headers.length);
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    if (String(values[i][keyIdx]) === String(keyValue)) {
      var rowArr = values[i];
      headers.forEach(function (h, idx) {
        if (dataByHeader && Object.prototype.hasOwnProperty.call(dataByHeader, h)) {
          rowArr[idx] = dataByHeader[h];
        }
      });

      var sheetRow = i + 2;
      sheet.getRange(sheetRow, 1, 1, headers.length).setValues([rowArr]);

      return { success: true, message: 'Data berhasil diupdate' };
    }
  }
  return { success: false, error: 'Data dengan ' + actualKeyColumn + ' = ' + keyValue + ' tidak ditemukan' };
}

function deleteRow_(sheetName, keyColumn, keyValue) {
  if (!sheetName || !keyColumn) return { success: false, error: 'Parameter tidak lengkap' };
  var sheet = getSheet_(sheetName);
  var headers = getHeaders_(sheet);
  var actualKeyColumn = sheetName === SHEET_NAME_PO ? poHeaderForKey_(keyColumn) : keyColumn;
  var keyIdx = headers.indexOf(actualKeyColumn);
  if (keyIdx === -1) return { success: false, error: 'Kolom kunci "' + actualKeyColumn + '" tidak ditemukan' };

  var lastRow = sheet.getLastRow();
  for (var r = 2; r <= lastRow; r++) {
    var val = sheet.getRange(r, keyIdx + 1).getValue();
    if (String(val) === String(keyValue)) {
      sheet.deleteRow(r);
      return { success: true, message: 'Data berhasil dihapus' };
    }
  }
  return { success: false, error: 'Data dengan ' + actualKeyColumn + ' = ' + keyValue + ' tidak ditemukan' };
}

// ----------------------------------------------------------------------------
// UPLOAD FILE (Unpriced PO PDF, dsb) KE GOOGLE DRIVE
// ----------------------------------------------------------------------------

function uploadFileToDrive_(base64Data, fileName, mimeType) {
  try {
    if (!base64Data || !fileName) return { success: false, error: 'File tidak lengkap' };
    var folder = getUploadFolder_();
    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: file.getUrl(), id: file.getId() };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getUploadFolder_() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('PO_UPLOAD_FOLDER_ID');
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      // folder id tersimpan tapi sudah tidak valid, buat ulang di bawah
    }
  }
  var folders = DriveApp.getFoldersByName('PO_Tracking_Uploads');
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('PO_Tracking_Uploads');
  props.setProperty('PO_UPLOAD_FOLDER_ID', folder.getId());
  return folder;
}
