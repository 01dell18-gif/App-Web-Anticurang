/**
 * Sistem Portal Ujian Google Form Anticurang & Dashboard Monitoring
 * Mengadopsi Arsitektur MPA Query Parameter Routing, Scriptlet Inclusion,
 * RBAC, SheetJS Excel Processing, & Dynamic Adaptive Print Engine.
 *
 * Spreadsheet Database: 1xZTTYRaCQQqjSI40XX9lFpz7lRIDxoAkpgE5TW72rZA
 */

const SPREADSHEET_ID = '1xZTTYRaCQQqjSI40XX9lFpz7lRIDxoAkpgE5TW72rZA';
const FIXED_DEPLOY_URL = 'https://script.google.com/macros/s/AKfycbxw20eF7UAvQv4zJJlMds32mUjU42KInSNiLgGKuZz9_sF76sOH4xhQEcvHkg9V-rO9/exec';

const SHEET_DB = 'DB_Anticurang';
const SHEET_MAPEL = 'DB_Mapel';
const SHEET_MURID = 'DB_Murid';
const SHEET_SETTINGS = 'Pengaturan';
const SHEET_MONITORING = 'Log_Monitoring';

const DB_HEADERS = [
  'No',
  'Mapel',
  'Kelas',
  'Link GFrom',
  'Tanggal',
  'Waktu Mulai',
  'Waktu Selesai',
  'Durasi (Menit)'
];
const MAPEL_HEADERS = [
  'No',
  'Nama Mapel'
];
const MURID_HEADERS = [
  'No',
  'Nama Murid',
  'NIS/NISN',
  'Kelas'
];
const SETTINGS_HEADERS = [
  'Nama Sekolah',
  'Logo',
  'Kop',
  'Folder ID Drive',
  'Sumber Data',
  'PIN Pengawas',
  'PIN Admin'
];
const MONITORING_HEADERS = [
  'ID Sesi',
  'Waktu Akses',
  'Nama Siswa',
  'NIS / Peserta',
  'Kelas Siswa',
  'No Soal',
  'Mapel',
  'Kelas Jadwal',
  'Tanggal Ujian',
  'Status',
  'Jumlah Pelanggaran',
  'Catatan Pelanggaran',
  'Waktu Selesai'
];

const DEFAULT_SETTINGS = {
  namaSekolah: 'PORTAL UJIAN ANTICURANG',
  logo: 'https://drive.google.com/file/d/1_4m0lkw99HvxaPXvOhy0UOudoPmiIRc-/view?usp=drive_link',
  kop: 'https://drive.google.com/file/d/1QXtnYwIrCc-pihDUoidjkRW5UVWykOn-/view?usp=drive_link',
  folderIdDrive: '',
  sumberData: 'Database Ujian & Pengawasan',
  pinPengawas: '654321',
  pinAdmin: '123456'
};

const DEFAULT_ADMIN_PIN = '123456';
const DEFAULT_PROCTOR_PIN = '654321';
const DEFAULT_MAX_VIOLATIONS = 3;
const ADMIN_TOKEN_TTL_SEC = 3600;
const PUBLIC_CACHE_KEY = 'portal_public_cache_v5';
const PUBLIC_CACHE_TTL_SEC = 15;

const ROUTE_MAP = {
  'portal': 'Portal',
  'jadwal': 'Portal',
  'admin': 'Monitoring',
  'pengawas': 'Monitoring',
  'monitoring': 'Monitoring',
  'laporan': 'Laporan',
  'kelola-soal': 'KelolaSoal',
  'data-mapel': 'DataMapel',
  'data-master': 'DataMapel',
  'mapel': 'DataMapel',
  'data-murid': 'DataMurid',
  'murid': 'DataMurid',
  'pengaturan': 'Pengaturan'
};

var _activePage = 'portal';
var _activeParams = {};

/**
 * Router & Controller utama berbasis Query Parameter (?page=...)
 */
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const requested = String(params.page || 'portal').trim().toLowerCase();
  const templateName = ROUTE_MAP[requested] || 'Portal';
  if (requested === 'jadwal') {
    _activePage = 'portal';
  } else if (requested === 'admin' || requested === 'pengawas') {
    _activePage = 'monitoring';
  } else if (requested === 'data-master') {
    _activePage = 'data-mapel';
  } else {
    _activePage = ROUTE_MAP[requested] ? requested : 'portal';
  }
  _activeParams = params;

  const settings = getSettings_();
  const t = HtmlService.createTemplateFromFile(templateName);
  t.currentPage = _activePage;
  t.params = _activeParams;
  t.appUrl = getAppUrl();
  t.initialAppTitle = settings.namaSekolah || DEFAULT_SETTINGS.namaSekolah;

  return t.evaluate()
    .setTitle((settings.namaSekolah || DEFAULT_SETTINGS.namaSekolah) + ' - Sistem Ujian Anticurang')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Mesin Kompilator Server-Side Template Inclusion (Scriptlet Pattern)
 */
function include(filename, data) {
  try {
    const t = HtmlService.createTemplateFromFile(filename);
    t.currentPage = _activePage;
    t.params = _activeParams;
    t.appUrl = getAppUrl();
    if (data && typeof data === 'object') {
      for (var k in data) {
        if (Object.prototype.hasOwnProperty.call(data, k)) {
          t[k] = data[k];
        }
      }
    }
    return t.evaluate().getContent();
  } catch (err) {
    return '<!-- Gagal memuat komponen "' + filename + '": ' + err.message + ' -->';
  }
}

/**
 * Mengembalikan URL Web App aktif:
 * 1. Override opsional CUSTOM_DEPLOY_URL (jika diset Admin di Pengaturan Sistem)
 * 2. ScriptApp.getService().getUrl() (otomatis mendeteksi URL deploy terbaru)
 * 3. FIXED_DEPLOY_URL (cadangan fallback)
 */
function getAppUrl() {
  try {
    const customUrl = String(PropertiesService.getScriptProperties().getProperty('CUSTOM_DEPLOY_URL') || '').trim();
    if (customUrl) return customUrl;
  } catch (e) {}
  try {
    const svcUrl = ScriptApp.getService().getUrl();
    if (svcUrl) return svcUrl;
  } catch (e) {}
  return FIXED_DEPLOY_URL;
}

/**
 * Menu opsional di Spreadsheet apabila script bersifat container-bound.
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Anticurang Web')
      .addItem('Inisialisasi Semua Sheet Database', 'initializeSpreadsheetStructure')
      .addToUi();
  } catch (err) {}
}

/**
 * Memastikan seluruh sheet (DB_Anticurang, Pengaturan, Log_Monitoring) tersedia beserta header.
 */
function initializeSpreadsheetStructure() {
  const ss = getSpreadsheet_();
  const dbSheet = ensureSheetWithHeaders_(ss, SHEET_DB, DB_HEADERS);
  ensureDbDurationColumnInSheet_(dbSheet);
  const mapelSheet = ensureSheetWithHeaders_(ss, SHEET_MAPEL, MAPEL_HEADERS);
  const muridSheet = ensureSheetWithHeaders_(ss, SHEET_MURID, MURID_HEADERS);
  const setSheet = ensureSheetWithHeaders_(ss, SHEET_SETTINGS, SETTINGS_HEADERS);
  const monSheet = ensureSheetWithHeaders_(ss, SHEET_MONITORING, MONITORING_HEADERS);

  if (setSheet.getLastRow() < 2) {
    setSheet.getRange(2, 1, 1, 7).setValues([[
      DEFAULT_SETTINGS.namaSekolah,
      DEFAULT_SETTINGS.logo,
      DEFAULT_SETTINGS.kop,
      DEFAULT_SETTINGS.folderIdDrive,
      DEFAULT_SETTINGS.sumberData,
      DEFAULT_SETTINGS.pinPengawas,
      getStoredAdminPin_()
    ]]);
  } else {
    ensurePinColumnsInSheet_(setSheet);
  }
  clearPortalCache_();
  return {
    ok: true,
    sheets: [dbSheet.getName(), mapelSheet.getName(), muridSheet.getName(), setSheet.getName(), monSheet.getName()]
  };
}

/**
 * Mengambil data publik untuk Portal Jadwal Ujian & header branding global.
 * Memanfaatkan CacheService singkat (15 detik) agar efisien saat diakses banyak siswa bersamaan.
 */
function getPublicPortalData(forceRefresh) {
  const cache = CacheService.getScriptCache();
  let cachedPayload = null;
  if (!forceRefresh) {
    try {
      const raw = cache.get(PUBLIC_CACHE_KEY);
      if (raw) cachedPayload = JSON.parse(raw);
    } catch (e) {}
  }

  let ss;
  let tz;
  let settings;
  let rows;
  let muridRows;
  let isFirstSetup;

  if (cachedPayload) {
    tz = cachedPayload.timeZone;
    settings = cachedPayload.settings;
    rows = cachedPayload.rows;
    muridRows = cachedPayload.muridRows || [];
    isFirstSetup = cachedPayload.isFirstSetup;
  } else {
    ss = getSpreadsheet_();
    tz = ss.getSpreadsheetTimeZone() || 'Asia/Singapore';
    settings = getSettings_(ss);
    rows = readExamRows_(ss, tz);
    muridRows = readMasterMuridRows_(ss);
    ensureSheetWithHeaders_(ss, SHEET_MONITORING, MONITORING_HEADERS);
    isFirstSetup = !hasCustomAdminPin_() && rows.length === 0;
    try {
      cache.put(PUBLIC_CACHE_KEY, JSON.stringify({
        timeZone: tz,
        settings: settings,
        rows: rows,
        muridRows: muridRows,
        isFirstSetup: isFirstSetup
      }), PUBLIC_CACHE_TTL_SEC);
    } catch (e) {}
  }

  const now = new Date();
  const config = getSecurityConfig_();

  const exams = rows.map(function (row) {
    const schedule = evaluateExamSchedule_(row, now, tz);
    return {
      rowIndex: row.rowIndex,
      no: row.no,
      mapel: row.mapel,
      kelas: row.kelas,
      hasLink: Boolean(row.link),
      tanggal: row.tanggal,
      hariLabel: formatHariTanggalIndonesia_(row.tanggal),
      tanggalDisplay: formatTanggalIndonesia_(row.tanggal),
      waktuMulai: row.waktuMulai,
      waktuSelesai: row.waktuSelesai,
      durasiMenit: row.durasiMenit || 0,
      status: schedule.status,
      statusLabel: schedule.statusLabel,
      startMs: schedule.startMs,
      endMs: schedule.endMs
    };
  });

  const students = (muridRows || []).map(function (m) {
    return {
      no: m.no,
      nama: m.namaMurid,
      nis: m.nis,
      kelas: m.kelas
    };
  });

  return {
    ok: true,
    serverNowMs: now.getTime(),
    timeZone: tz,
    todayIso: Utilities.formatDate(now, tz, 'yyyy-MM-dd'),
    isFirstTimeSetup: isFirstSetup,
    settings: enrichSettingsForClient_(settings),
    security: {
      maxViolations: config.maxViolations,
      requireFullscreen: config.requireFullscreen
    },
    exams: exams,
    students: students
  };
}

/**
 * Memulai sesi ujian anticurang, validasi jadwal, catat ke Log_Monitoring, dan kembalikan URL embed.
 */
function startExamSession(payload) {
  payload = payload || {};
  const idx = Number(payload.rowIndex);
  const studentName = String(payload.studentName || '').trim();
  const studentClass = String(payload.studentClass || '').trim();
  const studentNis = String(payload.studentNis || '').trim();

  if (!idx || idx < 2) {
    return { ok: false, message: 'Data soal tidak valid.' };
  }
  if (!studentName || studentName.length < 2) {
    return { ok: false, message: 'Nama lengkap siswa wajib diisi sebelum membuka soal.' };
  }

  const ss = getSpreadsheet_();
  const tz = ss.getSpreadsheetTimeZone() || 'Asia/Singapore';
  const now = new Date();
  const rows = readExamRows_(ss, tz);
  const target = rows.find(function (r) { return r.rowIndex === idx; });

  if (!target) {
    return { ok: false, message: 'Data ujian tidak ditemukan atau sudah dihapus.' };
  }
  if (!target.link) {
    return { ok: false, message: 'Link Google Form belum diisi untuk mata pelajaran ini.' };
  }

  const schedule = evaluateExamSchedule_(target, now, tz);
  if (schedule.status === 'upcoming') {
    return {
      ok: false,
      message: 'Ujian belum dapat dibuka. Jadwal mulai buka: ' + formatTanggalIndonesia_(target.tanggal) + ' pukul ' + target.waktuMulai + '.'
    };
  }
  if (schedule.status === 'ended') {
    return {
      ok: false,
      message: 'Batas akhir membuka soal telah berakhir pada ' + formatTanggalIndonesia_(target.tanggal) + ' pukul ' + target.waktuSelesai + '.'
    };
  }

  const normalized = resolveEmbeddableFormUrl_(target.link);
  if (!normalized.ok) {
    return { ok: false, message: normalized.message };
  }

  // Opsi B: Waktu Mulai & Waktu Selesai hanya jendela buka soal.
  // Selama dibuka sebelum Waktu Selesai, durasi penuh kolom H (menit) diberikan meski melewati Waktu Selesai.
  const durationMinutes = Number(target.durasiMenit) || 0;
  const sessionEndMs = durationMinutes > 0 ? (now.getTime() + durationMinutes * 60 * 1000) : null;

  const muridRows = readMasterMuridRows_(ss);
  const nameLower = studentName.toLowerCase();
  const reqClassLower = studentClass.toLowerCase();
  const matchedMurid = muridRows.find(function (m) {
    return m.namaMurid.toLowerCase() === nameLower && reqClassLower && m.kelas.toLowerCase() === reqClassLower;
  }) || muridRows.find(function (m) {
    return m.namaMurid.toLowerCase() === nameLower;
  }) || null;

  const finalName = matchedMurid ? matchedMurid.namaMurid : studentName;
  const finalNis = (matchedMurid && matchedMurid.nis) ? matchedMurid.nis : (studentNis || '-');
  const finalClass = (matchedMurid && matchedMurid.kelas)
    ? matchedMurid.kelas
    : ((target.kelas && target.kelas !== '-') ? target.kelas : (studentClass || '-'));

  const sessionId = Utilities.getUuid();
  const accessTimeStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const monSheet = ensureSheetWithHeaders_(ss, SHEET_MONITORING, MONITORING_HEADERS);
    const cols = getColumnMap_(monSheet, MONITORING_HEADERS);
    const nextRow = findNextAppendRow_(monSheet, MONITORING_HEADERS.length);
    const rowData = new Array(MONITORING_HEADERS.length).fill('');
    rowData[cols['ID Sesi'] - 1] = sessionId;
    rowData[cols['Waktu Akses'] - 1] = accessTimeStr;
    rowData[cols['Nama Siswa'] - 1] = finalName;
    rowData[cols['NIS / Peserta'] - 1] = finalNis;
    rowData[cols['Kelas Siswa'] - 1] = finalClass;
    rowData[cols['No Soal'] - 1] = target.no;
    rowData[cols['Mapel'] - 1] = target.mapel;
    rowData[cols['Kelas Jadwal'] - 1] = target.kelas;
    rowData[cols['Tanggal Ujian'] - 1] = target.tanggal || Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    rowData[cols['Status'] - 1] = 'Sedang Mengerjakan';
    rowData[cols['Jumlah Pelanggaran'] - 1] = 0;
    rowData[cols['Catatan Pelanggaran'] - 1] = '';
    rowData[cols['Waktu Selesai'] - 1] = '';

    const range = monSheet.getRange(nextRow, 1, 1, rowData.length);
    range.setNumberFormat('@');
    range.setValues([rowData]);
    monSheet.getRange(nextRow, cols['Jumlah Pelanggaran']).setNumberFormat('0').setValue(0);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  const config = getSecurityConfig_();
  return {
    ok: true,
    sessionId: sessionId,
    serverNowMs: now.getTime(),
    student: {
      name: finalName,
      nis: finalNis,
      kelas: finalClass
    },
    exam: {
      rowIndex: target.rowIndex,
      no: target.no,
      mapel: target.mapel,
      kelas: target.kelas,
      tanggal: target.tanggal,
      tanggalDisplay: formatTanggalIndonesia_(target.tanggal),
      waktuMulai: target.waktuMulai,
      waktuSelesai: target.waktuSelesai,
      durasiMenit: durationMinutes,
      sessionEndMs: sessionEndMs,
      endMs: sessionEndMs,
      accessStartMs: schedule.startMs,
      accessEndMs: schedule.endMs,
      embedUrl: normalized.embedUrl
    },
    security: {
      maxViolations: config.maxViolations,
      requireFullscreen: config.requireFullscreen
    }
  };
}

/**
 * Memperbarui log pelanggaran & status selesai sesi pada Sheet Log_Monitoring.
 */
function updateMonitoringSession(payload) {
  payload = payload || {};
  const sessionId = String(payload.sessionId || '').trim();
  if (!sessionId) return { ok: false, message: 'ID Sesi tidak valid.' };

  const status = String(payload.status || '').trim();
  const violationCount = Number(payload.violations);
  const logsText = String(payload.logsText || '').trim().slice(0, 1500);
  const isFinished = Boolean(payload.finished);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = getSpreadsheet_();
    const tz = ss.getSpreadsheetTimeZone() || 'Asia/Singapore';
    const monSheet = ensureSheetWithHeaders_(ss, SHEET_MONITORING, MONITORING_HEADERS);
    const cols = getColumnMap_(monSheet, MONITORING_HEADERS);
    const rowIdx = findMonitoringRowBySession_(monSheet, cols['ID Sesi'], sessionId);
    if (!rowIdx) {
      return { ok: false, message: 'Sesi monitoring tidak ditemukan.' };
    }

    if (status) {
      monSheet.getRange(rowIdx, cols['Status']).setNumberFormat('@').setValue(status);
    }
    if (!isNaN(violationCount) && violationCount >= 0) {
      monSheet.getRange(rowIdx, cols['Jumlah Pelanggaran']).setNumberFormat('0').setValue(violationCount);
    }
    if (logsText) {
      monSheet.getRange(rowIdx, cols['Catatan Pelanggaran']).setNumberFormat('@').setValue(logsText);
    }
    if (isFinished) {
      const finishTimeStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
      monSheet.getRange(rowIdx, cols['Waktu Selesai']).setNumberFormat('@').setValue(finishTimeStr);
    }
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Verifikasi PIN Login (memisahkan PIN Login Admin dan PIN Login Pengawas dari database).
 */
function verifyAdminPin(pin) {
  const input = String(pin || '').trim();
  if (!input) {
    return { ok: false, message: 'PIN wajib diisi.' };
  }

  const settings = getSettings_();
  const storedAdminPin = String(settings.pinAdmin || getStoredAdminPin_()).trim();
  const proctorPin = String(settings.pinPengawas || DEFAULT_PROCTOR_PIN).trim();

  let role = '';
  let displayName = '';

  if (input === storedAdminPin) {
    role = 'Admin';
    displayName = 'Administrator';
  } else if (input === proctorPin) {
    role = 'Pengawas';
    displayName = 'Pengawas Ruang';
  } else {
    return { ok: false, message: 'PIN Admin atau PIN Pengawas salah.' };
  }

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('admin_tok_' + token, role, ADMIN_TOKEN_TTL_SEC);
  return { ok: true, token: token, role: role, displayName: displayName };
}

/**
 * Mengecek apakah token sesi masih valid beserta role-nya saat berpindah antar halaman MPA.
 */
function validateAdminToken(token) {
  const role = getTokenRole_(token);
  return {
    ok: Boolean(role),
    role: role || '',
    displayName: role === 'Admin' ? 'Administrator' : (role === 'Pengawas' ? 'Pengawas Ruang' : '')
  };
}

/**
 * Verifikasi PIN Pengawas untuk membuka kunci sesi siswa yang terkena batas pelanggaran.
 * Terpisah dari PIN Admin: membaca langsung PIN Pengawas dari Database Sheet Pengaturan.
 */
function verifyProctorUnlockPin(pin, sessionId) {
  const input = String(pin || '').trim();
  const settings = getSettings_();
  const proctorPin = String(settings.pinPengawas || DEFAULT_PROCTOR_PIN).trim();
  if (!input || input !== proctorPin) {
    return { ok: false, message: 'PIN Pengawas salah. Hubungi Pengawas atau cek Pengaturan Sistem.' };
  }
  if (sessionId) {
    updateMonitoringSession({
      sessionId: sessionId,
      status: 'Dibuka Pengawas (Lanjut)',
      finished: false
    });
  }
  return { ok: true };
}

/**
 * Mengambil data dashboard sesuai role:
 * - Admin: DB_Anticurang, Pengaturan, Log_Monitoring
 * - Pengawas: Log_Monitoring & info Pengaturan (tanpa akses ubah data soal/pengaturan)
 */
function getAdminDashboardData(token) {
  const role = getTokenRole_(token);
  if (!role) {
    return { ok: false, authError: true, message: 'Sesi berakhir. Silakan login kembali.' };
  }

  const ss = getSpreadsheet_();
  const tz = ss.getSpreadsheetTimeZone() || 'Asia/Singapore';
  const now = new Date();
  const setSheet = ensureSheetWithHeaders_(ss, SHEET_SETTINGS, SETTINGS_HEADERS);
  ensurePinColumnsInSheet_(setSheet);
  const settings = getSettings_(ss);
  const config = getSecurityConfig_();
  const monitoring = readMonitoringRows_(ss);

  let exams = [];
  let masterMapel = [];
  let masterMurid = [];
  let masterKelas = [];
  if (role === 'Admin') {
    const rows = readExamRows_(ss, tz);
    masterMapel = readMasterMapelRows_(ss);
    masterMurid = readMasterMuridRows_(ss);
    masterKelas = deriveUniqueClassesFromMurid_(masterMurid);
    exams = rows.map(function (row) {
      const schedule = evaluateExamSchedule_(row, now, tz);
      return {
        rowIndex: row.rowIndex,
        no: row.no,
        mapel: row.mapel,
        kelas: row.kelas,
        link: row.link,
        tanggal: row.tanggal,
        hariLabel: formatHariTanggalIndonesia_(row.tanggal),
        tanggalDisplay: formatTanggalIndonesia_(row.tanggal),
        waktuMulai: row.waktuMulai,
        waktuSelesai: row.waktuSelesai,
        durasiMenit: row.durasiMenit || 0,
        status: schedule.status,
        statusLabel: schedule.statusLabel,
        isEditLink: /\/forms\/d\/[^/]+\/edit/i.test(row.link || ''),
        isSuspiciousLink: Boolean(row.link && !/docs\.google\.com\/forms\/|forms\.gle\//i.test(row.link))
      };
    });
  }

  return {
    ok: true,
    role: role,
    serverNowMs: now.getTime(),
    timeZone: tz,
    spreadsheetUrl: role === 'Admin' ? ss.getUrl() : '',
    settings: enrichSettingsForClient_(settings, true, role === 'Admin'),
    defaults: DEFAULT_SETTINGS,
    security: {
      maxViolations: config.maxViolations,
      requireFullscreen: config.requireFullscreen
    },
    exams: exams,
    masterMapel: masterMapel,
    masterMurid: masterMurid,
    masterKelas: masterKelas,
    connection: role === 'Admin' ? getConnectionInfo_(ss) : null,
    monitoring: monitoring
  };
}

function getConnectionInfo_(ss) {
  const props = PropertiesService.getScriptProperties();
  let ssId = '';
  let ssName = '';
  try { ssId = ss && typeof ss.getId === 'function' ? ss.getId() : ''; } catch (e) {}
  if (!ssId) {
    try { ssId = extractSpreadsheetIdAny_(ss && typeof ss.getUrl === 'function' ? ss.getUrl() : '') || ''; } catch (e) {}
  }
  try { ssName = ss && typeof ss.getName === 'function' ? ss.getName() : ''; } catch (e) {}
  return {
    activeSpreadsheetId: ssId || SPREADSHEET_ID,
    activeSpreadsheetName: ssName,
    customSpreadsheetId: String(props.getProperty('CUSTOM_SPREADSHEET_ID') || '').trim(),
    customDeployUrl: String(props.getProperty('CUSTOM_DEPLOY_URL') || '').trim(),
    activeDeployUrl: getAppUrl()
  };
}

/**
 * Menyiapkan 5 sheet struktur database pada spreadsheet aktif (untuk Admin dari UI Pengaturan).
 */
function initializeDatabaseByAdmin(token) {
  if (!isValidAdminToken_(token)) {
    return { ok: false, authError: true, message: 'Sesi admin berakhir. Silakan login kembali.' };
  }
  const out = initializeSpreadsheetStructure();
  return {
    ok: Boolean(out && out.ok),
    message: 'Struktur database berhasil disiapkan (' + ((out && out.sheets) ? out.sheets.length : 5) + ' tabel aktif).'
  };
}

/**
 * Menambahkan data master (DB_Mapel atau DB_Kelas) oleh Admin.
 */
function addMasterRecord(token, kind, payload) {
  if (!isValidAdminToken_(token)) {
    return { ok: false, authError: true, message: 'Sesi admin berakhir. Silakan login kembali.' };
  }
  const spec = getMasterSpec_(kind);
  if (!spec) {
    return { ok: false, message: 'Jenis data master tidak dikenal.' };
  }
  payload = payload || {};
  const val = String(payload.value || payload.namaMapel || payload.kelas || '').trim();
  const customNo = String(payload.no || '').trim();
  if (!val) {
    return { ok: false, message: spec.valueLabel + ' wajib diisi.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ensureSheetWithHeaders_(ss, spec.sheetName, spec.headers);
    const cols = getColumnMap_(sheet, spec.headers);
    const existing = readMasterRowsBySpec_(ss, spec);
    const isDuplicate = existing.some(function (item) {
      return String(item[spec.propName] || '').trim().toLowerCase() === val.toLowerCase();
    });
    if (isDuplicate) {
      return { ok: false, message: spec.valueLabel + ' "' + val + '" sudah ada di ' + spec.sheetName + '.' };
    }

    const nextRow = findNextAppendRow_(sheet, 2);
    const finalNo = customNo || computeNextNumber_(sheet, cols['No']);
    const noVal = (finalNo !== '' && !isNaN(Number(finalNo))) ? Number(finalNo) : finalNo;
    const valCol = resolveMasterValueCol_(sheet, cols, spec);
    sheet.getRange(nextRow, cols['No']).setValue(noVal);
    sheet.getRange(nextRow, valCol).setNumberFormat('@').setValue(val);
    SpreadsheetApp.flush();

    return {
      ok: true,
      message: spec.valueLabel + ' "' + val + '" berhasil ditambahkan.',
      items: readMasterRowsBySpec_(ss, spec)
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Memperbarui 1 baris data master (No dan/atau Nama Mapel / Kelas) oleh Admin.
 */
function updateMasterRecord(token, kind, payload) {
  if (!isValidAdminToken_(token)) {
    return { ok: false, authError: true, message: 'Sesi admin berakhir. Silakan login kembali.' };
  }
  const spec = getMasterSpec_(kind);
  if (!spec) {
    return { ok: false, message: 'Jenis data master tidak dikenal.' };
  }
  payload = payload || {};
  const idx = Number(payload.rowIndex);
  const val = String(payload.value || payload.namaMapel || payload.kelas || '').trim();
  const customNo = String(payload.no !== undefined ? payload.no : '').trim();
  if (!idx || idx < 2) {
    return { ok: false, message: 'Baris data tidak valid.' };
  }
  if (!val) {
    return { ok: false, message: spec.valueLabel + ' wajib diisi.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ensureSheetWithHeaders_(ss, spec.sheetName, spec.headers);
    if (idx > sheet.getLastRow()) {
      return { ok: false, message: 'Data sudah tidak ada di database.' };
    }
    const cols = getColumnMap_(sheet, spec.headers);
    const existing = readMasterRowsBySpec_(ss, spec);
    const isDuplicate = existing.some(function (item) {
      return item.rowIndex !== idx && String(item[spec.propName] || '').trim().toLowerCase() === val.toLowerCase();
    });
    if (isDuplicate) {
      return { ok: false, message: spec.valueLabel + ' "' + val + '" sudah digunakan.' };
    }

    const valCol = resolveMasterValueCol_(sheet, cols, spec);
    if (customNo !== '') {
      const noVal = !isNaN(Number(customNo)) ? Number(customNo) : customNo;
      sheet.getRange(idx, cols['No']).setValue(noVal);
    }
    sheet.getRange(idx, valCol).setNumberFormat('@').setValue(val);
    SpreadsheetApp.flush();

    return {
      ok: true,
      message: spec.valueLabel + ' "' + val + '" berhasil diperbarui.',
      items: readMasterRowsBySpec_(ss, spec)
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menghapus 1 baris data master dari DB_Mapel atau DB_Kelas oleh Admin.
 */
function deleteMasterRecord(token, kind, rowIndex) {
  if (!isValidAdminToken_(token)) {
    return { ok: false, authError: true, message: 'Sesi admin berakhir. Silakan login kembali.' };
  }
  const spec = getMasterSpec_(kind);
  if (!spec) {
    return { ok: false, message: 'Jenis data master tidak dikenal.' };
  }
  const idx = Number(rowIndex);
  if (!idx || idx < 2) {
    return { ok: false, message: 'Baris data tidak valid.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ensureSheetWithHeaders_(ss, spec.sheetName, spec.headers);
    if (idx > sheet.getLastRow()) {
      return { ok: false, message: 'Data sudah tidak ada di database.' };
    }
    const cols = getColumnMap_(sheet, spec.headers);
    sheet.deleteRow(idx);
    renumberNonFormulaNo_(sheet, cols['No'], 2);
    SpreadsheetApp.flush();

    return {
      ok: true,
      message: 'Data berhasil dihapus.',
      items: readMasterRowsBySpec_(ss, spec)
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menambahkan 1 data murid ke Sheet DB_Murid oleh Admin.
 */
function addMuridRecord(token, payload) {
  if (!isValidAdminToken_(token)) {
    return { ok: false, authError: true, message: 'Sesi admin berakhir. Silakan login kembali.' };
  }
  payload = payload || {};
  const no = String(payload.no || '').trim();
  const namaMurid = String(payload.namaMurid || payload.nama || '').trim();
  const nis = String(payload.nis || payload.nisn || '').trim();
  const kelas = String(payload.kelas || '').trim();
  if (!namaMurid) return { ok: false, message: 'Nama Murid wajib diisi.' };
  if (!kelas) return { ok: false, message: 'Kelas murid wajib diisi.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ensureSheetWithHeaders_(ss, SHEET_MURID, MURID_HEADERS);
    const cols = resolveMuridCols_(sheet);
    const existing = readMasterMuridRows_(ss);

    const dup = existing.some(function (m) {
      if (nis && nis !== '-' && m.nis && m.nis.toLowerCase() === nis.toLowerCase()) return true;
      return m.namaMurid.toLowerCase() === namaMurid.toLowerCase() && m.kelas.toLowerCase() === kelas.toLowerCase();
    });
    if (dup) {
      return { ok: false, message: 'Data murid atau NIS/NISN tersebut sudah terdaftar.' };
    }

    const nextRow = findNextAppendRow_(sheet, 4);
    const finalNo = no || computeNextNumber_(sheet, cols.no);
    const noVal = (finalNo !== '' && !isNaN(Number(finalNo))) ? Number(finalNo) : finalNo;
    sheet.getRange(nextRow, cols.no).setValue(noVal);
    sheet.getRange(nextRow, cols.nama).setNumberFormat('@').setValue(namaMurid);
    sheet.getRange(nextRow, cols.nis).setNumberFormat('@').setValue(nis);
    sheet.getRange(nextRow, cols.kelas).setNumberFormat('@').setValue(kelas);
    SpreadsheetApp.flush();
    clearPortalCache_();

    return {
      ok: true,
      message: 'Data murid "' + namaMurid + '" berhasil ditambahkan.',
      items: readMasterMuridRows_(ss)
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Memperbarui 1 baris data murid pada Sheet DB_Murid oleh Admin.
 */
function updateMuridRecord(token, payload) {
  if (!isValidAdminToken_(token)) {
    return { ok: false, authError: true, message: 'Sesi admin berakhir. Silakan login kembali.' };
  }
  payload = payload || {};
  const idx = Number(payload.rowIndex);
  const customNo = String(payload.no !== undefined ? payload.no : '').trim();
  const namaMurid = String(payload.namaMurid || payload.nama || '').trim();
  const nis = String(payload.nis !== undefined ? payload.nis : (payload.nisn !== undefined ? payload.nisn : '')).trim();
  const kelas = String(payload.kelas || '').trim();

  if (!idx || idx < 2) return { ok: false, message: 'Baris data tidak valid.' };
  if (!namaMurid) return { ok: false, message: 'Nama Murid wajib diisi.' };
  if (!kelas) return { ok: false, message: 'Kelas murid wajib diisi.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ensureSheetWithHeaders_(ss, SHEET_MURID, MURID_HEADERS);
    if (idx > sheet.getLastRow()) {
      return { ok: false, message: 'Data sudah tidak ada di database.' };
    }
    const cols = resolveMuridCols_(sheet);
    const existing = readMasterMuridRows_(ss);
    const dup = existing.some(function (m) {
      if (m.rowIndex === idx) return false;
      if (nis && nis !== '-' && m.nis && m.nis.toLowerCase() === nis.toLowerCase()) return true;
      return m.namaMurid.toLowerCase() === namaMurid.toLowerCase() && m.kelas.toLowerCase() === kelas.toLowerCase();
    });
    if (dup) {
      return { ok: false, message: 'Nama/Kelas atau NIS/NISN sudah dipakai data lain.' };
    }

    if (customNo !== '') {
      const noVal = !isNaN(Number(customNo)) ? Number(customNo) : customNo;
      sheet.getRange(idx, cols.no).setValue(noVal);
    }
    sheet.getRange(idx, cols.nama).setNumberFormat('@').setValue(namaMurid);
    sheet.getRange(idx, cols.nis).setNumberFormat('@').setValue(nis);
    sheet.getRange(idx, cols.kelas).setNumberFormat('@').setValue(kelas);
    SpreadsheetApp.flush();
    clearPortalCache_();

    return {
      ok: true,
      message: 'Data murid "' + namaMurid + '" berhasil diperbarui.',
      items: readMasterMuridRows_(ss)
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menghapus 1 baris data murid dari Sheet DB_Murid oleh Admin.
 */
function deleteMuridRecord(token, rowIndex) {
  if (!isValidAdminToken_(token)) {
    return { ok: false, authError: true, message: 'Sesi admin berakhir. Silakan login kembali.' };
  }
  const idx = Number(rowIndex);
  if (!idx || idx < 2) return { ok: false, message: 'Baris data tidak valid.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ensureSheetWithHeaders_(ss, SHEET_MURID, MURID_HEADERS);
    if (idx > sheet.getLastRow()) {
      return { ok: false, message: 'Data sudah tidak ada di database.' };
    }
    const cols = resolveMuridCols_(sheet);
    sheet.deleteRow(idx);
    renumberNonFormulaNo_(sheet, cols.no, 4);
    SpreadsheetApp.flush();
    clearPortalCache_();

    return {
      ok: true,
      message: 'Data murid berhasil dihapus.',
      items: readMasterMuridRows_(ss)
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Impor massal data murid dari Excel ke DB_Murid oleh Admin.
 */
function importMuridFromExcel(token, rowsArray, replaceExisting) {
  if (!isValidAdminToken_(token)) {
    return { ok: false, authError: true, message: 'Sesi admin berakhir. Silakan login kembali.' };
  }
  if (!Array.isArray(rowsArray) || !rowsArray.length) {
    return { ok: false, message: 'Data Excel kosong atau tidak terbaca.' };
  }
  const sanitized = [];
  for (var i = 0; i < rowsArray.length; i++) {
    const r = rowsArray[i] || {};
    const namaMurid = String(r.namaMurid || r.nama || '').trim();
    const nis = String(r.nis || r.nisn || '').trim();
    const kelas = String(r.kelas || '').trim();
    const no = String(r.no || '').trim();
    if (namaMurid && kelas) {
      sanitized.push({ no: no, namaMurid: namaMurid, nis: nis, kelas: kelas });
    }
  }
  if (!sanitized.length) {
    return { ok: false, message: 'Tidak ada baris valid (wajib memiliki Nama Murid dan Kelas).' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ensureSheetWithHeaders_(ss, SHEET_MURID, MURID_HEADERS);
    const cols = resolveMuridCols_(sheet);

    if (replaceExisting && sheet.getLastRow() >= 2) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }

    const startRow = findNextAppendRow_(sheet, 4);
    let nextNo = parseInt(computeNextNumber_(sheet, cols.no), 10) || 1;
    const width = Math.max(cols.no, cols.nama, cols.nis, cols.kelas, 4);
    const valuesMatrix = sanitized.map(function (item) {
      const rowArr = new Array(width).fill('');
      const finalNo = item.no || String(nextNo++);
      rowArr[cols.no - 1] = !isNaN(Number(finalNo)) ? Number(finalNo) : finalNo;
      rowArr[cols.nama - 1] = item.namaMurid;
      rowArr[cols.nis - 1] = item.nis;
      rowArr[cols.kelas - 1] = item.kelas;
      return rowArr;
    });

    sheet.getRange(startRow, cols.nis, valuesMatrix.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 1, valuesMatrix.length, width).setValues(valuesMatrix);
    SpreadsheetApp.flush();
    clearPortalCache_();

    return {
      ok: true,
      count: sanitized.length,
      message: sanitized.length + ' data murid berhasil diimpor.',
      items: readMasterMuridRows_(ss)
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Tambah atau perbarui baris data ujian di Sheet DB_Anticurang.
 */
function saveExamRecord(token, payload) {
  if (!isValidAdminToken_(token)) {
    return { ok: false, authError: true, message: 'Sesi admin berakhir. Silakan login kembali.' };
  }

  const validation = validateExamPayload_(payload);
  if (!validation.ok) {
    return validation;
  }
  const clean = validation.data;

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ensureSheetWithHeaders_(ss, SHEET_DB, DB_HEADERS);
    ensureDbDurationColumnInSheet_(sheet);
    const cols = getColumnMap_(sheet, DB_HEADERS);

    const isUpdate = Boolean(clean.rowIndex && clean.rowIndex >= 2);
    let targetRow = clean.rowIndex;

    if (isUpdate) {
      if (targetRow > sheet.getLastRow()) {
        return { ok: false, message: 'Baris data tidak ditemukan di spreadsheet.' };
      }
      if (!clean.no) {
        const existingNo = String(sheet.getRange(targetRow, cols['No']).getDisplayValue() || '').trim();
        clean.no = existingNo || String(targetRow - 1);
      }
    } else {
      targetRow = findNextAppendRow_(sheet, DB_HEADERS.length);
      if (!clean.no) {
        clean.no = computeNextNumber_(sheet, cols['No']);
      }
    }

    writeExamRowToSheet_(sheet, targetRow, cols, clean);
    SpreadsheetApp.flush();
    clearPortalCache_();

    return {
      ok: true,
      message: isUpdate ? 'Data soal berhasil diperbarui.' : 'Link soal baru berhasil ditambahkan.',
      rowIndex: targetRow
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Impor massal jadwal soal dari Excel (SheetJS) ke DB_Anticurang.
 */
function importExamsFromExcel(token, rowsArray, replaceExisting) {
  if (!isValidAdminToken_(token)) {
    return { ok: false, authError: true, message: 'Sesi admin berakhir. Silakan login kembali.' };
  }
  if (!Array.isArray(rowsArray) || !rowsArray.length) {
    return { ok: false, message: 'Data Excel kosong atau tidak terbaca.' };
  }

  const sanitized = [];
  for (var i = 0; i < rowsArray.length; i++) {
    const check = validateExamPayload_(rowsArray[i]);
    if (check.ok) {
      sanitized.push(check.data);
    }
  }
  if (!sanitized.length) {
    return { ok: false, message: 'Tidak ada baris valid (wajib memiliki Mapel, Kelas, dan Link GFrom).' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ensureSheetWithHeaders_(ss, SHEET_DB, DB_HEADERS);
    ensureDbDurationColumnInSheet_(sheet);
    const cols = getColumnMap_(sheet, DB_HEADERS);

    if (replaceExisting && sheet.getLastRow() >= 2) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }

    let startRow = findNextAppendRow_(sheet, DB_HEADERS.length);
    let nextNo = parseInt(computeNextNumber_(sheet, cols['No']), 10) || 1;

    sanitized.forEach(function (item, idx) {
      if (!item.no) {
        item.no = String(nextNo++);
      }
      writeExamRowToSheet_(sheet, startRow + idx, cols, item);
    });

    SpreadsheetApp.flush();
    clearPortalCache_();
    return {
      ok: true,
      count: sanitized.length,
      message: sanitized.length + ' jadwal link soal berhasil diimpor.'
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menghapus baris data ujian dari Sheet DB_Anticurang.
 */
function deleteExamRecord(token, rowIndex) {
  if (!isValidAdminToken_(token)) {
    return { ok: false, authError: true, message: 'Sesi admin berakhir. Silakan login kembali.' };
  }
  const idx = Number(rowIndex);
  if (!idx || idx < 2) {
    return { ok: false, message: 'Nomor baris tidak valid.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const sheet = ensureSheetWithHeaders_(ss, SHEET_DB, DB_HEADERS);
    if (idx > sheet.getLastRow()) {
      return { ok: false, message: 'Data sudah tidak ada di spreadsheet.' };
    }
    sheet.deleteRow(idx);
    renumberNonFormulaNo_(sheet, getColumnMap_(sheet, DB_HEADERS)['No']);
    SpreadsheetApp.flush();
    clearPortalCache_();

    return { ok: true, message: 'Data soal berhasil dihapus.' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menghapus satu baris atau mengosongkan seluruh Log_Monitoring.
 */
function clearMonitoringLogs(token, rowIndex) {
  if (!isValidAdminToken_(token)) {
    return { ok: false, authError: true, message: 'Sesi admin berakhir. Silakan login kembali.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const monSheet = ensureSheetWithHeaders_(ss, SHEET_MONITORING, MONITORING_HEADERS);
    const lastRow = monSheet.getLastRow();
    if (lastRow < 2) {
      return { ok: true, message: 'Log monitoring sudah kosong.' };
    }
    if (rowIndex && Number(rowIndex) >= 2 && Number(rowIndex) <= lastRow) {
      monSheet.deleteRow(Number(rowIndex));
      SpreadsheetApp.flush();
      return { ok: true, message: '1 baris log monitoring berhasil dihapus.' };
    }
    monSheet.deleteRows(2, lastRow - 1);
    SpreadsheetApp.flush();
    return { ok: true, message: 'Seluruh log monitoring berhasil dibersihkan.' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menyimpan Pengaturan (termasuk dukungan unggah gambar Logo/Kop ke Google Drive).
 */
function saveAppSettings(token, payload) {
  if (!isValidAdminToken_(token)) {
    return { ok: false, authError: true, message: 'Sesi admin berakhir. Silakan login kembali.' };
  }
  payload = payload || {};

  const namaSekolah = String(payload.namaSekolah || '').trim() || DEFAULT_SETTINGS.namaSekolah;
  let logo = String(payload.logo || '').trim() || DEFAULT_SETTINGS.logo;
  let kop = String(payload.kop || '').trim() || DEFAULT_SETTINGS.kop;
  const folderIdDrive = extractDriveIdAny_(payload.folderIdDrive) || String(payload.folderIdDrive || '').trim();
  const sumberData = String(payload.sumberData || '').trim() || DEFAULT_SETTINGS.sumberData;
  const pinPengawas = String(payload.pinPengawas || '').trim() || DEFAULT_PROCTOR_PIN;
  const pinAdminInput = String(payload.pinAdmin || payload.newPin || '').trim();
  if (pinAdminInput && pinAdminInput.length < 4) {
    return { ok: false, message: 'PIN Login Admin minimal 4 karakter.' };
  }

  const props = PropertiesService.getScriptProperties();
  let customSsIdToSave = null;
  if (payload.customSpreadsheetId !== undefined) {
    const rawSs = String(payload.customSpreadsheetId || '').trim();
    if (!rawSs) {
      customSsIdToSave = '';
    } else {
      const extractedSs = extractSpreadsheetIdAny_(rawSs);
      if (!extractedSs) {
        return { ok: false, message: 'Link atau ID Google Spreadsheet tidak valid.' };
      }
      try {
        SpreadsheetApp.openById(extractedSs);
      } catch (openErr) {
        return { ok: false, message: 'Spreadsheet baru tidak dapat diakses: ' + openErr.message };
      }
      customSsIdToSave = extractedSs;
    }
  }

  let customDeployToSave = null;
  if (payload.customDeployUrl !== undefined) {
    const rawDep = String(payload.customDeployUrl || '').trim();
    if (!rawDep) {
      customDeployToSave = '';
    } else if (!/^https?:\/\/.+/i.test(rawDep)) {
      return { ok: false, message: 'Link Deploy Web App harus diawali https://.' };
    } else {
      customDeployToSave = rawDep;
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (customSsIdToSave !== null) {
      props.setProperty('CUSTOM_SPREADSHEET_ID', customSsIdToSave);
    }
    if (customDeployToSave !== null) {
      props.setProperty('CUSTOM_DEPLOY_URL', customDeployToSave);
    }

    const ss = getSpreadsheet_();
    const sheet = ensureSheetWithHeaders_(ss, SHEET_SETTINGS, SETTINGS_HEADERS);
    if (customSsIdToSave) {
      ensureSheetWithHeaders_(ss, SHEET_DB, DB_HEADERS);
      ensureDbDurationColumnInSheet_(ss.getSheetByName(SHEET_DB));
      ensureSheetWithHeaders_(ss, SHEET_MAPEL, MAPEL_HEADERS);
      ensureSheetWithHeaders_(ss, SHEET_MURID, MURID_HEADERS);
      ensureSheetWithHeaders_(ss, SHEET_MONITORING, MONITORING_HEADERS);
    }
    const cols = getColumnMap_(sheet, SETTINGS_HEADERS);
    const oldSettings = getSettings_(ss);
    const pinAdmin = pinAdminInput || oldSettings.pinAdmin || getStoredAdminPin_();

    // Dual-mode ingestion: jika ada file Base64 untuk Logo / Kop, unggah ke Google Drive
    if (payload.logoBase64) {
      const uploadedLogo = uploadImageBase64ToDrive_(payload.logoBase64, 'Logo_' + Date.now(), folderIdDrive, oldSettings.logo);
      if (!uploadedLogo.ok) return uploadedLogo;
      logo = uploadedLogo.url;
    }
    if (payload.kopBase64) {
      const uploadedKop = uploadImageBase64ToDrive_(payload.kopBase64, 'Kop_' + Date.now(), folderIdDrive, oldSettings.kop);
      if (!uploadedKop.ok) return uploadedKop;
      kop = uploadedKop.url;
    }

    sheet.getRange(2, cols['Nama Sekolah']).setNumberFormat('@').setValue(namaSekolah);
    sheet.getRange(2, cols['Logo']).setNumberFormat('@').setValue(logo);
    sheet.getRange(2, cols['Kop']).setNumberFormat('@').setValue(kop);
    if (folderIdDrive) {
      ensureHeaderCell_(sheet, cols['Folder ID Drive'], 'Folder ID Drive');
      sheet.getRange(2, cols['Folder ID Drive']).setNumberFormat('@').setValue(folderIdDrive);
    }
    if (sumberData) {
      ensureHeaderCell_(sheet, cols['Sumber Data'], 'Sumber Data');
      sheet.getRange(2, cols['Sumber Data']).setNumberFormat('@').setValue(sumberData);
    }
    ensureHeaderCell_(sheet, cols['PIN Pengawas'], 'PIN Pengawas');
    sheet.getRange(2, cols['PIN Pengawas']).setNumberFormat('@').setValue(pinPengawas);
    ensureHeaderCell_(sheet, cols['PIN Admin'], 'PIN Admin');
    sheet.getRange(2, cols['PIN Admin']).setNumberFormat('@').setValue(pinAdmin);
    SpreadsheetApp.flush();

    const maxViol = Number(payload.maxViolations);
    if (!isNaN(maxViol) && maxViol >= 0 && maxViol <= 20) {
      props.setProperty('MAX_VIOLATIONS', String(Math.floor(maxViol)));
    }
    if (typeof payload.requireFullscreen === 'boolean') {
      props.setProperty('REQUIRE_FULLSCREEN', payload.requireFullscreen ? 'true' : 'false');
    }
    props.setProperty('ADMIN_PIN', pinAdmin);

    clearPortalCache_();

    return {
      ok: true,
      message: 'Pengaturan sistem, keamanan & koneksi database berhasil disimpan.',
      connection: getConnectionInfo_(ss),
      settings: enrichSettingsForClient_({
        namaSekolah: namaSekolah,
        logo: logo,
        kop: kop,
        folderIdDrive: folderIdDrive,
        sumberData: sumberData,
        pinPengawas: pinPengawas,
        pinAdmin: pinAdmin
      }, true, true)
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menguji dan menormalkan URL Google Form dari halaman Kelola Soal.
 */
function testNormalizeFormUrl(token, rawUrl) {
  if (!isValidAdminToken_(token)) {
    return { ok: false, authError: true, message: 'Sesi admin berakhir. Silakan login kembali.' };
  }
  return resolveEmbeddableFormUrl_(rawUrl);
}

/* =========================================================================
 * INTERNAL HELPERS
 * ========================================================================= */

function extractSpreadsheetIdAny_(val) {
  const s = String(val || '').trim();
  if (!s) return '';
  const m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  const mId = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (mId) return mId[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s;
  return '';
}

function getSpreadsheet_() {
  let customId = '';
  try {
    customId = String(PropertiesService.getScriptProperties().getProperty('CUSTOM_SPREADSHEET_ID') || '').trim();
  } catch (e) {}
  if (customId) {
    try {
      return SpreadsheetApp.openById(customId);
    } catch (eCustom) {
      throw new Error('Tidak dapat mengakses Database Spreadsheet ID ' + customId + ': ' + eCustom.message);
    }
  }
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (eActive) {}
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    throw new Error('Tidak dapat mengakses Spreadsheet ID ' + SPREADSHEET_ID + ': ' + e.message);
  }
}

function clearPortalCache_() {
  try {
    CacheService.getScriptCache().remove(PUBLIC_CACHE_KEY);
  } catch (e) {}
}

function ensureSheetWithHeaders_(ss, sheetName, expectedHeaders) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
  } else {
    const currentLastCol = Math.max(sheet.getLastColumn(), expectedHeaders.length);
    const row1 = sheet.getRange(1, 1, 1, currentLastCol).getDisplayValues()[0];
    const hasAnyHeader = row1.some(function (cell) { return String(cell).trim() !== ''; });
    if (!hasAnyHeader) {
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    }
  }
  return sheet;
}

function ensureHeaderCell_(sheet, colIdx, label) {
  const existing = String(sheet.getRange(1, colIdx).getDisplayValue() || '').trim();
  if (!existing) {
    sheet.getRange(1, colIdx).setValue(label);
  }
}

function getColumnMap_(sheet, defaultHeaders) {
  const lastCol = Math.max(sheet.getLastColumn(), defaultHeaders.length);
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const normalizedLookup = {};
  headerRow.forEach(function (label, idx) {
    const key = String(label || '').trim().toLowerCase();
    if (key && !(key in normalizedLookup)) {
      normalizedLookup[key] = idx + 1;
    }
  });

  const map = {};
  defaultHeaders.forEach(function (name, idx) {
    const key = name.toLowerCase();
    map[name] = normalizedLookup[key] || (idx + 1);
  });
  return map;
}

function ensurePinColumnsInSheet_(sheet) {
  try {
    const cols = getColumnMap_(sheet, SETTINGS_HEADERS);
    const proctorCol = cols['PIN Pengawas'];
    const adminCol = cols['PIN Admin'];

    if (!String(sheet.getRange(1, proctorCol).getDisplayValue() || '').trim()) {
      sheet.getRange(1, proctorCol).setValue('PIN Pengawas');
    }
    if (!String(sheet.getRange(1, adminCol).getDisplayValue() || '').trim()) {
      sheet.getRange(1, adminCol).setValue('PIN Admin');
    }

    if (sheet.getLastRow() >= 2) {
      if (!String(sheet.getRange(2, proctorCol).getDisplayValue() || '').trim()) {
        sheet.getRange(2, proctorCol).setNumberFormat('@').setValue(DEFAULT_PROCTOR_PIN);
      }
      if (!String(sheet.getRange(2, adminCol).getDisplayValue() || '').trim()) {
        sheet.getRange(2, adminCol).setNumberFormat('@').setValue(getStoredAdminPin_());
      }
    }
  } catch (e) {}
}

function ensureDbDurationColumnInSheet_(sheet) {
  try {
    const cols = getColumnMap_(sheet, DB_HEADERS);
    const durCol = cols['Durasi (Menit)'];
    if (!String(sheet.getRange(1, durCol).getDisplayValue() || '').trim()) {
      sheet.getRange(1, durCol).setValue('Durasi (Menit)');
    }
  } catch (e) {}
}

/**
 * Teknik Full Column Scan untuk membaca Pengaturan secara resilien.
 */
function getSettings_(optionalSs) {
  const ss = optionalSs || getSpreadsheet_();
  const sheet = ensureSheetWithHeaders_(ss, SHEET_SETTINGS, SETTINGS_HEADERS);
  const cols = getColumnMap_(sheet, SETTINGS_HEADERS);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return Object.assign({}, DEFAULT_SETTINGS, { pinAdmin: getStoredAdminPin_() });
  }

  const maxCol = Math.max(
    cols['Nama Sekolah'],
    cols['Logo'],
    cols['Kop'],
    cols['Folder ID Drive'],
    cols['Sumber Data'],
    cols['PIN Pengawas'],
    cols['PIN Admin'],
    7
  );
  const data = sheet.getRange(2, 1, lastRow - 1, maxCol).getDisplayValues();

  let namaSekolah = '';
  let logo = '';
  let kop = '';
  let folderIdDrive = '';
  let sumberData = '';
  let pinPengawas = '';
  let pinAdmin = '';

  for (var r = 0; r < data.length; r++) {
    if (!namaSekolah && data[r][cols['Nama Sekolah'] - 1]) {
      namaSekolah = String(data[r][cols['Nama Sekolah'] - 1]).trim();
    }
    if (!logo && data[r][cols['Logo'] - 1]) {
      logo = String(data[r][cols['Logo'] - 1]).trim();
    }
    if (!kop && data[r][cols['Kop'] - 1]) {
      kop = String(data[r][cols['Kop'] - 1]).trim();
    }
    if (!folderIdDrive && data[r][cols['Folder ID Drive'] - 1]) {
      folderIdDrive = String(data[r][cols['Folder ID Drive'] - 1]).trim();
    }
    if (!sumberData && data[r][cols['Sumber Data'] - 1]) {
      sumberData = String(data[r][cols['Sumber Data'] - 1]).trim();
    }
    if (!pinPengawas && data[r][cols['PIN Pengawas'] - 1]) {
      pinPengawas = String(data[r][cols['PIN Pengawas'] - 1]).trim();
    }
    if (!pinAdmin && data[r][cols['PIN Admin'] - 1]) {
      pinAdmin = String(data[r][cols['PIN Admin'] - 1]).trim();
    }
  }

  return {
    namaSekolah: namaSekolah || DEFAULT_SETTINGS.namaSekolah,
    logo: logo || DEFAULT_SETTINGS.logo,
    kop: kop || DEFAULT_SETTINGS.kop,
    folderIdDrive: folderIdDrive || DEFAULT_SETTINGS.folderIdDrive,
    sumberData: sumberData || DEFAULT_SETTINGS.sumberData,
    pinPengawas: pinPengawas || DEFAULT_PROCTOR_PIN,
    pinAdmin: pinAdmin || getStoredAdminPin_()
  };
}

function enrichSettingsForClient_(settings, includeProctorPin, includeAdminPin) {
  const logoRaw = settings.logo || DEFAULT_SETTINGS.logo;
  const kopRaw = settings.kop || DEFAULT_SETTINGS.kop;
  const isKopImg = /(https?:\/\/.*\.(?:png|jpg|jpeg|gif|webp|svg))|(drive\.google\.com)|(lh3\.googleusercontent\.com)/i.test(kopRaw);
  const out = {
    namaSekolah: settings.namaSekolah || DEFAULT_SETTINGS.namaSekolah,
    logoRaw: logoRaw,
    logoUrl: toDriveImageUrl_(logoRaw),
    logoFallbackUrl: toDriveThumbnailUrl_(logoRaw),
    kopRaw: kopRaw,
    isKopImage: isKopImg,
    kopUrl: isKopImg ? toDriveImageUrl_(kopRaw) : '',
    kopFallbackUrl: isKopImg ? toDriveThumbnailUrl_(kopRaw) : '',
    folderIdDrive: settings.folderIdDrive || '',
    sumberData: settings.sumberData || DEFAULT_SETTINGS.sumberData
  };
  if (includeProctorPin) {
    out.pinPengawas = settings.pinPengawas || DEFAULT_PROCTOR_PIN;
  }
  if (includeAdminPin) {
    out.pinAdmin = settings.pinAdmin || getStoredAdminPin_();
  }
  return out;
}

function uploadImageBase64ToDrive_(dataUrl, baseName, folderId, oldUrl) {
  try {
    const match = String(dataUrl || '').match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
    if (!match) {
      return { ok: false, message: 'Format file gambar tidak valid.' };
    }
    const mime = match[1];
    const ext = mime.split('/')[1].replace('jpeg', 'jpg').replace('svg+xml', 'svg');
    const bytes = Utilities.base64Decode(match[2]);
    const blob = Utilities.newBlob(bytes, mime, baseName + '.' + ext);

    let file;
    if (folderId) {
      const folder = DriveApp.getFolderById(folderId);
      file = folder.createFile(blob);
    } else {
      file = DriveApp.createFile(blob);
    }
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Root Folder Guard: hanya bersihkan file Drive lama jika bukan folder utama & bukan file default bawaan
    const oldId = extractDriveFileId_(oldUrl);
    if (oldId && oldId !== folderId && oldId !== '1_4m0lkw99HvxaPXvOhy0UOudoPmiIRc-' && oldId !== '1QXtnYwIrCc-pihDUoidjkRW5UVWykOn-') {
      try {
        DriveApp.getFileById(oldId).setTrashed(true);
      } catch (ignoreOld) {}
    }

    return {
      ok: true,
      url: 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=drive_link'
    };
  } catch (err) {
    return {
      ok: false,
      message: 'Gagal mengunggah gambar ke Google Drive: ' + err.message
    };
  }
}

function getMasterSpec_(kind) {
  const k = String(kind || '').trim().toLowerCase();
  if (k === 'mapel') {
    return {
      sheetName: SHEET_MAPEL,
      headers: MAPEL_HEADERS,
      headerColName: 'Nama Mapel',
      fallbackHeaderAliases: ['mapel', 'mata pelajaran', 'nama mapel'],
      valueLabel: 'Nama Mapel',
      propName: 'namaMapel'
    };
  }
  return null;
}

function resolveMasterValueCol_(sheet, cols, spec) {
  const lastCol = Math.max(sheet.getLastColumn(), 2);
  const row1 = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  for (var i = 0; i < row1.length; i++) {
    const label = String(row1[i] || '').trim().toLowerCase();
    if (spec.fallbackHeaderAliases.indexOf(label) !== -1) {
      return i + 1;
    }
  }
  return cols[spec.headerColName] || 2;
}

function readMasterMapelRows_(ss) {
  return readMasterRowsBySpec_(ss, getMasterSpec_('mapel'));
}

function resolveMuridCols_(sheet) {
  const cols = getColumnMap_(sheet, MURID_HEADERS);
  const lastCol = Math.max(sheet.getLastColumn(), 4);
  const row1 = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  let noCol = cols['No'];
  let namaCol = cols['Nama Murid'];
  let nisCol = cols['NIS/NISN'];
  let kelasCol = cols['Kelas'];

  for (var i = 0; i < row1.length; i++) {
    const lbl = String(row1[i] || '').trim().toLowerCase();
    if (!lbl) continue;
    if (lbl === 'no') noCol = i + 1;
    else if (['nama murid', 'nama siswa', 'nama lengkap', 'nama'].indexOf(lbl) !== -1) namaCol = i + 1;
    else if (['nis/nisn', 'nis / nisn', 'nisn', 'nis', 'no peserta', 'nis / no. peserta', 'nis / peserta'].indexOf(lbl) !== -1) nisCol = i + 1;
    else if (['kelas', 'nama kelas', 'rombel'].indexOf(lbl) !== -1) kelasCol = i + 1;
  }
  return { no: noCol, nama: namaCol, nis: nisCol, kelas: kelasCol };
}

function readMasterMuridRows_(ss) {
  const sheet = ensureSheetWithHeaders_(ss, SHEET_MURID, MURID_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const cols = resolveMuridCols_(sheet);
  const maxCol = Math.max(cols.no, cols.nama, cols.nis, cols.kelas, 4);
  const disp = sheet.getRange(2, 1, lastRow - 1, maxCol).getDisplayValues();
  const list = [];
  for (var i = 0; i < disp.length; i++) {
    const no = String(disp[i][cols.no - 1] || '').trim();
    const namaMurid = String(disp[i][cols.nama - 1] || '').trim();
    const nis = String(disp[i][cols.nis - 1] || '').trim();
    const kelas = String(disp[i][cols.kelas - 1] || '').trim();
    if (!namaMurid && !nis && !kelas) continue;
    list.push({
      rowIndex: i + 2,
      no: no || String(list.length + 1),
      namaMurid: namaMurid,
      nis: nis,
      kelas: kelas
    });
  }
  return list;
}

function deriveUniqueClassesFromMurid_(muridRows) {
  const seen = {};
  const out = [];
  (muridRows || []).forEach(function (m) {
    const k = String(m.kelas || '').trim();
    if (k && !seen[k]) {
      seen[k] = true;
      out.push({ kelas: k });
    }
  });
  return out;
}

function readMasterRowsBySpec_(ss, spec) {
  const sheet = ensureSheetWithHeaders_(ss, spec.sheetName, spec.headers);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const cols = getColumnMap_(sheet, spec.headers);
  const valCol = resolveMasterValueCol_(sheet, cols, spec);
  const maxCol = Math.max(cols['No'], valCol, 2);
  const disp = sheet.getRange(2, 1, lastRow - 1, maxCol).getDisplayValues();
  const list = [];
  for (var i = 0; i < disp.length; i++) {
    const no = String(disp[i][cols['No'] - 1] || '').trim();
    const val = String(disp[i][valCol - 1] || '').trim();
    if (!val) continue;
    const item = { rowIndex: i + 2, no: no || String(list.length + 1) };
    item[spec.propName] = val;
    list.push(item);
  }
  return list;
}

function readExamRows_(ss, tz) {
  const sheet = ensureSheetWithHeaders_(ss, SHEET_DB, DB_HEADERS);
  ensureDbDurationColumnInSheet_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const cols = getColumnMap_(sheet, DB_HEADERS);
  const maxCol = Math.max.apply(null, Object.keys(cols).map(function (k) { return cols[k]; }));
  const range = sheet.getRange(2, 1, lastRow - 1, maxCol);
  const rawValues = range.getValues();
  const dispValues = range.getDisplayValues();

  const items = [];
  for (var i = 0; i < rawValues.length; i++) {
    const raw = rawValues[i];
    const disp = dispValues[i];

    const noDisp = String(disp[cols['No'] - 1] || '').trim();
    const mapel = String(disp[cols['Mapel'] - 1] || '').trim();
    const kelas = String(disp[cols['Kelas'] - 1] || '').trim();
    const link = String(disp[cols['Link GFrom'] - 1] || '').trim();
    const tanggal = normalizeSheetDate_(raw[cols['Tanggal'] - 1], disp[cols['Tanggal'] - 1], tz);
    const waktuMulai = normalizeSheetTime_(raw[cols['Waktu Mulai'] - 1], disp[cols['Waktu Mulai'] - 1]);
    const waktuSelesai = normalizeSheetTime_(raw[cols['Waktu Selesai'] - 1], disp[cols['Waktu Selesai'] - 1]);
    const durasiMenit = parseDurationMinutes_(raw[cols['Durasi (Menit)'] - 1], disp[cols['Durasi (Menit)'] - 1]);

    if (!mapel && !kelas && !link && !tanggal && !waktuMulai && !waktuSelesai && !durasiMenit) {
      continue;
    }

    items.push({
      rowIndex: i + 2,
      no: noDisp || String(items.length + 1),
      mapel: mapel,
      kelas: kelas,
      link: link,
      tanggal: tanggal,
      waktuMulai: waktuMulai,
      waktuSelesai: waktuSelesai,
      durasiMenit: durasiMenit
    });
  }
  return items;
}

function readMonitoringRows_(ss) {
  const sheet = ensureSheetWithHeaders_(ss, SHEET_MONITORING, MONITORING_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const cols = getColumnMap_(sheet, MONITORING_HEADERS);
  const maxCol = Math.max.apply(null, Object.keys(cols).map(function (k) { return cols[k]; }));
  const dispValues = sheet.getRange(2, 1, lastRow - 1, maxCol).getDisplayValues();

  const list = [];
  for (var i = dispValues.length - 1; i >= 0; i--) {
    const r = dispValues[i];
    const sessionId = String(r[cols['ID Sesi'] - 1] || '').trim();
    const namaSiswa = String(r[cols['Nama Siswa'] - 1] || '').trim();
    if (!sessionId && !namaSiswa) continue;

    list.push({
      rowIndex: i + 2,
      sessionId: sessionId,
      waktuAkses: String(r[cols['Waktu Akses'] - 1] || '').trim(),
      namaSiswa: namaSiswa,
      nis: String(r[cols['NIS / Peserta'] - 1] || '').trim(),
      kelasSiswa: String(r[cols['Kelas Siswa'] - 1] || '').trim(),
      noSoal: String(r[cols['No Soal'] - 1] || '').trim(),
      mapel: String(r[cols['Mapel'] - 1] || '').trim(),
      kelasJadwal: String(r[cols['Kelas Jadwal'] - 1] || '').trim(),
      tanggalUjian: String(r[cols['Tanggal Ujian'] - 1] || '').trim(),
      status: String(r[cols['Status'] - 1] || '').trim(),
      pelanggaran: parseInt(r[cols['Jumlah Pelanggaran'] - 1], 10) || 0,
      catatanPelanggaran: String(r[cols['Catatan Pelanggaran'] - 1] || '').trim(),
      waktuSelesai: String(r[cols['Waktu Selesai'] - 1] || '').trim()
    });
  }
  return list;
}

function findMonitoringRowBySession_(sheet, sessionCol, sessionId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const colRange = sheet.getRange(2, sessionCol, lastRow - 1, 1);
  if (typeof colRange.createTextFinder === 'function') {
    const found = colRange.createTextFinder(sessionId).matchEntireCell(true).findNext();
    return found ? found.getRow() : null;
  }
  const vals = colRange.getDisplayValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() === sessionId) return i + 2;
  }
  return null;
}

function normalizeSheetDate_(rawVal, dispVal, tz) {
  const disp = String(dispVal || '').trim();
  if (!disp && !(rawVal instanceof Date)) return '';
  if (rawVal instanceof Date && !isNaN(rawVal.getTime()) && rawVal.getFullYear() >= 1970) {
    return Utilities.formatDate(rawVal, tz, 'yyyy-MM-dd');
  }
  let m = disp.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (m) return m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]);
  m = disp.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (m) return m[3] + '-' + pad2_(m[2]) + '-' + pad2_(m[1]);
  return disp;
}

function normalizeSheetTime_(rawVal, dispVal) {
  const disp = String(dispVal || '').trim();
  if (!disp && !(rawVal instanceof Date)) return '';
  const normalizedDisp = disp.replace(/\./g, ':');
  let m = normalizedDisp.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i);
  if (m) {
    let hour = parseInt(m[1], 10);
    const min = parseInt(m[2], 10);
    const ampm = m[3] ? m[3].toUpperCase() : '';
    if (ampm === 'PM' && hour < 12) hour += 12;
    if (ampm === 'AM' && hour === 12) hour = 0;
    if (hour >= 0 && hour <= 23 && min >= 0 && min <= 59) {
      return pad2_(hour) + ':' + pad2_(min);
    }
  }
  if (rawVal instanceof Date && !isNaN(rawVal.getTime())) {
    return pad2_(rawVal.getHours()) + ':' + pad2_(rawVal.getMinutes());
  }
  return disp;
}

function evaluateExamSchedule_(row, now, tz) {
  if (!row.tanggal && !row.waktuMulai && !row.waktuSelesai) {
    return { status: 'active', statusLabel: 'Tersedia', startMs: null, endMs: null };
  }
  const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const nowMinutes = timeStringToMinutes_(Utilities.formatDate(now, tz, 'HH:mm'));
  const dateStr = row.tanggal || todayStr;
  const startStr = row.waktuMulai || '00:00';
  const endStr = row.waktuSelesai || '23:59';
  const startMin = timeStringToMinutes_(startStr);
  const endMin = timeStringToMinutes_(endStr);
  const startMs = parseDateTimeInTzMs_(dateStr, startStr, tz);
  const endMs = parseDateTimeInTzMs_(dateStr, endStr, tz);

  if (dateStr > todayStr) return { status: 'upcoming', statusLabel: 'Belum Dimulai', startMs: startMs, endMs: endMs };
  if (dateStr < todayStr) return { status: 'ended', statusLabel: 'Sudah Berakhir', startMs: startMs, endMs: endMs };
  if (startMin !== null && nowMinutes < startMin) return { status: 'upcoming', statusLabel: 'Belum Dimulai', startMs: startMs, endMs: endMs };
  if (endMin !== null && nowMinutes > endMin) return { status: 'ended', statusLabel: 'Sudah Berakhir', startMs: startMs, endMs: endMs };
  return { status: 'active', statusLabel: 'Sedang Berlangsung', startMs: startMs, endMs: endMs };
}

function parseDateTimeInTzMs_(dateStr, timeStr, tz) {
  const dm = String(dateStr || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const tm = String(timeStr || '').match(/^(\d{2}):(\d{2})$/);
  if (!dm || !tm) return null;
  const offset = Utilities.formatDate(new Date(), tz, 'XXX');
  const iso = dm[1] + '-' + dm[2] + '-' + dm[3] + 'T' + tm[1] + ':' + tm[2] + ':00' + offset;
  const parsed = new Date(iso);
  return isNaN(parsed.getTime()) ? null : parsed.getTime();
}

function validateExamPayload_(payload) {
  if (!payload || typeof payload !== 'object') {
    return { ok: false, message: 'Data kiriman tidak valid.' };
  }
  const mapel = String(payload.mapel || '').trim();
  const kelas = String(payload.kelas || '').trim();
  let link = String(payload.link || '').trim();
  if (link && !/^https?:\/\//i.test(link) && /^[a-z0-9.-]+\.[a-z]{2,}/i.test(link)) {
    link = 'https://' + link;
  }
  const tanggal = normalizeSheetDate_(null, payload.tanggal, 'Asia/Singapore');
  const waktuMulai = normalizeSheetTime_(null, payload.waktuMulai);
  const waktuSelesai = normalizeSheetTime_(null, payload.waktuSelesai);
  const rawDurasiInput = payload.durasiMenit !== undefined ? payload.durasiMenit : (payload.durasi !== undefined ? payload.durasi : '');
  const durasiText = String(rawDurasiInput === null || rawDurasiInput === undefined ? '' : rawDurasiInput).trim();
  let durasiMenit = 0;
  if (durasiText !== '' && durasiText !== '0') {
    if (/^-/.test(durasiText)) {
      return { ok: false, message: 'Durasi (Kolom H) tidak boleh bernilai negatif.' };
    }
    durasiMenit = parseDurationMinutes_(rawDurasiInput, durasiText);
    if (!durasiMenit) {
      return { ok: false, message: 'Durasi (Kolom H) harus berupa angka menit positif (misal: 30), atau kosongkan jika fleksibel.' };
    }
    if (durasiMenit > 1440) {
      return { ok: false, message: 'Durasi maksimal adalah 1440 menit (24 jam).' };
    }
  }
  const no = String(payload.no || '').trim();
  const rowIndex = payload.rowIndex ? Number(payload.rowIndex) : null;

  if (!mapel) return { ok: false, message: 'Nama Mata Pelajaran (Mapel) wajib diisi.' };
  if (!kelas) return { ok: false, message: 'Kelas wajib diisi.' };
  if (!link || !/^https?:\/\/.+/i.test(link)) return { ok: false, message: 'Link GFrom wajib berupa URL valid (https://...).' };
  if (tanggal && !/^\d{4}-\d{2}-\d{2}$/.test(tanggal)) return { ok: false, message: 'Format Tanggal harus YYYY-MM-DD.' };
  if (waktuMulai && !/^\d{2}:\d{2}$/.test(waktuMulai)) return { ok: false, message: 'Format Waktu Mulai harus HH:MM.' };
  if (waktuSelesai && !/^\d{2}:\d{2}$/.test(waktuSelesai)) return { ok: false, message: 'Format Waktu Selesai harus HH:MM.' };
  if (waktuMulai && waktuSelesai && timeStringToMinutes_(waktuSelesai) <= timeStringToMinutes_(waktuMulai)) {
    return { ok: false, message: 'Waktu Selesai harus lebih akhir daripada Waktu Mulai.' };
  }

  return {
    ok: true,
    data: {
      rowIndex: rowIndex,
      no: no,
      mapel: mapel,
      kelas: kelas,
      link: link,
      tanggal: tanggal,
      waktuMulai: waktuMulai,
      waktuSelesai: waktuSelesai,
      durasiMenit: durasiMenit
    }
  };
}

function writeExamRowToSheet_(sheet, rowIndex, cols, data) {
  const noStr = String(data.no || '').trim();
  const noVal = (noStr !== '' && !isNaN(Number(noStr))) ? Number(noStr) : noStr;
  sheet.getRange(rowIndex, cols['No']).setValue(noVal);
  sheet.getRange(rowIndex, cols['Mapel']).setNumberFormat('@').setValue(data.mapel);
  sheet.getRange(rowIndex, cols['Kelas']).setNumberFormat('@').setValue(data.kelas);
  sheet.getRange(rowIndex, cols['Link GFrom']).setNumberFormat('@').setValue(data.link);
  sheet.getRange(rowIndex, cols['Tanggal']).setNumberFormat('@').setValue(data.tanggal);
  sheet.getRange(rowIndex, cols['Waktu Mulai']).setNumberFormat('@').setValue(data.waktuMulai);
  sheet.getRange(rowIndex, cols['Waktu Selesai']).setNumberFormat('@').setValue(data.waktuSelesai);
  ensureHeaderCell_(sheet, cols['Durasi (Menit)'], 'Durasi (Menit)');
  if (data.durasiMenit && data.durasiMenit > 0) {
    sheet.getRange(rowIndex, cols['Durasi (Menit)']).setNumberFormat('0').setValue(data.durasiMenit);
  } else {
    sheet.getRange(rowIndex, cols['Durasi (Menit)']).setNumberFormat('@').setValue('');
  }
}

function parseDurationMinutes_(rawVal, dispVal) {
  if (typeof rawVal === 'number' && !isNaN(rawVal) && isFinite(rawVal)) {
    const mins = Math.round(rawVal);
    return mins > 0 ? mins : 0;
  }
  const disp = String(dispVal !== undefined ? dispVal : (rawVal != null ? rawVal : '')).trim();
  if (!disp) return 0;
  const m = disp.match(/(\d+(?:[.,]\d+)?)/);
  if (!m) return 0;
  const num = Math.round(Number(m[1].replace(',', '.')));
  return (!isNaN(num) && isFinite(num) && num > 0) ? num : 0;
}

function findNextAppendRow_(sheet, minCols) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;
  const maxCol = Math.max(sheet.getLastColumn(), minCols || DB_HEADERS.length);
  const values = sheet.getRange(2, 1, lastRow - 1, maxCol).getDisplayValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i].some(function (cell) { return String(cell).trim() !== ''; })) return i + 3;
  }
  return 2;
}

function computeNextNumber_(sheet, noCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return '1';
  const vals = sheet.getRange(2, noCol, lastRow - 1, 1).getValues();
  let maxNo = 0;
  vals.forEach(function (r) {
    const n = parseInt(r[0], 10);
    if (!isNaN(n) && n > maxNo) maxNo = n;
  });
  return String(maxNo + 1);
}

function renumberNonFormulaNo_(sheet, noCol, minCols) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, noCol, lastRow - 1, 1);
  const formulas = range.getFormulas();
  const maxCol = Math.max(sheet.getLastColumn(), minCols || DB_HEADERS.length);
  const allRows = sheet.getRange(2, 1, lastRow - 1, maxCol).getDisplayValues();
  let seq = 1;
  for (var i = 0; i < allRows.length; i++) {
    const hasData = allRows[i].slice(1).some(function (c) { return String(c).trim() !== ''; });
    if (hasData && !formulas[i][0]) {
      sheet.getRange(i + 2, noCol).setValue(seq++);
    }
  }
}

function resolveEmbeddableFormUrl_(rawUrl) {
  let url = String(rawUrl || '').trim();
  if (!url) return { ok: false, message: 'URL Google Form masih kosong.' };
  if (!/^https?:\/\//i.test(url)) url = 'https://' + url;
  if (!/^https?:\/\/.+\..+/i.test(url)) return { ok: false, message: 'Format URL Google Form tidak valid.' };
  if (/^https?:\/\/forms\.gle\//i.test(url)) url = followRedirects_(url);
  if (/docs\.google\.com\/forms\/d\/[^/]+\/edit/i.test(url)) {
    url = url.replace(/\/edit(\?.*)?$/i, '/viewform');
  }
  if (/docs\.google\.com\/forms\//i.test(url) && url.indexOf('embedded=true') === -1) {
    url += (url.indexOf('?') === -1 ? '?' : '&') + 'embedded=true';
  }
  return { ok: true, embedUrl: url };
}

function followRedirects_(initialUrl) {
  let currentUrl = initialUrl;
  for (var hop = 0; hop < 4; hop++) {
    try {
      const resp = UrlFetchApp.fetch(currentUrl, { followRedirects: false, muteHttpExceptions: true });
      const code = resp.getResponseCode();
      if (code >= 300 && code < 400) {
        const headers = resp.getHeaders();
        const loc = headers.Location || headers.location;
        if (!loc) break;
        currentUrl = loc;
      } else break;
    } catch (e) { break; }
  }
  return currentUrl;
}

function extractDriveFileId_(url) {
  const s = String(url || '').trim();
  if (!s) return '';
  let m = s.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = s.match(/lh3\.googleusercontent\.com\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  return '';
}

function extractDriveIdAny_(url) {
  const s = String(url || '').trim();
  if (!s) return '';
  let m = s.match(/(?:id=|file\/d\/|folders\/)([a-zA-Z0-9_-]{15,})/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{15,}$/.test(s)) return s;
  return '';
}

function toDriveImageUrl_(url) {
  const id = extractDriveFileId_(url);
  return id ? ('https://lh3.googleusercontent.com/d/' + id) : String(url || '').trim();
}

function toDriveThumbnailUrl_(url) {
  const id = extractDriveFileId_(url);
  return id ? ('https://drive.google.com/thumbnail?id=' + id + '&sz=w1200') : String(url || '').trim();
}

function formatTanggalIndonesia_(isoDate) {
  const m = String(isoDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return isoDate || 'Fleksibel';
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  return parseInt(m[3], 10) + ' ' + (months[parseInt(m[2], 10) - 1] || m[2]) + ' ' + m[1];
}

function formatHariTanggalIndonesia_(isoDate) {
  const m = String(isoDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return 'Jadwal Fleksibel (Setiap Hari)';
  const dt = new Date(Date.UTC(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10)));
  const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  return days[dt.getUTCDay()] + ', ' + parseInt(m[3], 10) + ' ' + months[dt.getUTCMonth()] + ' ' + m[1];
}

function timeStringToMinutes_(hhmm) {
  const m = String(hhmm || '').match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
}

function pad2_(val) {
  const n = parseInt(val, 10);
  return n < 10 ? '0' + n : String(n);
}

function getStoredAdminPin_() {
  const stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  return stored ? String(stored).trim() : DEFAULT_ADMIN_PIN;
}

function hasCustomAdminPin_() {
  const stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  return Boolean(stored && String(stored).trim() !== DEFAULT_ADMIN_PIN);
}

function getSecurityConfig_() {
  const props = PropertiesService.getScriptProperties();
  const maxViolRaw = props.getProperty('MAX_VIOLATIONS');
  const reqFsRaw = props.getProperty('REQUIRE_FULLSCREEN');
  const maxViol = maxViolRaw !== null ? parseInt(maxViolRaw, 10) : DEFAULT_MAX_VIOLATIONS;
  return {
    maxViolations: isNaN(maxViol) ? DEFAULT_MAX_VIOLATIONS : maxViol,
    requireFullscreen: reqFsRaw === null ? true : reqFsRaw !== 'false'
  };
}

function getTokenRole_(token) {
  const tok = String(token || '').trim();
  if (!tok) return '';
  const val = CacheService.getScriptCache().get('admin_tok_' + tok);
  if (val === 'Admin' || val === '1') return 'Admin';
  if (val === 'Pengawas') return 'Pengawas';
  return '';
}

function isValidAdminToken_(token) {
  return getTokenRole_(token) === 'Admin';
}

function isValidStaffToken_(token) {
  const r = getTokenRole_(token);
  return r === 'Admin' || r === 'Pengawas';
}
