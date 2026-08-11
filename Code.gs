// ============================================================
// APLIKASI PENDATAAN DUSUN 1 - DESA WARNASARI
// VERSI FINAL - DENGAN SEMUA PERBAIKAN KEAMANAN & FITUR
// ============================================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Pendataan Dusun 1')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Dusun 1 - Tools')
      .addItem('Perbaiki Header & Kolom Sheet (Auto Repair)', 'repairAllSheetHeadersAndData')
      .addItem('Perbaiki Data RW (jalankan sekali)', 'repairRWData')
      .addItem('Reset Users (buat ulang semua user)', 'resetUsers')
      .addItem('Migrasi Password ke Hash (jalankan sekali)', 'migrateToHash')
      .addItem('Lihat Audit Log', 'showAuditLog')
      .addToUi();
  } catch(e){}
}

// ============================================================
// KONSTANTA KEAMANAN
// ============================================================

var SECURITY = {
  MAX_LOGIN_ATTEMPTS: 5,
  LOCKOUT_DURATION: 15,
  SESSION_TIMEOUT: 30,
  SESSION_CACHE_PREFIX: 'sess_',
  HASH_ALGORITHM: Utilities.DigestAlgorithm.SHA_256
};

// ============================================================
// KONSTANTA GLOBAL
// ============================================================

// Daftar RW tunggal — digunakan oleh generatePDF_Kadus, getRWStats, getDropdown
var RW_LIST = ['02', '03', '04', '12', '13', '15'];

// Title map untuk PDF dan Empty page — digunakan bersama
var SHEET_TITLE_MAP = {
  'DataWarga':      'DATA WARGA',
  'DataDatang':     'DATA PENDUDUK DATANG',
  'DataPergi':      'DATA PENDUDUK PERGI',
  'DataMeninggal':  'DATA PENDUDUK MENINGGAL',
  'DataLahir':      'DATA KELAHIRAN',
  'DataPengontrak': 'DATA PENGONTRAK / PENDUDUK SEMENTARA',
  'LaporanBulanan': 'LAPORAN BULANAN'
};

// ============================================================
// HELPER PDF — CSS, HEADER, FOOTER (menghilangkan duplikasi)
// ============================================================

/**
 * Mengembalikan string CSS standar untuk semua PDF laporan.
 * @param {string} [pageSize] - 'landscape' atau 'portrait'. Default 'landscape'.
 */
function _buildPdfStyles(pageSize) {
  var ps = pageSize || 'landscape';
  var css = '';
  css += '@page { size: A4 ' + ps + '; margin: 12mm 10mm 12mm 10mm; }';
  css += '* { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color-adjust: exact !important; }';
  css += 'html, body { width: 100%; margin: 0; padding: 0; background: #ffffff; }';
  css += 'body { font-family: "Times New Roman", Arial, sans-serif; font-size: 10px; padding: 14px 18px; box-sizing: border-box; }';
  css += '.header { text-align: center; margin-bottom: 10px; }';
  css += '.header h1 { font-size: 16px; font-weight: bold; margin: 0; letter-spacing: 1px; color: #000000; }';
  css += '.header h2 { font-size: 13px; font-weight: bold; margin: 3px 0; }';
  css += '.header .periode { font-size: 11px; font-weight: bold; margin: 3px 0; }';
  css += '.total-box { border: 1px solid #1e3a5f; background-color: #f1f5f9 !important; padding: 6px 10px; margin: 8px 0; text-align: center; font-weight: bold; width: 100%; box-sizing: border-box; -webkit-print-color-adjust: exact !important; }';
  css += '.total-box span { margin: 0 8px; font-size: 10px; display: inline-block; }';
  css += 'table { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 9.5px; }';
  css += 'th { background-color: #1e3a5f !important; color: #ffffff !important; padding: 6px 4px; border: 1px solid #1e3a5f; text-align: center; font-weight: bold; font-size: 9.5px; -webkit-print-color-adjust: exact !important; }';
  css += 'td { padding: 5px 4px; border: 1px solid #000; text-align: center; font-size: 9.5px; }';
  css += 'tr:nth-child(even) { background-color: #f8fafc !important; -webkit-print-color-adjust: exact !important; }';
  css += 'tr.total-row { background-color: #e2e8f0 !important; font-weight: bold; -webkit-print-color-adjust: exact !important; }';
  css += '@media screen and (max-width: 600px) {';
  css += '  body { padding: 8px 6px; }';
  css += '  .header h1 { font-size: 13px; }';
  css += '  .header h2 { font-size: 11px; }';
  css += '  .header .periode { font-size: 9.5px; }';
  css += '  .total-box { padding: 4px 2px; font-size: 8px; margin: 4px 0; }';
  css += '  .total-box span { margin: 0 3px; font-size: 8px; }';
  css += '  table { font-size: 7.5px; margin-top: 4px; }';
  css += '  th { padding: 3px 1px; font-size: 7.5px; }';
  css += '  td { padding: 3px 1px; font-size: 7.5px; }';
  css += '}';
  css += '@media print { .no-print { display: none; } }';
  return css;
}

/**
 * Mengembalikan atribut inline style untuk <th> di PDF (menghilangkan var thS/thStyle duplikat).
 */
function _thStyle() {
  return 'bgcolor="#1e3a5f" style="background-color:#1e3a5f !important;color:#ffffff !important;-webkit-print-color-adjust:exact !important;"';
}

/**
 * Membangun scope title berdasarkan rwVal dan rtVal.
 * Contoh: 'DUSUN 1 RW 02 RT 3' atau 'DUSUN 1 RW 02' atau 'DUSUN 1'
 */
function _buildScopeTitle(rwVal, rtVal) {
  if (rtVal) {
    var formattedRW = String(rwVal).length === 1 ? '0' + rwVal : String(rwVal);
    return 'DUSUN 1 RW ' + formattedRW + ' RT ' + rtVal;
  } else if (rwVal) {
    var formattedRW2 = String(rwVal).length === 1 ? '0' + rwVal : String(rwVal);
    return 'DUSUN 1 RW ' + formattedRW2;
  }
  return 'DUSUN 1';
}

/**
 * Membangun teks filter dari objek filter untuk ditampilkan di PDF.
 */
function _buildFilterText(filter) {
  var text = '';
  if (!filter) return text;
  if (filter.rw)     text += ' RW ' + escapeHtmlServer(filter.rw);
  if (filter.rt)     text += ' RT ' + escapeHtmlServer(filter.rt);
  if (filter.bulan)  text += ' Bulan ' + escapeHtmlServer(getBulanName2(filter.bulan));
  if (filter.tahun)  text += ' ' + escapeHtmlServer(filter.tahun);
  if (filter.search) text += ' Cari: ' + escapeHtmlServer(filter.search);
  return text.trim();
}

/**
 * Membangun HTML footer standar PDF (dicetak oleh, tanggal, copyright).
 */
function _buildPdfFooter(u, now) {
  var d = now || new Date();
  var blnNames = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  var tanggalStr = d.getDate() + ' ' + blnNames[d.getMonth()] + ' ' + d.getFullYear();
  var html = '<div style="margin-top:15px;padding-top:6px;border-top:1px solid #ccc;font-size:8.5px;color:#4b5563;display:flex;justify-content:space-between;align-items:center;">';
  html += '<span>Dicetak: ' + d.toLocaleString('id-ID') + ' | Desa Warnasari, ' + tanggalStr + '</span>';
  html += '<span><strong>&copy; 2026 KKN 06 Desa Warnasari Ikopin University</strong></span>';
  html += '</div>';
  return html;
}

/**
 * Memeriksa akses RT/RW terhadap data tertentu.
 * Mengembalikan null jika akses diizinkan, atau objek {action, message} jika ditolak.
 */
function _checkDataAccess(u, data, actionName) {
  if (!u) return { success: false, message: 'Sesi tidak valid.' };
  if (canAccessAllData(u)) return null;

  var dataRW = normalizeRW(data.RW || data.rw || '');
  var dataRT = normalizeRT(data.RT || data.rt || '');
  var userRW = normalizeRW(u.rw || '');
  var userRT = normalizeRT(u.rt || '');

  if (u.role === 'rt') {
    if (dataRW !== userRW || dataRT !== userRT) {
      writeAuditLog(u.username, u.role, actionName + '_FAILED', 'Akses ditolak: hanya RT ' + u.rt + ' (ditemukan RT ' + dataRT + ')');
      return { success: false, message: 'Akses Ditolak: Akun Anda hanya diizinkan mengelola data di RT ' + u.rt + ' (RW ' + u.rw + ').' };
    }
  } else if (u.role === 'rw') {
    if (dataRW !== userRW) {
      writeAuditLog(u.username, u.role, actionName + '_FAILED', 'Akses ditolak: hanya RW ' + u.rw);
      return { success: false, message: 'Akses Ditolak: Akun Anda hanya diizinkan mengelola data di RW ' + u.rw + '.' };
    }
  }
  return null;
}


// ============================================================
// SERVER-SIDE HTML ESCAPE (untuk PDF & Audit Log output)
// ============================================================

function escapeHtmlServer(text) {
  if (text === null || text === undefined) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;');
}

// ============================================================
// INPUT SANITIZATION (anti formula injection di spreadsheet)
// ============================================================

function sanitizeInput(value) {
  if (value === null || value === undefined) return '';
  var str = String(value).trim();
  // Strip formula prefixes yang bisa di-execute oleh spreadsheet
  if (/^[=+\-@]/.test(str)) {
    str = "'" + str;
  }
  return str;
}

function sanitizeDataObject(data, skipFields) {
  var skip = skipFields || [];
  var sanitized = {};
  for (var key in data) {
    if (data.hasOwnProperty(key)) {
      if (key === '_rowIndex' || skip.indexOf(key) !== -1) {
        sanitized[key] = data[key];
      } else {
        sanitized[key] = sanitizeInput(data[key]);
      }
    }
  }
  return sanitized;
}

// ============================================================
// VERIFIKASI USER (re-verify terhadap Users sheet)
// ============================================================

function verifyUser(u) {
  if (!u || !u.username || !u.role) return null;
  var users = getUsers();
  var real = users[u.username];
  if (!real) return null;
  if (real.role !== u.role) return null;
  return {
    username: u.username,
    role: real.role,
    dusun: real.dusun,
    rw: real.rw,
    rt: real.rt,
    label: real.label
  };
}

// ============================================================
// KOLOM YANG HARUS DIPAKSA JADI TEXT
// ============================================================

var TEXT_COLUMNS = {
  'DataWarga': ['No KK','NIK','RW','RT','No. Rumah'],
  'DataDatang': ['No KK','NIK','RW','RT','No. Rumah'],
  'DataPergi': ['No KK','NIK','RW','RT','No. Rumah'],
  'DataMeninggal': ['No KK','NIK','RW','RT','No. Rumah'],
  'DataLahir': ['No KK','RW','RT','No. Rumah'],
  'DataPengontrak': ['NIK','No HP','RW','RT','No. Rumah'],
  'LaporanBulanan': ['RW','RT'],
  'Users': ['RW','RT'],
  'Backup': ['RW','RT'],
  'AuditLog': [],
  'LoginAttempts': []
};

function applyTextFormatToColumn(sheet, headers, colNames, numRows) {
  try {
    var rows = numRows || Math.max(sheet.getMaxRows() - 1, 1);
    colNames.forEach(function(cn){
      var idx = headers.indexOf(cn);
      if(idx > -1){
        sheet.getRange(2, idx + 1, rows, 1).setNumberFormat('@');
      }
    });
  } catch(e){ console.error('applyTextFormatToColumn error:', e); }
}

function normalizeRW(rw) {
  var r = String(rw == null ? '' : rw).trim();
  if(r === '') return r;
  if(/^\d+$/.test(r) && r.length === 1) return '0' + r;
  return r;
}

function normalizeNIK(nik){
  return String(nik == null ? '' : nik).trim();
}

function normalizeRT(rt) {
  var r = String(rt == null ? '' : rt).trim();
  if (r === '') return r;
  if (/^\d+$/.test(r)) {
    return String(parseInt(r, 10));
  }
  return r;
}

function checkNikExist(nik, currentNik) {
  try {
    if (!nik || String(nik).trim() === '') return { exist: false };
    var normNik = normalizeNIK(nik);
    if (currentNik && normalizeNIK(currentNik) === normNik) {
      return { exist: false };
    }
    var existingWarga = getData('DataWarga');
    for (var i = 0; i < existingWarga.length; i++) {
      var w = existingWarga[i];
      if (normalizeNIK(w.NIK) === normNik) {
        return {
          exist: true,
          nama: w['Nama Lengkap'] || w.Nama || 'Warga lain',
          rt: w.RT,
          rw: w.RW,
          dusun: w.Dusun || 'Dusun 1',
          noKK: w['No. KK'] || w['No KK'] || ''
        };
      }
    }
    return { exist: false };
  } catch (e) {
    return { exist: false, error: e.message };
  }
}

function checkNameDobExist(nama, tglLahir, currentNik) {
  try {
    if (!nama || !tglLahir || String(nama).trim() === '') return { exist: false };
    var normNama = String(nama).trim().toUpperCase();
    var normTgl = formatDate(tglLahir);
    var normCurrentNik = currentNik ? normalizeNIK(currentNik) : '';

    var existingWarga = getData('DataWarga');
    for (var i = 0; i < existingWarga.length; i++) {
      var w = existingWarga[i];
      if (normCurrentNik && normalizeNIK(w.NIK) === normCurrentNik) continue;

      var wNama = String(w['Nama Lengkap'] || w.Nama || '').trim().toUpperCase();
      var wTgl = formatDate(w['Tanggal Lahir'] || w.TglLahir || '');

      if (wNama === normNama && wTgl === normTgl && wNama !== '') {
        return {
          exist: true,
          nama: w['Nama Lengkap'] || w.Nama,
          nik: w.NIK,
          rt: w.RT,
          rw: w.RW
        };
      }
    }
    return { exist: false };
  } catch (e) {
    return { exist: false, error: e.message };
  }
}

// ============================================================
// PASSWORD HASH DENGAN SALT
// ============================================================

function generateSalt() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()';
  var salt = '';
  for (var i = 0; i < 16; i++) {
    salt += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return salt;
}

function hashPasswordWithSalt(password, salt) {
  var salted = password + salt;
  var digest = Utilities.computeDigest(
    SECURITY.HASH_ALGORITHM,
    salted,
    Utilities.Charset.UTF_8
  );
  
  var hash = '';
  for (var i = 0; i < digest.length; i++) {
    var byte = digest[i];
    if (byte < 0) byte += 256;
    var hex = byte.toString(16);
    if (hex.length === 1) hex = '0' + hex;
    hash += hex;
  }
  return hash;
}

function verifyPasswordWithSalt(inputPassword, storedHash, salt) {
  var inputHash = hashPasswordWithSalt(inputPassword, salt);
  return inputHash === storedHash;
}

function generateRandomPassword() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%';
  var password = '';
  for (var i = 0; i < 12; i++) {
    password += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return password;
}

// ============================================================
// SESSION MANAGEMENT (CacheService — persisten antar request)
// ============================================================

function createSession(username) {
  var token = Utilities.getUuid();
  var cache = CacheService.getScriptCache();
  var sessionData = {
    username: username,
    created: new Date().toISOString(),
    lastActivity: new Date().toISOString()
  };
  // Cache expire sesuai session timeout (dalam detik)
  cache.put(SECURITY.SESSION_CACHE_PREFIX + token, JSON.stringify(sessionData), SECURITY.SESSION_TIMEOUT * 60);
  return token;
}

function validateSession(token) {
  if (!token) return false;
  var cache = CacheService.getScriptCache();
  var raw = cache.get(SECURITY.SESSION_CACHE_PREFIX + token);
  if (!raw) return false;
  try {
    var session = JSON.parse(raw);
    var now = new Date();
    var lastActivity = new Date(session.lastActivity);
    var diffMinutes = (now - lastActivity) / (1000 * 60);
    if (diffMinutes > SECURITY.SESSION_TIMEOUT) {
      cache.remove(SECURITY.SESSION_CACHE_PREFIX + token);
      return false;
    }
    // Perpanjang session
    session.lastActivity = now.toISOString();
    cache.put(SECURITY.SESSION_CACHE_PREFIX + token, JSON.stringify(session), SECURITY.SESSION_TIMEOUT * 60);
    return true;
  } catch(e) {
    return false;
  }
}

function extendSession(token) {
  if (!token) return false;
  var cache = CacheService.getScriptCache();
  var raw = cache.get(SECURITY.SESSION_CACHE_PREFIX + token);
  if (!raw) return false;
  try {
    var session = JSON.parse(raw);
    session.lastActivity = new Date().toISOString();
    cache.put(SECURITY.SESSION_CACHE_PREFIX + token, JSON.stringify(session), SECURITY.SESSION_TIMEOUT * 60);
    return true;
  } catch(e) {
    return false;
  }
}

function pingSession(username) {
  try {
    var users = getUsers();
    if (users[username]) {
      return {success: true};
    }
    return {success: false};
  } catch(e) {
    return {success: false};
  }
}

function endSession(token) {
  if (!token) return false;
  var cache = CacheService.getScriptCache();
  cache.remove(SECURITY.SESSION_CACHE_PREFIX + token);
  return true;
}

// ============================================================
// INISIALISASI SHEET
// ============================================================

function initAuditLogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('AuditLog');
  if(!sheet) {
    sheet = ss.insertSheet('AuditLog');
    var headers = ['Timestamp','Username','Role','Aksi','Detail','IP Address','Session ID'];
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.getRange(1,1,1,headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function initLoginAttemptsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('LoginAttempts');
  if(!sheet) {
    sheet = ss.insertSheet('LoginAttempts');
    var headers = ['Username','Attempts','LastAttempt','LockedUntil','IP Address'];
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.getRange(1,1,1,headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function initBackupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Backup');
  if(!sheet) {
    sheet = ss.insertSheet('Backup');
    var headers = ['Tanggal Hapus','Sheet Asal','No KK','NIK','Nama Lengkap','Dusun','RW','RT','Jenis Kelamin','Tempat Lahir','Tanggal Lahir','Status Perkawinan','Pekerjaan','Pendidikan','Agama','Tahun','Bulan','Data Lengkap JSON','Dihapus Oleh'];
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.getRange(1,1,1,headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange(2, 7, 1, 1).setNumberFormat('@');
    sheet.getRange(2, 8, 1, 1).setNumberFormat('@');
  }
  return sheet;
}

// ============================================================
// AUDIT LOG FUNGSI
// ============================================================

function writeAuditLog(username, role, aksi, detail, sessionId) {
  try {
    var sheet = initAuditLogSheet();
    var now = new Date();
    var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    var row = [timestamp, username || 'system', role || 'system', aksi, detail || '', 'Web App', sessionId || ''];
    sheet.appendRow(row);

    // Auto-pruning:jika log melebihi 20.000 baris, hapus 5.000 baris tertua agar spreadsheet tetap kencang & aman
    if (sheet.getLastRow() > 20000) {
      sheet.deleteRows(2, 5000);
    }
    return true;
  } catch(e) {
    console.error('Audit log error:', e);
    return false;
  }
}

function showAuditLog() {
  try {
    var sheet = initAuditLogSheet();
    var data = sheet.getDataRange().getValues();
    if(data.length < 2) {
      SpreadsheetApp.getUi().alert('Belum ada data audit log.');
      return;
    }
    
    var html = '<html><head><style>';
    html += 'body{font-family:Arial;padding:10px}';
    html += 'table{width:100%;border-collapse:collapse;font-size:11px}';
    html += 'th{background:#1e3a5f;color:#fff;padding:6px;text-align:left}';
    html += 'td{padding:4px;border-bottom:1px solid #ddd}';
    html += 'tr:hover{background:#f0f2f5}';
    html += '.close{background:#ef4444;color:#fff;border:none;padding:6px 16px;border-radius:4px;cursor:pointer;float:right}';
    html += '</style></head><body>';
    html += '<h2>📋 Audit Log</h2>';
    html += '<button class="close" onclick="google.script.host.close()">Tutup</button>';
    html += '<br><br>';
    html += '<table>';
    html += '<tr><th>Timestamp</th><th>Username</th><th>Role</th><th>Aksi</th><th>Detail</th></tr>';
    
    for(var i=1; i<Math.min(data.length, 101); i++) {
      var row = data[i];
      html += '<tr>';
      html += '<td>' + escapeHtmlServer(row[0]||'') + '</td>';
      html += '<td>' + escapeHtmlServer(row[1]||'') + '</td>';
      html += '<td>' + escapeHtmlServer(row[2]||'') + '</td>';
      html += '<td>' + escapeHtmlServer(row[3]||'') + '</td>';
      html += '<td>' + escapeHtmlServer(row[4]||'') + '</td>';
      html += '</tr>';
    }
    
    if(data.length > 101) {
      html += '<tr><td colspan="5" style="text-align:center;color:#999;">... dan ' + (data.length-101) + ' data lainnya. Lihat di sheet AuditLog.</td></tr>';
    }
    
    html += '</table>';
    html += '<br><p style="color:#999;font-size:10px;">Total: ' + (data.length-1) + ' log entries</p>';
    html += '</body></html>';
    
    var ui = SpreadsheetApp.getUi();
    var output = HtmlService.createHtmlOutput(html)
      .setWidth(900)
      .setHeight(500);
    ui.showModalDialog(output, 'Audit Log');
    
  } catch(e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
  }
}

// ============================================================
// PROTEKSI BRUTE FORCE
// ============================================================

function getLoginAttempts(username) {
  try {
    var sheet = initLoginAttemptsSheet();
    var data = sheet.getDataRange().getValues();
    
    for(var i=1; i<data.length; i++) {
      if(String(data[i][0]||'').trim() === username) {
        return {
          rowIndex: i+1,
          attempts: Number(data[i][1]||0),
          lastAttempt: data[i][2] ? new Date(data[i][2]) : null,
          lockedUntil: data[i][3] ? new Date(data[i][3]) : null
        };
      }
    }
    return null;
  } catch(e) {
    return null;
  }
}

function updateLoginAttempts(username, success) {
  try {
    var sheet = initLoginAttemptsSheet();
    var data = sheet.getDataRange().getValues();
    var now = new Date();
    var nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    
    var rowIdx = -1;
    for(var i=1; i<data.length; i++) {
      if(String(data[i][0]||'').trim() === username) {
        rowIdx = i+1;
        break;
      }
    }
    
    if(success) {
      if(rowIdx === -1) {
        sheet.appendRow([username, 0, nowStr, '', '']);
      } else {
        sheet.getRange(rowIdx, 2).setValue(0);
        sheet.getRange(rowIdx, 3).setValue(nowStr);
        sheet.getRange(rowIdx, 4).setValue('');
        sheet.getRange(rowIdx, 5).setValue('');
      }
      return;
    }
    
    var attempts = 1;
    var lockedUntil = '';
    
    if(rowIdx !== -1) {
      attempts = Number(data[rowIdx-1][1]||0) + 1;
      
      if(attempts >= SECURITY.MAX_LOGIN_ATTEMPTS) {
        var lockTime = new Date(now.getTime() + (SECURITY.LOCKOUT_DURATION * 60 * 1000));
        lockedUntil = Utilities.formatDate(lockTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      }
      
      sheet.getRange(rowIdx, 2).setValue(attempts);
      sheet.getRange(rowIdx, 3).setValue(nowStr);
      sheet.getRange(rowIdx, 4).setValue(lockedUntil);
    } else {
      if(attempts >= SECURITY.MAX_LOGIN_ATTEMPTS) {
        var lockTime2 = new Date(now.getTime() + (SECURITY.LOCKOUT_DURATION * 60 * 1000));
        lockedUntil = Utilities.formatDate(lockTime2, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      }
      sheet.appendRow([username, attempts, nowStr, lockedUntil, '']);
    }
  } catch(e) {
    console.error('Update login attempts error:', e);
  }
}

function isUserLocked(username) {
  try {
    var attempts = getLoginAttempts(username);
    if(!attempts || !attempts.lockedUntil) return false;
    
    var now = new Date();
    if(now < attempts.lockedUntil) {
      return true;
    }
    
    var sheet = initLoginAttemptsSheet();
    sheet.getRange(attempts.rowIndex, 2).setValue(0);
    sheet.getRange(attempts.rowIndex, 4).setValue('');
    return false;
  } catch(e) {
    return false;
  }
}

// ============================================================
// BACKUP DATA
// ============================================================

function backupData(sheetName, rowData, deletedBy) {
  try {
    var sheet = initBackupSheet();
    var now = new Date();
    var tanggalHapus = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    
    var row = [
      tanggalHapus,
      sheetName,
      rowData['No KK'] || '',
      rowData['NIK'] || '',
      rowData['Nama Lengkap'] || rowData['Nama Bayi'] || '',
      rowData['Dusun'] || '',
      normalizeRW(rowData['RW'] || ''),
      String(rowData['RT'] || ''),
      rowData['Jenis Kelamin'] || '',
      rowData['Tempat Lahir'] || '',
      rowData['Tanggal Lahir'] || '',
      rowData['Status Perkawinan'] || '',
      rowData['Pekerjaan'] || '',
      rowData['Pendidikan'] || '',
      rowData['Agama'] || '',
      rowData['Tahun'] || '',
      rowData['Bulan'] || '',
      JSON.stringify(rowData),
      deletedBy || 'system'
    ];
    
    sheet.appendRow(row);
    
    var lastRow = sheet.getLastRow();
    if(lastRow > 1) {
      sheet.getRange(lastRow, 7, 1, 1).setNumberFormat('@');
      sheet.getRange(lastRow, 8, 1, 1).setNumberFormat('@');
    }
    
    return true;
  } catch(e) {
    console.error('Backup error:', e);
    return false;
  }
}

function getBackupData() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Backup');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    var headers = data[0];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var hasContent = row.some(function(c){ return String(c).trim()!==''; });
      if(!hasContent) continue;
      var obj = {};
      headers.forEach(function(h, idx) {
        obj[h] = String(row[idx] || '');
      });
      obj._rowIndex = i - 1;
      result.push(obj);
    }
    return result;
  } catch(e) {
    return [];
  }
}

function restoreBackupData(rowIndex, u) {
  try {
    u = verifyUser(u);
    if (!u || u.role !== 'superadmin') {
      return {success: false, message: 'Hanya SuperAdmin yang bisa restore backup'};
    }
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Backup');
    if (!sheet) return {success: false, message: 'Sheet Backup tidak ditemukan'};
    
    var data = sheet.getDataRange().getValues();
    if (rowIndex + 1 >= data.length) {
      return {success: false, message: 'Data backup tidak ditemukan'};
    }
    
    var row = data[rowIndex + 1];
    var jsonData = JSON.parse(row[17]);
    var sheetName = row[1];
    
    var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!targetSheet) {
      return {success: false, message: 'Sheet asal tidak ditemukan: ' + sheetName};
    }
    
    var headers = getHeaders(sheetName);
    var newRow = headers.map(function(h) {
      return jsonData[h] || '';
    });
    targetSheet.appendRow(newRow);
    
    sheet.deleteRow(rowIndex + 2);
    
    writeAuditLog(u.username, u.role, 'RESTORE_BACKUP', 'Restore data dari backup ke ' + sheetName);
    return {success: true, message: 'Data berhasil direstore'};
  } catch(e) {
    writeAuditLog(u ? u.username : 'unknown', u ? u.role : 'unknown', 'RESTORE_BACKUP_ERROR', e.message);
    return {success: false, message: e.message};
  }
}

// ============================================================
// USER & PASSWORD - DENGAN SALT
// ============================================================

function getUsers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    var h = ['Username','Password','Role','Dusun','RW','RT','Label','Salt'];
    sheet.getRange(1,1,1,h.length).setValues([h]);
    sheet.getRange(1,1,1,h.length).setFontWeight('bold');
    
    // ⚠️ PORTFOLIO NOTE: Default passwords below are placeholder examples only.
    // In production, replace with strong unique passwords before running initSetup().
    var superPassword = 'YOUR_SUPER_ADMIN_PASSWORD';

    // Format: [username, password, role, dusun, rw, rt, label]
    var defaultUsers = [
      ['superadmin', superPassword,              'superadmin', '',        '', '',  'Super Admin'],
      ['kadus1',     'YOUR_KADUS_PASSWORD',      'kadus',      'Dusun 1', '', '',  'Kepala Dusun 1'],
      ['rw13',       'YOUR_RW_PASSWORD',         'rw',         'Dusun 1', '13', '', 'RW 13'],
      ['rw04',       'YOUR_RW_PASSWORD',         'rw',         'Dusun 1', '04', '', 'RW 04'],
      ['rw02',       'YOUR_RW_PASSWORD',         'rw',         'Dusun 1', '02', '', 'RW 02'],
      ['rw03',       'YOUR_RW_PASSWORD',         'rw',         'Dusun 1', '03', '', 'RW 03'],
      ['rw12',       'YOUR_RW_PASSWORD',         'rw',         'Dusun 1', '12', '', 'RW 12'],
      ['rw15',       'YOUR_RW_PASSWORD',         'rw',         'Dusun 1', '15', '', 'RW 15'],
      ['rt1301',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '13', '1', 'RT 1 RW 13'],
      ['rt1302',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '13', '2', 'RT 2 RW 13'],
      ['rt1303',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '13', '3', 'RT 3 RW 13'],
      ['rt1304',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '13', '4', 'RT 4 RW 13'],
      ['rt1305',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '13', '5', 'RT 5 RW 13'],
      ['rt0401',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '04', '1', 'RT 1 RW 04'],
      ['rt0402',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '04', '2', 'RT 2 RW 04'],
      ['rt0403',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '04', '3', 'RT 3 RW 04'],
      ['rt0201',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '02', '1', 'RT 1 RW 02'],
      ['rt0202',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '02', '2', 'RT 2 RW 02'],
      ['rt0203',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '02', '3', 'RT 3 RW 02'],
      ['rt0204',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '02', '4', 'RT 4 RW 02'],
      ['rt0301',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '03', '1', 'RT 1 RW 03'],
      ['rt0302',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '03', '2', 'RT 2 RW 03'],
      ['rt1201',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '12', '1', 'RT 1 RW 12'],
      ['rt1202',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '12', '2', 'RT 2 RW 12'],
      ['rt1203',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '12', '3', 'RT 3 RW 12'],
      ['rt1501',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '15', '1', 'RT 1 RW 15'],
      ['rt1502',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '15', '2', 'RT 2 RW 15'],
      ['rt1503',     'YOUR_RT_PASSWORD',         'rt',         'Dusun 1', '15', '3', 'RT 3 RW 15']
    ];
    
    var d = [];
    var plainPasswords = [];
    defaultUsers.forEach(function(u) {
      var salt = generateSalt();
      var hashed = hashPasswordWithSalt(u[1], salt);
      d.push([u[0], hashed, u[2], u[3], u[4], u[5], u[6], salt]);
      plainPasswords.push({ username: u[0], password: u[1], label: u[6], role: u[2], rw: u[4], rt: u[5] });
    });
    
    sheet.getRange(2, 5, d.length, 1).setNumberFormat('@');
    sheet.getRange(2, 6, d.length, 1).setNumberFormat('@');
    sheet.getRange(2, 1, d.length, 8).setValues(d);
    
    console.log('🔐 Semua password user berpola telah digenerate dan dikirim via PDF email.');
    
    try {
      var email = '';
      try { email = Session.getEffectiveUser().getEmail(); } catch(e1) {}
      if (!email) { try { email = Session.getActiveUser().getEmail(); } catch(e2) {} }
      if (!email) { try { email = ss.getOwner().getEmail(); } catch(e3) {} }
      if (email) {
        // Buat Dokumen PDF HTML untuk Lampiran Email
        var pdfHtml = '<html><head><meta charset="UTF-8"><style>';
        pdfHtml += '@page { size: A4 portrait; margin: 12mm 10mm 12mm 10mm; }';
        pdfHtml += 'body { font-family: Arial, sans-serif; font-size: 11px; padding: 10px; color: #1f2937; }';
        pdfHtml += '.header { text-align: center; margin-bottom: 15px; border-bottom: 2px solid #1e3a5f; padding-bottom: 10px; }';
        pdfHtml += '.header h1 { font-size: 16px; color: #1e3a5f; margin: 0; font-weight: bold; }';
        pdfHtml += '.header h2 { font-size: 12px; color: #4b5563; margin: 4px 0 0 0; }';
        pdfHtml += '.info { background: #f1f5f9; border-left: 4px solid #1e3a5f; padding: 8px 12px; margin-bottom: 15px; font-size: 10px; line-height: 1.4; }';
        pdfHtml += 'table { width: 100%; border-collapse: collapse; margin-top: 10px; }';
        pdfHtml += 'th { background-color: #1e3a5f; color: #ffffff; padding: 8px 6px; border: 1px solid #1e3a5f; text-align: center; font-size: 10px; }';
        pdfHtml += 'td { padding: 6px; border: 1px solid #cbd5e1; text-align: center; font-size: 10px; }';
        pdfHtml += 'tr:nth-child(even) { background-color: #f8fafc; }';
        pdfHtml += '.pass-code { font-family: monospace; font-weight: bold; color: #0f766e; font-size: 11px; background: #ccfbf1; padding: 2px 6px; border-radius: 4px; }';
        pdfHtml += '.footer { margin-top: 20px; font-size: 9px; color: #64748b; text-align: center; border-top: 1px solid #e2e8f0; padding-top: 8px; }';
        pdfHtml += '</style></head><body>';

        pdfHtml += '<div class="header">';
        pdfHtml += '<h1>DAFTAR AKUN & PASSWORD PENGURUS</h1>';
        pdfHtml += '<h2>APLIKASI PENDATAAN DUSUN 1 - DESA WARNASARI</h2>';
        pdfHtml += '</div>';

        pdfHtml += '<div class="info">';
        pdfHtml += '<strong>📌 Catatan Penting untuk Admin:</strong><br>';
        pdfHtml += '1. Password di bawah ini menggunakan pola <strong>Warnasari + Nomor RT/RW</strong> agar ramah dan mudah diingat oleh pengurus.<br>';
        pdfHtml += '2. Pengurus cukup melakukan login 1 kali pada aplikasi Android/Web, selanjutnya sistem akan mengingat akun mereka secara otomatis.';
        pdfHtml += '</div>';

        pdfHtml += '<table><thead><tr>';
        pdfHtml += '<th style="width:5%;">NO</th>';
        pdfHtml += '<th style="width:25%;">NAMA JABATAN</th>';
        pdfHtml += '<th style="width:20%;">USERNAME</th>';
        pdfHtml += '<th style="width:30%;">PASSWORD AWAL</th>';
        pdfHtml += '<th style="width:20%;">ROLE / WILAYAH</th>';
        pdfHtml += '</tr></thead><tbody>';

        plainPasswords.forEach(function(p, idx) {
          var wilayah = p.role.toUpperCase();
          if (p.rw) wilayah += ' (RW ' + p.rw + (p.rt ? ' RT ' + p.rt : '') + ')';
          pdfHtml += '<tr>';
          pdfHtml += '<td>' + (idx + 1) + '</td>';
          pdfHtml += '<td style="text-align:left;font-weight:bold;">' + escapeHtmlServer(p.label) + '</td>';
          pdfHtml += '<td><code style="font-weight:bold;">' + escapeHtmlServer(p.username) + '</code></td>';
          pdfHtml += '<td><span class="pass-code">' + escapeHtmlServer(p.password) + '</span></td>';
          pdfHtml += '<td>' + escapeHtmlServer(wilayah) + '</td>';
          pdfHtml += '</tr>';
        });

        pdfHtml += '</tbody></table>';

        pdfHtml += '<div class="footer">';
        pdfHtml += 'Dicetak: ' + new Date().toLocaleString('id-ID') + ' | <strong>&copy; 2026 KKN 06 Desa Warnasari Ikopin University</strong>';
        pdfHtml += '</div>';
        pdfHtml += '</body></html>';

        var blob = Utilities.newBlob(pdfHtml, 'text/html', 'Daftar_Akun_Warnasari.html');
        var pdfBlob = blob.getAs('application/pdf').setName('Daftar_Akun_Pengurus_Dusun1_Warnasari.pdf');

        MailApp.sendEmail({
          to: email,
          subject: '📄 PDF Daftar Akun & Password Pengurus Dusun 1 - Desa Warnasari',
          body: 'Halo Admin,\n\nTerlampir file PDF berisi daftar lengkap username dan password awal seluruh pengurus RT/RW Dusun 1 Desa Warnasari.\n\nPassword dibuat dengan pola mudah (Warnasari + Nomor RT/RW) agar ramah bagi pengurus desa.\n\nSilakan unduh atau cetak lampiran PDF di bawah ini untuk dibagikan ke masing-masing pengurus.\n\nSalam,\nTim KKN 06 Desa Warnasari Ikopin University',
          attachments: [pdfBlob]
        });
      }
    } catch(e) {
      console.log('Gagal kirim PDF email:', e);
    }
  }
  
  var vals = sheet.getDataRange().getValues();
  if(vals.length<2) return {};
  var users = {};
  for(var i=1;i<vals.length;i++){
    var u = String(vals[i][0]||'').trim();
    if(u) users[u] = {
      passwordHash:String(vals[i][1]||'').trim(),
      role:String(vals[i][2]||'').trim(),
      dusun:String(vals[i][3]||'').trim(),
      rw:normalizeRW(vals[i][4]),
      rt:String(vals[i][5]||'').trim(),
      label:String(vals[i][6]||'').trim()||u,
      salt:String(vals[i][7]||'').trim()
    };
  }
  return users;
}

// ============================================================
// LOGIN DENGAN SALT + BRUTE FORCE PROTECTION
// ============================================================

function login(username, password) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    return {success: false, message: 'Server sibuk, coba lagi.'};
  }
  try {
    // Sanitasi input login
    username = String(username || '').trim();
    password = String(password || '');
    if (!username || !password) {
      return {success: false, message: 'Username dan password harus diisi'};
    }

    if(isUserLocked(username)) {
      var attempts = getLoginAttempts(username);
      var remaining = Math.ceil((attempts.lockedUntil - new Date()) / 60000);
      writeAuditLog(username, 'unknown', 'LOGIN_BLOCKED', 'User terkunci, sisa ' + remaining + ' menit');
      return {
        success: false, 
        message: 'Akun terkunci! Coba lagi setelah ' + remaining + ' menit.',
        locked: true,
        remaining: remaining
      };
    }
    
    var users = getUsers();
    if(users[username]) {
      if(verifyPasswordWithSalt(password, users[username].passwordHash, users[username].salt)){
        updateLoginAttempts(username, true);
        var u = users[username];
        var token = createSession(username);
        writeAuditLog(username, u.role, 'LOGIN_SUCCESS', 'Login berhasil', token);
        return {
          success:true, 
          username:username, 
          role:u.role, 
          dusun:u.dusun||null, 
          rw:u.rw||null, 
          rt:u.rt||null, 
          label:u.label||username,
          sessionTimeout: SECURITY.SESSION_TIMEOUT,
          token: token
        };
      }
    }
    
    updateLoginAttempts(username, false);
    var attempts2 = getLoginAttempts(username);
    var remainingAttempts = SECURITY.MAX_LOGIN_ATTEMPTS - (attempts2 ? attempts2.attempts : 0);
    
    writeAuditLog(username, 'unknown', 'LOGIN_FAILED', 'Percobaan login gagal, sisa ' + remainingAttempts + ' attempts');
    
    if(remainingAttempts <= 0) {
      return {
        success: false, 
        message: 'Terlalu banyak percobaan gagal! Akun terkunci ' + SECURITY.LOCKOUT_DURATION + ' menit.',
        locked: true
      };
    }
    
    return {
      success: false, 
      message: 'Username atau password salah! Sisa percobaan: ' + remainingAttempts,
      remainingAttempts: remainingAttempts
    };
  } catch(e){ 
    writeAuditLog('system', 'system', 'LOGIN_ERROR', e.message);
    return {success:false, message:e.message}; 
  } finally {
    lock.releaseLock();
  }
}

function logout(username, role, token) {
  if (token) {
    endSession(token);
  }
  writeAuditLog(username, role, 'LOGOUT', 'User logout');
  return {success: true};
}

// ============================================================
// CEK AKSES
// ============================================================

function isSuperAdmin(u) {
  return u && u.role === 'superadmin';
}

function isKadus(u) {
  return u && u.role === 'kadus';
}

function canAccessAllData(u) {
  return u && (u.role === 'superadmin' || u.role === 'kadus');
}

// ============================================================
// VALIDASI TERPUSAT
// ============================================================

function validateData(sheetName, data, isEdit, existingRow) {
  var errors = [];
  
  if (data.Dusun && data.Dusun !== 'Dusun 1') {
    errors.push('Data harus Dusun 1');
  }
  
  if (data.NIK && data.NIK.trim() !== '' && data.NIK.trim() !== 'Belum Ada' && data.NIK.trim() !== '-') {
    if (!isValidNIK(data.NIK)) {
      errors.push('NIK harus 16 digit angka');
    }
  } else if (sheetName === 'DataWarga' || sheetName === 'DataDatang') {
    errors.push('NIK wajib diisi');
  }
  
  if (data['No KK'] && data['No KK'].trim() !== '') {
    if (!isValidKK(data['No KK'])) {
      errors.push('No KK harus 16 digit angka');
    }
  }
  
  if (sheetName === 'DataWarga' && data.NIK && data.NIK.trim() !== '') {
    var existingWarga = getData('DataWarga');
    var currentNik = isEdit && existingRow ? normalizeNIK(existingRow.NIK) : '';
    var dupNik = existingWarga.some(function(r) {
      if (isEdit && r._rowIndex === existingRow._rowIndex) return false;
      return r.NIK === normalizeNIK(data.NIK);
    });
    if (dupNik) {
      errors.push('NIK ' + data.NIK + ' sudah terdaftar sebagai warga lain! NIK tidak boleh ganda.');
    }
  }
  
  if (sheetName === 'DataDatang' && data.NIK && data.NIK.trim() !== '') {
    var existingWarga2 = getData('DataWarga');
    var currentNikDatang = isEdit && existingRow ? normalizeNIK(existingRow.NIK) : '';
    var newNikDatang = normalizeNIK(data.NIK);
    if (!isEdit || currentNikDatang !== newNikDatang) {
      var dupNik2 = existingWarga2.some(function(r) {
        if (isEdit && r.NIK === currentNikDatang) return false;
        return r.NIK === newNikDatang;
      });
      if (dupNik2) {
        errors.push('NIK ' + data.NIK + ' sudah terdaftar sebagai warga! NIK tidak boleh ganda.');
      }
    }
  }
  
  if (sheetName === 'DataLahir') {
    if (data['Berat Badan'] && isNaN(Number(data['Berat Badan']))) {
      errors.push('Berat badan harus berupa angka');
    }
    if (data['Panjang Badan'] && isNaN(Number(data['Panjang Badan']))) {
      errors.push('Panjang badan harus berupa angka');
    }
    if (!data['NIK Bayi'] || data['NIK Bayi'].trim() === '') {
      data['NIK Bayi'] = 'Belum Ada';
    } else if (data['NIK Bayi'].trim() !== 'Belum Ada' && data['NIK Bayi'].trim() !== '-' && !/^\d{16}$/.test(data['NIK Bayi'].trim())) {
      errors.push('NIK Bayi harus 16 digit angka (atau dikosongkan jika belum ada)');
    }
  }
  
  if ((sheetName === 'DataPergi' || sheetName === 'DataMeninggal') && data.NIK) {
    if (data.NIK !== 'Belum Ada') {
      var wargaAda = getData('DataWarga').some(function(r) {
        return r.NIK === normalizeNIK(data.NIK);
      });
      var bayiAda = getData('DataLahir').some(function(r) {
        return (r['NIK Bayi'] && r['NIK Bayi'] === data.NIK);
      });
      if (!wargaAda && !bayiAda) {
        errors.push('NIK tidak ditemukan pada Data Warga maupun Data Kelahiran!');
      }
    }
  }
  
  return errors;
}

// ============================================================
// RESET USERS (MENU TOOLS)
// ============================================================

function resetUsers() {
  try {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      '⚠️ PERINGATAN!',
      'Fungsi ini akan MENGHAPUS SEMUA USER yang ada dan membuat ulang.\n\n' +
      'Semua password akan direset ke default.\n\n' +
      'Yakin ingin melanjutkan?',
      ui.ButtonSet.YES_NO
    );
    
    if (response != ui.Button.YES) {
      ui.alert('Dibatalkan.');
      return {success:false, message:'Dibatalkan'};
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Users');
    if (sheet) ss.deleteSheet(sheet);
    
    var attemptsSheet = ss.getSheetByName('LoginAttempts');
    if (attemptsSheet) ss.deleteSheet(attemptsSheet);
    
    getUsers();
    writeAuditLog('system', 'system', 'RESET_USERS', 'Users berhasil direset dengan salt');
    
    ui.alert('✅ Users berhasil direset!\n\nSemua password sudah terenkripsi dengan SALT.\n\nPassword SuperAdmin telah dikirim ke email Anda.');
    return {success:true, message:'Users berhasil direset!'};
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    return {success:false, message:e.message};
  }
}

// ============================================================
// MIGRASI PASSWORD KE HASH DENGAN SALT
// ============================================================

function migrateToHash() {
  try {
    var sheet = getSheet('Users');
    var data = sheet.getDataRange().getValues();
    if(data.length < 2) {
      SpreadsheetApp.getUi().alert('Tidak ada data user untuk dimigrasi.');
      return {success:false, message:'Tidak ada data user'};
    }
    
    var headers = data[0];
    var passIdx = headers.indexOf('Password');
    var saltIdx = headers.indexOf('Salt');
    
    if(passIdx === -1) {
      SpreadsheetApp.getUi().alert('Kolom Password tidak ditemukan!');
      return {success:false, message:'Kolom Password tidak ditemukan'};
    }
    
    if(saltIdx === -1) {
      sheet.insertColumnAfter(passIdx + 1);
      sheet.getRange(1, passIdx + 2).setValue('Salt');
      headers = sheet.getDataRange().getValues()[0];
      saltIdx = headers.indexOf('Salt');
      SpreadsheetApp.getUi().alert('Kolom Salt ditambahkan.');
    }
    
    var updated = 0;
    for(var i=1; i<data.length; i++) {
      var row = data[i];
      var password = String(row[passIdx] || '').trim();
      var salt = String(row[saltIdx] || '').trim();
      
      if(password.length === 64 && /^[a-f0-9]{64}$/.test(password) && salt) {
        continue;
      }
      
      if(password) {
        var newSalt = salt || generateSalt();
        var hashed = hashPasswordWithSalt(password, newSalt);
        sheet.getRange(i+1, passIdx+1).setValue(hashed);
        sheet.getRange(i+1, saltIdx+1).setValue(newSalt);
        updated++;
      }
    }
    
    writeAuditLog('system', 'system', 'MIGRATE_HASH', updated + ' password di-hash dengan salt');
    
    SpreadsheetApp.getUi().alert('✅ Migrasi selesai!\n\n' + updated + ' password berhasil di-hash dengan salt.');
    return {success:true, message:updated + ' password di-hash dengan salt'};
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
    return {success:false, message:e.message};
  }
}

// ============================================================
// INIT SHEETS DENGAN KONFIRMASI
// ============================================================

function initSheets() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert(
    '⚠️ PERINGATAN!',
    'Fungsi ini akan MENGHAPUS SEMUA SHEET yang ada dan membuat ulang.\n\n' +
    'Semua data akan HILANG!\n\n' +
    'Yakin ingin melanjutkan?',
    ui.ButtonSet.YES_NO
  );
  
  if (response != ui.Button.YES) {
    ui.alert('Dibatalkan. Data aman.');
    return {success: false, message: 'Dibatalkan oleh user'};
  }
  
  var response2 = ui.alert(
    'Konfirmasi Terakhir',
    'SEMUA DATA AKAN HILANG!\n\n' +
    'Pastikan Anda sudah backup data penting.\n\n' +
    'LANJUTKAN?',
    ui.ButtonSet.YES_NO
  );
  
  if (response2 != ui.Button.YES) {
    ui.alert('Dibatalkan. Data aman.');
    return {success: false, message: 'Dibatalkan oleh user'};
  }
  
  try {
    createEmergencyBackup();
  } catch(e) {
    ui.alert('Error backup: ' + e.message);
    return {success: false, message: e.message};
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = {
    'DataWarga': ['No KK','NIK','Nama Lengkap','Dusun','RW','RT','No. Rumah','Jenis Kelamin','Tempat Lahir','Tanggal Lahir','Status Perkawinan','Pekerjaan','Pendidikan','Agama','Golongan Darah','Hubungan Keluarga','Nama Ayah','Nama Ibu','Tahun','Bulan'],
    'DataDatang': ['No KK','NIK','Nama Lengkap','Dusun','RW','RT','No. Rumah','Jenis Kelamin','Tempat Lahir','Tanggal Lahir','Status Perkawinan','Pekerjaan','Pendidikan','Agama','Golongan Darah','Hubungan Keluarga','Nama Ayah','Nama Ibu','Tanggal Datang','Asal Daerah','Alasan','Tahun','Bulan'],
    'DataPergi': ['No KK','NIK','Nama Lengkap','Dusun','RW','RT','No. Rumah','Tanggal Pergi','Tujuan','Alasan','Tahun','Bulan'],
    'DataMeninggal': ['No KK','NIK','Nama Lengkap','Dusun','RW','RT','No. Rumah','Tanggal Meninggal','Tempat Meninggal','Penyebab','Tahun','Bulan'],
    'DataLahir': ['Nama Bayi','Jenis Kelamin','Tanggal Lahir','Tempat Lahir','Berat Badan','Panjang Badan','Dusun','RW','RT','No. Rumah','Nama Ayah','Nama Ibu','No KK','NIK Bayi','Tahun','Bulan'],
    'DataPengontrak': ['NIK','Nama Lengkap','Jenis Kelamin','No HP','Alamat KTP Asal','Dusun','RW','RT','No. Rumah','Nama Pemilik Rumah','Tanggal Mulai','Durasi','Status','Tahun','Bulan'],
    'LaporanBulanan': ['Bulan','Tahun','Dusun','RW','RT','Jumlah Awal','Datang','Pergi','Meninggal','Lahir','Jumlah Akhir','Laki-laki','Perempuan','KK']
  };

  Object.keys(config).forEach(function(n){
    var s = ss.getSheetByName(n);
    if(s) ss.deleteSheet(s);
    s = ss.insertSheet(n);
    var h = config[n];
    s.getRange(1,1,1,h.length).setValues([h]);
    s.getRange(1,1,1,h.length).setFontWeight('bold');
    s.setFrozenRows(1);
    if(TEXT_COLUMNS[n]) applyTextFormatToColumn(s, h, TEXT_COLUMNS[n], 2000);
  });
  
  initBackupSheet();
  initAuditLogSheet();
  initLoginAttemptsSheet();
  getUsers();
  
  writeAuditLog('system', 'system', 'INIT_SHEETS', 'Aplikasi diinisialisasi (data direset)');
  
  ui.alert('✅ Inisialisasi selesai! Semua sheet telah dibuat ulang.');
  return {success:true};
}

function createEmergencyBackup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var backupName = 'EMERGENCY_BACKUP_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  var backup = ss.copy(backupName);
  
  var props = PropertiesService.getScriptProperties();
  props.setProperty('EMERGENCY_BACKUP_ID', backup.getId());
  props.setProperty('EMERGENCY_BACKUP_NAME', backupName);
  
  return {success: true, id: backup.getId(), name: backupName};
}

// ============================================================
// FUNGSI BANTU
// ============================================================

function getHeaders(n){
  var h = {
    'DataWarga':['No KK','NIK','Nama Lengkap','Dusun','RW','RT','No. Rumah','Jenis Kelamin','Tempat Lahir','Tanggal Lahir','Status Perkawinan','Pekerjaan','Pendidikan','Agama','Golongan Darah','Hubungan Keluarga','Nama Ayah','Nama Ibu','Tahun','Bulan'],
    'DataDatang':['No KK','NIK','Nama Lengkap','Dusun','RW','RT','No. Rumah','Jenis Kelamin','Tempat Lahir','Tanggal Lahir','Status Perkawinan','Pekerjaan','Pendidikan','Agama','Golongan Darah','Hubungan Keluarga','Nama Ayah','Nama Ibu','Tanggal Datang','Asal Daerah','Alasan','Tahun','Bulan'],
    'DataPergi':['No KK','NIK','Nama Lengkap','Dusun','RW','RT','No. Rumah','Tanggal Pergi','Tujuan','Alasan','Tahun','Bulan'],
    'DataMeninggal':['No KK','NIK','Nama Lengkap','Dusun','RW','RT','No. Rumah','Tanggal Meninggal','Tempat Meninggal','Penyebab','Tahun','Bulan'],
    'DataLahir':['Nama Bayi','Jenis Kelamin','Tanggal Lahir','Tempat Lahir','Berat Badan','Panjang Badan','Dusun','RW','RT','No. Rumah','Nama Ayah','Nama Ibu','No KK','NIK Bayi','Tahun','Bulan'],
    'DataPengontrak':['NIK','Nama Lengkap','Jenis Kelamin','No HP','Alamat KTP Asal','Dusun','RW','RT','No. Rumah','Nama Pemilik Rumah','Tanggal Mulai','Durasi','Status','Tahun','Bulan'],
    'LaporanBulanan':['Bulan','Tahun','Dusun','RW','RT','Jumlah Awal','Datang','Pergi','Meninggal','Lahir','Jumlah Akhir','Laki-laki','Perempuan','KK']
  };
  return h[n]||[];
}

function ensureSheetHeaders(s, sheetName) {
  try {
    var expected = getHeaders(sheetName);
    if (!expected || expected.length === 0) return;
    var lastCol = s.getLastColumn();
    if (lastCol === 0) {
      s.getRange(1, 1, 1, expected.length).setValues([expected]).setFontWeight('bold');
      s.setFrozenRows(1);
      return;
    }
    var current = s.getRange(1, 1, 1, Math.max(lastCol, expected.length)).getValues()[0];
    var isSame = expected.every(function(h, idx) { return current[idx] === h; });
    if (!isSame) {
      s.getRange(1, 1, 1, expected.length).setValues([expected]).setFontWeight('bold');
      s.setFrozenRows(1);
    }
  } catch (e) {}
}

function repairAllSheetHeadersAndData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetsToRepair = ['DataWarga', 'DataDatang', 'DataPergi', 'DataMeninggal', 'DataLahir'];

    sheetsToRepair.forEach(function(sheetName) {
      var s = ss.getSheetByName(sheetName);
      if (!s) return;
      var canonical = getHeaders(sheetName);
      if (!canonical || canonical.length === 0) return;

      var range = s.getDataRange();
      var values = range.getValues();
      if (values.length < 1) return;

      var currentHeaders = values[0].map(function(h) { return String(h).trim(); });
      var dataRows = values.slice(1);

      var repairedRows = [];
      dataRows.forEach(function(row) {
        if (!row || row.every(function(c) { return String(c).trim() === ''; })) return;

        var rowObj = {};
        currentHeaders.forEach(function(hName, idx) {
          if (hName) rowObj[hName] = row[idx];
        });

        // Ensure input time (Tahun & Bulan) is preserved
        if (!rowObj['Tahun'] || String(rowObj['Tahun']).trim() === '') {
          rowObj['Tahun'] = '2026';
        }
        if (!rowObj['Bulan'] || String(rowObj['Bulan']).trim() === '') {
          rowObj['Bulan'] = '8';
        }

        var cleanRow = canonical.map(function(col) {
          var val = rowObj[col];
          if (val instanceof Date || (typeof val === 'string' && (val.indexOf('GMT+') !== -1 || val.indexOf('Waktu Indonesia') !== -1 || /^[A-Za-z]{3}\s+[A-Za-z]{3}/.test(val)))) {
            val = formatDate(val);
          } else if (col.indexOf('Tanggal') !== -1 || col.indexOf('Tgl') !== -1) {
            val = formatDate(val);
          }
          return (val !== undefined && val !== null) ? val : '';
        });
        repairedRows.push(cleanRow);
      });

      s.clear();
      s.getRange(1, 1, 1, canonical.length).setValues([canonical]).setFontWeight('bold');
      s.setFrozenRows(1);
      if (TEXT_COLUMNS[sheetName]) {
        applyTextFormatToColumn(s, canonical, TEXT_COLUMNS[sheetName], Math.max(2000, repairedRows.length + 10));
      }

      if (repairedRows.length > 0) {
        s.getRange(2, 1, repairedRows.length, canonical.length).setValues(repairedRows);
      }
    });

    // Synchronize DataLahir & DataDatang updates to DataWarga
    try {
      var sLahir = ss.getSheetByName('DataLahir');
      var sWarga = ss.getSheetByName('DataWarga');
      if (sLahir && sWarga) {
        var hLahir = getHeaders('DataLahir');
        var hWarga = getHeaders('DataWarga');
        var vLahir = sLahir.getDataRange().getValues();
        var vWarga = sWarga.getDataRange().getValues();

        if (vLahir.length > 1 && vWarga.length > 1) {
          var nameIdxLahir = hLahir.indexOf('Nama Bayi');
          var kkIdxLahir = hLahir.indexOf('No KK');
          var rwIdxLahir = hLahir.indexOf('RW');
          var rtIdxLahir = hLahir.indexOf('RT');
          var nikIdxLahir = hLahir.indexOf('NIK Bayi');

          var nameIdxWarga = hWarga.indexOf('Nama Lengkap');
          var kkIdxWarga = hWarga.indexOf('No KK');
          var rwIdxWarga = hWarga.indexOf('RW');
          var rtIdxWarga = hWarga.indexOf('RT');
          var nikIdxWarga = hWarga.indexOf('NIK');

          for (var iL = 1; iL < vLahir.length; iL++) {
            var rowL = vLahir[iL];
            var namaL = String(rowL[nameIdxLahir] || '').trim().toLowerCase();
            var kkL = String(rowL[kkIdxLahir] || '').trim();
            var rwL = normalizeRW(rowL[rwIdxLahir]);
            var rtL = String(rowL[rtIdxLahir] || '').trim();
            var nikL = String(rowL[nikIdxLahir] || '').trim();

            for (var iW = 1; iW < vWarga.length; iW++) {
              var rowW = vWarga[iW];
              var namaW = String(rowW[nameIdxWarga] || '').trim().toLowerCase();
              var kkW = String(rowW[kkIdxWarga] || '').trim();

              if (namaL && namaW === namaL && (kkL ? kkW === kkL : true)) {
                if (rwL) sWarga.getRange(iW + 1, rwIdxWarga + 1).setNumberFormat('@').setValue(rwL);
                if (rtL) sWarga.getRange(iW + 1, rtIdxWarga + 1).setNumberFormat('@').setValue(rtL);
                if (!nikL || nikL === '-' || nikL.indexOf('32042026') === 0 || nikL === 'Belum Ada') {
                  sWarga.getRange(iW + 1, nikIdxWarga + 1).setNumberFormat('@').setValue('Belum Ada');
                  sLahir.getRange(iL + 1, nikIdxLahir + 1).setNumberFormat('@').setValue('Belum Ada');
                }
                break;
              }
            }
          }
        }
      }
    } catch (errSync) { }

    clearDataCache();
    updateLaporan();
    return { success: true, message: 'Semua sheet dan kolom berhasil diperbaiki' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getSheet(n){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(n);
  if(!s){
    s = ss.insertSheet(n);
    var h = getHeaders(n);
    if(h.length){
      s.getRange(1,1,1,h.length).setValues([h]);
      s.getRange(1,1,1,h.length).setFontWeight('bold');
      s.setFrozenRows(1);
      if(TEXT_COLUMNS[n]) applyTextFormatToColumn(s, h, TEXT_COLUMNS[n], 2000);
    }
  } else {
    ensureSheetHeaders(s, n);
  }
  return s;
}

var DATA_CACHE = {};

function clearDataCache() {
  DATA_CACHE = {};
}

function getData(n){
  if (DATA_CACHE[n]) return DATA_CACHE[n];
  try {
    var s = getSheet(n);
    var v = s.getDataRange().getValues();
    if(v.length<2) return [];
    var h = v[0];
    var result = [];
    for(var i=1; i<v.length; i++){
      var row = v[i];
      var hasContent = row.some(function(c){ return String(c).trim()!==''; });
      if(!hasContent) continue;
      var o = {};
      h.forEach(function(h2, j) {
        var rawVal = row[j];
        if (rawVal instanceof Date || (typeof rawVal === 'string' && (rawVal.indexOf('GMT+') !== -1 || rawVal.indexOf('Waktu Indonesia') !== -1 || /^[A-Za-z]{3}\s+[A-Za-z]{3}/.test(rawVal)))) {
          o[h2] = formatDate(rawVal);
        } else if (h2.indexOf('Tanggal') !== -1 || h2.indexOf('Tgl') !== -1) {
          o[h2] = formatDate(rawVal);
        } else {
          o[h2] = String(rawVal || '');
        }
      });
      if(o.hasOwnProperty('RW')) o.RW = normalizeRW(o.RW);
      if(o.hasOwnProperty('NIK')) o.NIK = normalizeNIK(o.NIK);
      o._rowIndex = i - 1;
      result.push(o);
    }
    DATA_CACHE[n] = result;
    return result;
  } catch(e){ return []; }
}

function getDropdown() {
  return {
    dusun:['Dusun 1'],
    rw:['13','04','02','03','12','15'],
    rtMap:{'13':['1','2','3','4','5'],'04':['1','2','3'],'02':['1','2','3','4'],'03':['1','2'],'12':['1','2','3'],'15':['1','2','3']},
    rt:['1','2','3','4','5'],
    jenisKelamin:['LAKI-LAKI','PEREMPUAN'],
    statusPerkawinan:['BELUM KAWIN','KAWIN','KAWIN TERCATAT','KAWIN BELUM TERCATAT','CERAI HIDUP','CERAI MATI'],
    pendidikan: [
      'TIDAK / BELUM SEKOLAH',
      'BELUM TAMAT SD / SEDERAJAT',
      'TAMAT SD / SEDERAJAT',
      'SLTP / SEDERAJAT',
      'SLTA / SEDERAJAT',
      'DIPLOMA I / II',
      'AKADEMI / DIPLOMA III',
      'DIPLOMA IV / STRATA I',
      'STRATA II',
      'STRATA III'
    ],
    agama:['ISLAM','KRISTEN','KATHOLIK','HINDU','BUDDHA','KONGHUCU','KEPERCAYAAN KEPADA TUHAN YME'],
    pekerjaan: [
      'BELUM/TIDAK BEKERJA',
      'MENGURUS RUMAH TANGGA',
      'PELAJAR/MAHASISWA',
      'PNS',
      'TNI/POLRI',
      'KARYAWAN SWASTA',
      'WIRASWASTA',
      'PETANI/PEKEBUN',
      'BURUH HARIAN LEPAS',
      'NELAYAN',
      'PEDAGANG',
      'GURU',
      'PENSIUNAN',
      'LAINNYA'
    ],
    alasanDatang:['Pekerjaan','Pendidikan','Pernikahan','Keluarga','Lainnya'],
    alasanPergi:['Pekerjaan','Pendidikan','Pernikahan','Keluarga','Lainnya'],
    penyebabMeninggal:['Sakit','Kecelakaan','Usia Lanjut','Lainnya'],
    hubunganKeluarga:['KEPALA KELUARGA','KEPALA RUMAH TANGGA','SUAMI','ISTRI','ANAK','CUCU','ORANG TUA','MERTUA','FAMILI LAIN','PEMBANTU','LAINNYA'],
    golonganDarah:['A','B','AB','O','A+','A-','B+','B-','AB+','AB-','O+','O-','TIDAK TAHU'],
    kewarganegaraan:['WNI','WNA'],
    statusPengontrak:['Aktif','Selesai','Pindah']
  };
}

// ============================================================
// CEK NIK DUPLIKAT
// ============================================================

function checkNikExists(nik, excludeNik) {
  try {
    var nikTrim = normalizeNIK(nik);
    var exclude = normalizeNIK(excludeNik);
    if (!nikTrim) return null;
    if (exclude && nikTrim === exclude) return null;
    var d = getData('DataWarga');
    for (var i = 0; i < d.length; i++) {
      var r = d[i];
      if (normalizeNIK(r.NIK) === nikTrim) {
        return {
          exists: true,
          nama: r['Nama Lengkap'] || r.Nama || 'Warga lain',
          rt: r.RT,
          rw: r.RW
        };
      }
    }
    return null;
  } catch (e) { return null; }
}

// ============================================================
// AUTOCOMPLETE NIK
// ============================================================

function searchNikList(u, searchTerm) {
  try {
    u = verifyUser(u);
    if (!u) return [];

    var term = (searchTerm || '').trim().toLowerCase();
    
    // 1. Ambil data dari DataWarga (Dusun 1)
    var dWarga = getData('DataWarga').filter(function(r){ 
      return r.Dusun === 'Dusun 1'; 
    });
    
    if(!canAccessAllData(u)){
      if(u.role === 'rw') {
        dWarga = dWarga.filter(function(r){ return String(r.RW) === String(u.rw); });
      } else if(u.role === 'rt') {
        dWarga = dWarga.filter(function(r){ return String(r.RW) === String(u.rw) && String(r.RT) === String(u.rt); });
      }
    }
    
    if(term) {
      dWarga = dWarga.filter(function(w){
        var nik = (w.NIK || '').toLowerCase();
        var nama = (w['Nama Lengkap'] || '').toLowerCase();
        var noKK = (w['No KK'] || '').toLowerCase();
        return nik.indexOf(term) !== -1 || 
               nama.indexOf(term) !== -1 || 
               noKK.indexOf(term) !== -1;
      });
    }

    var res = [];
    var existingKeys = {}; // Key unik untuk mendeteksi duplikat

    // Masukkan data dari DataWarga
    dWarga.forEach(function(w){ 
      if(w.NIK && w.NIK.trim()) {
        var nikClean = w.NIK.trim();
        var namaClean = (w['Nama Lengkap'] || '').trim().toLowerCase();
        var rwClean = String(w.RW || '').trim();
        var rtClean = String(w.RT || '').trim();

        // Key unik: NIK jika valid, atau Nama_RW_RT jika Belum Ada NIK
        var uniqueKey = (nikClean !== 'Belum Ada') ? nikClean : (namaClean + '_' + rwClean + '_' + rtClean);
        existingKeys[uniqueKey] = true;

        res.push({
          nik: nikClean, 
          nama: w['Nama Lengkap'] || '', 
          noKK: w['No KK'] || '', 
          dusun: w.Dusun || '', 
          rw: w.RW || '', 
          rt: w.RT || '',
          isBayi: false
        });
      }
    });

    // 2. Ambil data dari DataLahir (Dusun 1)
    var dLahir = getData('DataLahir').filter(function(r){ 
      return r.Dusun === 'Dusun 1'; 
    });

    if(!canAccessAllData(u)){
      if(u.role === 'rw') {
        dLahir = dLahir.filter(function(r){ return String(r.RW) === String(u.rw); });
      } else if(u.role === 'rt') {
        dLahir = dLahir.filter(function(r){ return String(r.RW) === String(u.rw) && String(r.RT) === String(u.rt); });
      }
    }

    if(term) {
      dLahir = dLahir.filter(function(b){
        var nikB = (b['NIK Bayi'] || '').toLowerCase();
        var namaB = (b['Nama Bayi'] || '').toLowerCase();
        var ibuB = (b['Nama Ibu'] || '').toLowerCase();
        var ayahB = (b['Nama Ayah'] || '').toLowerCase();
        var kkB = (b['No KK'] || '').toLowerCase();
        return nikB.indexOf(term) !== -1 || 
               namaB.indexOf(term) !== -1 || 
               ibuB.indexOf(term) !== -1 || 
               ayahB.indexOf(term) !== -1 || 
               kkB.indexOf(term) !== -1;
      });
    }

    // Masukkan data dari DataLahir (Hanya jika belum ada di DataWarga)
    dLahir.forEach(function(b){
      var displayNik = (b['NIK Bayi'] && b['NIK Bayi'].trim() && b['NIK Bayi'] !== 'Belum Ada') ? b['NIK Bayi'].trim() : 'Belum Ada';
      var namaBayi = (b['Nama Bayi'] || '').trim().toLowerCase();
      var rwClean = String(b.RW || '').trim();
      var rtClean = String(b.RT || '').trim();

      var uniqueKey = (displayNik !== 'Belum Ada') ? displayNik : (namaBayi + '_' + rwClean + '_' + rtClean);

      if (!existingKeys[uniqueKey]) {
        existingKeys[uniqueKey] = true;
        var namaLabel = (b['Nama Bayi'] || 'Bayi Ny.') + (b['Nama Ibu'] ? ' (Ibu: ' + b['Nama Ibu'] + ')' : '');
        res.push({
          nik: displayNik,
          nama: namaLabel,
          noKK: b['No KK'] || '',
          dusun: b.Dusun || 'Dusun 1',
          rw: b.RW || '',
          rt: b.RT || '',
          isBayi: true,
          namaBayiAsli: b['Nama Bayi'] || ''
        });
      }
    });

    if(res.length > 50) {
      res = res.slice(0, 50);
    }
    
    return res;
  } catch(e) {
    console.error('searchNikList error:', e);
    return [];
  }
}

function getWargaByNIK(nik) {
  var d = getData('DataWarga');
  var nikTrim = normalizeNIK(nik);
  for(var i=0;i<d.length;i++){ 
    if(d[i].NIK === nikTrim) return d[i]; 
  }
  return null;
}

// ============================================================
// VALIDASI FORMAT
// ============================================================

function isValidNIK(nik) {
  return /^\d{16}$/.test(String(nik||'').trim());
}

function isValidKK(kk) {
  return /^\d{16}$/.test(String(kk||'').trim());
}

// ============================================================
// FORMAT TANGGAL
// ============================================================

function formatDate(dateStr) {
  if (!dateStr) return '';
  if (dateStr instanceof Date) {
    if (isNaN(dateStr.getTime())) return '';
    var day = String(dateStr.getDate()).padStart(2, '0');
    var month = String(dateStr.getMonth() + 1).padStart(2, '0');
    var year = dateStr.getFullYear();
    return day + '/' + month + '/' + year;
  }
  var s = String(dateStr).trim();
  if (!s || s === '-' || s === 'null' || s === 'undefined') return '';

  // Clean JS Date string formats like "Fri Aug 07 2026 00:00:00 GMT+0700 (Waktu Indonesia Barat)"
  if (s.indexOf('GMT+') !== -1 || s.indexOf('Waktu Indonesia') !== -1 || s.indexOf('00:00:00') !== -1 || /^[A-Za-z]{3}\s+[A-Za-z]{3}\s+\d{2}/.test(s)) {
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      var day = String(d.getDate()).padStart(2, '0');
      var month = String(d.getMonth() + 1).padStart(2, '0');
      var year = d.getFullYear();
      return day + '/' + month + '/' + year;
    }
  }

  // Handle YYYY-MM-DD or YYYY/MM/DD
  var matchYMD = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (matchYMD) {
    var y = matchYMD[1];
    var m = ('0' + matchYMD[2]).slice(-2);
    var d = ('0' + matchYMD[3]).slice(-2);
    return d + '/' + m + '/' + y;
  }

  // Handle DD/MM/YYYY or DD-MM-YYYY
  var matchDMY = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (matchDMY) {
    var d = ('0' + matchDMY[1]).slice(-2);
    var m = ('0' + matchDMY[2]).slice(-2);
    var y = matchDMY[3];
    return d + '/' + m + '/' + y;
  }

  try {
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      var day = String(d.getDate()).padStart(2, '0');
      var month = String(d.getMonth() + 1).padStart(2, '0');
      var year = d.getFullYear();
      return day + '/' + month + '/' + year;
    }
  } catch (e) {}

  return s;
}

// ============================================================
// FILTER TERPUSAT
// ============================================================

function applyFilters(data, filter, u) {
  if (!data || data.length === 0) return data;
  
  u = verifyUser(u);
  if (!u) return [];

  var result = data;
  
  if (!canAccessAllData(u)) {
    if (u.role === 'rw') {
      result = result.filter(function(r) {
        return r.Dusun === 'Dusun 1' && String(r.RW) === String(u.rw);
      });
    } else if (u.role === 'rt') {
      result = result.filter(function(r) {
        return r.Dusun === 'Dusun 1' && 
               String(r.RW) === String(u.rw) && 
               String(r.RT) === String(u.rt);
      });
    }
  }
  
  if (filter) {
    if (filter.rw && filter.rw !== '') {
      result = result.filter(function(r) {
        return String(r.RW) === String(normalizeRW(filter.rw));
      });
    }
    if (filter.rt && filter.rt !== '') {
      result = result.filter(function(r) {
        return String(r.RT) === String(filter.rt);
      });
    }
    if (filter.bulan && filter.bulan !== '') {
      result = result.filter(function(r) {
        return String(r.Bulan) === String(filter.bulan);
      });
    }
    if (filter.tahun && filter.tahun !== '') {
      result = result.filter(function(r) {
        return String(r.Tahun) === String(filter.tahun);
      });
    }
    if (filter.search && filter.search !== '') {
      var search = filter.search.toLowerCase().trim();
      result = result.filter(function(r) {
        var nik = (r.NIK || '').toLowerCase();
        var nama = (r['Nama Lengkap'] || '').toLowerCase();
        var noKK = (r['No KK'] || '').toLowerCase();
        return nik.indexOf(search) !== -1 ||
               nama.indexOf(search) !== -1 ||
               noKK.indexOf(search) !== -1;
      });
    }
  }
  
  return result;
}

// ============================================================
// DATA DENGAN FILTER
// ============================================================

function getFilteredData(n, u, filter) {
  if (n === 'LaporanBulanan') {
    try { updateLaporan(); } catch (e) { }
  }
  var allData = getData(n);
  if(!allData || allData.length === 0) return [];

  var dateFields = {
    'DataWarga': 'Tanggal Lahir',
    'DataDatang': 'Tanggal Datang',
    'DataPergi': 'Tanggal Pergi',
    'DataMeninggal': 'Tanggal Meninggal',
    'DataLahir': 'Tanggal Lahir'
  };

  if(dateFields[n]){
    var field = dateFields[n];
    allData = allData.map(function(row){
      if(row[field]){
        row[field] = formatDate(row[field]);
      }
      return row;
    });
  }

  return applyFilters(allData, filter, u);
}

function getFilterOptions(u) {
  u = verifyUser(u);
  if (!u) return {dusun:['Dusun 1'], rw:[], rt:[]};

  var allWarga = getData('DataWarga').filter(function(r){ return r.Dusun === 'Dusun 1'; });

  // Bangun map RW -> set of RT secara dinamis dari data aktual
  var rwMap = {};
  allWarga.forEach(function(r){
    if(!r.RW) return;
    var rw = normalizeRW(r.RW);
    if(!rwMap[rw]) rwMap[rw] = {};
    if(r.RT) rwMap[rw][String(r.RT).trim()] = true;
  });

  // Urutkan RW secara numerik
  var allRW = Object.keys(rwMap).sort(function(a,b){ return parseInt(a,10) - parseInt(b,10); });

  // Kumpulkan semua RT unik dari semua RW, diurutkan numerik
  var rtSet = {};
  Object.keys(rwMap).forEach(function(rw){
    Object.keys(rwMap[rw]).forEach(function(rt){ rtSet[rt] = true; });
  });
  var allRT = Object.keys(rtSet).sort(function(a,b){ return parseInt(a,10) - parseInt(b,10); });

  // Fallback jika data kosong
  if(allRW.length === 0) allRW = ['02','03','04','12','13','15'];
  if(allRT.length === 0) allRT = ['1','2','3','4','5'];

  var rw = allRW, rt = allRT;

  if(canAccessAllData(u)){
    // KADUS / superadmin: lihat semua
  } else if(u.role === 'rw'){
    rw = [normalizeRW(u.rw)];
    rt = rwMap[normalizeRW(u.rw)] ? Object.keys(rwMap[normalizeRW(u.rw)]).sort(function(a,b){ return parseInt(a,10)-parseInt(b,10); }) : allRT;
  } else if(u.role === 'rt'){
    rw = [normalizeRW(u.rw)];
    rt = [String(u.rt)];
  }

  return {dusun:['Dusun 1'], rw:rw, rt:rt};
}

function getNikList(u) {
  u = verifyUser(u);
  if (!u) return [];

  var d = getData('DataWarga').filter(function(r){ return r.Dusun==='Dusun 1'; });
  
  if(canAccessAllData(u)){
  } else if(u.role==='rw') {
    d = d.filter(function(r){ return String(r.RW)===String(u.rw); });
  } else if(u.role==='rt') {
    d = d.filter(function(r){ return String(r.RW)===String(u.rw) && String(r.RT)===String(u.rt); });
  }
  
  var res = [];
  d.forEach(function(w){ if(w.NIK && w.NIK.trim()) res.push({nik:w.NIK, nama:w['Nama Lengkap'], noKK:w['No KK'], dusun:w.Dusun, rw:w.RW, rt:w.RT, noRumah:w['No. Rumah']||''}); });
  return res;
}

// ============================================================
// CRUD - DENGAN VALIDASI
// ============================================================

function writeRowSafely(sheetName, sheetObj, headers, rowIndex, rowDataObj) {
  var range = sheetObj.getRange(rowIndex, 1, 1, headers.length);
  var textCols = TEXT_COLUMNS[sheetName] || [];
  headers.forEach(function(colName, i){
    if(textCols.indexOf(colName) !== -1){
      range.getCell(1, i+1).setNumberFormat('@');
    }
  });
  var values = headers.map(function(k){ return rowDataObj[k] || ''; });
  range.setValues([values]);
}

function formatDataObject(data) {
  if (!data || typeof data !== 'object') return data;
  Object.keys(data).forEach(function (key) {
    var val = data[key];
    if (typeof val === 'string' && val.trim()) {
      var s = val.trim();
      if (key.indexOf('Nama') !== -1) {
        data[key] = s.toUpperCase();
      } else if (key !== 'No KK' && key !== 'NIK' && key !== 'NIK Bayi' && key !== 'RW' && key !== 'RT' && 
                 key !== 'No. Rumah' && key !== 'No HP' && key.indexOf('Tanggal') === -1 && 
                 key.indexOf('Tgl') === -1 && key !== 'Tahun' && key !== 'Bulan' && key !== '_rowIndex') {
        data[key] = s.toLowerCase().replace(/(?:^|\s|-|\/)\S/g, function (a) { return a.toUpperCase(); });
      }
    }
  });
  return data;
}

function addData(n, data, u) {
  try {
    // Verifikasi user terhadap Users sheet
    u = verifyUser(u);
    if (!u) return {success: false, message: 'Sesi tidak valid. Silakan login ulang.'};

    // Format Kapitalisasi
    data = formatDataObject(data);

    // Sanitasi input terhadap formula injection
    var safeFields = ['RW', 'RT', 'NIK', 'No KK', 'Tahun', 'Bulan'];
    data = sanitizeDataObject(data, safeFields);

    var errors = validateData(n, data, false, null);
    if (errors.length > 0) {
      writeAuditLog(u.username, u.role, 'ADD_DATA_FAILED', errors.join('; '));
      return {success: false, message: errors.join('; ')};
    }

    if(data.RW) data.RW = normalizeRW(data.RW);
    if(data.RT) data.RT = String(data.RT).trim();
    if(data.hasOwnProperty('NIK')) data.NIK = normalizeNIK(data.NIK);

    // Cek akses RT/RW
    var accessErr = _checkDataAccess(u, data, 'ADD_DATA');
    if (accessErr) return accessErr;

    var now = new Date();
    var tahun = String(now.getFullYear());
    var bulan = String(now.getMonth() + 1);

    data['Tahun'] = tahun;
    data['Bulan'] = bulan;

    var s = getSheet(n);
    var h = getHeaders(n);
    var newRowIndex = s.getLastRow() + 1;
    writeRowSafely(n, s, h, newRowIndex, data);

    if(n==='DataDatang'){
      var ws = getSheet('DataWarga');
      var wh = getHeaders('DataWarga');
      var wd = {
        'No KK':data['No KK']||'',
        'NIK':data['NIK']||'',
        'Nama Lengkap':data['Nama Lengkap']||'',
        'Dusun':'Dusun 1',
        'RW':data['RW']||'',
        'RT':data['RT']||'',
        'No. Rumah':data['No. Rumah']||'',
        'Jenis Kelamin':data['Jenis Kelamin']||'',
        'Tempat Lahir':data['Tempat Lahir']||'',
        'Tanggal Lahir':data['Tanggal Lahir']||'',
        'Status Perkawinan':data['Status Perkawinan']||'',
        'Pekerjaan':data['Pekerjaan']||'',
        'Pendidikan':data['Pendidikan']||'',
        'Agama':data['Agama']||'',
        'Hubungan Keluarga':data['Hubungan Keluarga']||'',
        'Tahun': tahun,
        'Bulan': bulan
      };
      writeRowSafely('DataWarga', ws, wh, ws.getLastRow() + 1, wd);
    }
    else if(n==='DataPergi' || n==='DataMeninggal'){
      var nik = data['NIK'];
      if(nik){
        var wd2 = getData('DataWarga');
        var idx = -1;
        for(var i=0;i<wd2.length;i++){ if(wd2[i].NIK===nik){ idx=i; break; } }
        if(idx!==-1){ 
          var ws2 = getSheet('DataWarga');
          backupData('DataWarga', wd2[idx], u.username);
          ws2.deleteRow(wd2[idx]._rowIndex + 2); 
        }
      }
    }
    else if(n==='DataLahir'){
      var ws3 = getSheet('DataWarga');
      var wh3 = getHeaders('DataWarga');
      
      var rawNikBayi = (data['NIK Bayi'] || '').trim();
      var nikBayi = (rawNikBayi && rawNikBayi !== '-' && rawNikBayi !== 'null') ? rawNikBayi : 'Belum Ada';
      data['NIK Bayi'] = nikBayi;
      
      var wd3 = {
        'No KK':data['No KK']||'',
        'NIK':nikBayi,
        'Nama Lengkap':data['Nama Bayi']||'',
        'Dusun':'Dusun 1',
        'RW':data['RW']||'',
        'RT':data['RT']||'',
        'No. Rumah':data['No. Rumah']||'',
        'Jenis Kelamin':data['Jenis Kelamin']||'',
        'Tempat Lahir':data['Tempat Lahir']||'',
        'Tanggal Lahir':data['Tanggal Lahir']||'',
        'Status Perkawinan':'Belum Kawin',
        'Pekerjaan':'BELUM/TIDAK BEKERJA',
        'Pendidikan':'TIDAK / BELUM SEKOLAH',
        'Agama':'Islam',
        'Hubungan Keluarga':'Anak',
        'Tahun': tahun,
        'Bulan': bulan
      };
      writeRowSafely('DataWarga', ws3, wh3, ws3.getLastRow() + 1, wd3);
    }

    clearDataCache();
    updateLaporan();
    writeAuditLog(u.username, u.role, 'ADD_DATA', n + ': ' + (data['Nama Lengkap'] || data['Nama Bayi'] || ''));
    return {success:true, message:'Data berhasil ditambahkan'};
  } catch(e){ 
    writeAuditLog(u.username, u.role, 'ADD_DATA_ERROR', e.message);
    return {success:false, message:e.message}; 
  }
}

function processDatangKeluarga(payload, u) {
  try {
    u = verifyUser(u);
    if (!u) return { success: false, message: 'Sesi tidak valid.' };
    if (!payload || !payload.noKK || !payload.members || payload.members.length === 0) {
      return { success: false, message: 'Data keluarga tidak lengkap.' };
    }

    var noKK = normalizeNIK(payload.noKK);
    var tglDatang = formatDate(payload.tglDatang || new Date());
    var asal = payload.asal || '';
    var alasan = payload.alasan || 'Pindah Alamat';
    var rw = normalizeRW(payload.rw || u.rw);
    var rt = String(payload.rt || u.rt).trim();
    var dusun = payload.dusun || 'Dusun 1';
    var noRumah = payload.noRumah || '';

    if (!canAccessAllData(u)) {
      if (u.role === 'rt' && (rw !== u.rw || rt !== u.rt)) {
        return { success: false, message: 'Hanya bisa menambah untuk RT ' + u.rt };
      }
      if (u.role === 'rw' && rw !== u.rw) {
        return { success: false, message: 'Hanya bisa menambah untuk RW ' + u.rw };
      }
    }

    var now = new Date();
    var bulan = String(now.getMonth() + 1);
    var tahun = String(now.getFullYear());

    var sheetDatang = getSheet('DataDatang');
    var headersDatang = getHeaders('DataDatang');
    var sheetWarga = getSheet('DataWarga');
    var headersWarga = getHeaders('DataWarga');

    // Ambil NIK terdaftar di DataWarga
    var existingWarga = getData('DataWarga');
    var existingNikMap = {};
    existingWarga.forEach(function(w) {
      if (w.NIK) {
        existingNikMap[normalizeNIK(w.NIK)] = w['Nama Lengkap'] || w.Nama || 'Warga lain';
      }
    });

    var seenBatchNik = {};
    for (var i = 0; i < payload.members.length; i++) {
      var m = payload.members[i];
      var nikNorm = normalizeNIK(m.nik);
      if (!nikNorm || nikNorm.length < 16) {
        return { success: false, message: 'NIK pada Anggota #' + (i + 1) + ' harus 16 digit angka valid!' };
      }
      if (seenBatchNik[nikNorm]) {
        return { success: false, message: 'NIK ' + m.nik + ' terdeteksi ganda/sama dalam daftar keluarga ini!' };
      }
      seenBatchNik[nikNorm] = true;
      if (existingNikMap[nikNorm]) {
        return { success: false, message: 'NIK ' + m.nik + ' (' + m.nama + ') sudah terdaftar di Data Warga atas nama: ' + existingNikMap[nikNorm] + '! NIK wajib unik.' };
      }
    }

    var count = 0;
    payload.members.forEach(function(m) {
      if (!m.nik || !m.nama) return;
      var nik = normalizeNIK(m.nik);
      var nama = String(m.nama).trim();

      var dataDatangObj = {
        'No KK': noKK,
        'NIK': nik,
        'Nama Lengkap': nama,
        'Dusun': dusun,
        'RW': rw,
        'RT': rt,
        'No. Rumah': noRumah,
        'Jenis Kelamin': m.jk || 'Laki-laki',
        'Tempat Lahir': m.tmpLahir || '',
        'Tanggal Lahir': formatDate(m.tglLahir || ''),
        'Status Perkawinan': m.statusPerkawinan || '',
        'Pekerjaan': m.pekerjaan || '',
        'Pendidikan': m.pendidikan || '',
        'Agama': m.agama || 'Islam',
        'Hubungan Keluarga': m.hubungan || 'Anggota',
        'Nama Ayah': m.namaAyah || '-',
        'Nama Ibu': m.namaIbu || '-',
        'Tanggal Datang': tglDatang,
        'Asal Daerah': asal,
        'Alasan': alasan,
        'Tahun': tahun,
        'Bulan': bulan
      };
      dataDatangObj = formatDataObject(dataDatangObj);
      writeRowSafely('DataDatang', sheetDatang, headersDatang, sheetDatang.getLastRow() + 1, dataDatangObj);

      var dataWargaObj = {
        'No KK': noKK,
        'NIK': nik,
        'Nama Lengkap': nama,
        'Dusun': dusun,
        'RW': rw,
        'RT': rt,
        'No. Rumah': noRumah,
        'Jenis Kelamin': m.jk || 'Laki-laki',
        'Tempat Lahir': m.tmpLahir || '',
        'Tanggal Lahir': formatDate(m.tglLahir || ''),
        'Status Perkawinan': m.statusPerkawinan || '',
        'Pekerjaan': m.pekerjaan || '',
        'Pendidikan': m.pendidikan || '',
        'Agama': m.agama || 'Islam',
        'Hubungan Keluarga': m.hubungan || 'Anggota',
        'Nama Ayah': m.namaAyah || '-',
        'Nama Ibu': m.namaIbu || '-',
        'Tahun': tahun,
        'Bulan': bulan
      };
      dataWargaObj = formatDataObject(dataWargaObj);
      writeRowSafely('DataWarga', sheetWarga, headersWarga, sheetWarga.getLastRow() + 1, dataWargaObj);
      count++;
    });

    clearDataCache();
    updateLaporan();
    writeAuditLog(u.username, u.role, 'DATANG_KELUARGA', 'Datang 1 Keluarga (' + count + ' jiwa) KK: ' + noKK);
    return { success: true, count: count, message: 'Berhasil mendaftarkan ' + count + ' anggota keluarga ke Data Datang & Data Warga.' };
  } catch (e) {
    writeAuditLog(u.username, u.role, 'DATANG_KELUARGA_ERROR', e.message);
    return { success: false, message: e.message };
  }
}

function processPindahKeluarga(noKK, dateVal, destinationVal, reasonVal, u) {
  try {
    u = verifyUser(u);
    if (!u) return { success: false, message: 'Sesi tidak valid.' };
    if (!noKK || String(noKK).trim() === '') {
      return { success: false, message: 'No KK wajib diisi.' };
    }

    var cleanKK = normalizeNIK(noKK);
    var allWarga = getData('DataWarga');
    var familyMembers = allWarga.filter(function(r) {
      return normalizeNIK(r['No KK']) === cleanKK;
    });

    if (familyMembers.length === 0) {
      return { success: false, message: 'Tidak ditemukan anggota keluarga dengan No KK: ' + noKK };
    }

    var now = new Date();
    var bulan = String(now.getMonth() + 1);
    var tahun = String(now.getFullYear());
    var formattedDate = formatDate(dateVal || now);

    var sheetPergi = getSheet('DataPergi');
    var headersPergi = getHeaders('DataPergi');
    var sheetWarga = getSheet('DataWarga');

    var count = 0;
    familyMembers.forEach(function(member) {
      var dataPergi = {
        'No KK': member['No KK'] || '',
        'NIK': member['NIK'] || '',
        'Nama Lengkap': member['Nama Lengkap'] || '',
        'Dusun': member['Dusun'] || 'Dusun 1',
        'RW': member['RW'] || '',
        'RT': member['RT'] || '',
        'No. Rumah': member['No. Rumah'] || '',
        'Tanggal Pergi': formattedDate,
        'Tujuan': destinationVal || '',
        'Alasan': reasonVal || 'Pindah Satu Keluarga',
        'Tahun': tahun,
        'Bulan': bulan
      };
      writeRowSafely('DataPergi', sheetPergi, headersPergi, sheetPergi.getLastRow() + 1, dataPergi);
      backupData('DataWarga', member, u.username);
      count++;
    });

    var rowsToDelete = familyMembers.map(function(m) { return m._rowIndex + 2; }).sort(function(a, b) { return b - a; });
    rowsToDelete.forEach(function(rIdx) {
      sheetWarga.deleteRow(rIdx);
    });

    clearDataCache();
    updateLaporan();
    writeAuditLog(u.username, u.role, 'PINDAH_KELUARGA', 'Pindah 1 Keluarga (' + count + ' jiwa) KK: ' + noKK);
    return { success: true, count: count, message: 'Berhasil memindahkan ' + count + ' anggota keluarga ke Data Pergi.' };
  } catch (e) {
    writeAuditLog(u.username, u.role, 'PINDAH_KELUARGA_ERROR', e.message);
    return { success: false, message: e.message };
  }
}

function syncAllCorrelatedSheets(editedSheetName, data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var targetNik = normalizeNIK(data.NIK || data['NIK Bayi'] || '');
    var targetNama = (data['Nama Lengkap'] || data['Nama Bayi'] || '').trim().toLowerCase();
    var targetNoKk = normalizeNIK(data['No KK'] || '');

    var rwNorm = normalizeRW(data.RW);
    var rtClean = String(data.RT || '').trim();

    // 1. Jika di-edit di DataLahir atau DataDatang -> Sinkronkan ke DataWarga
    if (editedSheetName === 'DataDatang' || editedSheetName === 'DataLahir') {
      var wsW = ss.getSheetByName('DataWarga');
      if (wsW) {
        var wdW = getData('DataWarga');
        var wIdx = -1;
        for (var iW = 0; iW < wdW.length; iW++) {
          var wNik = normalizeNIK(wdW[iW].NIK || '');
          var wNama = (wdW[iW]['Nama Lengkap'] || '').trim().toLowerCase();
          var wNoKk = normalizeNIK(wdW[iW]['No KK'] || '');

          if (targetNik && targetNik.toUpperCase() !== 'BELUM ADA' && targetNik.toUpperCase() !== 'BELUMADA' && wNik === targetNik) {
            wIdx = wdW[iW]._rowIndex;
            break;
          }
          if (targetNama && wNama === targetNama && (targetNoKk ? wNoKk === targetNoKk : true)) {
            wIdx = wdW[iW]._rowIndex;
            break;
          }
        }
        if (wIdx !== -1) {
          var whW = getHeaders('DataWarga');
          var rowW = wIdx + 2;
          var mapDataW = {
            'No KK': data['No KK'],
            'NIK': (targetNik && targetNik.toUpperCase() !== 'BELUM ADA' && targetNik.toUpperCase() !== 'BELUMADA') ? targetNik : 'Belum Ada',
            'Nama Lengkap': data['Nama Bayi'] || data['Nama Lengkap'],
            'Dusun': 'Dusun 1',
            'RW': rwNorm,
            'RT': rtClean,
            'No. Rumah': data['No. Rumah'],
            'Jenis Kelamin': data['Jenis Kelamin'],
            'Tempat Lahir': data['Tempat Lahir'],
            'Tanggal Lahir': formatDate(data['Tanggal Lahir'])
          };
          for (var colName in mapDataW) {
            if (mapDataW[colName] !== undefined) {
              var cIdx = whW.indexOf(colName);
              if (cIdx !== -1) {
                var cellW = wsW.getRange(rowW, cIdx + 1);
                if (TEXT_COLUMNS['DataWarga'] && TEXT_COLUMNS['DataWarga'].indexOf(colName) !== -1) cellW.setNumberFormat('@');
                cellW.setValue(mapDataW[colName] || '');
              }
            }
          }
        }
      }
    }

    // 2. Jika di-edit di DataWarga -> Sinkronkan ke DataLahir, DataDatang, DataPergi, DataMeninggal, DataPengontrak
    if (editedSheetName === 'DataWarga') {
      var correlatedSheets = ['DataLahir', 'DataDatang', 'DataPergi', 'DataMeninggal', 'DataPengontrak'];
      correlatedSheets.forEach(function (cSheetName) {
        var sObj = ss.getSheetByName(cSheetName);
        if (!sObj) return;
        var cData = getData(cSheetName);
        var cHeaders = getHeaders(cSheetName);

        var nameColHeader = (cSheetName === 'DataLahir') ? 'Nama Bayi' : 'Nama Lengkap';
        var nikColHeader = (cSheetName === 'DataLahir') ? 'NIK Bayi' : 'NIK';

        for (var iC = 0; iC < cData.length; iC++) {
          var cNik = normalizeNIK(cData[iC][nikColHeader] || cData[iC].NIK || '');
          var cNama = (cData[iC][nameColHeader] || cData[iC]['Nama Lengkap'] || '').trim().toLowerCase();
          var cNoKk = normalizeNIK(cData[iC]['No KK'] || '');

          var isMatch = false;
          if (targetNik && targetNik.toUpperCase() !== 'BELUM ADA' && targetNik.toUpperCase() !== 'BELUMADA' && cNik === targetNik) {
            isMatch = true;
          } else if (targetNama && cNama === targetNama && (targetNoKk ? cNoKk === targetNoKk : true)) {
            isMatch = true;
          }

          if (isMatch) {
            var rowC = cData[iC]._rowIndex + 2;
            var updateFields = {
              'No KK': data['No KK'],
              'RW': rwNorm,
              'RT': rtClean,
              'No. Rumah': data['No. Rumah'],
              'Jenis Kelamin': data['Jenis Kelamin'],
              'Tempat Lahir': data['Tempat Lahir'],
              'Tanggal Lahir': formatDate(data['Tanggal Lahir'])
            };
            updateFields[nameColHeader] = data['Nama Lengkap'] || data['Nama Bayi'];
            if (targetNik && targetNik.toUpperCase() !== 'BELUM ADA') updateFields[nikColHeader] = targetNik;

            for (var fKey in updateFields) {
              var colIdxC = cHeaders.indexOf(fKey);
              if (colIdxC !== -1) {
                var cellC = sObj.getRange(rowC, colIdxC + 1);
                if (TEXT_COLUMNS[cSheetName] && TEXT_COLUMNS[cSheetName].indexOf(fKey) !== -1) cellC.setNumberFormat('@');
                cellC.setValue(updateFields[fKey] || '');
              }
            }
          }
        }
      });
    }
  } catch (errSync) {
    console.error('syncAllCorrelatedSheets error:', errSync);
  }
}

function updateData(n, idx, data, u) {
  try {
    // Verifikasi user terhadap Users sheet
    u = verifyUser(u);
    if (!u) return {success: false, message: 'Sesi tidak valid. Silakan login ulang.'};

    // Format Kapitalisasi
    data = formatDataObject(data);

    // Sanitasi input terhadap formula injection
    var safeFields = ['RW', 'RT', 'NIK', 'No KK', 'Tahun', 'Bulan'];
    data = sanitizeDataObject(data, safeFields);

    var allData = getData(n);
    var existingRow = null;
    for(var i=0; i<allData.length; i++) {
      if(allData[i]._rowIndex === idx) {
        existingRow = allData[i];
        break;
      }
    }
    if(!existingRow) {
      return {success:false, message:'Data tidak ditemukan'};
    }
    
    var errors = validateData(n, data, true, existingRow);
    if (errors.length > 0) {
      writeAuditLog(u.username, u.role, 'UPDATE_DATA_FAILED', errors.join('; '));
      return {success: false, message: errors.join('; ')};
    }

    if(data.RW) data.RW = normalizeRW(data.RW);
    if(data.RT) data.RT = String(data.RT).trim();
    if(data.hasOwnProperty('NIK')) data.NIK = normalizeNIK(data.NIK);

    // Cek akses RT/RW
    var accessErr2 = _checkDataAccess(u, data, 'UPDATE_DATA');
    if (accessErr2) return accessErr2;

    var s = getSheet(n);
    var h = getHeaders(n);
    var row = idx+2;

    var existingRow2 = s.getRange(row,1,1,h.length).getValues()[0];

    var tahunIdx = h.indexOf('Tahun');
    var bulanIdx = h.indexOf('Bulan');
    if(tahunIdx>-1 && existingRow2[tahunIdx]) data['Tahun'] = existingRow2[tahunIdx];
    if(bulanIdx>-1 && existingRow2[bulanIdx]) data['Bulan'] = existingRow2[bulanIdx];

    if(tahunIdx>-1 && (!data['Tahun'] || data['Tahun']==='')) {
      var dValTahun = data['Tanggal Datang'] || data['Tanggal Pergi'] || data['Tanggal Meninggal'] || data['Tanggal Lahir'];
      var dObjTahun = dValTahun ? new Date(dValTahun) : new Date();
      data['Tahun'] = String(isNaN(dObjTahun.getTime()) ? new Date().getFullYear() : dObjTahun.getFullYear());
    }
    if(bulanIdx>-1 && (!data['Bulan'] || data['Bulan']==='')) {
      var dValBulan = data['Tanggal Datang'] || data['Tanggal Pergi'] || data['Tanggal Meninggal'] || data['Tanggal Lahir'];
      var dObjBulan = dValBulan ? new Date(dValBulan) : new Date();
      data['Bulan'] = String(isNaN(dObjBulan.getTime()) ? (new Date().getMonth() + 1) : (dObjBulan.getMonth() + 1));
    }

    var textCols = TEXT_COLUMNS[n] || [];
    h.forEach(function(k,i){
      var cell = s.getRange(row,i+1);
      if(textCols.indexOf(k)!==-1) cell.setNumberFormat('@');
      cell.setValue(data[k]||'');
    });

    syncAllCorrelatedSheets(n, data);

    clearDataCache();
    updateLaporan();
    writeAuditLog(u.username, u.role, 'UPDATE_DATA', n + ': ' + (data['Nama Lengkap'] || data['Nama Bayi'] || ''));
    return {success:true, message:'Data berhasil diupdate'};
  } catch(e){ 
    writeAuditLog(u.username, u.role, 'UPDATE_DATA_ERROR', e.message);
    return {success:false, message:e.message}; 
  }
}

function deleteData(n, idx, u) {
  try {
    // Verifikasi user terhadap Users sheet
    u = verifyUser(u);
    if (!u) return {success: false, message: 'Sesi tidak valid. Silakan login ulang.'};

    var all = getData(n);
    var data = null;
    for(var i=0;i<all.length;i++){
      if(all[i]._rowIndex === idx){ data = all[i]; break; }
    }
    if(!data) return {success:false, message:'Data tidak ditemukan'};

    if(data.Dusun!=='Dusun 1') {
      writeAuditLog(u.username, u.role, 'DELETE_DATA_FAILED', 'Data bukan Dusun 1');
      return {success:false, message:'Data harus Dusun 1'};
    }
    
    // Cek akses RT/RW
    var accessErr3 = _checkDataAccess(u, data, 'DELETE_DATA');
    if (accessErr3) return accessErr3;
    
    backupData(n, data, u.username);
    
    var s = getSheet(n);
    s.deleteRow(idx+2);

    if(n==='DataDatang' && data.NIK){
      var wdW = getData('DataWarga');
      var wIdx = -1;
      for(var iW=0; iW<wdW.length; iW++){
        if(wdW[iW].NIK === normalizeNIK(data.NIK)){
          wIdx = wdW[iW]._rowIndex;
          break;
        }
      }
      if(wIdx !== -1){
        var wsW = getSheet('DataWarga');
        backupData('DataWarga', wdW[wIdx], u.username);
        wsW.deleteRow(wIdx + 2);
      }
    }

    clearDataCache();
    updateLaporan();
    writeAuditLog(u.username, u.role, 'DELETE_DATA', n + ': ' + (data['Nama Lengkap'] || data['Nama Bayi'] || '') + ' (dibackup)');
    return {success:true, message:'Data berhasil dihapus dan dibackup'};
  } catch(e){ 
    writeAuditLog(u.username, u.role, 'DELETE_DATA_ERROR', e.message);
    return {success:false, message:e.message}; 
  }
}

// ============================================================
// LAPORAN BULANAN
// ============================================================

function updateLaporan() {
  try {
    var s = getSheet('LaporanBulanan');

    var warga = getData('DataWarga').filter(function(r){ return r.Dusun === 'Dusun 1'; });
    var datang = getData('DataDatang').filter(function(r){ return r.Dusun === 'Dusun 1'; });
    var pergi = getData('DataPergi').filter(function(r){ return r.Dusun === 'Dusun 1'; });
    var meninggal = getData('DataMeninggal').filter(function(r){ return r.Dusun === 'Dusun 1'; });
    var lahir = getData('DataLahir').filter(function(r){ return r.Dusun === 'Dusun 1'; });

    var groupKeys = {};
    function regKey(rw,rt){
      var normRw = normalizeRW(rw);
      var normRt = String(rt || '').trim();
      var k = normRw + '|' + normRt;
      if(!groupKeys[k]) groupKeys[k] = {RW: normRw, RT: normRt};
      return k;
    }

    var akhirSekarang = {}, lakiSekarang = {}, perempuanSekarang = {}, kkSekarang = {};
    warga.forEach(function(w){
      if(!w.RW) return; // hanya skip jika tidak ada RW, RT boleh kosong
      var k = regKey(w.RW, w.RT || '0'); // default RT '0' jika kosong
      akhirSekarang[k] = (akhirSekarang[k]||0) + 1;
      // FIX: gunakan case-insensitive agar cocok dengan data huruf besar (LAKI-LAKI, PEREMPUAN)
      var jkNorm = String(w['Jenis Kelamin'] || '').trim().toLowerCase();
      if(jkNorm === 'laki-laki' || jkNorm === 'l') lakiSekarang[k] = (lakiSekarang[k]||0)+1;
      if(jkNorm === 'perempuan' || jkNorm === 'p') perempuanSekarang[k] = (perempuanSekarang[k]||0)+1;
      if(!kkSekarang[k]) kkSekarang[k] = new Set();
      if(w['No KK']) kkSekarang[k].add(w['No KK']);
    });

    var events = {};
    function parseYearMonth(row, dateFieldName) {
      // Prioritas 1: baca dari field Bulan dan Tahun (waktu penginputan)
      var y = row.Tahun ? String(row.Tahun).trim() : '';
      var m = row.Bulan ? String(row.Bulan).trim() : '';
      if (y && m) {
        var mNum = parseInt(m, 10);
        if (isNaN(mNum)) {
          var bulanNames = ['januari','februari','maret','april','mei','juni','juli','agustus','september','oktober','november','desember'];
          var idx = bulanNames.indexOf(m.toLowerCase());
          if (idx !== -1) mNum = idx + 1;
        }
        if (mNum >= 1 && mNum <= 12) {
          return y + '-' + String(mNum).padStart(2, '0');
        }
      }

      // Prioritas 2: baca dari field tanggal event jika Bulan/Tahun di sheet lama belum terisi
      if (dateFieldName && row[dateFieldName]) {
        var dObj = row[dateFieldName] instanceof Date
          ? row[dateFieldName]
          : new Date(row[dateFieldName]);
        if (!isNaN(dObj.getTime())) {
          return dObj.getFullYear() + '-' + String(dObj.getMonth() + 1).padStart(2, '0');
        }
      }

      var now = new Date();
      return now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
    }

    function addEvent(row, tipe, dateFieldName){
      if(!row.RW) return; // hanya skip jika tidak ada RW
      var k = regKey(row.RW, row.RT || '0'); // default RT '0' jika kosong
      var ym = parseYearMonth(row, dateFieldName);
      if(!events[k]) events[k] = {};
      if(!events[k][ym]) events[k][ym] = {datang:0,pergi:0,meninggal:0,lahir:0};
      events[k][ym][tipe]++;
    }
    datang.forEach(function(r){ addEvent(r, 'datang', 'Tanggal Datang'); });
    pergi.forEach(function(r){ addEvent(r, 'pergi', 'Tanggal Pergi'); });
    meninggal.forEach(function(r){ addEvent(r, 'meninggal', 'Tanggal Meninggal'); });
    lahir.forEach(function(r){ addEvent(r, 'lahir', 'Tanggal Lahir'); });

    var now = new Date();
    var curYM = now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0');

    var results = [];

    Object.keys(groupKeys).forEach(function(k){
      var g = groupKeys[k];
      var evMap = events[k] || {};
      var allYM = Object.keys(evMap);
      if(allYM.indexOf(curYM)===-1) allYM.push(curYM);
      allYM.sort();
      var earliestYM = allYM[0];

      var ymList = [];
      var parts = earliestYM.split('-');
      var y = Number(parts[0]), m = Number(parts[1]);
      var ny = now.getFullYear(), nm = now.getMonth()+1;
      var guard = 0;
      while((y < ny || (y===ny && m<=nm)) && guard < 1200){
        ymList.push(y+'-'+String(m).padStart(2,'0'));
        m++;
        if(m>12){ m=1; y++; }
        guard++;
      }
      if(ymList.length===0) ymList = [curYM];

      var akhirNow = akhirSekarang[k]||0;
      var lakiNow = lakiSekarang[k]||0;
      var perempuanNow = perempuanSekarang[k]||0;
      var kkNow = kkSekarang[k] ? kkSekarang[k].size : 0;

      var akhirMap = {};
      akhirMap[ymList[ymList.length-1]] = akhirNow;
      for(var i=ymList.length-1;i>=1;i--){
        var ymNow2 = ymList[i];
        var ymPrev = ymList[i-1];
        var ev = evMap[ymNow2] || {datang:0,pergi:0,meninggal:0,lahir:0};
        var akhirIni = akhirMap[ymNow2] !== undefined ? akhirMap[ymNow2] : akhirNow;
        var awalIni = akhirIni - ev.datang + ev.pergi + ev.meninggal - ev.lahir;
        akhirMap[ymPrev] = awalIni;
      }

      ymList.forEach(function(ym){
        var ev = evMap[ym] || {datang:0,pergi:0,meninggal:0,lahir:0};
        var akhirBln = akhirMap[ym] !== undefined ? akhirMap[ym] : akhirNow;
        var awalBln = akhirBln - ev.datang + ev.pergi + ev.meninggal - ev.lahir;
        var p2 = ym.split('-');
        var isSekarang = (ym===curYM);
        results.push([
          Number(p2[1]), Number(p2[0]), 'Dusun 1', g.RW, g.RT,
          Math.max(0, awalBln), ev.datang, ev.pergi, ev.meninggal, ev.lahir,
          Math.max(0, akhirBln),
          isSekarang ? lakiNow : '',
          isSekarang ? perempuanNow : '',
          isSekarang ? kkNow : ''
        ]);
      });
    });

    var lastRow = s.getLastRow();
    if(lastRow>1) s.getRange(2,1,lastRow-1, s.getMaxColumns()).clearContent();

    if(results.length){
      s.getRange(2, 4, results.length, 1).setNumberFormat('@');
      s.getRange(2, 5, results.length, 1).setNumberFormat('@');
      s.getRange(2,1,results.length, results[0].length).setValues(results);
    }

  } catch(e){
    console.error('Laporan error:', e);
  }
}

function getLaporan(u, filter) {
  try {
    updateLaporan();
    var s = getSheet('LaporanBulanan');
    var data = s.getDataRange().getValues();
    
    if(data.length < 2) {
      var now = new Date();
      var curB = String(now.getMonth()+1);
      var curT = String(now.getFullYear());
      
      return {
        data: [{
          'Bulan': curB,
          'Tahun': curT,
          'Dusun': 'Dusun 1',
          'RW': '00',
          'RT': '0',
          'Jumlah Awal': 0,
          'Datang': 0,
          'Pergi': 0,
          'Meninggal': 0,
          'Lahir': 0,
          'Jumlah Akhir': 0,
          'Laki-laki': 0,
          'Perempuan': 0,
          'KK': 0
        }],
        totals: {awal:0, datang:0, pergi:0, meninggal:0, lahir:0, akhir:0, laki:0, perempuan:0, kk:0}
      };
    }

    var rows = data.slice(1).filter(function(row){
      return row.some(function(c){ return String(c).trim()!==''; });
    }).map(function(row){
      // Normalisasi Bulan dan Tahun: parseInt menghindari float "8.0" vs "8"
      var bulanRaw = String(row[0] || '').trim();
      var tahunRaw = String(row[1] || '').trim();
      var bulanNum = parseInt(bulanRaw, 10);
      var tahunNum = parseInt(tahunRaw, 10);
      return {
        'Bulan': isNaN(bulanNum) ? bulanRaw : String(bulanNum),
        'Tahun': isNaN(tahunNum) ? tahunRaw : String(tahunNum),
        'Dusun': String(row[2] || ''),
        'RW': normalizeRW(row[3] || ''),
        'RT': String(row[4] || '').trim(),
        'Jumlah Awal': Number(row[5] || 0),
        'Datang': Number(row[6] || 0),
        'Pergi': Number(row[7] || 0),
        'Meninggal': Number(row[8] || 0),
        'Lahir': Number(row[9] || 0),
        'Jumlah Akhir': Number(row[10] || 0),
        'Laki-laki': Number(row[11] || 0),
        'Perempuan': Number(row[12] || 0),
        'KK': Number(row[13] || 0)
      };
    });

    var current = rows.filter(function(r){
      return r.Dusun === 'Dusun 1';
    });

    var filtered = applyFilters(current, filter, u);

    var hasBulanTahunFilter = filter && ((filter.bulan && filter.bulan !== '') || (filter.tahun && filter.tahun !== ''));
    if (hasBulanTahunFilter) {
      if (filter.bulan && filter.bulan !== '') {
        filtered = filtered.filter(function(r) {
          return String(r.Bulan) === String(filter.bulan);
        });
      }
      if (filter.tahun && filter.tahun !== '') {
        filtered = filtered.filter(function(r) {
          return String(r.Tahun) === String(filter.tahun);
        });
      }
    } else {
      var now = new Date();
      var curB = String(now.getMonth()+1), curT = String(now.getFullYear());
      filtered = filtered.filter(function(r) {
        return String(r.Bulan) === curB && String(r.Tahun) === curT;
      });
    }

    var totals = {
      awal: 0, datang: 0, pergi: 0, meninggal: 0,
      lahir: 0, akhir: 0, laki: 0, perempuan: 0, kk: 0
    };

    filtered.forEach(function(r) {
      totals.awal += Number(r['Jumlah Awal']) || 0;
      totals.datang += Number(r['Datang']) || 0;
      totals.pergi += Number(r['Pergi']) || 0;
      totals.meninggal += Number(r['Meninggal']) || 0;
      totals.lahir += Number(r['Lahir']) || 0;
      totals.akhir += Number(r['Jumlah Akhir']) || 0;
      totals.laki += Number(r['Laki-laki']) || 0;
      totals.perempuan += Number(r['Perempuan']) || 0;
      totals.kk += Number(r['KK']) || 0;
    });

    return {
      data: filtered,
      totals: totals
    };
  } catch(e) {
    console.error('getLaporan error:', e);
    var now = new Date();
    return {
      data: [],
      totals: {awal:0, datang:0, pergi:0, meninggal:0, lahir:0, akhir:0, laki:0, perempuan:0, kk:0}
    };
  }
}

// ============================================================
// STATISTIK & VISUALISASI
// ============================================================

function getRWStats() {
  var w = getData('DataWarga').filter(function(r){ return r.Dusun==='Dusun 1'; });
  // 6 RW tetap yang selalu tampil (hardcoded seperti referensi)
  var stats = {'13':0, '04':0, '02':0, '03':0, '12':0, '15':0};
  w.forEach(function(item){
    if(item.RW) {
      var rw = normalizeRW(item.RW);
      if(rw) stats[rw] = (stats[rw] || 0) + 1;
    }
  });
  return stats;
}

function getVisualData(u) {
  try {
    var w = getData('DataWarga').filter(function(r){ return r.Dusun==='Dusun 1'; });
    
    if(!canAccessAllData(u)){
      if(u.role==='rw') {
        w = w.filter(function(r){ return String(r.RW)===String(u.rw); });
      } else if(u.role==='rt') {
        w = w.filter(function(r){ return String(r.RW)===String(u.rw) && String(r.RT)===String(u.rt); });
      }
    }

    // Gunakan getRWStats yang sudah punya 6 RW tetap sebagai baseline
    var stats = getRWStats();

    // Urutkan: 02, 03, 04, 12, 13, 15 (numerik)
    var sortedRWKeys = Object.keys(stats).sort(function(a, b){
      return parseInt(a, 10) - parseInt(b, 10);
    });
    var rwLabels = sortedRWKeys.map(function(k){ return 'RW ' + k; });
    var rwValues = sortedRWKeys.map(function(k){ return stats[k] || 0; });

    var datang = getData('DataDatang').filter(function(r){ return r.Dusun==='Dusun 1'; });
    var pergi = getData('DataPergi').filter(function(r){ return r.Dusun==='Dusun 1'; });
    var meninggal = getData('DataMeninggal').filter(function(r){ return r.Dusun==='Dusun 1'; });
    var lahir = getData('DataLahir').filter(function(r){ return r.Dusun==='Dusun 1'; });

    // Filter tambahan jika role bukan admin/superadmin
    if(!canAccessAllData(u)){
      if(u.role==='rw') {
        datang    = datang.filter(function(r){ return String(r.RW)===String(u.rw); });
        pergi     = pergi.filter(function(r){ return String(r.RW)===String(u.rw); });
        meninggal = meninggal.filter(function(r){ return String(r.RW)===String(u.rw); });
        lahir     = lahir.filter(function(r){ return String(r.RW)===String(u.rw); });
      } else if(u.role==='rt') {
        datang    = datang.filter(function(r){ return String(r.RW)===String(u.rw) && String(r.RT)===String(u.rt); });
        pergi     = pergi.filter(function(r){ return String(r.RW)===String(u.rw) && String(r.RT)===String(u.rt); });
        meninggal = meninggal.filter(function(r){ return String(r.RW)===String(u.rw) && String(r.RT)===String(u.rt); });
        lahir     = lahir.filter(function(r){ return String(r.RW)===String(u.rw) && String(r.RT)===String(u.rt); });
      }
    }

    return {
      rwStats: stats,
      rwLabels: rwLabels,
      rwValues: rwValues,
      totalWarga: w.length,
      totalDatang: datang.length,
      totalPergi: pergi.length,
      totalMeninggal: meninggal.length,
      totalLahir: lahir.length
    };
  } catch(e){
    return {
      rwStats: {'13':0, '04':0, '02':0, '03':0, '12':0, '15':0},
      rwLabels: ['RW 02','RW 03','RW 04','RW 12','RW 13','RW 15'],
      rwValues: [0,0,0,0,0,0],
      totalWarga: 0, totalDatang: 0, totalPergi: 0, totalMeninggal: 0, totalLahir: 0
    };
  }
}

// ============================================================
// REPAIR DATA LAMA
// ============================================================

function repairRWData() {
  var sheetsToFix = ['DataWarga','DataDatang','DataPergi','DataMeninggal','DataLahir','LaporanBulanan','Users','Backup'];
  var report = [];

  sheetsToFix.forEach(function(n){
    var s = getSheet(n);
    if(!s) return;
    var h = getHeaders(n).length ? getHeaders(n) : (n==='Users' ? ['Username','Password','Role','Dusun','RW','RT','Label'] : (n==='Backup' ? ['Tanggal Hapus','Sheet Asal','No KK','NIK','Nama Lengkap','Dusun','RW','RT','Jenis Kelamin','Tempat Lahir','Tanggal Lahir','Status Perkawinan','Pekerjaan','Pendidikan','Agama','Tahun','Bulan','Data Lengkap JSON','Dihapus Oleh'] : []));
    if(!h.length) return;

    var rwIdx = h.indexOf('RW');
    if(rwIdx === -1) return;

    var lastRow = s.getLastRow();
    if(lastRow < 2) return;

    s.getRange(2, rwIdx+1, lastRow-1, 1).setNumberFormat('@');

    var vals = s.getRange(2, rwIdx+1, lastRow-1, 1).getValues();
    var fixed = 0;
    var newVals = vals.map(function(row){
      var original = String(row[0]||'');
      var norm = normalizeRW(original);
      if(norm !== original) fixed++;
      return [norm];
    });
    s.getRange(2, rwIdx+1, lastRow-1, 1).setValues(newVals);
    report.push(n + ': ' + fixed + ' baris diperbaiki');
  });

  updateLaporan();
  writeAuditLog('system', 'system', 'REPAIR_RW', 'Data RW diperbaiki');

  try {
    SpreadsheetApp.getUi().alert('Perbaikan selesai!\n\n' + report.join('\n'));
  } catch(e){
    console.log(report.join('\n'));
  }
  return report;
}

// ============================================================
// MANAJEMEN USER DENGAN SALT
// ============================================================

function getUserManagementData(adminUser) {
  try {
    if (!adminUser || !isSuperAdmin(adminUser)) {
      return {success: false, message: 'Hanya SuperAdmin yang bisa mengakses'};
    }
    
    var users = getUsers();
    var userList = [];
    for (var username in users) {
      if (users.hasOwnProperty(username)) {
        var u = users[username];
        userList.push({
          username: username,
          role: u.role,
          dusun: u.dusun || '',
          rw: u.rw || '',
          rt: u.rt || '',
          label: u.label || username,
          hasPassword: u.passwordHash ? true : false
        });
      }
    }
    
    userList.sort(function(a, b) {
      var roleOrder = {superadmin: 0, kadus: 1, rw: 2, rt: 3};
      return (roleOrder[a.role] || 99) - (roleOrder[b.role] || 99);
    });
    
    return {success: true, users: userList};
  } catch(e) {
    return {success: false, message: e.message};
  }
}

function addUser(username, password, role, dusun, rw, rt, label, adminUser) {
  try {
    adminUser = verifyUser(adminUser);
    if(!adminUser || !isSuperAdmin(adminUser)) {
      writeAuditLog(adminUser ? adminUser.username : 'unknown', adminUser ? adminUser.role : 'unknown', 'ADD_USER_FAILED', 'Akses ditolak: bukan SuperAdmin');
      return {success:false, message:'Hanya SuperAdmin yang bisa menambah user!'};
    }
    
    if (!username || username.trim() === '') {
      return {success:false, message:'Username tidak boleh kosong!'};
    }
    
    username = username.trim();
    
    var sheet = getSheet('Users');
    var lastRow = sheet.getLastRow();
    var newRow = lastRow + 1;
    
    var existing = getUsers();
    if(existing[username]) {
      writeAuditLog(adminUser.username, adminUser.role, 'ADD_USER_FAILED', 'Username ' + username + ' sudah terdaftar');
      return {success:false, message:'Username sudah terdaftar!'};
    }
    
    if(!password || password.length < 6) {
      return {success:false, message:'Password minimal 6 karakter!'};
    }
    
    if(rw) rw = normalizeRW(rw);
    
    var salt = generateSalt();
    var hashedPassword = hashPasswordWithSalt(password, salt);
    
    var rowData = [username, hashedPassword, role, dusun || 'Dusun 1', rw || '', rt || '', label || username, salt];
    sheet.getRange(newRow, 1, 1, 8).setValues([rowData]);
    
    sheet.getRange(newRow, 5, 1, 1).setNumberFormat('@');
    sheet.getRange(newRow, 6, 1, 1).setNumberFormat('@');
    
    writeAuditLog(adminUser.username, adminUser.role, 'ADD_USER', 'User ' + username + ' (' + role + ') ditambahkan dengan salt');
    return {success:true, message:'User berhasil ditambahkan!'};
  } catch(e) {
    writeAuditLog(adminUser.username, adminUser.role, 'ADD_USER_ERROR', e.message);
    return {success:false, message:e.message};
  }
}

function editUser(oldUsername, newUsername, password, role, dusun, rw, rt, label, adminUser) {
  try {
    adminUser = verifyUser(adminUser);
    if(!adminUser || !isSuperAdmin(adminUser)) {
      writeAuditLog(adminUser ? adminUser.username : 'unknown', adminUser ? adminUser.role : 'unknown', 'EDIT_USER_FAILED', 'Akses ditolak: bukan SuperAdmin');
      return {success:false, message:'Hanya SuperAdmin yang bisa mengedit user!'};
    }
    
    oldUsername = oldUsername.trim();
    newUsername = newUsername.trim();
    
    if(oldUsername === 'superadmin' && newUsername !== 'superadmin') {
      return {success:false, message:'Tidak bisa mengubah username SuperAdmin!'};
    }
    
    var sheet = getSheet('Users');
    var data = sheet.getDataRange().getValues();
    
    var rowIdx = -1;
    for(var i=1; i<data.length; i++) {
      if(String(data[i][0]||'').trim() === oldUsername) {
        rowIdx = i+1;
        break;
      }
    }
    
    if(rowIdx === -1) {
      return {success:false, message:'User tidak ditemukan!'};
    }
    
    if(oldUsername !== newUsername) {
      var existing = getUsers();
      if(existing[newUsername]) {
        return {success:false, message:'Username baru sudah terdaftar!'};
      }
    }
    
    if(rw) rw = normalizeRW(rw);
    
    var newPassword;
    var newSalt = data[rowIdx-1][7] || '';
    
    if(password && password.length > 0) {
      if(password.length < 6) {
        return {success:false, message:'Password minimal 6 karakter!'};
      }
      newSalt = generateSalt();
      newPassword = hashPasswordWithSalt(password, newSalt);
    } else {
      newPassword = data[rowIdx-1][1];
      newSalt = data[rowIdx-1][7] || '';
    }
    
    var rowData = [newUsername, newPassword, role, dusun || 'Dusun 1', rw || '', rt || '', label || newUsername, newSalt];
    
    sheet.getRange(rowIdx, 5, 1, 1).setNumberFormat('@');
    sheet.getRange(rowIdx, 6, 1, 1).setNumberFormat('@');
    
    sheet.getRange(rowIdx, 1, 1, 8).setValues([rowData]);
    
    writeAuditLog(adminUser.username, adminUser.role, 'EDIT_USER', 'User ' + oldUsername + ' diupdate menjadi ' + newUsername);
    return {success:true, message:'User berhasil diupdate!'};
  } catch(e) {
    writeAuditLog(adminUser.username, adminUser.role, 'EDIT_USER_ERROR', e.message);
    return {success:false, message:e.message};
  }
}

function deleteUser(username, adminUser) {
  try {
    adminUser = verifyUser(adminUser);
    if(!adminUser || !isSuperAdmin(adminUser)) {
      writeAuditLog(adminUser ? adminUser.username : 'unknown', adminUser ? adminUser.role : 'unknown', 'DELETE_USER_FAILED', 'Akses ditolak: bukan SuperAdmin');
      return {success:false, message:'Hanya SuperAdmin yang bisa menghapus user!'};
    }
    
    username = username.trim();
    
    if(username === 'superadmin') {
      writeAuditLog(adminUser.username, adminUser.role, 'DELETE_USER_FAILED', 'Mencoba menghapus SuperAdmin');
      return {success:false, message:'Tidak bisa menghapus SuperAdmin!'};
    }
    
    var sheet = getSheet('Users');
    var data = sheet.getDataRange().getValues();
    
    var rowIdx = -1;
    for(var i=1; i<data.length; i++) {
      if(String(data[i][0]||'').trim() === username) {
        rowIdx = i+1;
        break;
      }
    }
    
    if(rowIdx === -1) {
      return {success:false, message:'User tidak ditemukan!'};
    }
    
    sheet.deleteRow(rowIdx);
    writeAuditLog(adminUser.username, adminUser.role, 'DELETE_USER', 'User ' + username + ' dihapus');
    return {success:true, message:'User berhasil dihapus!'};
  } catch(e) {
    writeAuditLog(adminUser.username, adminUser.role, 'DELETE_USER_ERROR', e.message);
    return {success:false, message:e.message};
  }
}

function getAllUsers(adminUser) {
  try {
    adminUser = verifyUser(adminUser);
    if(!adminUser || !isSuperAdmin(adminUser)) {
      return [];
    }
    
    var sheet = getSheet('Users');
    var data = sheet.getDataRange().getValues();
    if(data.length < 2) return [];
    
    var users = [];
    for(var i=1; i<data.length; i++) {
      users.push({
        username: String(data[i][0]||'').trim(),
        role: String(data[i][2]||'').trim(),
        dusun: String(data[i][3]||'').trim(),
        rw: normalizeRW(data[i][4]),
        rt: String(data[i][5]||'').trim(),
        label: String(data[i][6]||'').trim()
      });
    }
    return users;
  } catch(e) {
    return [];
  }
}

function resetPassword(username, newPassword, adminUser) {
  try {
    adminUser = verifyUser(adminUser);
    if(!adminUser || !isSuperAdmin(adminUser)) {
      writeAuditLog(adminUser ? adminUser.username : 'unknown', adminUser ? adminUser.role : 'unknown', 'RESET_PASSWORD_FAILED', 'Akses ditolak: bukan SuperAdmin');
      return {success:false, message:'Hanya SuperAdmin yang bisa reset password!'};
    }
    
    username = username.trim();
    
    if(!newPassword || newPassword.length < 6) {
      return {success:false, message:'Password minimal 6 karakter!'};
    }
    
    var sheet = getSheet('Users');
    var data = sheet.getDataRange().getValues();
    
    var rowIdx = -1;
    for(var i=1; i<data.length; i++) {
      if(String(data[i][0]||'').trim() === username) {
        rowIdx = i+1;
        break;
      }
    }
    
    if(rowIdx === -1) {
      return {success:false, message:'User tidak ditemukan!'};
    }
    
    var salt = generateSalt();
    var hashedPassword = hashPasswordWithSalt(newPassword, salt);
    
    sheet.getRange(rowIdx, 2).setValue(hashedPassword);
    sheet.getRange(rowIdx, 8).setValue(salt);
    
    writeAuditLog(adminUser.username, adminUser.role, 'RESET_PASSWORD', 'Password ' + username + ' direset dengan salt baru');
    return {success:true, message:'Password berhasil direset!'};
  } catch(e) {
    writeAuditLog(adminUser.username, adminUser.role, 'RESET_PASSWORD_ERROR', e.message);
    return {success:false, message:e.message};
  }
}

function changeMyPassword(username, oldPassword, newPassword, confirmPassword) {
  try {
    var users = getUsers();
    if (!users[username]) {
      return {success: false, message: 'User tidak ditemukan'};
    }
    
    if (!verifyPasswordWithSalt(oldPassword, users[username].passwordHash, users[username].salt)) {
      writeAuditLog(username, users[username].role, 'CHANGE_PASSWORD_FAILED', 'Password lama salah');
      return {success: false, message: 'Password lama salah!'};
    }
    
    if (newPassword.length < 6) {
      return {success: false, message: 'Password minimal 6 karakter!'};
    }
    
    if (newPassword !== confirmPassword) {
      return {success: false, message: 'Konfirmasi password tidak cocok!'};
    }
    
    var sheet = getSheet('Users');
    var data = sheet.getDataRange().getValues();
    var rowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === username) {
        rowIdx = i + 1;
        break;
      }
    }
    
    if (rowIdx === -1) {
      return {success: false, message: 'User tidak ditemukan'};
    }
    
    var salt = generateSalt();
    var hashedPassword = hashPasswordWithSalt(newPassword, salt);
    
    sheet.getRange(rowIdx, 2).setValue(hashedPassword);
    sheet.getRange(rowIdx, 8).setValue(salt);
    
    writeAuditLog(username, users[username].role, 'CHANGE_PASSWORD', 'Password diubah dengan salt baru');
    return {success: true, message: 'Password berhasil diubah!'};
  } catch(e) {
    return {success: false, message: e.message};
  }
}

// ============================================================
// PDF FUNGSI - LENGKAP
// ============================================================

function generatePDF(sheetName, filter, u) {
  try {
    u = verifyUser(u);
    if (!u) return {success: false, message: 'Sesi tidak valid.'};

    if(sheetName === 'LaporanBulanan'){
      var lap = getLaporan(u, filter);
      if(!lap || !lap.data || !lap.data.length){
        writeAuditLog(u.username, u.role, 'EXPORT_PDF', 'Laporan Bulanan - Data Kosong');
        return generateEmptyPDF(sheetName, filter, u);
      }
      writeAuditLog(u.username, u.role, 'EXPORT_PDF', 'Laporan Bulanan');
      return generatePDF_All(sheetName, lap.data, filter, u);
    }
    var rows = getFilteredData(sheetName, u, filter);
    if(!rows || !rows.length){
      writeAuditLog(u.username, u.role, 'EXPORT_PDF', sheetName + ' - Data Kosong');
      return generateEmptyPDF(sheetName, filter, u);
    }
    writeAuditLog(u.username, u.role, 'EXPORT_PDF', sheetName);
    return generatePDF_All(sheetName, rows, filter, u);
  } catch(e){
    writeAuditLog(u.username, u.role, 'EXPORT_PDF_ERROR', e.message);
    return {success:false, message:e.message};
  }
}

function generateEmptyPDF(sheetName, filter, u) {
  try {
    var filterText = '';
    if(filter){
      if(filter.rw) filterText += ' RW ' + escapeHtmlServer(filter.rw);
      if(filter.rt) filterText += ' RT ' + escapeHtmlServer(filter.rt);
      if(filter.bulan) filterText += ' Bulan ' + escapeHtmlServer(getBulanName2(filter.bulan));
      if(filter.tahun) filterText += ' ' + escapeHtmlServer(filter.tahun);
      if(filter.search) filterText += ' Cari: ' + escapeHtmlServer(filter.search);
    }

    var titleMap = {
      'DataWarga': 'DATA WARGA',
      'DataDatang': 'DATA PENDUDUK DATANG',
      'DataPergi': 'DATA PENDUDUK PERGI',
      'DataMeninggal': 'DATA PENDUDUK MENINGGAL',
      'DataLahir': 'DATA KELAHIRAN',
      'DataPengontrak': 'DATA PENGONTRAK / PENDUDUK SEMENTARA',
      'LaporanBulanan': 'LAPORAN BULANAN'
    };

    var html = '<html><head><meta charset="UTF-8"><style>';
    html += '@page { size: A4 landscape; margin: 10mm 8mm 10mm 8mm; }';
    html += 'body { font-family: Arial, sans-serif; text-align: center; padding-top: 80px; }';
    html += 'h1 { color: #1e3a5f; font-size: 28px; margin-bottom: 10px; }';
    html += 'h2 { color: #4b5563; font-size: 18px; font-weight: normal; margin-bottom: 10px; }';
    html += '.icon { font-size: 72px; color: #d1d5db; margin-bottom: 20px; }';
    html += '.msg { color: #6b7280; font-size: 16px; margin: 5px 0; }';
    html += '.filter { color: #9ca3af; font-size: 12px; margin-top: 20px; }';
    html += '.footer { margin-top: 60px; color: #9ca3af; font-size: 11px; border-top: 1px solid #e5e7eb; padding-top: 15px; }';
    html += '</style></head><body>';
    html += '<div class="icon">📄</div>';
    html += '<h1>' + (titleMap[sheetName] || sheetName.toUpperCase()) + '</h1>';
    html += '<h2>Dusun 1 - Desa Warnasari</h2>';
    html += '<p class="msg">Tidak ditemukan data untuk periode yang dipilih.</p>';
    html += '<p class="filter">Filter: ' + filterText + '</p>';
    html += '<p class="footer">Dicetak: ' + new Date().toLocaleString('id-ID') + ' | <strong>&copy; 2026 KKN 06 Desa Warnasari Ikopin University — Seluruh Hak Cipta Dilindungi</strong></p>';
    html += '</body></html>';
    
    var blob = Utilities.newBlob(html, 'text/html', 'empty.html');
    var pdfBlob = blob.getAs('application/pdf');
    return {success: true, blob: pdfBlob.getBytes(), html: html, name: 'Laporan_Kosong_' + new Date().toISOString().slice(0,10) + '.pdf'};
  } catch(e) {
    return {success: false, message: e.message};
  }
}


function generatePDF_All(sheetName, data, filter, u) {
  try {
    var rows = data || [];
    var headers = getHeaders(sheetName);
    var filterText = _buildFilterText(filter);

    var rwVal = (filter && filter.rw) ? filter.rw : (u ? u.rw : '');
    var rtVal = (filter && filter.rt) ? filter.rt : (u ? u.rt : '');
    var scopeTitle = ' ' + _buildScopeTitle(rwVal, rtVal);

    var html = '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><style>';
    html += _buildPdfStyles('landscape');
    html += '</style></head><body>';

    var baseTitle = SHEET_TITLE_MAP[sheetName] || sheetName.toUpperCase();
    html += '<div class="header">';
    html += '<h1>' + baseTitle + scopeTitle + '</h1>';
    html += '<h2>DESA WARNASARI, KEC. PANGALENGAN, KAB. BANDUNG</h2>';
    html += '<p class="sub">PROVINSI JAWA BARAT | DICETAK: ' + new Date().toLocaleDateString('id-ID', {day:'numeric',month:'long',year:'numeric'}).toUpperCase() + '</p>';
    html += '</div>';

    html += '<table>';
    html += '<thead><tr>';
    headers.forEach(function(h){
      html += '<th bgcolor="#1e3a5f" style="background-color:#1e3a5f !important;color:#ffffff !important;-webkit-print-color-adjust:exact !important;">' + h + '</th>';
    });
    html += '</tr></thead><tbody>';

    if(rows && rows.length > 0){
      rows.forEach(function(row){
        html += '<tr>';
        headers.forEach(function(h){
          var val = row[h] !== undefined && row[h] !== null ? String(row[h]) : '';
          if(val.length > 40) val = val.substring(0, 37) + '...';
          html += '<td>' + escapeHtmlServer(val) + '</td>';
        });
        html += '</tr>';
      });
    } else {
      html += '<tr><td colspan="' + colCount + '" style="text-align:center;padding:20px;color:#999;">Tidak ada data</td></tr>';
    }

    html += '</tbody></table>';
    if (sheetName === 'LaporanBulanan') {
      html += getLaporanEventTablesHtml(u, filter);
    }
    html += _buildPdfFooter(u, new Date());

    html += '</body></html>';

    var blob = Utilities.newBlob(html, 'text/html', sheetName + '.html');
    var pdfBlob = blob.getAs('application/pdf');
    return {success:true, blob:pdfBlob.getBytes(), html:html, name:'Laporan_' + sheetName + '_' + new Date().toISOString().slice(0,10) + '_' + new Date().getTime() + '.pdf'};
  } catch(e){

    return {success:false, message:e.message};
  }
}

// ============================================================
// BUKU INDUK PENDUDUK - FORMAT RESMI
// ============================================================

function generatePDF_BukuInduk(filter, u) {
  try {
    u = verifyUser(u);
    if (!u) return { success: false, message: 'Sesi kedaluwarsa, silakan login ulang.' };

    writeAuditLog(u.username, u.role, 'EXPORT_PDF_BUKU_INDUK', 'Buku Induk Penduduk');

    filter = filter || {};
    if (u.role === 'rt') {
      filter.rw = u.rw;
      filter.rt = u.rt;
    } else if (u.role === 'rw') {
      filter.rw = u.rw;
    }

    var rows = getFilteredData('DataWarga', u, filter) || [];

    var rwVal = (filter && filter.rw) ? filter.rw : (u ? u.rw : '');
    var rtVal = (filter && filter.rt) ? filter.rt : (u ? u.rt : '');

    var scopeTitle = '';
    if (rtVal) {
      var formattedRW = String(rwVal).length === 1 ? '0' + rwVal : rwVal;
      scopeTitle = 'DUSUN 1 RW ' + formattedRW + ' RT ' + rtVal;
    } else if (rwVal) {
      var formattedRW = String(rwVal).length === 1 ? '0' + rwVal : rwVal;
      scopeTitle = 'DUSUN 1 RW ' + formattedRW;
    } else {
      scopeTitle = 'DUSUN 1';
    }

    var now = new Date();
    var datePrintedStr = now.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });

    var html = '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><style>';
    html += '@page { size: A4 landscape; margin: 10mm 8mm 10mm 8mm; }';
    html += '* { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color-adjust: exact !important; }';
    html += 'html, body { width: 100%; margin: 0; padding: 0; background: #ffffff; }';
    html += 'body { font-family: "Times New Roman", Arial, sans-serif; font-size: 9.5px; padding: 12px 15px; box-sizing: border-box; }';
    html += '.header { text-align: center; margin-bottom: 10px; }';
    html += '.header h1 { font-size: 16px; color: #000000; margin: 0; font-weight: bold; letter-spacing: 0.5px; }';
    html += '.header h2 { font-size: 12px; color: #000000; margin: 2px 0; font-weight: bold; }';
    html += '.header .sub { font-size: 10px; color: #374151; margin-top: 2px; font-weight: bold; }';
    html += '.total-box { border: 1px solid #1e3a5f; background-color: #f1f5f9 !important; padding: 5px 10px; margin: 6px 0; text-align: center; font-weight: bold; width: 100%; box-sizing: border-box; -webkit-print-color-adjust: exact !important; }';
    html += '.total-box span { margin: 0 10px; font-size: 9.5px; display: inline-block; }';
    html += 'table { width: 100%; border-collapse: collapse; margin-top: 6px; font-size: 9px; }';
    html += 'th { background-color: #1e3a5f !important; color: #ffffff !important; padding: 5px 3px; border: 1px solid #1e3a5f; text-align: center; font-weight: bold; font-size: 9px; -webkit-print-color-adjust: exact !important; }';
    html += 'td { padding: 4px 3px; border: 1px solid #000000; text-align: center; font-size: 8.5px; word-break: break-word; }';
    html += 'tr:nth-child(even) { background-color: #f8fafc !important; -webkit-print-color-adjust: exact !important; }';
    html += '.footer { text-align: right; font-size: 8px; color: #6b7280; margin-top: 10px; border-top: 1px solid #e5e7eb; padding-top: 6px; }';
    html += '</style></head><body>';

    html += '<div class="header">';
    html += '<h1>BUKU INDUK PENDUDUK ' + scopeTitle + '</h1>';
    html += '<h2>DESA WARNASARI, KEC. PANGALENGAN, KAB. BANDUNG</h2>';
    html += '<p class="sub">PROVINSI JAWA BARAT</p>';
    html += '</div>';

    // Summary counts
    var totalWarga = rows.length;
    var totalLaki = 0;
    var totalPerempuan = 0;
    var kkSet = {};
    rows.forEach(function(r) {
      var jk = String(r['Jenis Kelamin'] || '').trim().toLowerCase();
      if (jk === 'laki-laki' || jk === 'l') totalLaki++;
      else if (jk === 'perempuan' || jk === 'p') totalPerempuan++;

      var nkk = String(r['No KK'] || r['NO KK'] || '').trim();
      if (nkk && nkk !== '-') kkSet[nkk] = true;
    });
    var totalKK = Object.keys(kkSet).length;

    html += '<div class="total-box">';
    html += '<span>Total Penduduk: ' + totalWarga.toLocaleString('id-ID') + ' Jiwa</span>';
    html += '<span>Laki-Laki: ' + totalLaki.toLocaleString('id-ID') + ' Jiwa</span>';
    html += '<span>Perempuan: ' + totalPerempuan.toLocaleString('id-ID') + ' Jiwa</span>';
    html += '<span>Jumlah KK: ' + totalKK.toLocaleString('id-ID') + ' KK</span>';
    html += '</div>';

    var thStyle = _thStyle();

    html += '<table>';
    html += '<thead><tr>';
    html += '<th ' + thStyle + ' style="width:3%;">NO</th>';
    html += '<th ' + thStyle + ' style="width:15%;">NAMA LENGKAP</th>';
    html += '<th ' + thStyle + ' style="width:12%;">NIK</th>';
    html += '<th ' + thStyle + ' style="width:12%;">NO KK</th>';
    html += '<th ' + thStyle + ' style="width:4%;">JK</th>';
    html += '<th ' + thStyle + ' style="width:14%;">TEMPAT, TANGGAL LAHIR</th>';
    html += '<th ' + thStyle + ' style="width:7%;">AGAMA</th>';
    html += '<th ' + thStyle + ' style="width:9%;">PENDIDIKAN</th>';
    html += '<th ' + thStyle + ' style="width:9%;">PEKERJAAN</th>';
    html += '<th ' + thStyle + ' style="width:8%;">STATUS NIKAH</th>';
    html += '<th ' + thStyle + ' style="width:7%;">RT / RW</th>';
    html += '</tr></thead><tbody>'; // FIX: hapus duplikat tag thead/tbody

    if (rows && rows.length > 0) {
      rows.forEach(function(r, idx) {
        var jkChar = (String(r['Jenis Kelamin']||'').trim().toLowerCase().startsWith('l')) ? 'L' : 'P';
        var tglLahirFormatted = formatDateIndoServer(r['Tanggal Lahir']);
        var tmpLahirUpper = String(r['Tempat Lahir'] || '-').toUpperCase();
        var ttl = escapeHtmlServer(tmpLahirUpper) + ', ' + tglLahirFormatted;
        var rtRwStr = 'RT ' + escapeHtmlServer(r['RT']||'-') + ' / RW ' + escapeHtmlServer(r['RW']||'-');

        html += '<tr>';
        html += '<td style="text-align:center;">' + (idx + 1) + '</td>';
        html += '<td style="text-align:left;font-weight:bold;">' + escapeHtmlServer(String(r['Nama Lengkap'] || '-').toUpperCase()) + '</td>';
        html += '<td style="text-align:center;">' + escapeHtmlServer(r['NIK'] || '-') + '</td>';
        html += '<td style="text-align:center;">' + escapeHtmlServer(r['No KK'] || '-') + '</td>';
        html += '<td style="text-align:center;font-weight:bold;">' + jkChar + '</td>';
        html += '<td style="text-align:left;">' + ttl + '</td>';
        html += '<td style="text-align:center;">' + escapeHtmlServer(String(r['Agama'] || '-').toUpperCase()) + '</td>';
        html += '<td style="text-align:center;">' + escapeHtmlServer(String(r['Pendidikan'] || '-').toUpperCase()) + '</td>';
        html += '<td style="text-align:left;">' + escapeHtmlServer(String(r['Pekerjaan'] || '-').toUpperCase()) + '</td>';
        html += '<td style="text-align:center;">' + escapeHtmlServer(String(r['Status Perkawinan'] || '-').toUpperCase()) + '</td>';
        html += '<td style="text-align:center;">' + rtRwStr + '</td>';
        html += '</tr>';
      });
    } else {
      html += '<tr><td colspan="11" style="text-align:center;padding:20px;color:#999;">Tidak ada data warga</td></tr>';
    }

    html += '</tbody></table>';

    // Signature section
    html += '<div style="margin-top:20px;display:flex;justify-content:space-between;page-break-inside:avoid;font-size:9.5px;">';
    html += '<div style="text-align:center;width:40%;">';
    html += '<p style="margin:0;">Mengetahui,</p>';
    html += '<p style="margin:2px 0 45px 0;font-weight:bold;">Ketua RW ' + escapeHtmlServer(rwVal || '...') + '</p>';
    html += '<p style="margin:0;font-weight:bold;text-decoration:underline;">( ........................................... )</p>';
    html += '</div>';

    html += '<div style="text-align:center;width:40%;">';
    html += '<p style="margin:0;">Warnasari, ' + datePrintedStr + '</p>';
    html += '<p style="margin:2px 0 45px 0;font-weight:bold;">Ketua RT ' + escapeHtmlServer(rtVal || '...') + '</p>';
    html += '<p style="margin:0;font-weight:bold;text-decoration:underline;">( ........................................... )</p>';
    html += '</div>';
    html += '</div>';

    html += '<div class="footer" style="margin-top:12px;padding-top:6px;border-top:1px solid #ccc;font-size:8.5px;color:#4b5563;display:flex;justify-content:space-between;align-items:center;">';
    html += '<span>Buku Induk Penduduk | Dicetak oleh: ' + escapeHtmlServer(u.username) + ' (' + escapeHtmlServer(u.role.toUpperCase()) + ') | ' + new Date().toLocaleString('id-ID') + '</span>';
    html += '<span><strong>&copy; 2026 KKN 06 Desa Warnasari Ikopin University</strong></span>';
    html += '</div>';

    html += '</body></html>';

    var blob = Utilities.newBlob(html, 'text/html', 'Buku_Induk_Penduduk.html');
    var pdfBlob = blob.getAs('application/pdf');
    return {
      success: true,
      blob: pdfBlob.getBytes(),
      html: html,
      name: 'Buku_Induk_Penduduk_' + (scopeTitle.replace(/\s+/g, '_')) + '_' + new Date().toISOString().slice(0, 10) + '.pdf'
    };
  } catch (e) {
    writeAuditLog(u ? u.username : 'GUEST', u ? u.role : 'GUEST', 'EXPORT_PDF_BUKU_INDUK_ERROR', e.message);
    return { success: false, message: e.message };
  }
}

// ============================================================
// PDF RW - FORMAT RESMI
// ============================================================

function generatePDF_RW(data, filter, u) {
  try {
    if(u.role === 'rt'){
      writeAuditLog(u.username, u.role, 'EXPORT_PDF_RW_DENIED', 'Role RT tidak boleh format RW');
      return {success:false, message:'Format PDF RW tidak tersedia untuk role RT.'};
    }
    
    writeAuditLog(u.username, u.role, 'EXPORT_PDF_RW', 'PDF RW');
    var rows = data || [];

    if(rows.length === 0) {
      return generateEmptyPDF('LaporanBulanan', filter, u);
    }

    var totalAwal = 0, totalDatang = 0, totalPergi = 0, totalLahir = 0, totalMeninggal = 0, totalAkhir = 0, totalKK = 0, totalLaki = 0, totalPerempuan = 0;

    rows.forEach(function(r) {
      var awal = Number(r['Jumlah Awal'] || r[5] || 0);
      var datang = Number(r['Datang'] || r[6] || 0);
      var pergi = Number(r['Pergi'] || r[7] || 0);
      var meninggal = Number(r['Meninggal'] || r[8] || 0);
      var lahir = Number(r['Lahir'] || r[9] || 0);
      var akhir = Number(r['Jumlah Akhir'] || r[10] || 0);
      var kk = Number(r['KK'] || r[13] || 0);
      var laki = Number(r['Laki-laki'] || r[11] || 0);
      var perempuan = Number(r['Perempuan'] || r[12] || 0);

      totalAwal += awal;
      totalDatang += datang;
      totalPergi += pergi;
      totalMeninggal += meninggal;
      totalLahir += lahir;
      totalAkhir += akhir;
      totalKK += kk;
      totalLaki += laki;
      totalPerempuan += perempuan;
    });

    var rwDisplay = '';
    if (filter && filter.rw) {
      rwDisplay = filter.rw;
    } else if (rows.length > 0) {
      var firstRow = rows[0];
      rwDisplay = firstRow.RW || firstRow[3] || '';
    } else {
      rwDisplay = 'XX';
    }
    
    if (rwDisplay && rwDisplay.length === 1) {
      rwDisplay = '0' + rwDisplay;
    }

    var now = new Date();
    var periodeBulan = (filter && filter.bulan) ? filter.bulan : String(now.getMonth() + 1);
    var periodeTahun = (filter && filter.tahun) ? filter.tahun : String(now.getFullYear());
    var periodeText = getBulanName2(periodeBulan) + ' ' + periodeTahun;

    var html = '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><style>';
    html += _buildPdfStyles('landscape');
    html += '</style></head><body>';

    html += '<div class="header">';
    html += '<h1>DESA WARNASARI</h1>';
    html += '<h2>LAPORAN KEPENDUDUKAN DUSUN 1 RW ' + rwDisplay + '</h2>';
    html += '<p class="periode">PERIODE: ' + periodeText.toUpperCase() + '</p>';
    html += '</div>';

    html += '<div class="total-box">';
    html += '<span>Total Penduduk Awal: ' + totalAwal.toLocaleString() + ' Jiwa</span>';
    html += '<span>Warga Datang: ' + totalDatang.toLocaleString() + ' Jiwa</span>';
    html += '<span>Warga Pergi: ' + totalPergi.toLocaleString() + ' Jiwa</span>';
    html += '<span>Warga Lahir: ' + totalLahir.toLocaleString() + ' Jiwa</span>';
    html += '<span>Warga Meninggal: ' + totalMeninggal.toLocaleString() + ' Jiwa</span>';
    html += '</div>';

    var thS = 'bgcolor="#1e3a5f" style="background-color:#1e3a5f !important;color:#ffffff !important;-webkit-print-color-adjust:exact !important;"';
    html += '<table>';
    html += '<thead><tr>';
    html += '<th ' + thS + ' style="width:12%;">UNIT RT/RW</th>';
    html += '<th ' + thS + ' style="width:10%;">JUMLAH AWAL</th>';
    html += '<th ' + thS + ' style="width:9%;">DATANG</th>';
    html += '<th ' + thS + ' style="width:9%;">PERGI</th>';
    html += '<th ' + thS + ' style="width:9%;">LAHIR</th>';
    html += '<th ' + thS + ' style="width:9%;">MENINGGAL</th>';
    html += '<th ' + thS + ' style="width:10%;">JUMLAH AKHIR</th>';
    html += '<th ' + thS + ' style="width:9%;">KK</th>';
    html += '<th ' + thS + ' style="width:11%;">LAKI-LAKI</th>';
    html += '<th ' + thS + ' style="width:12%;">PEREMPUAN</th>';
    html += '</tr></thead><tbody>';

    rows.forEach(function(r) {
      var rt = r.RT || r[4] || '-';
      var rw = r.RW || r[3] || filter.rw || '-';
      var unit = 'RT ' + rt + ' / RW ' + rw;
      
      var awal = Number(r['Jumlah Awal'] || r[5] || 0);
      var datang = Number(r['Datang'] || r[6] || 0);
      var pergi = Number(r['Pergi'] || r[7] || 0);
      var meninggal = Number(r['Meninggal'] || r[8] || 0);
      var lahir = Number(r['Lahir'] || r[9] || 0);
      var akhir = Number(r['Jumlah Akhir'] || r[10] || 0);
      var kk = Number(r['KK'] || r[13] || 0);
      var laki = Number(r['Laki-laki'] || r[11] || 0);
      var perempuan = Number(r['Perempuan'] || r[12] || 0);

      html += '<tr>';
      html += '<td>' + unit + '</td>';
      html += '<td>' + awal.toLocaleString() + '</td>';
      html += '<td>' + datang.toLocaleString() + '</td>';
      html += '<td>' + pergi.toLocaleString() + '</td>';
      html += '<td>' + lahir.toLocaleString() + '</td>';
      html += '<td>' + meninggal.toLocaleString() + '</td>';
      html += '<td><strong>' + akhir.toLocaleString() + '</strong></td>';
      html += '<td>' + kk.toLocaleString() + '</td>';
      html += '<td>' + laki.toLocaleString() + '</td>';
      html += '<td>' + perempuan.toLocaleString() + '</td>';
      html += '</tr>';
    });

    html += '<tr class="total-row">';
    html += '<td><strong>JUMLAH</strong></td>';
    html += '<td><strong>' + totalAwal.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalDatang.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalPergi.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalLahir.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalMeninggal.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalAkhir.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalKK.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalLaki.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalPerempuan.toLocaleString() + '</strong></td>';
    html += '</tr>';

    html += '</tbody></table>';

    html += '<div style="text-align:right;margin-top:6px;font-weight:bold;font-size:11px;">';
    html += 'Total Penduduk Akhir: ' + totalAkhir.toLocaleString() + ' Jiwa';
    html += ' | Total KK: ' + totalKK.toLocaleString();
    html += ' | Laki-laki: ' + totalLaki.toLocaleString();
    html += ' | Perempuan: ' + totalPerempuan.toLocaleString();
    html += '</div>';

    html += getLaporanEventTablesHtml(u, filter);

    html += _buildPdfFooter(u, new Date());

    html += '</body></html>';

    var blob = Utilities.newBlob(html, 'text/html', 'laporan.html');
    var pdfBlob = blob.getAs('application/pdf');
    return {
      success: true, 
      blob: pdfBlob.getBytes(), 
      html: html,
      name: 'Laporan_RW_' + rwDisplay + '_' + new Date().toISOString().slice(0,10) + '_' + new Date().getTime() + '.pdf'
    };
    
  } catch(e){
    writeAuditLog(u.username, u.role, 'EXPORT_PDF_RW_ERROR', e.message);
    return {success:false, message:e.message};
  }
}

// ============================================================
// EXPORT PDF KADUS (REKAPITULASI DUSUN 1 PER RW)
// ============================================================

function generatePDF_Kadus(filter, u) {
  try {
    u = verifyUser(u);
    if (!u) return {success:false, message:'Session tidak valid'};
    if (!isSuperAdmin(u) && !isKadus(u)) {
      return {success:false, message:'Hanya SuperAdmin dan Kadus yang dapat mencetak Laporan Kadus!'};
    }

    var reportData = getLaporan(u, filter);
    var rows = reportData.data || [];

    // Kelompokkan data per RW (RW 02, RW 03, RW 04, RW 12, RW 13, RW 15)
    var listRW = RW_LIST.slice(); // gunakan konstanta global, bukan hardcode
    var rwGroups = {};
    listRW.forEach(function(rwKey) {
      rwGroups[rwKey] = {
        rw: rwKey,
        awal: 0, datang: 0, pergi: 0, meninggal: 0, lahir: 0, akhir: 0, kk: 0, laki: 0, perempuan: 0
      };
    });

    rows.forEach(function(r) {
      var rw = normalizeRW(r.RW || r[3] || '');
      if (!rwGroups[rw]) {
        rwGroups[rw] = {
          rw: rw,
          awal: 0, datang: 0, pergi: 0, meninggal: 0, lahir: 0, akhir: 0, kk: 0, laki: 0, perempuan: 0
        };
        listRW.push(rw);
      }
      rwGroups[rw].awal += Number(r['Jumlah Awal'] || 0);
      rwGroups[rw].datang += Number(r['Datang'] || 0);
      rwGroups[rw].pergi += Number(r['Pergi'] || 0);
      rwGroups[rw].meninggal += Number(r['Meninggal'] || 0);
      rwGroups[rw].lahir += Number(r['Lahir'] || 0);
      rwGroups[rw].akhir += Number(r['Jumlah Akhir'] || 0);
      rwGroups[rw].kk += Number(r['KK'] || 0);
      rwGroups[rw].laki += Number(r['Laki-laki'] || 0);
      rwGroups[rw].perempuan += Number(r['Perempuan'] || 0);
    });

    var totalAwal = 0, totalDatang = 0, totalPergi = 0, totalLahir = 0, totalMeninggal = 0, totalAkhir = 0, totalKK = 0, totalLaki = 0, totalPerempuan = 0;

    var now = new Date();
    var periodeBulan = (filter && filter.bulan) ? filter.bulan : String(now.getMonth() + 1);
    var periodeTahun = (filter && filter.tahun) ? filter.tahun : String(now.getFullYear());
    var periodeText = getBulanName2(periodeBulan) + ' ' + periodeTahun;

    var html = '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><style>';
    html += _buildPdfStyles('landscape');
    html += '</style></head><body>';

    html += '<div class="header">';
    html += '<h1>DESA WARNASARI</h1>';
    html += '<h2>LAPORAN KEPENDUDUKAN DUSUN 1 (REKAPITULASI DUSUN PER RW)</h2>';
    html += '<p class="periode">PERIODE: ' + periodeText.toUpperCase() + '</p>';
    html += '</div>';

    // Rincian per RW
    listRW.forEach(function(rwKey) {
      var item = rwGroups[rwKey];
      totalAwal += item.awal;
      totalDatang += item.datang;
      totalPergi += item.pergi;
      totalMeninggal += item.meninggal;
      totalLahir += item.lahir;
      totalAkhir += item.akhir;
      totalKK += item.kk;
      totalLaki += item.laki;
      totalPerempuan += item.perempuan;
    });

    html += '<div class="total-box">';
    html += '<span>Total Penduduk Awal: ' + totalAwal.toLocaleString() + ' Jiwa</span>';
    html += '<span>Warga Datang: ' + totalDatang.toLocaleString() + ' Jiwa</span>';
    html += '<span>Warga Pergi: ' + totalPergi.toLocaleString() + ' Jiwa</span>';
    html += '<span>Warga Lahir: ' + totalLahir.toLocaleString() + ' Jiwa</span>';
    html += '<span>Warga Meninggal: ' + totalMeninggal.toLocaleString() + ' Jiwa</span>';
    html += '</div>';

    var thS = 'bgcolor="#1e3a5f" style="background-color:#1e3a5f !important;color:#ffffff !important;-webkit-print-color-adjust:exact !important;"';
    html += '<table>';
    html += '<thead><tr>';
    html += '<th ' + thS + ' style="width:14%;">UNIT RW</th>';
    html += '<th ' + thS + ' style="width:10%;">JUMLAH AWAL</th>';
    html += '<th ' + thS + ' style="width:9%;">DATANG</th>';
    html += '<th ' + thS + ' style="width:9%;">PERGI</th>';
    html += '<th ' + thS + ' style="width:9%;">LAHIR</th>';
    html += '<th ' + thS + ' style="width:9%;">MENINGGAL</th>';
    html += '<th ' + thS + ' style="width:10%;">JUMLAH AKHIR</th>';
    html += '<th ' + thS + ' style="width:9%;">KK</th>';
    html += '<th ' + thS + ' style="width:11%;">LAKI-LAKI</th>';
    html += '<th ' + thS + ' style="width:10%;">PEREMPUAN</th>';
    html += '</tr></thead><tbody>';

    listRW.forEach(function(rwKey) {
      var item = rwGroups[rwKey];
      html += '<tr>';
      html += '<td><strong>RW ' + rwKey + '</strong></td>';
      html += '<td>' + item.awal.toLocaleString() + '</td>';
      html += '<td>' + item.datang.toLocaleString() + '</td>';
      html += '<td>' + item.pergi.toLocaleString() + '</td>';
      html += '<td>' + item.lahir.toLocaleString() + '</td>';
      html += '<td>' + item.meninggal.toLocaleString() + '</td>';
      html += '<td><strong>' + item.akhir.toLocaleString() + '</strong></td>';
      html += '<td>' + item.kk.toLocaleString() + '</td>';
      html += '<td>' + item.laki.toLocaleString() + '</td>';
      html += '<td>' + item.perempuan.toLocaleString() + '</td>';
      html += '</tr>';
    });

    html += '<tr class="total-row">';
    html += '<td><strong>JUMLAH DUSUN 1</strong></td>';
    html += '<td><strong>' + totalAwal.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalDatang.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalPergi.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalLahir.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalMeninggal.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalAkhir.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalKK.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalLaki.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalPerempuan.toLocaleString() + '</strong></td>';
    html += '</tr>';

    html += '</tbody></table>';

    html += '<div style="text-align:right;margin-top:6px;font-weight:bold;font-size:11px;">';
    html += 'Total Penduduk Akhir Dusun 1: ' + totalAkhir.toLocaleString() + ' Jiwa';
    html += ' | Total KK: ' + totalKK.toLocaleString();
    html += ' | Laki-laki: ' + totalLaki.toLocaleString();
    html += ' | Perempuan: ' + totalPerempuan.toLocaleString();
    html += '</div>';

    html += getLaporanEventTablesHtml(u, filter);

    var now2 = new Date();
    var tanggal = now2.getDate();
    var bulan = now2.getMonth() + 1;
    var tahun = now2.getFullYear();
    var bulanNama = getBulanName2(String(bulan));
    var tanggalStr = tanggal + ' ' + bulanNama + ' ' + tahun;

    html += '<div style="margin-top:15px;padding-top:6px;border-top:1px solid #ccc;font-size:8.5px;color:#4b5563;display:flex;justify-content:space-between;align-items:center;">';
    html += '<span>Dicetak: ' + now2.toLocaleString('id-ID') + ' | Desa Warnasari, ' + tanggalStr + '</span>';
    html += '<span><strong>&copy; 2026 KKN 06 Desa Warnasari Ikopin University</strong></span>';
    html += '</div>';

    html += '</body></html>';

    var blob = Utilities.newBlob(html, 'text/html', 'laporan_kadus.html');
    var pdfBlob = blob.getAs('application/pdf');
    writeAuditLog(u.username, u.role, 'EXPORT_PDF_KADUS', 'Rekap Dusun 1');
    return {
      success: true, 
      blob: pdfBlob.getBytes(), 
      html: html,
      name: 'Laporan_Dusun1_Kadus_' + new Date().toISOString().slice(0,10) + '_' + new Date().getTime() + '.pdf'
    };

  } catch(e) {
    writeAuditLog(u.username, u.role, 'EXPORT_PDF_KADUS_ERROR', e.message);
    return {success:false, message:e.message};
  }
}

// ============================================================
// PDF RT - FORMAT RESMI
// ============================================================

function generatePDF_RT(data, filter, u) {
  try {
    writeAuditLog(u.username, u.role, 'EXPORT_PDF_RT', 'PDF RT');
    
    var rows = data || [];

    if(rows.length === 0) {
      return generateEmptyPDF('LaporanBulanan', filter, u);
    }

    var filteredRows = rows;
    if (filter && filter.rt) {
      filteredRows = rows.filter(function(r) {
        return String(r.RT || r[4] || '') === String(filter.rt);
      });
    }

    if (filteredRows.length > 1 && filter && filter.rw) {
      filteredRows = filteredRows.filter(function(r) {
        return String(r.RW || r[3] || '') === String(filter.rw);
      });
    }

    var r = filteredRows[0] || rows[0];
    
    var rwDisplay = '';
    var rtDisplay = '';
    
    if (filter && filter.rw) {
      rwDisplay = filter.rw;
    } else if (r) {
      rwDisplay = r.RW || r[3] || '';
    } else {
      rwDisplay = 'XX';
    }
    
    if (filter && filter.rt) {
      rtDisplay = filter.rt;
    } else if (r) {
      rtDisplay = r.RT || r[4] || '';
    } else {
      rtDisplay = 'X';
    }
    
    if (rwDisplay && rwDisplay.length === 1) {
      rwDisplay = '0' + rwDisplay;
    }
    
    var totalAwal = Number(r['Jumlah Awal'] || r[5] || 0);
    var totalDatang = Number(r['Datang'] || r[6] || 0);
    var totalPergi = Number(r['Pergi'] || r[7] || 0);
    var totalMeninggal = Number(r['Meninggal'] || r[8] || 0);
    var totalLahir = Number(r['Lahir'] || r[9] || 0);
    var totalAkhir = Number(r['Jumlah Akhir'] || r[10] || 0);
    var totalKK = Number(r['KK'] || r[13] || 0);
    var totalLaki = Number(r['Laki-laki'] || r[11] || 0);
    var totalPerempuan = Number(r['Perempuan'] || r[12] || 0);

    var now = new Date();
    var periodeBulan = (filter && filter.bulan) ? filter.bulan : String(now.getMonth() + 1);
    var periodeTahun = (filter && filter.tahun) ? filter.tahun : String(now.getFullYear());
    var periodeText = getBulanName2(periodeBulan) + ' ' + periodeTahun;

    var html = '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><style>';
    html += _buildPdfStyles('landscape');
    html += '</style></head><body>';

    html += '<div class="header">';
    html += '<h1>DESA WARNASARI</h1>';
    html += '<h2>LAPORAN KEPENDUDUKAN DUSUN 1 RW ' + rwDisplay + ' RT ' + rtDisplay + '</h2>';
    html += '<p class="periode">PERIODE: ' + periodeText.toUpperCase() + '</p>';
    html += '</div>';

    html += '<div class="total-box">';
    html += '<span>Total Penduduk Awal: ' + totalAwal.toLocaleString() + ' Jiwa</span>';
    html += '<span>Warga Datang: ' + totalDatang.toLocaleString() + ' Jiwa</span>';
    html += '<span>Warga Pergi: ' + totalPergi.toLocaleString() + ' Jiwa</span>';
    html += '<span>Warga Lahir: ' + totalLahir.toLocaleString() + ' Jiwa</span>';
    html += '<span>Warga Meninggal: ' + totalMeninggal.toLocaleString() + ' Jiwa</span>';
    html += '</div>';

    var thS = 'bgcolor="#1e3a5f" style="background-color:#1e3a5f !important;color:#ffffff !important;-webkit-print-color-adjust:exact !important;"';
    html += '<table>';
    html += '<thead><tr>';
    html += '<th ' + thS + ' style="width:12%;">UNIT RT/RW</th>';
    html += '<th ' + thS + ' style="width:10%;">JUMLAH AWAL</th>';
    html += '<th ' + thS + ' style="width:9%;">DATANG</th>';
    html += '<th ' + thS + ' style="width:9%;">PERGI</th>';
    html += '<th ' + thS + ' style="width:9%;">LAHIR</th>';
    html += '<th ' + thS + ' style="width:9%;">MENINGGAL</th>';
    html += '<th ' + thS + ' style="width:10%;">JUMLAH AKHIR</th>';
    html += '<th ' + thS + ' style="width:9%;">KK</th>';
    html += '<th ' + thS + ' style="width:11%;">LAKI-LAKI</th>';
    html += '<th ' + thS + ' style="width:12%;">PEREMPUAN</th>';
    html += '</tr></thead><tbody>';

    var unit = 'RT ' + rtDisplay + ' / RW ' + rwDisplay;
    
    html += '<tr>';
    html += '<td><strong>' + unit + '</strong></td>';
    html += '<td><strong>' + totalAwal.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalDatang.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalPergi.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalLahir.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalMeninggal.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalAkhir.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalKK.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalLaki.toLocaleString() + '</strong></td>';
    html += '<td><strong>' + totalPerempuan.toLocaleString() + '</strong></td>';
    html += '</tr>';

    html += '</tbody></table>';

    html += '<div style="text-align:right;margin-top:6px;font-weight:bold;font-size:11px;">';
    html += 'Total Penduduk Akhir: ' + totalAkhir.toLocaleString() + ' Jiwa';
    html += ' | Total KK: ' + totalKK.toLocaleString();
    html += ' | Laki-laki: ' + totalLaki.toLocaleString();
    html += ' | Perempuan: ' + totalPerempuan.toLocaleString();
    html += '</div>';

    html += getLaporanEventTablesHtml(u, filter);

    var now2 = new Date();
    var tanggal = now2.getDate();
    var bulan = now2.getMonth() + 1;
    var tahun = now2.getFullYear();
    var bulanNama = getBulanName2(String(bulan));
    var tanggalStr = tanggal + ' ' + bulanNama + ' ' + tahun;
    
    html += '<div style="margin-top:15px;font-size:8.5px;color:#4b5563;border-top:1px solid #ccc;padding-top:6px;display:flex;justify-content:space-between;align-items:center;">';
    html += '<span>Dicetak: ' + now2.toLocaleString('id-ID') + ' | Desa Warnasari, ' + tanggalStr + '</span>';
    html += '<span><strong>&copy; 2026 KKN 06 Desa Warnasari Ikopin University</strong></span>';
    html += '</div>';

    html += '</body></html>';

    var blob = Utilities.newBlob(html, 'text/html', 'laporan.html');
    var pdfBlob = blob.getAs('application/pdf');
    return {
      success: true, 
      blob: pdfBlob.getBytes(), 
      html: html,
      name: 'Laporan_RT_RW' + rwDisplay + '_RT' + rtDisplay + '_' + new Date().toISOString().slice(0,10) + '_' + new Date().getTime() + '.pdf'
    };

    
  } catch(e){
    writeAuditLog(u.username, u.role, 'EXPORT_PDF_RT_ERROR', e.message);
    return {success:false, message:e.message};
  }
}

function getBulanName2(bulan) {
  var nama = ['', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  return nama[Number(bulan)] || bulan;
}

// ============================================================
// DOWNLOAD EXCEL & CSV
// ============================================================

function downloadSpreadsheet(u) {
  try {
    // Hanya superadmin dan kadus yang bisa download seluruh spreadsheet
    u = verifyUser(u);
    if (!u) return {success: false, message: 'Sesi tidak valid. Silakan login ulang.'};
    if (!canAccessAllData(u)) {
      writeAuditLog(u.username, u.role, 'DOWNLOAD_EXCEL_DENIED', 'Akses ditolak: role ' + u.role);
      return {success: false, message: 'Hanya SuperAdmin dan Kadus yang bisa download Excel.'};
    }
    writeAuditLog(u.username, u.role, 'DOWNLOAD_EXCEL', 'Download Excel');
    var url = SpreadsheetApp.getActiveSpreadsheet().getUrl(); 
    return {success:true, url:url.replace('/edit','/export?format=xlsx')}; 
  }
  catch(e){ 
    writeAuditLog(u ? u.username : 'unknown', u ? u.role : 'unknown', 'DOWNLOAD_EXCEL_ERROR', e.message);
    return {success:false, message:e.message}; 
  }
}

function exportToCSV(sheetName, filter, u) {
  try {
    u = verifyUser(u);
    if (!u) return {success: false, message: 'Sesi tidak valid.'};

    var data = getFilteredData(sheetName, u, filter);
    if (!data || data.length === 0) {
      return {success: false, message: 'Tidak ada data'};
    }
    
    var headers = getHeaders(sheetName);
    var csv = headers.join(',') + '\n';
    
    data.forEach(function(row) {
      var values = headers.map(function(h) {
        var val = String(row[h] || '');
        if (val.includes(',') || val.includes('"') || val.includes('\n')) {
          val = '"' + val.replace(/"/g, '""') + '"';
        }
        return val;
      });
      csv += values.join(',') + '\n';
    });
    
    writeAuditLog(u.username, u.role, 'EXPORT_CSV', sheetName);
    return {
      success: true, 
      csv: csv, 
      name: sheetName + '_' + new Date().toISOString().slice(0,10) + '.csv'
    };
  } catch(e) {
    writeAuditLog(u.username, u.role, 'EXPORT_CSV_ERROR', e.message);
    return {success: false, message: e.message};
  }
}

function getLaporanDetailEvents(u, filter) {
  u = verifyUser(u);
  if (!u) return { datang: [], pergi: [], lahir: [], meninggal: [] };

  var f = {};
  if (filter) {
    if (filter.rw) f.rw = filter.rw;
    if (filter.rt) f.rt = filter.rt;
    if (filter.search) f.search = filter.search;
    if (filter.bulan) f.bulan = filter.bulan;
    if (filter.tahun) f.tahun = filter.tahun;
  }

  var hasBulanTahunFilter = filter && ((filter.bulan && filter.bulan !== '') || (filter.tahun && filter.tahun !== ''));
  if (!hasBulanTahunFilter) {
    var now = new Date();
    f.bulan = String(now.getMonth() + 1);
    f.tahun = String(now.getFullYear());
  }

  return {
    datang: getFilteredData('DataDatang', u, f) || [],
    pergi: getFilteredData('DataPergi', u, f) || [],
    lahir: getFilteredData('DataLahir', u, f) || [],
    meninggal: getFilteredData('DataMeninggal', u, f) || []
  };
}

function formatDateIndoServer(dStr) {
  if (!dStr) return '-';
  var d = (dStr instanceof Date) ? dStr : new Date(dStr);
  if (isNaN(d.getTime())) return String(dStr);
  var day = String(d.getDate()).padStart(2, '0');
  var month = String(d.getMonth() + 1).padStart(2, '0');
  var year = d.getFullYear();
  return day + '/' + month + '/' + year;
}

function getLaporanEventTablesHtml(u, filter) {
  var events = getLaporanDetailEvents(u, filter);
  var html = '';

  var thStyle = 'bgcolor="#1e3a5f" style="background-color:#1e3a5f !important;color:#ffffff !important;padding:4px 3px;border:1px solid #1e3a5f;text-align:center;font-weight:bold;-webkit-print-color-adjust:exact !important;"';

  if (events.datang && events.datang.length > 0) {
    html += '<div style="margin-top:16px;page-break-inside:avoid;">';
    html += '<h3 style="font-size:11px;font-weight:bold;margin:8px 0 4px;color:#1e3a5f;background-color:#f1f5f9 !important;border:1px solid #cbd5e1;padding:4px 8px;border-radius:4px;-webkit-print-color-adjust:exact !important;">📋 RINCIAN PENDUDUK DATANG (' + events.datang.length + ' Orang)</h3>';
    html += '<table style="width:100%;border-collapse:collapse;margin-top:3px;font-size:8.5px;">';
    html += '<thead><tr>';
    html += '<th ' + thStyle + ' style="width:4%;">NO</th>';
    html += '<th ' + thStyle + ' style="width:14%;">NO KK</th>';
    html += '<th ' + thStyle + ' style="width:14%;">NIK</th>';
    html += '<th ' + thStyle + ' style="width:18%;">NAMA LENGKAP</th>';
    html += '<th ' + thStyle + ' style="width:10%;">RT / RW</th>';
    html += '<th ' + thStyle + ' style="width:10%;">NO. RUMAH</th>';
    html += '<th ' + thStyle + ' style="width:12%;">TANGGAL DATANG</th>';
    html += '<th ' + thStyle + ' style="width:18%;">ASAL DAERAH</th>';
    html += '</tr></thead><tbody>';
    events.datang.forEach(function(r, idx) {
      var nr = r['No. Rumah'] || r['No Rumah'] || '-';
      var tglFormatted = formatDateIndoServer(r['Tanggal Datang']);
      html += '<tr>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + (idx+1) + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + escapeHtmlServer(r['No KK']||'-') + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + escapeHtmlServer(r['NIK']||'-') + '</td>';
      html += '<td style="border:1px solid #000;text-align:left;padding:3px;font-weight:bold;">' + escapeHtmlServer(String(r['Nama Lengkap']||'-').toUpperCase()) + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">RT ' + escapeHtmlServer(r['RT']||'-') + ' / RW ' + escapeHtmlServer(r['RW']||'-') + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + escapeHtmlServer(String(nr).toUpperCase()) + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + escapeHtmlServer(tglFormatted) + '</td>';
      html += '<td style="border:1px solid #000;text-align:left;padding:3px;">' + escapeHtmlServer(String(r['Asal Daerah']||'-').toUpperCase()) + '</td>';
      html += '</tr>';
    });
    html += '</tbody></table></div>';
  }

  if (events.pergi && events.pergi.length > 0) {
    html += '<div style="margin-top:16px;page-break-inside:avoid;">';
    html += '<h3 style="font-size:11px;font-weight:bold;margin:8px 0 4px;color:#000000;background-color:#f1f5f9 !important;border:1px solid #000000;padding:4px 8px;border-radius:4px;-webkit-print-color-adjust:exact !important;">📋 RINCIAN PENDUDUK PERGI (' + events.pergi.length + ' Orang)</h3>';
    html += '<table style="width:100%;border-collapse:collapse;margin-top:3px;font-size:8.5px;">';
    html += '<thead><tr>';
    html += '<th ' + thStyle + ' style="width:4%;">NO</th>';
    html += '<th ' + thStyle + ' style="width:14%;">NO KK</th>';
    html += '<th ' + thStyle + ' style="width:14%;">NIK</th>';
    html += '<th ' + thStyle + ' style="width:18%;">NAMA LENGKAP</th>';
    html += '<th ' + thStyle + ' style="width:10%;">RT / RW</th>';
    html += '<th ' + thStyle + ' style="width:10%;">NO. RUMAH</th>';
    html += '<th ' + thStyle + ' style="width:12%;">TANGGAL PERGI</th>';
    html += '<th ' + thStyle + ' style="width:18%;">TUJUAN</th>';
    html += '</tr></thead><tbody>';
    events.pergi.forEach(function(r, idx) {
      var nr = r['No. Rumah'] || r['No Rumah'] || '-';
      var tglFormatted = formatDateIndoServer(r['Tanggal Pergi']);
      html += '<tr>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + (idx+1) + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + escapeHtmlServer(r['No KK']||'-') + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + escapeHtmlServer(r['NIK']||'-') + '</td>';
      html += '<td style="border:1px solid #000;text-align:left;padding:3px;font-weight:bold;">' + escapeHtmlServer(String(r['Nama Lengkap']||'-').toUpperCase()) + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">RT ' + escapeHtmlServer(r['RT']||'-') + ' / RW ' + escapeHtmlServer(r['RW']||'-') + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + escapeHtmlServer(String(nr).toUpperCase()) + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + escapeHtmlServer(tglFormatted) + '</td>';
      html += '<td style="border:1px solid #000;text-align:left;padding:3px;">' + escapeHtmlServer(String(r['Tujuan']||'-').toUpperCase()) + '</td>';
      html += '</tr>';
    });
    html += '</tbody></table></div>';
  }

  if (events.lahir && events.lahir.length > 0) {
    html += '<div style="margin-top:16px;page-break-inside:avoid;">';
    html += '<h3 style="font-size:11px;font-weight:bold;margin:8px 0 4px;color:#000000;background-color:#f1f5f9 !important;border:1px solid #000000;padding:4px 8px;border-radius:4px;-webkit-print-color-adjust:exact !important;">📋 RINCIAN DATA KELAHIRAN (' + events.lahir.length + ' Bayi)</h3>';
    html += '<table style="width:100%;border-collapse:collapse;margin-top:3px;font-size:8.5px;">';
    html += '<thead><tr>';
    html += '<th ' + thStyle + ' style="width:4%;">NO</th>';
    html += '<th ' + thStyle + ' style="width:16%;">NAMA BAYI</th>';
    html += '<th ' + thStyle + ' style="width:8%;">JK</th>';
    html += '<th ' + thStyle + ' style="width:12%;">TANGGAL LAHIR</th>';
    html += '<th ' + thStyle + ' style="width:14%;">TEMPAT LAHIR</th>';
    html += '<th ' + thStyle + ' style="width:14%;">NAMA AYAH</th>';
    html += '<th ' + thStyle + ' style="width:14%;">NAMA IBU</th>';
    html += '<th ' + thStyle + ' style="width:10%;">RT / RW</th>';
    html += '<th ' + thStyle + ' style="width:8%;">NO. RUMAH</th>';
    html += '</tr></thead><tbody>';
    events.lahir.forEach(function(r, idx) {
      var nr = r['No. Rumah'] || r['No Rumah'] || '-';
      var tglFormatted = formatDateIndoServer(r['Tanggal Lahir']);
      html += '<tr>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + (idx+1) + '</td>';
      html += '<td style="border:1px solid #000;text-align:left;padding:3px;font-weight:bold;">' + escapeHtmlServer(String(r['Nama Bayi']||'-').toUpperCase()) + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + escapeHtmlServer(String(r['Jenis Kelamin']||'-').toUpperCase()) + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + escapeHtmlServer(tglFormatted) + '</td>';
      html += '<td style="border:1px solid #000;text-align:left;padding:3px;">' + escapeHtmlServer(String(r['Tempat Lahir']||'-').toUpperCase()) + '</td>';
      html += '<td style="border:1px solid #000;text-align:left;padding:3px;">' + escapeHtmlServer(String(r['Nama Ayah']||'-').toUpperCase()) + '</td>';
      html += '<td style="border:1px solid #000;text-align:left;padding:3px;">' + escapeHtmlServer(String(r['Nama Ibu']||'-').toUpperCase()) + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">RT ' + escapeHtmlServer(r['RT']||'-') + ' / RW ' + escapeHtmlServer(r['RW']||'-') + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + escapeHtmlServer(String(nr).toUpperCase()) + '</td>';
      html += '</tr>';
    });
    html += '</tbody></table></div>';
  }

  if (events.meninggal && events.meninggal.length > 0) {
    html += '<div style="margin-top:16px;page-break-inside:avoid;">';
    html += '<h3 style="font-size:11px;font-weight:bold;margin:8px 0 4px;color:#000000;background-color:#f1f5f9 !important;border:1px solid #000000;padding:4px 8px;border-radius:4px;-webkit-print-color-adjust:exact !important;">📋 RINCIAN PENDUDUK MENINGGAL (' + events.meninggal.length + ' Orang)</h3>';
    html += '<table style="width:100%;border-collapse:collapse;margin-top:3px;font-size:8.5px;">';
    html += '<thead><tr>';
    html += '<th ' + thStyle + ' style="width:4%;">NO</th>';
    html += '<th ' + thStyle + ' style="width:14%;">NO KK</th>';
    html += '<th ' + thStyle + ' style="width:14%;">NIK</th>';
    html += '<th ' + thStyle + ' style="width:18%;">NAMA LENGKAP</th>';
    html += '<th ' + thStyle + ' style="width:10%;">RT / RW</th>';
    html += '<th ' + thStyle + ' style="width:10%;">NO. RUMAH</th>';
    html += '<th ' + thStyle + ' style="width:12%;">TANGGAL MENINGGAL</th>';
    html += '<th ' + thStyle + ' style="width:18%;">PENYEBAB</th>';
    html += '</tr></thead><tbody>';
    events.meninggal.forEach(function(r, idx) {
      var nr = r['No. Rumah'] || r['No Rumah'] || '-';
      var tglFormatted = formatDateIndoServer(r['Tanggal Meninggal']);
      html += '<tr>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + (idx+1) + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + escapeHtmlServer(r['No KK']||'-') + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + escapeHtmlServer(r['NIK']||'-') + '</td>';
      html += '<td style="border:1px solid #000;text-align:left;padding:3px;font-weight:bold;">' + escapeHtmlServer(String(r['Nama Lengkap']||'-').toUpperCase()) + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">RT ' + escapeHtmlServer(r['RT']||'-') + ' / RW ' + escapeHtmlServer(r['RW']||'-') + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + escapeHtmlServer(String(nr).toUpperCase()) + '</td>';
      html += '<td style="border:1px solid #000;text-align:center;padding:3px;">' + escapeHtmlServer(tglFormatted) + '</td>';
      html += '<td style="border:1px solid #000;text-align:left;padding:3px;">' + escapeHtmlServer(String(r['Penyebab']||'-').toUpperCase()) + '</td>';
      html += '</tr>';
    });
    html += '</tbody></table></div>';
  }

  return html;
}

/**
 * Mengambil seluruh anggota keluarga berdasarkan No KK untuk Salinan Kartu Keluarga
 */
function getKartuKeluargaData(noKK) {
  try {
    if (!noKK) return { success: false, message: 'Nomor KK tidak valid' };
    var cleanNoKK = String(noKK).trim();
    
    var allWarga = getData('DataWarga');
    var members = [];
    
    for (var i = 0; i < allWarga.length; i++) {
      var row = allWarga[i];
      if (String(row['No KK'] || '').trim() === cleanNoKK) {
        members.push(row);
      }
    }
    
    if (members.length === 0) {
      return { success: false, message: 'Data keluarga dengan No KK ' + cleanNoKK + ' tidak ditemukan.' };
    }
    
    // Urutkan anggota: Kepala Keluarga / Rumah Tangga paling atas (1), diikuti Suami/Istri, Anak, dst.
    members.sort(function(a, b) {
      var hubA = String(a['Hubungan Keluarga'] || '').toLowerCase();
      var hubB = String(b['Hubungan Keluarga'] || '').toLowerCase();
      
      var scoreA = 99;
      if (hubA.indexOf('kepala') !== -1) scoreA = 1;
      else if (hubA.indexOf('suami') !== -1) scoreA = 2;
      else if (hubA.indexOf('istri') !== -1) scoreA = 3;
      else if (hubA.indexOf('anak') !== -1) scoreA = 4;
      
      var scoreB = 99;
      if (hubB.indexOf('kepala') !== -1) scoreB = 1;
      else if (hubB.indexOf('suami') !== -1) scoreB = 2;
      else if (hubB.indexOf('istri') !== -1) scoreB = 3;
      else if (hubB.indexOf('anak') !== -1) scoreB = 4;
      
      return scoreA - scoreB;
    });
    
    var headOfFamily = members[0] || {};
    var rawDusun = String(headOfFamily['Dusun'] || '1').trim();
    var dusunVal = (rawDusun.toLowerCase().indexOf('dusun') === 0) ? rawDusun : ('Dusun ' + rawDusun);
    var nrRaw = String(headOfFamily['No. Rumah'] || headOfFamily['Alamat'] || '').trim();

    var alamatClean = '';
    var nrUpper = nrRaw.toUpperCase().trim();
    if (!nrUpper || nrUpper === '-' || nrUpper === 'NULL' || nrUpper === 'UNDEFINED') {
      alamatClean = ('KMP. WARNASARI ' + dusunVal.toUpperCase()).trim();
    } else if (
      nrUpper.indexOf('KMP') !== -1 ||
      nrUpper.indexOf('KP') !== -1 ||
      nrUpper.indexOf('KAMPUNG') !== -1 ||
      nrUpper.indexOf('WARNASARI') !== -1 ||
      nrUpper.indexOf('DUSUN') !== -1 ||
      nrUpper.indexOf('DSN') !== -1 ||
      nrUpper.indexOf('JL') !== -1 ||
      nrUpper.indexOf('JALAN') !== -1 ||
      nrUpper.indexOf('GG') !== -1 ||
      nrUpper.indexOf('GANG') !== -1
    ) {
      alamatClean = nrUpper;
    } else if (/^(NO\.?\s*)?\d+[A-Z]?$/i.test(nrUpper)) {
      var houseNum = nrUpper.replace(/^NO\.?\s*/i, '');
      alamatClean = ('KMP. WARNASARI ' + dusunVal.toUpperCase() + ' NO. ' + houseNum).trim();
    } else {
      alamatClean = nrUpper;
    }

    var rtVal = String(headOfFamily['RT'] || '-').replace(/rt/i, '').trim();
    var rwVal = String(headOfFamily['RW'] || '-').replace(/rw/i, '').trim();
    if (rtVal && rtVal.length === 1) rtVal = '0' + rtVal;
    if (rwVal && rwVal.length === 1) rwVal = '0' + rwVal;

    var info = {
      noKK: cleanNoKK,
      kepalaKeluarga: headOfFamily['Nama Lengkap'] || '-',
      alamat: alamatClean,
      rt: rtVal,
      rw: rwVal,
      dusun: dusunVal,
      desa: 'WARNASARI',
      kecamatan: 'PANGALENGAN',
      kabupaten: 'BANDUNG',
      kodePos: '40378',
      provinsi: 'JAWA BARAT',
      totalAnggota: members.length
    };
    
    return {
      success: true,
      info: info,
      members: members
    };
  } catch(e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

/**
 * Membuat PDF resmi Salinan Kartu Keluarga berdasarkan No KK
 */
function generatePDF_KartuKeluarga(noKK, u) {
  try {
    if (!noKK) return { success: false, message: 'Nomor KK tidak valid' };
    var res = getKartuKeluargaData(noKK);
    if (!res || !res.success) return res;

    var info = res.info;
    var members = res.members;
    
    if (u && u.username) {
      writeAuditLog(u.username, u.role, 'EXPORT_PDF_KK', 'PDF Salinan KK ' + info.noKK);
    }

    var html = '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><style>';
    html += '@page { size: A4 landscape; margin: 10mm 10mm 10mm 10mm; }';
    html += '* { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; box-sizing: border-box; }';
    html += 'body { font-family: Arial, sans-serif; font-size: 10px; padding: 15px; color: #000000; background: #ffffff; }';
    html += '.title-box { text-align: center; margin-bottom: 15px; border-bottom: 2px solid #000; padding-bottom: 8px; }';
    html += '.title-box h1 { font-size: 18px; font-weight: bold; margin: 0; text-transform: uppercase; letter-spacing: 1px; }';
    html += '.title-box .sub { font-size: 12px; font-weight: bold; margin-top: 4px; }';
    html += '.meta-grid { display: table; width: 100%; margin-bottom: 15px; font-size: 10.5px; line-height: 1.6; }';
    html += '.meta-col { display: table-cell; width: 50%; vertical-align: top; }';
    html += '.meta-row { margin-bottom: 2px; }';
    html += '.meta-label { display: inline-block; width: 130px; font-weight: bold; }';
    html += 'table { width: 100%; border-collapse: collapse; margin-bottom: 15px; font-size: 9.5px; }';
    html += 'th { background-color: #f1f5f9 !important; color: #000 !important; border: 1px solid #000; padding: 5px; text-align: left; font-weight: bold; text-transform: uppercase; }';
    html += 'td { border: 1px solid #000; padding: 4px 5px; text-align: left; font-size: 9px; }';
    html += '.sig-box { display: table; width: 100%; margin-top: 25px; font-size: 10.5px; }';
    html += '.sig-col { display: table-cell; width: 50%; text-align: center; vertical-align: top; }';
    html += '.footer-copy { position: fixed; bottom: 0px; left: 0px; right: 0px; width: 100%; border-top: 1px solid #cbd5e1; padding-top: 6px; text-align: center; font-size: 8.5px; color: #475569; background: #ffffff; }';
    html += '</style></head><body>';

    html += '<div class="title-box">';
    html += '<h1>SALINAN KARTU KELUARGA</h1>';
    html += '<div class="sub">No. ' + escapeHtmlServer(info.noKK) + '</div>';
    html += '</div>';

    html += '<div class="meta-grid">';
    html += '<div class="meta-col">';
    html += '<div class="meta-row"><span class="meta-label">ALAMAT</span>: ' + escapeHtmlServer(info.alamat) + '</div>';
    html += '<div class="meta-row"><span class="meta-label">RT / RW</span>: ' + escapeHtmlServer(info.rt) + ' / ' + escapeHtmlServer(info.rw) + '</div>';
    html += '<div class="meta-row"><span class="meta-label">DESA / KELURAHAN</span>: WARNASARI</div>';
    html += '<div class="meta-row"><span class="meta-label">KECAMATAN</span>: PANGALENGAN</div>';
    html += '</div>';
    html += '<div class="meta-col">';
    html += '<div class="meta-row"><span class="meta-label">KABUPATEN</span>: BANDUNG</div>';
    html += '<div class="meta-row"><span class="meta-label">KODE POS</span>: 40378</div>';
    html += '<div class="meta-row"><span class="meta-label">PROVINSI</span>: JAWA BARAT</div>';
    html += '<div class="meta-row"><span class="meta-label">JUMLAH ANGGOTA</span>: ' + info.totalAnggota + ' Jiwa</div>';
    html += '</div>';
    html += '</div>';

    // Tabel 1: Demografi & Pekerjaan + Golongan Darah
    html += '<table><thead><tr>';
    html += '<th style="text-align:center;width:30px;">NO</th><th>NAMA LENGKAP</th><th>NIK</th><th>JENIS KELAMIN</th><th>TEMPAT LAHIR</th><th>TANGGAL LAHIR</th><th>AGAMA</th><th>PENDIDIKAN</th><th>JENIS PEKERJAAN</th><th>GOLONGAN DARAH</th>';
    html += '</tr></thead><tbody>';
    members.forEach(function(m, idx) {
      var rawNik = String(m['NIK'] || '-').trim();
      if (rawNik.indexOf('e+') !== -1 || rawNik.indexOf('E+') !== -1) {
        rawNik = Number(rawNik).toFixed(0);
      }
      var tglLahirVal = formatDateIndoServer(m['Tanggal Lahir']);
      var golDarah = String(m['Golongan Darah'] || m['Gol. Darah'] || 'TIDAK TAHU').toUpperCase();

      html += '<tr>';
      html += '<td style="text-align:center;">' + (idx+1) + '</td>';
      html += '<td style="font-weight:bold;">' + escapeHtmlServer(String(m['Nama Lengkap']||'-').toUpperCase()) + '</td>';
      html += '<td style="font-family:Arial,sans-serif;color:#000000;">' + escapeHtmlServer(rawNik) + '</td>';
      html += '<td>' + escapeHtmlServer(String(m['Jenis Kelamin']||'-').toUpperCase()) + '</td>';
      html += '<td>' + escapeHtmlServer(String(m['Tempat Lahir']||'-').toUpperCase()) + '</td>';
      html += '<td>' + escapeHtmlServer(tglLahirVal) + '</td>';
      html += '<td>' + escapeHtmlServer(String(m['Agama']||'ISLAM').toUpperCase()) + '</td>';
      html += '<td>' + escapeHtmlServer(String(m['Pendidikan']||'-').toUpperCase()) + '</td>';
      html += '<td>' + escapeHtmlServer(String(m['Pekerjaan']||'-').toUpperCase()) + '</td>';
      html += '<td>' + escapeHtmlServer(golDarah) + '</td>';
      html += '</tr>';
    });
    html += '</tbody></table>';

    // Tabel 2: Status Perkawinan & Hubungan Keluarga + Tanggal Perkawinan, Paspor, KITAS
    html += '<table><thead><tr>';
    html += '<th style="text-align:center;width:30px;">NO</th><th>STATUS PERKAWINAN</th><th style="text-align:center;">TANGGAL PERKAWINAN</th><th>STATUS HUBUNGAN DALAM KELUARGA</th><th style="text-align:center;">KEWARGANEGARAAN</th><th style="text-align:center;">NO. PASPOR</th><th style="text-align:center;">NO. KITAS/KITAP</th><th>NAMA AYAH</th><th>NAMA IBU</th>';
    html += '</tr></thead><tbody>';
    members.forEach(function(m, idx) {
      var tglKawin = m['Tanggal Perkawinan'] ? formatDateIndoServer(m['Tanggal Perkawinan']) : '-';
      var pasporVal = String(m['No. Paspor'] || m['Paspor'] || '-').toUpperCase();
      var kitasVal = String(m['No. KITAS'] || m['No. KITAP'] || m['KITAS/KITAP'] || '-').toUpperCase();

      html += '<tr>';
      html += '<td style="text-align:center;">' + (idx+1) + '</td>';
      html += '<td>' + escapeHtmlServer(String(m['Status Perkawinan']||'-').toUpperCase()) + '</td>';
      html += '<td style="text-align:center;">' + escapeHtmlServer(tglKawin) + '</td>';
      html += '<td style="font-weight:bold;">' + escapeHtmlServer(String(m['Hubungan Keluarga']||'-').toUpperCase()) + '</td>';
      html += '<td style="text-align:center;">WNI</td>';
      html += '<td style="text-align:center;">' + escapeHtmlServer(pasporVal) + '</td>';
      html += '<td style="text-align:center;">' + escapeHtmlServer(kitasVal) + '</td>';
      html += '<td>' + escapeHtmlServer(String(m['Nama Ayah']||'-').toUpperCase()) + '</td>';
      html += '<td>' + escapeHtmlServer(String(m['Nama Ibu']||'-').toUpperCase()) + '</td>';
      html += '</tr>';
    });
    html += '</tbody></table>';

    // Signatures
    var d = new Date();
    var blnIndo = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    var tglCetak = d.getDate() + ' ' + blnIndo[d.getMonth()] + ' ' + d.getFullYear();

    html += '<div class="sig-box">';
    html += '<div class="sig-col">';
    html += '<div style="font-weight:bold;margin-bottom:50px;">KEPALA KELUARGA,</div>';
    html += '<div style="font-weight:bold;text-transform:uppercase;">' + escapeHtmlServer(String(info.kepalaKeluarga).toUpperCase()) + '</div>';
    html += '</div>';
    html += '<div class="sig-col">';
    html += '<div style="font-weight:bold;">Warnasari, ' + tglCetak + '</div>';
    html += '<div style="font-weight:bold;margin-bottom:50px;">KEPALA DESA WARNASARI,</div>';
    html += '<div style="border-bottom:1px solid #000;width:180px;margin:0 auto;height:1px;"></div>';
    html += '</div>';
    html += '</div>';

    html += '<div class="footer-copy">';
    html += '&copy; 2026 Tim KKN 06 Ikopin University &bull; Pemerintah Desa Warnasari. Seluruh Hak Cipta Dilindungi.';
    html += '</div>';

    html += '</body></html>';

    var blob = Utilities.newBlob(html, 'text/html', 'Salinan_KK_' + info.noKK + '.html').getAs('application/pdf');
    var bytes = blob.getBytes();
    
    return {
      success: true,
      blob: bytes,
      name: 'Salinan_KK_' + info.noKK + '.pdf',
      html: html
    };
  } catch(e) {
    return { success: false, message: 'Error generate PDF KK: ' + e.message };
  }
}

// ============================================================
// BATCH INSERT 1 KK UNTUK DATA WARGA
// ============================================================

function saveWargaKeluargaBatch(user, payload) {
  try {
    user = verifyUser(user);
    if (!user) return { success: false, message: 'Sesi tidak valid' };
    if (!payload || !payload.noKK || !payload.members || payload.members.length === 0) {
      return { success: false, message: 'Data keluarga/anggota tidak lengkap' };
    }

    var sheet = getSheet('DataWarga');
    if (!sheet) return { success: false, message: 'Sheet DataWarga tidak ditemukan' };

    var noKK = String(payload.noKK || '').trim();
    var rw = normalizeRW(payload.rw || user.rw || '13');
    var rt = normalizeRT(payload.rt || user.rt || '1');

    if (!canAccessAllData(user)) {
      if (user.role === 'rt') {
        rw = normalizeRW(user.rw);
        rt = normalizeRT(user.rt);
      } else if (user.role === 'rw') {
        rw = normalizeRW(user.rw);
      }
    }

    var accessErr = _checkDataAccess(user, { RW: rw, RT: rt }, 'SAVE_BATCH_KK');
    if (accessErr) return accessErr;

    var dusun = payload.dusun || 'Dusun 1';
    var noRumah = String(payload.noRumah || '').trim();

    // Validasi maksimal 1 Kepala Keluarga
    var kepalaKeluargaCount = 0;
    for (var k = 0; k < payload.members.length; k++) {
      var hubCheck = String(payload.members[k].hub || payload.members[k].hubungan || '').trim();
      if (hubCheck === 'Kepala Keluarga') {
        kepalaKeluargaCount++;
      }
    }
    if (kepalaKeluargaCount > 1) {
      return { success: false, message: 'Dalam 1 Kartu Keluarga (KK) hanya boleh ada 1 Kepala Keluarga!' };
    }

    // Ambil daftar NIK & Nama+DOB terdaftar di DataWarga
    var existingWarga = getData('DataWarga');
    var existingNikMap = {};
    var existingNameDobMap = {};
    existingWarga.forEach(function(w) {
      if (w.NIK) {
        existingNikMap[normalizeNIK(w.NIK)] = w['Nama Lengkap'] || w.Nama || 'Warga lain';
      }
      var wNama = String(w['Nama Lengkap'] || w.Nama || '').trim().toUpperCase();
      var wTgl = formatDate(w['Tanggal Lahir'] || w.TglLahir || '');
      if (wNama && wTgl) {
        existingNameDobMap[wNama + '_' + wTgl] = w.NIK;
      }
    });

    var seenBatchNik = {};
    for (var i = 0; i < payload.members.length; i++) {
      var m = payload.members[i];
      var nikNorm = normalizeNIK(m.nik);

      if (!nikNorm || nikNorm.length < 16) {
        return { success: false, message: 'NIK pada Anggota #' + (i + 1) + ' (' + (m.nama || '') + ') harus 16 digit angka valid!' };
      }

      // Cek ganda internal dalam formulir ini
      if (seenBatchNik[nikNorm]) {
        return { success: false, message: 'NIK ' + m.nik + ' terdeteksi ganda/sama dalam daftar formulir ini (Anggota #' + (i + 1) + ')!' };
      }
      seenBatchNik[nikNorm] = true;

      // Cek ganda dengan database DataWarga
      if (existingNikMap[nikNorm]) {
        return { success: false, message: 'NIK ' + m.nik + ' (' + m.nama + ') sudah terdaftar di database atas nama: ' + existingNikMap[nikNorm] + '! NIK wajib unik.' };
      }

      // Cek ganda Nama + Tanggal Lahir (Pencegahan typo NIK)
      var normNama = String(m.nama || '').trim().toUpperCase();
      var normTgl = formatDate(m.tglLahir || m.tanggalLahir || '');
      if (normNama && normTgl && existingNameDobMap[normNama + '_' + normTgl]) {
        var matchNik = existingNameDobMap[normNama + '_' + normTgl];
        if (matchNik !== nikNorm) {
          return { success: false, message: 'Warga bernama ' + m.nama + ' dengan Tanggal Lahir ' + m.tglLahir + ' sudah terdaftar di database dengan NIK ' + matchNik + '! Periksa kembali NIK yang dimasukkan.' };
        }
      }
    }

    var rowsToAdd = [];
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var now = new Date();
    var curTahun = String(now.getFullYear());
    var curBulan = String(now.getMonth() + 1);

    payload.members.forEach(function(m) {
      var row = [];
      headers.forEach(function(h) {
        var key = String(h).trim();
        if (key === 'No. KK' || key === 'No KK') row.push(noKK);
        else if (key === 'NIK') row.push(String(m.nik || '').trim());
        else if (key === 'Nama Lengkap' || key === 'Nama') row.push(String(m.nama || '').trim().toUpperCase());
        else if (key === 'Dusun') row.push(dusun);
        else if (key === 'RW') row.push(rw);
        else if (key === 'RT') row.push(rt);
        else if (key === 'Detail No. Rumah' || key === 'No. Rumah' || key === 'Alamat') row.push(noRumah);
        else if (key === 'Jenis Kelamin') row.push(m.jk || 'Laki-laki');
        else if (key === 'Hubungan Keluarga') row.push(m.hub || 'Anak');
        else if (key === 'Tempat Lahir') row.push(m.tempatLahir || '');
        else if (key === 'Tanggal Lahir') row.push(m.tglLahir || '');
        else if (key === 'Agama') row.push(m.agama || 'Islam');
        else if (key === 'Pendidikan') row.push(m.pendidikan || 'SMA');
        else if (key === 'Pekerjaan') row.push(m.pekerjaan || 'Wiraswasta');
        else if (key === 'Status Perkawinan') row.push(m.statusPerkawinan || 'Belum Kawin');
        else if (key === 'Nama Ayah') row.push(m.namaAyah || '-');
        else if (key === 'Nama Ibu') row.push(m.namaIbu || '-');
        else if (key === 'Status') row.push('Aktif');
        else if (key === 'Tahun') row.push(curTahun);
        else if (key === 'Bulan') row.push(curBulan);
        else row.push('');
      });
      rowsToAdd.push(row);
    });

    if (rowsToAdd.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, headers.length).setValues(rowsToAdd);
    }

    clearDataCache();
    updateLaporan();

    writeAuditLog(user.username, user.role, 'TAMBAH_KK_BATCH', 'Menambahkan ' + rowsToAdd.length + ' warga dalam 1 KK (No. KK: ' + noKK + ')');

    return {
      success: true,
      count: rowsToAdd.length,
      message: 'Berhasil menambahkan ' + rowsToAdd.length + ' anggota keluarga!'
    };
  } catch(e) {
    return { success: false, message: 'Gagal menyimpan KK: ' + e.message };
  }
}

function tryFixAndParseJSON(jsonStr) {
  try {
    return JSON.parse(jsonStr);
  } catch (e) {
    var fixed = jsonStr.replace(/[\r\n\t]+/g, ' ');
    fixed = fixed.replace(/,\s*([\}\]])/g, '$1');

    try {
      return JSON.parse(fixed);
    } catch (e2) {
      var openBraces = (fixed.match(/\{/g) || []).length;
      var closeBraces = (fixed.match(/\}/g) || []).length;
      var openBrackets = (fixed.match(/\[/g) || []).length;
      var closeBrackets = (fixed.match(/\]/g) || []).length;

      if (fixed.lastIndexOf('"') > fixed.lastIndexOf('}')) {
        fixed += '"';
      }
      while (closeBrackets < openBrackets) {
        fixed += ']';
        closeBrackets++;
      }
      while (closeBraces < openBraces) {
        fixed += '}';
        closeBraces++;
      }

      return JSON.parse(fixed);
    }
  }
}

// ============================================================
// AI SCAN KARTU KELUARGA (GEMINI VISION OCR)
// ============================================================

function scanKKWithAI(base64Data, mimeType, userApiKey, u) {
  try {
    u = verifyUser(u);
    if (!u) return { success: false, message: 'Sesi tidak valid.' };

    var apiKey = (userApiKey && userApiKey.trim()) ? userApiKey.trim() : '';
    if (!apiKey) {
      try {
        apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
      } catch (errKey) { }
    }

    if (!apiKey) {
      return {
        success: false,
        needApiKey: true,
        message: 'API Key Gemini diperlukan. Dapatkan API Key gratis di https://aistudio.google.com/app/apikey'
      };
    }

    if (!base64Data) {
      return { success: false, message: 'Data file tidak boleh kosong.' };
    }

    if (base64Data.indexOf('base64,') !== -1) {
      base64Data = base64Data.split('base64,')[1];
    }

    mimeType = mimeType || 'image/jpeg';

    var promptText = "Ekstrak seluruh informasi dari dokumen Kartu Keluarga (KK) ini secara teliti. Jika namaAyah atau namaIbu tidak terbaca/tidak ada, gunakan '-'. " +
      "HANYA KELUARKAN SATU OBJEK JSON VALID. TANPA KATA PENGANTAR, TANPA MARKDOWN.\n" +
      "Format JSON:\n" +
      "{\n" +
      '  "noKK": "...",\n' +
      '  "namaKepalaKeluarga": "...",\n' +
      '  "alamat": "...",\n' +
      '  "rt": "...",\n' +
      '  "rw": "...",\n' +
      '  "anggota": [\n' +
      "    {\n" +
      '      "nama": "...",\n' +
      '      "nik": "...",\n' +
      '      "jenisKelamin": "Laki-laki / Perempuan",\n' +
      '      "tempatLahir": "...",\n' +
      '      "tanggalLahir": "YYYY-MM-DD",\n' +
      '      "agama": "Islam / Kristen / Katolik / Hindu / Buddha / Konghucu",\n' +
      '      "pendidikan": "SD / SMP / SMA / D3 / S1 / S2 / S3 / Tidak Sekolah",\n' +
      '      "pekerjaan": "...",\n' +
      '      "statusPerkawinan": "Belum Kawin / Kawin / Cerai Hidup / Cerai Mati",\n' +
      '      "hubunganKeluarga": "Kepala Keluarga / Kepala Rumah Tangga / Istri / Anak / Lainnya",\n' +
      '      "namaAyah": "...",\n' +
      '      "namaIbu": "..."\n' +
      "    }\n" +
      "  ]\n" +
      "}";

    var payload = {
      "contents": [{
        "parts": [
          { "text": promptText },
          {
            "inlineData": {
              "mimeType": mimeType,
              "data": base64Data
            }
          }
        ]
      }],
      "generationConfig": {
        "temperature": 0.1,
        "maxOutputTokens": 8192
      }
    };

    // STEP 1: Cek ke Google AI Studio model Gemini versi apa saja yang aktif untuk API Key ini
    var modelsToTry = [];
    var listErrorMsg = "";
    
    try {
      var listUrl = "https://generativelanguage.googleapis.com/v1beta/models?key=" + apiKey;
      var listResp = UrlFetchApp.fetch(listUrl, { "muteHttpExceptions": true });
      var listCode = listResp.getResponseCode();
      var listText = listResp.getContentText();

      if (listCode === 200) {
        var listJson = JSON.parse(listText);
        if (listJson.models && listJson.models.length > 0) {
          listJson.models.forEach(function(m) {
            var supportsGenerate = !m.supportedGenerationMethods || m.supportedGenerationMethods.indexOf('generateContent') !== -1;
            if (supportsGenerate) {
              var name = m.name ? m.name.replace(/^models\//, '') : '';
              if (name && modelsToTry.indexOf(name) === -1) {
                modelsToTry.push(name);
              }
            }
          });
        }
      } else {
        var errObj = {};
        try { errObj = JSON.parse(listText); } catch(e) {}
        listErrorMsg = (errObj.error && errObj.error.message) ? errObj.error.message : listText;
        return { success: false, message: "API Key Gemini Error / ListModels (" + listCode + "): " + listErrorMsg };
      }
    } catch (e) {
      listErrorMsg = e.message;
    }

    // Prioritaskan model flash yang cepat dan multimodal
    modelsToTry.sort(function(a, b) {
      if (a.indexOf('flash') !== -1 && b.indexOf('flash') === -1) return -1;
      if (a.indexOf('flash') === -1 && b.indexOf('flash') !== -1) return 1;
      return 0;
    });

    if (modelsToTry.length === 0) {
      modelsToTry = ['gemini-1.5-flash-8b', 'gemini-1.5-flash-001', 'gemini-2.0-flash-exp', 'gemini-2.0-flash'];
    }

    // STEP 2: Panggil model yang sudah terverifikasi aktif dari daftar
    var responseCode = 0;
    var responseText = "";
    var lastError = "";

    for (var i = 0; i < modelsToTry.length; i++) {
      var modelName = modelsToTry[i];
      var url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelName + ":generateContent?key=" + apiKey;
      var options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
      };

      try {
        var response = UrlFetchApp.fetch(url, options);
        responseCode = response.getResponseCode();
        responseText = response.getContentText();

        if (responseCode === 200) {
          break;
        } else {
          lastError = "[" + modelName + "]: " + responseText;
        }
      } catch (e) {
        lastError = "[" + modelName + "]: " + e.message;
      }
    }

    if (responseCode !== 200) {
      var errObj = {};
      try { errObj = JSON.parse(responseText || lastError); } catch (e) { }
      var errMsg = (errObj.error && errObj.error.message) ? errObj.error.message : (responseText || lastError);
      return { success: false, message: "Gagal analisis (" + responseCode + "): " + errMsg };
    }

    var result = JSON.parse(responseText);
    var candidateText = "";
    if (result.candidates && result.candidates[0] && result.candidates[0].content && result.candidates[0].content.parts) {
      candidateText = result.candidates[0].content.parts[0].text;
    }

    if (!candidateText) {
      return { success: false, message: "AI tidak mengembalikan hasil." };
    }

    var firstBrace = candidateText.indexOf('{');
    var lastBrace = candidateText.lastIndexOf('}');
    if (firstBrace === -1 || lastBrace === -1 || lastBrace <= firstBrace) {
      return { success: false, message: "Respon AI bukan format JSON valid." };
    }

    var cleanJson = candidateText.substring(firstBrace, lastBrace + 1);
    var parsedData = tryFixAndParseJSON(cleanJson);

    writeAuditLog(u.username, u.role, 'SCAN_KK_SUCCESS', 'Berhasil scan KK: ' + (parsedData.noKK || 'Tanpa No KK'));
    return { success: true, data: parsedData };

  } catch (e) {
    writeAuditLog(u ? u.username : 'GUEST', u ? u.role : 'GUEST', 'SCAN_KK_ERROR', e.message);
    return { success: false, message: 'Gagal memproses Scan KK: ' + e.message };
  }
}

function saveScannedKKData(membersList, headerInfo, u) {
  try {
    u = verifyUser(u);
    if (!u) return { success: false, message: 'Sesi tidak valid.' };

    if (!membersList || !membersList.length) {
      return { success: false, message: 'Tidak ada data anggota keluarga untuk disimpan.' };
    }

    var checkRW = normalizeRW(headerInfo ? headerInfo.rw : (u ? u.rw : ''));
    var checkRT = String(headerInfo ? headerInfo.rt : (u ? u.rt : '')).trim();

    if (!canAccessAllData(u)) {
      if (u.role === 'rt') { checkRW = u.rw; checkRT = u.rt; }
      else if (u.role === 'rw') { checkRW = u.rw; }
    }

    if (!checkRW || !checkRT) {
      return { success: false, message: 'Peringatan: Wilayah RT dan RW wajib dipilih sebelum menyimpan data KK.' };
    }

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) {
      return { success: false, message: 'Server sedang sibuk, coba beberapa saat lagi.' };
    }

    var sheet = getSheet('DataWarga');
    if (!sheet) {
      lock.releaseLock();
      return { success: false, message: 'Sheet DataWarga tidak ditemukan.' };
    }

    var headers = getHeaders('DataWarga');
    var existingWarga = getData('DataWarga');
    var existingNikMap = {};
    existingWarga.forEach(function (w) {
      if (w.NIK) {
        existingNikMap[normalizeNIK(w.NIK)] = {
          rowIndex: w._rowIndex + 2,
          rw: normalizeRW(w.RW),
          rt: normalizeRT(w.RT),
          nama: w['Nama Lengkap'] || w.Nama || 'Warga lain'
        };
      }
    });

    // Pengecekan awal: Jika ada NIK yang terdaftar di RT lain dan user adalah Akun RT, tolak seluruh batch!
    for (var chkIdx = 0; chkIdx < membersList.length; chkIdx++) {
      var chkMem = membersList[chkIdx];
      var chkNik = normalizeNIK(chkMem.nik || '');
      if (!chkNik) continue;

      var chkRW = normalizeRW(chkMem.rw || (headerInfo ? headerInfo.rw : '') || u.rw || '');
      var chkRT = normalizeRT(chkMem.rt || (headerInfo ? headerInfo.rt : '') || u.rt || '');

      if (!canAccessAllData(u)) {
        if (u.role === 'rt') { chkRW = normalizeRW(u.rw); chkRT = normalizeRT(u.rt); }
        else if (u.role === 'rw') { chkRW = normalizeRW(u.rw); }
      }

      var existingRec = existingNikMap[chkNik];
      if (existingRec) {
        var isDiffRW = (existingRec.rw !== chkRW);
        var isDiffRT = (existingRec.rt !== chkRT);

        if (!canAccessAllData(u)) {
          if (u.role === 'rt' && (isDiffRT || isDiffRW)) {
            lock.releaseLock();
            return {
              success: false,
              message: 'Akses Ditolak: NIK ' + chkNik + ' (' + (chkMem.nama || existingRec.nama) + ') sudah terdaftar di RT ' + existingRec.rt + '/RW ' + existingRec.rw + '. Akun RT dilarang menimpa data RT/RW lain.'
            };
          } else if (u.role === 'rw' && isDiffRW) {
            lock.releaseLock();
            return {
              success: false,
              message: 'Akses Ditolak: NIK ' + chkNik + ' (' + (chkMem.nama || existingRec.nama) + ') sudah terdaftar di RW ' + existingRec.rw + ' (RT ' + existingRec.rt + '). Akun RW ' + u.rw + ' dilarang menimpa data RW lain.'
            };
          }
        }
      }
    }

    var countSaved = 0;
    var countUpdated = 0;
    var countInserted = 0;
    var now = new Date();
    var curYear = String(now.getFullYear());
    var curMonth = String(now.getMonth() + 1);

    membersList.forEach(function (m) {
      var nikNorm = normalizeNIK(m.nik || '');
      if (!nikNorm) return;

      var rw = normalizeRW(m.rw || (headerInfo ? headerInfo.rw : '') || u.rw || '');
      var rt = normalizeRT(m.rt || (headerInfo ? headerInfo.rt : '') || u.rt || '');

      if (!canAccessAllData(u)) {
        if (u.role === 'rt') {
          rw = normalizeRW(u.rw);
          rt = normalizeRT(u.rt);
        } else if (u.role === 'rw') {
          rw = normalizeRW(u.rw);
        }
      }

      var rowData = {
        'No KK': normalizeNIK(m.noKK || (headerInfo ? headerInfo.noKK : '') || ''),
        'NIK': nikNorm,
        'Nama Lengkap': sanitizeInput(m.nama || ''),
        'Dusun': 'Dusun 1',
        'RW': rw,
        'RT': rt,
        'No. Rumah': sanitizeInput(m.noRumah || (headerInfo ? headerInfo.alamat : '') || ''),
        'Jenis Kelamin': sanitizeInput(m.jenisKelamin || 'Laki-laki'),
        'Tempat Lahir': sanitizeInput(m.tempatLahir || ''),
        'Tanggal Lahir': formatDate(m.tanggalLahir || ''),
        'Status Perkawinan': sanitizeInput(m.statusPerkawinan || ''),
        'Pekerjaan': sanitizeInput(m.pekerjaan || ''),
        'Pendidikan': sanitizeInput(m.pendidikan || ''),
        'Agama': sanitizeInput(m.agama || 'Islam'),
        'Hubungan Keluarga': sanitizeInput(m.hubunganKeluarga || 'Anggota'),
        'Nama Ayah': sanitizeInput(m.namaAyah || '-'),
        'Nama Ibu': sanitizeInput(m.namaIbu || '-'),
        'Tahun': curYear,
        'Bulan': curMonth
      };

      rowData = formatDataObject(rowData);

      var existingRec = existingNikMap[nikNorm];
      if (existingRec && existingRec.rowIndex) {
        writeRowSafely('DataWarga', sheet, headers, existingRec.rowIndex, rowData);
        countUpdated++;
      } else {
        writeRowSafely('DataWarga', sheet, headers, sheet.getLastRow() + 1, rowData);
        countInserted++;
      }
      countSaved++;
    });

    applyTextFormatToColumn(sheet, headers, TEXT_COLUMNS['DataWarga']);
    clearDataCache();
    lock.releaseLock();

    writeAuditLog(u.username, u.role, 'SCAN_KK_SAVE', 'Scan KK Simpan/Timpa (' + countInserted + ' baru, ' + countUpdated + ' ditimpa) KK: ' + (headerInfo ? headerInfo.noKK : ''));

    try { updateLaporan(); } catch (e) { }

    return {
      success: true,
      count: countSaved,
      inserted: countInserted,
      updated: countUpdated,
      message: 'Berhasil memproses ' + countSaved + ' warga (' + countInserted + ' baru, ' + countUpdated + ' data ditimpa/diupdate).'
    };

  } catch (e) {
    try { LockService.getScriptLock().releaseLock(); } catch (err) { }
    return { success: false, message: 'Gagal menyimpan data scan: ' + e.message };
  }
}

function saveScannedKKAsDatang(membersList, headerInfo, datangInfo, u) {
  try {
    u = verifyUser(u);
    if (!u) return { success: false, message: 'Sesi tidak valid.' };

    if (!membersList || !membersList.length) {
      return { success: false, message: 'Tidak ada data anggota keluarga untuk disimpan.' };
    }

    var checkRW = normalizeRW(headerInfo ? headerInfo.rw : (u ? u.rw : ''));
    var checkRT = String(headerInfo ? headerInfo.rt : (u ? u.rt : '')).trim();

    if (!canAccessAllData(u)) {
      if (u.role === 'rt') { checkRW = u.rw; checkRT = u.rt; }
      else if (u.role === 'rw') { checkRW = u.rw; }
    }

    if (!checkRW || !checkRT) {
      return { success: false, message: 'Peringatan: Wilayah RT dan RW wajib dipilih sebelum menyimpan data KK.' };
    }

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) {
      return { success: false, message: 'Server sedang sibuk, coba beberapa saat lagi.' };
    }

    var sheetDatang = getSheet('DataDatang');
    var headersDatang = getHeaders('DataDatang');
    var sheetWarga = getSheet('DataWarga');
    var headersWarga = getHeaders('DataWarga');

    if (!sheetDatang || !sheetWarga) {
      lock.releaseLock();
      return { success: false, message: 'Sheet data tidak ditemukan.' };
    }

    var existingWarga = getData('DataWarga');
    var existingNikMap = {};
    existingWarga.forEach(function (w) {
      if (w.NIK) {
        existingNikMap[normalizeNIK(w.NIK)] = {
          rowIndex: w._rowIndex + 2,
          rw: normalizeRW(w.RW),
          rt: normalizeRT(w.RT),
          nama: w['Nama Lengkap'] || w.Nama || 'Warga lain'
        };
      }
    });

    // Check batch for RT permission
    for (var chkIdx = 0; chkIdx < membersList.length; chkIdx++) {
      var chkMem = membersList[chkIdx];
      var chkNik = normalizeNIK(chkMem.nik || '');
      if (!chkNik) continue;

      var chkRW = normalizeRW(chkMem.rw || (headerInfo ? headerInfo.rw : '') || u.rw || '');
      var chkRT = normalizeRT(chkMem.rt || (headerInfo ? headerInfo.rt : '') || u.rt || '');

      if (!canAccessAllData(u)) {
        if (u.role === 'rt') { chkRW = normalizeRW(u.rw); chkRT = normalizeRT(u.rt); }
        else if (u.role === 'rw') { chkRW = normalizeRW(u.rw); }
      }

      var existingRec = existingNikMap[chkNik];
      if (existingRec) {
        var isDiffRW = (existingRec.rw !== chkRW);
        var isDiffRT = (existingRec.rt !== chkRT);

        if (!canAccessAllData(u)) {
          if (u.role === 'rt' && (isDiffRT || isDiffRW)) {
            lock.releaseLock();
            return {
              success: false,
              message: 'Akses Ditolak: NIK ' + chkNik + ' (' + (chkMem.nama || existingRec.nama) + ') sudah terdaftar di RT ' + existingRec.rt + '/RW ' + existingRec.rw + '. Akun RT dilarang menimpa data RT/RW lain.'
            };
          } else if (u.role === 'rw' && isDiffRW) {
            lock.releaseLock();
            return {
              success: false,
              message: 'Akses Ditolak: NIK ' + chkNik + ' (' + (chkMem.nama || existingRec.nama) + ') sudah terdaftar di RW ' + existingRec.rw + ' (RT ' + existingRec.rt + '). Akun RW ' + u.rw + ' dilarang menimpa data RW lain.'
            };
          }
        }
      }
    }

    var now = new Date();
    var curYear = String(now.getFullYear());
    var curMonth = String(now.getMonth() + 1);
    var tglDatang = formatDate(datangInfo ? datangInfo.tglDatang : now);
    var asal = datangInfo ? (datangInfo.asal || '') : '';
    var alasan = datangInfo ? (datangInfo.alasan || 'Pindah Alamat') : 'Pindah Alamat';

    var countSaved = 0;
    membersList.forEach(function (m) {
      var nikNorm = normalizeNIK(m.nik || '');
      if (!nikNorm) return;

      var rw = normalizeRW(m.rw || (headerInfo ? headerInfo.rw : '') || u.rw || '');
      var rt = normalizeRT(m.rt || (headerInfo ? headerInfo.rt : '') || u.rt || '');

      if (!canAccessAllData(u)) {
        if (u.role === 'rt') { rw = normalizeRW(u.rw); rt = normalizeRT(u.rt); }
        else if (u.role === 'rw') { rw = normalizeRW(u.rw); }
      }

      var noRumah = sanitizeInput(m.noRumah || (headerInfo ? headerInfo.alamat : '') || '');

      var dataDatangObj = {
        'No KK': normalizeNIK(m.noKK || (headerInfo ? headerInfo.noKK : '') || ''),
        'NIK': nikNorm,
        'Nama Lengkap': sanitizeInput(m.nama || ''),
        'Dusun': 'Dusun 1',
        'RW': rw,
        'RT': rt,
        'No. Rumah': noRumah,
        'Jenis Kelamin': sanitizeInput(m.jenisKelamin || 'Laki-laki'),
        'Tempat Lahir': sanitizeInput(m.tempatLahir || ''),
        'Tanggal Lahir': formatDate(m.tanggalLahir || ''),
        'Status Perkawinan': sanitizeInput(m.statusPerkawinan || ''),
        'Pekerjaan': sanitizeInput(m.pekerjaan || ''),
        'Pendidikan': sanitizeInput(m.pendidikan || ''),
        'Agama': sanitizeInput(m.agama || 'Islam'),
        'Hubungan Keluarga': sanitizeInput(m.hubunganKeluarga || 'Anggota'),
        'Nama Ayah': sanitizeInput(m.namaAyah || '-'),
        'Nama Ibu': sanitizeInput(m.namaIbu || '-'),
        'Tanggal Datang': tglDatang,
        'Asal Daerah': asal,
        'Alasan': alasan,
        'Tahun': curYear,
        'Bulan': curMonth
      };
      dataDatangObj = formatDataObject(dataDatangObj);
      writeRowSafely('DataDatang', sheetDatang, headersDatang, sheetDatang.getLastRow() + 1, dataDatangObj);

      var dataWargaObj = {
        'No KK': normalizeNIK(m.noKK || (headerInfo ? headerInfo.noKK : '') || ''),
        'NIK': nikNorm,
        'Nama Lengkap': sanitizeInput(m.nama || ''),
        'Dusun': 'Dusun 1',
        'RW': rw,
        'RT': rt,
        'No. Rumah': noRumah,
        'Jenis Kelamin': sanitizeInput(m.jenisKelamin || 'Laki-laki'),
        'Tempat Lahir': sanitizeInput(m.tempatLahir || ''),
        'Tanggal Lahir': formatDate(m.tanggalLahir || ''),
        'Status Perkawinan': sanitizeInput(m.statusPerkawinan || ''),
        'Pekerjaan': sanitizeInput(m.pekerjaan || ''),
        'Pendidikan': sanitizeInput(m.pendidikan || ''),
        'Agama': sanitizeInput(m.agama || 'Islam'),
        'Hubungan Keluarga': sanitizeInput(m.hubunganKeluarga || 'Anggota'),
        'Nama Ayah': sanitizeInput(m.namaAyah || '-'),
        'Nama Ibu': sanitizeInput(m.namaIbu || '-'),
        'Tahun': curYear,
        'Bulan': curMonth
      };
      dataWargaObj = formatDataObject(dataWargaObj);

      var existingRec = existingNikMap[nikNorm];
      if (existingRec && existingRec.rowIndex) {
        writeRowSafely('DataWarga', sheetWarga, headersWarga, existingRec.rowIndex, dataWargaObj);
      } else {
        writeRowSafely('DataWarga', sheetWarga, headersWarga, sheetWarga.getLastRow() + 1, dataWargaObj);
      }

      countSaved++;
    });

    applyTextFormatToColumn(sheetDatang, headersDatang, TEXT_COLUMNS['DataDatang']);
    applyTextFormatToColumn(sheetWarga, headersWarga, TEXT_COLUMNS['DataWarga']);
    clearDataCache();
    lock.releaseLock();

    writeAuditLog(u.username, u.role, 'SCAN_KK_DATANG_SAVE', 'Menyimpan ' + countSaved + ' warga datang dari Scan KK: ' + (headerInfo ? headerInfo.noKK : ''));
    try { updateLaporan(); } catch (e) { }

    return {
      success: true,
      message: 'Berhasil mendaftarkan ' + countSaved + ' warga datang ke Data Datang & Data Warga.'
    };

  } catch (e) {
    try { LockService.getScriptLock().releaseLock(); } catch (err) { }
    return { success: false, message: 'Gagal menyimpan data scan warga datang: ' + e.message };
  }
}

function testFetchAuthorization() {
  UrlFetchApp.fetch("https://www.google.com");
}

// ============================================================
// EKSPOR & IMPOR MASTER DATA (UNTUK MIGRASI DAN BACKUP)
// ============================================================

function exportMasterDataJson(u) {
  try {
    u = verifyUser(u);
    if (!u) return { success: false, message: 'Sesi tidak valid. Silakan login ulang.' };
    if (!isSuperAdmin(u) && !isKadus(u)) {
      return { success: false, message: 'Hanya SuperAdmin dan Kadus yang dapat mengekspor Master Data!' };
    }

    var sheetNames = ['DataWarga', 'DataDatang', 'DataPergi', 'DataMeninggal', 'DataLahir', 'DataPengontrak'];
    var masterPackage = {
      version: '1.0',
      appName: 'Warnasari Data',
      dusun: u.dusun || 'Dusun 1',
      exportedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      exportedBy: u.username,
      data: {}
    };

    var totalRows = 0;
    sheetNames.forEach(function(sName) {
      var rows = getData(sName) || [];
      var cleanRows = rows.map(function(r) {
        var copy = {};
        for (var k in r) {
          if (r.hasOwnProperty(k) && k !== '_rowIndex') {
            copy[k] = r[k];
          }
        }
        if (!copy['Dusun'] || copy['Dusun'].trim() === '') {
          copy['Dusun'] = u.dusun || 'Dusun 1';
        }
        return copy;
      });
      masterPackage.data[sName] = cleanRows;
      totalRows += cleanRows.length;
    });

    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    var filename = 'MASTER_DATA_WARNASARI_DUSUN1_' + nowStr + '.json';

    writeAuditLog(u.username, u.role, 'EXPORT_MASTER_DATA', 'Ekspor ' + totalRows + ' data master ke JSON');

    return {
      success: true,
      filename: filename,
      jsonContent: JSON.stringify(masterPackage, null, 2),
      totalRecords: totalRows
    };
  } catch (e) {
    writeAuditLog(u ? u.username : 'unknown', u ? u.role : 'unknown', 'EXPORT_MASTER_ERROR', e.message);
    return { success: false, message: 'Gagal mengekspor data master: ' + e.message };
  }
}

function importMasterDataJson(u, jsonString) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { success: false, message: 'Server sibuk, coba lagi.' };
  }

  try {
    u = verifyUser(u);
    if (!u) {
      lock.releaseLock();
      return { success: false, message: 'Sesi tidak valid.' };
    }
    if (!isSuperAdmin(u) && !isKadus(u)) {
      lock.releaseLock();
      return { success: false, message: 'Hanya SuperAdmin dan Kadus yang dapat mengimpor Master Data!' };
    }

    if (!jsonString) {
      lock.releaseLock();
      return { success: false, message: 'File JSON kosong atau tidak valid.' };
    }

    var payload = JSON.parse(jsonString);
    if (!payload || !payload.data) {
      lock.releaseLock();
      return { success: false, message: 'Format file JSON tidak dikenali sebagai paket Master Data.' };
    }

    var sheetNames = ['DataWarga', 'DataDatang', 'DataPergi', 'DataMeninggal', 'DataLahir', 'DataPengontrak'];
    var importStats = {};
    var totalImported = 0;

    sheetNames.forEach(function(sName) {
      var rowsToImport = payload.data[sName];
      if (!rowsToImport || !Array.isArray(rowsToImport) || rowsToImport.length === 0) {
        importStats[sName] = 0;
        return;
      }

      var targetSheet = getSheet(sName);
      var headers = getHeaders(sName);
      var textCols = TEXT_COLUMNS[sName] || [];

      var existingData = getData(sName);
      var existingNikMap = {};
      existingData.forEach(function(exRow) {
        var exNik = normalizeNIK(exRow.NIK || exRow['NIK Bayi'] || exRow['No. Paspor']);
        if (exNik && exNik !== 'BELUMADA' && exNik !== '-') {
          existingNikMap[exNik] = exRow._rowIndex + 2;
        }
      });

      var countAdd = 0;
      var countUpdate = 0;

      rowsToImport.forEach(function(rawRow) {
        var rowData = sanitizeDataObject(rawRow, ['RW', 'RT', 'NIK', 'No KK', 'Tahun', 'Bulan']);
        if (rowData.RW) rowData.RW = normalizeRW(rowData.RW);
        if (rowData.RT) rowData.RT = String(rowData.RT).trim();
        if (rowData.NIK) rowData.NIK = normalizeNIK(rowData.NIK);
        if (!rowData.Dusun || rowData.Dusun.trim() === '') rowData.Dusun = u.dusun || 'Dusun 1';

        rowData = formatDataObject(rowData);

        var targetNik = normalizeNIK(rowData.NIK || rowData['NIK Bayi']);
        var existingRowIdx = (targetNik && targetNik !== 'BELUMADA' && targetNik !== '-') ? existingNikMap[targetNik] : null;

        if (existingRowIdx) {
          writeRowSafely(sName, targetSheet, headers, existingRowIdx, rowData);
          countUpdate++;
        } else {
          var nextRow = targetSheet.getLastRow() + 1;
          writeRowSafely(sName, targetSheet, headers, nextRow, rowData);
          if (targetNik && targetNik !== 'BELUMADA' && targetNik !== '-') {
            existingNikMap[targetNik] = nextRow;
          }
          countAdd++;
        }
      });

      applyTextFormatToColumn(targetSheet, headers, textCols);
      importStats[sName] = countAdd + countUpdate;
      totalImported += (countAdd + countUpdate);
    });

    clearDataCache();
    lock.releaseLock();
    try { updateLaporan(); } catch (e) { }

    writeAuditLog(u.username, u.role, 'IMPORT_MASTER_DATA', 'Berhasil impor ' + totalImported + ' data master dari JSON');

    return {
      success: true,
      message: 'Berhasil mengimpor ' + totalImported + ' data warga!',
      stats: importStats
    };

  } catch (e) {
    try { LockService.getScriptLock().releaseLock(); } catch (err) { }
    writeAuditLog(u ? u.username : 'unknown', u ? u.role : 'unknown', 'IMPORT_MASTER_ERROR', e.message);
    return { success: false, message: 'Gagal mengimpor data master: ' + e.message };
  }
}


