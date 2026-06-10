// ============================================================
//  CODE.GS — Google Apps Script Portal Akademik
//  Ahmad Musyafa, SKM, MPH | FETP Indonesia
//  Gabungan: Code.gs + Statistik.gs
// ============================================================

// ══════════════════════════════════════════════════════════════
//  BAGIAN 1 — PENGATURAN UTAMA & DRIVE
// ══════════════════════════════════════════════════════════════

const MAIN_FOLDER_ID = '1QypRuHDKCKBdygpNvw6iubvrVaS4uQBu';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Portal Akademik | Ahmad Musyafa, SKM, MPH')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Ambil daftar kategori (folder utama) ────────────────────
function getCategories(isAdmin) {
  try {
    const root    = DriveApp.getFolderById(MAIN_FOLDER_ID);
    const folders = root.getFolders();
    let list = [];
    while (folders.hasNext()) {
      const f        = folders.next();
      const isHidden = f.getDescription() === 'HIDDEN';
      if (isAdmin || !isHidden) {
        list.push({ id: f.getId(), name: f.getName(), isHidden: isHidden });
      }
    }
    return list.sort((a, b) => a.name.localeCompare(b.name));
  } catch (e) {
    console.error('getCategories error:', e);
    return [];
  }
}

// ── Ambil isi folder (sub-folder + file) ────────────────────
function getFolderContent(folderId, isAdmin) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    let data = { name: folder.getName(), subFolders: [], files: [] };

    // Sub-folder
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      const sf       = subFolders.next();
      const isHidden = sf.getDescription() === 'HIDDEN';
      if (isAdmin || !isHidden) {
        data.subFolders.push({ id: sf.getId(), name: sf.getName(), isHidden: isHidden });
      }
    }

    // File
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f         = files.next();
      const fId       = f.getId();
      const fName     = f.getName();
      const fNameLow  = fName.toLowerCase();
      const mime      = f.getMimeType();
      const isHidden  = f.getDescription() === 'HIDDEN';

      const isSlide = (mime === MimeType.GOOGLE_SLIDES || fNameLow.endsWith('.pptx') || fNameLow.endsWith('.ppt'));
      const isDoc   = (mime === MimeType.GOOGLE_DOCS   || fNameLow.endsWith('.docx') || fNameLow.endsWith('.doc'));
      const isSheet = (mime === MimeType.GOOGLE_SHEETS || fNameLow.endsWith('.xlsx') || fNameLow.endsWith('.xls'));
      const isPDF   = (mime === MimeType.PDF);

      // Pastikan file bisa diakses siapa saja dengan link
      try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}

      // rm=minimal menyembunyikan toolbar Google (tombol download & buka tab baru)
      let embedUrl = 'https://drive.google.com/file/d/' + fId + '/preview?rm=minimal';
      if (isSlide) embedUrl = 'https://docs.google.com/presentation/d/' + fId + '/embed?rm=minimal&loop=false&delayms=0';
      if (isDoc)   embedUrl = 'https://docs.google.com/document/d/'     + fId + '/preview?rm=minimal';
      if (isSheet) embedUrl = 'https://docs.google.com/spreadsheets/d/' + fId + '/preview?rm=minimal';

      let downloadUrl = null;
      if (isSlide)      downloadUrl = 'https://docs.google.com/presentation/d/'  + fId + '/export/pdf';
      else if (isDoc)   downloadUrl = 'https://docs.google.com/document/d/'      + fId + '/export?format=pdf';
      else if (isSheet) downloadUrl = 'https://docs.google.com/spreadsheets/d/'  + fId + '/export?format=pdf';
      else if (isPDF)   downloadUrl = 'https://drive.google.com/uc?export=download&id=' + fId;

      if (isAdmin || !isHidden) {
        data.files.push({
          id         : fId,
          name       : fName,
          isPDF      : isPDF,
          isHidden   : isHidden,
          embedUrl   : embedUrl,
          downloadUrl: isAdmin ? downloadUrl : null
        });
      }
    }
    return data;
  } catch (e) {
    console.error('getFolderContent error:', e);
    return null;
  }
}

// ── Fungsi Admin: Rename / Toggle Sembunyikan ───────────────
function updateItem(id, type, newName, toggleHide) {
  try {
    const item = (type === 'folder') ? DriveApp.getFolderById(id) : DriveApp.getFileById(id);
    if (newName)              item.setName(newName);
    if (toggleHide !== undefined) item.setDescription(toggleHide ? 'HIDDEN' : '');
    return { success: true };
  } catch (e) {
    return { success: false, msg: e.toString() };
  }
}

// ── Fungsi Admin: Hapus ke Sampah ──────────────────────────
function deleteItem(id, type) {
  try {
    if (type === 'folder') DriveApp.getFolderById(id).setTrashed(true);
    else                   DriveApp.getFileById(id).setTrashed(true);
    return { success: true };
  } catch (e) {
    return { success: false, msg: e.toString() };
  }
}

// ── Fungsi Admin: Buat Folder Baru ─────────────────────────
function createNewFolder(name, parentId) {
  try {
    DriveApp.getFolderById(parentId || MAIN_FOLDER_ID).createFolder(name);
    return { success: true };
  } catch (e) {
    return { success: false, msg: e.toString() };
  }
}

// ── Fungsi Admin: Upload File ───────────────────────────────
function uploadMaterialFile(data, fileName, folderId) {
  try {
    const folder      = DriveApp.getFolderById(folderId);
    const contentType = data.substring(5, data.indexOf(';'));
    const bytes       = Utilities.base64Decode(data.split(',')[1]);
    const blob        = Utilities.newBlob(bytes, contentType, fileName);
    const file        = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ── Autentikasi Admin ───────────────────────────────────────
function checkAuth(u, p) {
  // Ganti username & password sesuai kebutuhan
  return (u === 'ahmadm' && p === 'ahmad001')
    ? { authorized: true }
    : { authorized: false };
}

// ── Sinkronisasi Izin Massal (jalankan manual dari editor) ──
function sinkronisasiIzinMassal() {
  try {
    const root = DriveApp.getFolderById(MAIN_FOLDER_ID);
    if (root) perbaikiIzinRecursive_(root);
  } catch (e) {
    console.error('sinkronisasiIzinMassal error:', e);
  }
}

function perbaikiIzinRecursive_(folder) {
  if (!folder) return;
  try {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (file) file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) perbaikiIzinRecursive_(subFolders.next());
  } catch (e) {
    console.warn('perbaikiIzinRecursive_ warning:', e);
  }
}


// ══════════════════════════════════════════════════════════════
//  BAGIAN 2 — STATISTIK KUNJUNGAN
//  Dapat diakses oleh SEMUA PENGGUNA (bukan hanya admin)
// ══════════════════════════════════════════════════════════════

// ── Konfigurasi Sheet Statistik ─────────────────────────────
const STATS_SHEET_ID   = '1lD-GM_OxPVTHbCcsp3A8NvYmxQBT7G5B3DEdSJShivs';
const STATS_SHEET_NAME = 'Kunjungan';

// ── Helper: Ambil atau buat sheet ───────────────────────────
function getOrCreateSheet_() {
  const ss    = SpreadsheetApp.openById(STATS_SHEET_ID);
  let   sheet = ss.getSheetByName(STATS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(STATS_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Tanggal', 'Halaman', 'Nama', 'Email', 'Role']);
    sheet.getRange(1, 1, 1, 6)
         .setFontWeight('bold')
         .setBackground('#002d5d')
         .setFontColor('white');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Catat Kunjungan — dipanggil dari Index.html ─────────────
//    Terbuka untuk semua pengguna (user maupun admin)
function catatKunjungan(halaman, nama, email, role) {
  try {
    const sheet = getOrCreateSheet_();
    const now   = new Date();
    const tz    = Session.getScriptTimeZone();
    const tgl   = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    sheet.appendRow([
      now,
      tgl,
      halaman || 'Beranda',
      nama    || 'Anonim',
      email   || '',
      role    || 'User'
    ]);
    return { success: true };
  } catch (e) {
    console.error('catatKunjungan error:', e);
    return { success: false, msg: e.toString() };
  }
}

// ── Helper: konversi nilai kolom tanggal → string yyyy-MM-dd ──
// Google Sheets sering auto-convert string tanggal menjadi Date object
// sehingga toString().substring(0,10) menghasilkan "Tue May 12" bukan "2026-05-12"
function toTglStr_(val, tz) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, tz, 'yyyy-MM-dd');
  return String(val).substring(0, 10);
}

// ── Ambil Statistik Ringkasan — terbuka untuk semua pengguna ─
function getStatistik(startDate, endDate, mode) {
  try {
    const sheet = getOrCreateSheet_();
    const data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return emptyStats_();

    const rows  = data.slice(1);
    const tz    = Session.getScriptTimeZone();
    const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    // Untuk mode tahunan: ambil semua data dari awal agar semua tahun muncul
    const start = (mode === 'tahunan') ? '2000-01-01' : (startDate || '2000-01-01');
    const end   = endDate || today;

    // Filter rentang tanggal — pakai r[0] (Timestamp) konsisten agar tidak timezone shift
    const filtered = rows.filter(function(r) {
      const ts = r[0];
      if (!ts) return false;
      const d   = (ts instanceof Date) ? ts : new Date(ts);
      const tgl = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      return tgl >= start && tgl <= end;
    });

    // Kunjungan hari ini — gunakan r[0] (Timestamp) bukan r[1] (Tanggal)
    // karena Sheets menyimpan tanggal sebagai midnight UTC yang bisa bergeser zona waktu
    const hariIni = rows.filter(function(r) {
      const ts = r[0];
      if (!ts) return false;
      const d = (ts instanceof Date) ? ts : new Date(ts);
      return Utilities.formatDate(d, tz, 'yyyy-MM-dd') === today;
    }).length;

    // Grafik: per jam (harian) | per tahun (tahunan) | per tanggal (lainnya)
    const isTahunan  = (mode === 'tahunan');
    const isHarian   = (mode === 'harian');
    const perTanggal = {};
    filtered.forEach(function(r) {
      const ts = r[0];
      if (!ts) return;
      const d  = (ts instanceof Date) ? ts : new Date(ts);
      let key;
      if (isHarian)   key = Utilities.formatDate(d, tz, 'HH');         // 00-23
      else if (isTahunan) key = Utilities.formatDate(d, tz, 'yyyy');   // 2025, 2026...
      else            key = Utilities.formatDate(d, tz, 'yyyy-MM-dd'); // per hari
      if (key) perTanggal[key] = (perTanggal[key] || 0) + 1;
    });

    // Halaman terpopuler
    const perHalaman = {};
    filtered.forEach(function(r) {
      const h = r[2] || 'Beranda';
      perHalaman[h] = (perHalaman[h] || 0) + 1;
    });
    const halamanTop = Object.entries(perHalaman)
      .sort(function(a, b) { return b[1] - a[1]; })
      .slice(0, 5)
      .map(function(e) { return { nama: e[0], jml: e[1] }; });

    // 10 kunjungan terakhir (dari SEMUA data, bukan hanya filtered)
    const terakhir = rows.slice(-10).reverse().map(function(r) {
      const ts = r[0];
      let waktuStr = '-';
      try {
        if (ts instanceof Date) waktuStr = Utilities.formatDate(ts, tz, 'dd MMM, HH.mm');
        else if (ts)            waktuStr = Utilities.formatDate(new Date(ts), tz, 'dd MMM, HH.mm');
      } catch (ex) {
        waktuStr = ts ? ts.toString() : '-';
      }
      return {
        waktu  : waktuStr,
        halaman: r[2] || '-',
        nama   : r[3] || 'Anonim',
        email  : r[4] || '',
        role   : r[5] || 'User'
      };
    });

    // Kegiatan unik
    const kegiatanSet = {};
    filtered.forEach(function(r) {
      const k = (r[2] || '').split(' > ')[0];
      if (k) kegiatanSet[k] = true;
    });

    return {
      hariIni      : hariIni,
      totalPeriode : filtered.length,
      totalKegiatan: Object.keys(kegiatanSet).length,
      grafik       : perTanggal,
      halamanTop   : halamanTop,
      terakhir     : terakhir
    };
  } catch (e) {
    console.error('getStatistik error:', e);
    return emptyStats_();
  }
}

// ── Helper: objek statistik kosong ──────────────────────────
function emptyStats_() {
  return {
    hariIni      : 0,
    totalPeriode : 0,
    totalKegiatan: 0,
    grafik       : {},
    halamanTop   : [],
    terakhir     : []
  };
}


// ══════════════════════════════════════════════════════════════
//  BAGIAN 3 — FUNGSI TEST (Jalankan manual dari editor GAS)
// ══════════════════════════════════════════════════════════════

function testCatatKunjungan() {
  const hasil = catatKunjungan('Test Halaman', 'Ahmad Musyafa', 'syafakkpsmd@gmail.com', 'Admin');
  Logger.log('catatKunjungan: ' + JSON.stringify(hasil));
}

function testGetStatistik() {
  const stats = getStatistik('', '');
  Logger.log('getStatistik → hariIni=' + stats.hariIni + ', total=' + stats.totalPeriode);
  Logger.log('halamanTop: ' + JSON.stringify(stats.halamanTop));
}
