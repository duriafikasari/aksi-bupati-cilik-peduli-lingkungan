/**
 * PAPAN SKOR ONLINE - AKSI BUPATI CILIK PEDULI LINGKUNGAN
 * Menerima data murid dari gim dan menyimpannya di Google Sheets (satu baris per murid, diperbarui otomatis).
 * Data lengkap (termasuk No. Absen) hanya terlihat di Google Sheets. Gim hanya dapat membaca daftar ringkas lewat ?aksi=papan.
 * Langkah pemasangan: lihat PETUNJUK-GOOGLE-SHEETS.txt
 */
var NAMA_SHEET = 'PapanSkor';
var KEPALA = ['Kunci', 'Nama', 'Asal Sekolah', 'Kelas', 'No. Absen', 'Karakter',
  'Zona Sampah', 'Zona Energi', 'Zona Tata Wilayah', 'Zona Transportasi', 'Zona Sekolah',
  'Total', 'Zona Selesai', 'Waktu Selesai', 'Diperbarui'];

function teks_(v, maks) {
  var s = String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim().substring(0, maks);
  if (/^[=+\-@]/.test(s)) s = "'" + s;          // cegah rumus tersisip
  return s;
}
function angka_(v, maks) {
  var n = Number(v);
  if (!isFinite(n)) return '';
  return Math.max(0, Math.min(maks, Math.round(n)));
}
function jawab_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function lembar_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(NAMA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(NAMA_SHEET);
    sh.appendRow(KEPALA);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, KEPALA.length).setFontWeight('bold');
  }
  return sh;
}

// Daftar untuk papan skor bersama di gim: GET ...?aksi=papan
// Hanya berisi nama, asal sekolah, kelas, nilai per zona, total. TIDAK memuat No. Absen, kunci, atau waktu.
// Ubah TAMPILKAN_DAFTAR menjadi false untuk menutup daftar ini (Sheet tetap menerima data).
var TAMPILKAN_DAFTAR = true;
var MAKS_BARIS_DAFTAR = 300;

function doGet(e) {
  if (e && e.parameter && e.parameter.aksi === 'papan') {
    if (!TAMPILKAN_DAFTAR) return jawab_({ ok: false, error: 'daftar ditutup' });
    var cache = CacheService.getScriptCache();
    var tersimpan = cache.get('papan');
    if (tersimpan) return ContentService.createTextOutput(tersimpan).setMimeType(ContentService.MimeType.JSON);
    var sh = lembar_(), n = sh.getLastRow(), daftar = [];
    if (n > 1) {
      var v = sh.getRange(2, 1, n - 1, KEPALA.length).getValues();
      v.forEach(function (r) {
        daftar.push({ nama: String(r[1]).replace(/^'/, ''), sekolah: String(r[2]).replace(/^'/, ''), kelas: String(r[3]),
          zona: [r[6], r[7], r[8], r[9], r[10]].map(function (x) { return x === '' ? null : Number(x); }),
          total: Number(r[11]) || 0, selesai: String(r[12]) });
      });
      daftar.sort(function (a, b) { return b.total - a.total; });
      daftar = daftar.slice(0, MAKS_BARIS_DAFTAR);
    }
    var teksJson = JSON.stringify({ ok: true, daftar: daftar });
    try { cache.put('papan', teksJson, 20); } catch (x) { /* abaikan */ }
    return ContentService.createTextOutput(teksJson).setMimeType(ContentService.MimeType.JSON);
  }
  return jawab_({ ok: true, info: 'Papan Skor Aksi Bupati Cilik' });
}

function doPost(e) {
  var kunci = LockService.getScriptLock();
  try {
    kunci.waitLock(20000);
    var d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var nama = teks_(d.name, 60);
    if (!nama) return jawab_({ ok: false, error: 'nama kosong' });
    var sekolah = teks_(d.school, 60), kelas = teks_(d.cls, 10), absen = teks_(d.roll, 6);
    var id = (nama + '|' + sekolah + '|' + kelas + '|' + absen).toLowerCase();
    var sk = Array.isArray(d.scores) ? d.scores : [];
    var z = [0, 1, 2, 3, 4].map(function (i) { return sk[i] == null ? '' : angka_(sk[i], 100); });
    var total = z.reduce(function (a, b) { return a + (b === '' ? 0 : b); }, 0);
    var selesai = z.filter(function (x) { return x !== ''; }).length;
    var baris = [id, nama, sekolah, kelas, absen, teks_(d.char, 40)].concat(z)
      .concat([total, selesai + '/5', teks_(d.finishedAt || '', 40), new Date()]);
    var sh = lembar_();
    var n = sh.getLastRow();
    var ids = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues() : [];
    var pos = -1;
    for (var i = 0; i < ids.length; i++) { if (ids[i][0] === id) { pos = i + 2; break; } }
    if (pos < 0) { sh.appendRow(baris); }
    else { sh.getRange(pos, 1, 1, baris.length).setValues([baris]); }
    return jawab_({ ok: true });
  } catch (err) {
    return jawab_({ ok: false, error: String(err) });
  } finally {
    try { kunci.releaseLock(); } catch (x) { /* abaikan */ }
  }
}
