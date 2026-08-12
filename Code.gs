// ============================================================
// Code.gs - Google Apps Script Backend
// Sistem Absensi SMKN Rakit Kulim - Barcode + GPS
// ============================================================

// ==================== KONFIGURASI ==========================
const CONFIG = {
  SHEET_SISWA:   "DataSiswa",
  SHEET_ABSEN:   "Absensi",
  SHEET_SETTING: "Pengaturan",
  SHEET_ADMIN:   "Admin",
  SHEET_LIBUR:   "HariLibur",
  SHEET_JADWAL:  "JadwalHari",
  RADIUS_METER:  100,
};

// ==================== MAIN HANDLER ==========================
function doGet(e) {
  const action = e.parameter.action;
  let result;
  try {
    switch (action) {
      case "getSetting":    result = getSetting(); break;
      case "getSiswa":      result = getSiswaList(); break;
      case "getAbsensi":    result = getAbsensi(e.parameter); break;
      case "getRekapBulan": result = getRekapBulan(e.parameter); break;
      case "getLibur":      result = getLibur(); break;
      case "getJadwal":     result = getJadwal(); break;
      case "getRekapHarian": result = getRekapHarian(e.parameter); break;
      case "getRiwayatSiswa": result = getRiwayatSiswa(e.parameter); break;
      case "cekStatusHariIni": result = cekStatusAbsenHariIni(e.parameter); break;
      default: result = { status: "error", message: "Action tidak dikenal" };
    }
  } catch (err) {
    result = { status: "error", message: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body, result;
  try {
    body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case "absen":         result = absenSiswa(body); break;
      case "absenPulang":   result = absenPulangSiswa(body); break;
      case "absenManual":   result = absenManual(body); break;
      case "loginAdmin":    result = loginAdmin(body); break;
      case "daftarSiswa":   result = daftarSiswa(body); break;
      case "updateSiswa":   result = updateSiswa(body); break;
      case "hapusSiswa":    result = hapusSiswa(body); break;
      case "cetakUlangBarcode": result = cetakUlangBarcodeSiswa(body); break;
      case "saveSetting":   result = saveSetting(body); break;
      case "hapusAbsensi":  result = hapusAbsensi(body); break;
      case "tambahLibur":   result = tambahLibur(body); break;
      case "hapusLibur":    result = hapusLibur(body); break;
      case "saveJadwal":    result = saveJadwal(body); break;
      default: result = { status: "error", message: "Action tidak dikenal" };
    }
  } catch (err) {
    result = { status: "error", message: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// Paksa titik lokasi (TILOK) sekolah ke koordinat permanen berikut,
// menimpa nilai LatSekolah/LngSekolah yang mungkin sudah ada di sheet
// Pengaturan. Jalankan SEKALI secara manual dari editor Apps Script
// (pilih fungsi ini di dropdown, lalu klik Run) setelah deploy ulang,
// supaya sheet yang sudah lama jalan ikut ter-update. Aman dijalankan
// berkali-kali (idempotent).
function setLokasiSekolahPermanen() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SETTING);
  if (!sh) return { status: "error", message: "Sheet Pengaturan tidak ditemukan. Jalankan initSheets() dulu." };
  return saveSetting({ data: {
    LatSekolah: "-0.626301381459804",
    LngSekolah: "102.23851663181422"
  }});
}

// ==================== INISIALISASI SHEET ====================
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Sheet DataSiswa
  let sh = ss.getSheetByName(CONFIG.SHEET_SISWA);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_SISWA);
    sh.appendRow(["NIS","Nama","Kelas","Barcode","Tanggal_Daftar"]);
    sh.getRange(1,1,1,5).setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");
  } else {
    migrasiKolomBarcode(sh);
  }

  // Sheet Absensi
  sh = ss.getSheetByName(CONFIG.SHEET_ABSEN);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_ABSEN);
    sh.appendRow(["ID","NIS","Nama","Kelas","Tanggal","Waktu","Status","Catatan",
                  "Waktu_Pulang","Lat_Pulang","Lng_Pulang","Jarak_Pulang","Status_Pulang"]);
    sh.getRange(1,1,1,13).setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");
  } else {
    migrasiKolomPulang(sh);
    hapusKolomGPSLama(sh);
  }

  // Sheet Pengaturan
  sh = ss.getSheetByName(CONFIG.SHEET_SETTING);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_SETTING);
    sh.appendRow(["Key","Value"]);
    sh.appendRow(["NamaSekolah","SMKN Rakit Kulim"]);
    sh.appendRow(["Alamat","Jl. Pendidikan No. 1"]);
    sh.appendRow(["LatSekolah","-0.626301381459804"]);
    sh.appendRow(["LngSekolah","102.23851663181422"]);
    sh.appendRow(["RadiusMeter","100"]);
    sh.getRange(sh.getLastRow()+1,2,3,1).setNumberFormat("@"); // paksa teks utk 3 baris Jam* berikut, cegah auto-convert ke Date
    sh.appendRow(["JamMasuk","07:00"]);
    sh.appendRow(["JamPulang","15:00"]);
    sh.appendRow(["JamTelat","07:30"]);
    sh.appendRow(["AkurasiMaxMeter","50"]);
    sh.getRange(1,1,1,2).setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");
  } else {
    migrasiSettingBaru(sh);
  }

  // Sheet Admin
  sh = ss.getSheetByName(CONFIG.SHEET_ADMIN);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_ADMIN);
    sh.appendRow(["Username","Password","Nama","Role"]);
    sh.appendRow(["admin","admin123","Administrator","superadmin"]);
    sh.getRange(1,1,1,4).setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");
  }

  // Sheet HariLibur
  sh = ss.getSheetByName(CONFIG.SHEET_LIBUR);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_LIBUR);
    sh.appendRow(["ID","TglMulai","TglAkhir","Keterangan","DibuatOleh"]);
    sh.getRange(1,1,1,5).setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");
  }

  // Sheet JadwalHari — jadwal khusus per hari
  sh = ss.getSheetByName(CONFIG.SHEET_JADWAL);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_JADWAL);
    sh.appendRow(["Hari","NamaHari","JamMasuk","JamTelat","JamPulang","Aktif"]);
    // Default: Senin s/d Sabtu, isi dari setting umum kecuali Senin & Jumat
    const defaults = [
      ["0","Minggu","07:00","07:30","15:00","false"],
      ["1","Senin","07:15","07:45","15:00","true"],
      ["2","Selasa","07:00","07:30","15:00","true"],
      ["3","Rabu","07:00","07:30","15:00","true"],
      ["4","Kamis","07:00","07:30","15:00","true"],
      ["5","Jumat","07:00","07:30","11:30","true"],
      ["6","Sabtu","07:00","07:30","13:00","true"],
    ];
    sh.getRange("C:E").setNumberFormat("@"); // kolom JamMasuk/JamTelat/JamPulang dipaksa teks agar tidak auto-convert jadi Date
    defaults.forEach(row => sh.appendRow(row));
    sh.getRange(1,1,1,6).setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");
  }

  return { status: "ok", message: "Semua sheet berhasil diinisialisasi" };
}

// Ganti kolom lama "Foto_Descriptor" (kolom D) jadi "Barcode" pada sheet
// DataSiswa lama yang masih pakai sistem wajah, supaya sheet lama otomatis
// ikut ke sistem barcode tanpa perlu diedit manual. Aman dijalankan berkali-kali.
function migrasiKolomBarcode(sh) {
  sh = sh || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SISWA);
  if (!sh) return;
  const header = sh.getRange(1,1,1, Math.max(4, sh.getLastColumn())).getValues()[0];
  if (header[3] === "Foto_Descriptor" || header[3] === "") {
    sh.getRange(1,4).setValue("Barcode");
  }
}

// Barcode siswa = NIS/NISN itu sendiri. Disamakan persis dengan kartu fisik
// "Kartu Absen" yang sudah dicetak sebelumnya (barcode di kartu = angka NISN),
// supaya kartu yang sudah ada tetap valid dan tidak perlu dicetak ulang.
function buatBarcodeSiswa(nis) {
  return nis.toString();
}

// Tambahkan kolom Waktu_Pulang dst ke sheet Absensi lama yang belum punya.
// Aman dijalankan berkali-kali (idempotent). Jalankan manual sekali dari editor
// Apps Script jika sheet Absensi sudah ada sebelum update absen-pulang ini.
function migrasiKolomPulang(sh) {
  sh = sh || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSEN);
  if (!sh) return { status: "error", message: "Sheet Absensi tidak ditemukan" };
  const header = sh.getRange(1,1,1, Math.max(11, sh.getLastColumn())).getValues()[0];
  const kolomBaru = ["Waktu_Pulang","Lat_Pulang","Lng_Pulang","Jarak_Pulang","Status_Pulang"];
  if (header.indexOf("Waktu_Pulang") === -1) {
    const startCol = 12; // kolom L
    sh.getRange(1, startCol, 1, kolomBaru.length).setValues([kolomBaru])
      .setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");
  }
  return { status: "ok", message: "Migrasi kolom pulang selesai" };
}

// Tambahkan key setting baru (AkurasiMaxMeter) ke sheet Pengaturan lama
// yang belum punya. Aman dijalankan berkali-kali. (ThresholdWajah sudah
// tidak dipakai sejak absen pindah ke barcode, dibiarkan saja di sheet
// lama kalau masih ada, tidak mengganggu.)
function migrasiSettingBaru(sh) {
  sh = sh || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SETTING);
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  const keys = data.map(r => r[0]);
  if (keys.indexOf("AkurasiMaxMeter") === -1) sh.appendRow(["AkurasiMaxMeter","50"]);
}

// Menghapus PERMANEN kolom-kolom detail GPS yang sudah tidak dipakai lagi:
// Latitude, Longitude, Jarak_Meter (absen masuk), Akurasi_Masuk, Akurasi_Pulang,
// LuarTilok_Masuk, LuarTilok_Pulang, AkurasiRendah_Masuk, AkurasiRendah_Pulang.
// Dicari berdasarkan NAMA header (bukan nomor kolom tetap) supaya tetap aman
// dijalankan pada sheet lama yang urutan kolomnya mungkin berbeda-beda, lalu
// dihapus dari KANAN ke KIRI supaya index kolom yang belum diproses tidak
// ikut bergeser saat penghapusan berlangsung. Aman dijalankan berkali-kali
// (idempotent) — kalau kolomnya sudah tidak ada, tidak melakukan apa-apa.
function hapusKolomGPSLama(sh) {
  sh = sh || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSEN);
  if (!sh) return;
  // Paksa kolom Tanggal (E) selalu teks polos supaya Sheets tidak lagi
  // otomatis mengonversinya jadi Date di kemudian hari (root cause data
  // absensi "hilang" dari dashboard karena perbandingan tanggal gagal).
  sh.getRange("E:E").setNumberFormat("@");
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return;
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const kolomHapus = ["Latitude","Longitude","Jarak_Meter","Akurasi_Masuk","Akurasi_Pulang",
                       "LuarTilok_Masuk","LuarTilok_Pulang","AkurasiRendah_Masuk","AkurasiRendah_Pulang"];
  for (let c = header.length - 1; c >= 0; c--) {
    if (kolomHapus.indexOf(header[c]) !== -1) {
      sh.deleteColumn(c + 1);
    }
  }
}

// ==================== JADWAL HARI ===========================
function getJadwal() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_JADWAL);
  if (!sh) {
    // Kembalikan default hardcode jika sheet belum ada
    return { status: "ok", data: [
      {hari:"0",namaHari:"Minggu",jamMasuk:"07:00",jamTelat:"07:30",jamPulang:"15:00",aktif:false},
      {hari:"1",namaHari:"Senin",jamMasuk:"07:15",jamTelat:"07:45",jamPulang:"15:00",aktif:true},
      {hari:"2",namaHari:"Selasa",jamMasuk:"07:00",jamTelat:"07:30",jamPulang:"15:00",aktif:true},
      {hari:"3",namaHari:"Rabu",jamMasuk:"07:00",jamTelat:"07:30",jamPulang:"15:00",aktif:true},
      {hari:"4",namaHari:"Kamis",jamMasuk:"07:00",jamTelat:"07:30",jamPulang:"15:00",aktif:true},
      {hari:"5",namaHari:"Jumat",jamMasuk:"07:00",jamTelat:"07:30",jamPulang:"11:30",aktif:true},
      {hari:"6",namaHari:"Sabtu",jamMasuk:"07:00",jamTelat:"07:30",jamPulang:"13:00",aktif:true},
    ]};
  }
  const rows = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0].toString()) continue;
    list.push({
      hari:      rows[i][0].toString(),
      namaHari:  rows[i][1].toString(),
      jamMasuk:  jamToString(rows[i][2]),
      jamTelat:  jamToString(rows[i][3]),
      jamPulang: jamToString(rows[i][4]),
      aktif:     rows[i][5].toString() === "true"
    });
  }
  return { status: "ok", data: list };
}

function saveJadwal(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_JADWAL);
  if (!sh) return { status: "error", message: "Sheet JadwalHari tidak ditemukan. Jalankan initSheets() dulu." };
  const rows = sh.getDataRange().getValues();
  // body.data = array of {hari, jamMasuk, jamTelat, jamPulang, aktif}
  for (const item of body.data) {
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === item.hari.toString()) {
        sh.getRange(i+1, 3, 1, 3).setNumberFormat("@"); // paksa teks dulu sebelum ditulis
        sh.getRange(i+1, 3).setValue(item.jamMasuk);
        sh.getRange(i+1, 4).setValue(item.jamTelat);
        sh.getRange(i+1, 5).setValue(item.jamPulang);
        sh.getRange(i+1, 6).setValue(item.aktif ? "true" : "false");
        break;
      }
    }
  }
  return { status: "ok", message: "Jadwal berhasil disimpan" };
}

// ==================== HARI LIBUR ============================
function getLibur() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LIBUR);
  if (!sh) return { status: "ok", data: [] };
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    list.push({
      id:          data[i][0].toString(),
      tglMulai:    data[i][1].toString(),
      tglAkhir:    data[i][2].toString(),
      keterangan:  data[i][3].toString()
    });
  }
  return { status: "ok", data: list };
}

function tambahLibur(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LIBUR);
  if (!sh) return { status: "error", message: "Sheet HariLibur tidak ditemukan. Jalankan initSheets() dulu." };
  const id = "LBR" + new Date().getTime();
  sh.appendRow([id, body.tglMulai, body.tglAkhir, body.keterangan, "admin"]);
  return { status: "ok", message: "Hari libur ditambahkan" };
}

function hapusLibur(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LIBUR);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === body.id.toString()) {
      sh.deleteRow(i + 1);
      return { status: "ok", message: "Hari libur dihapus" };
    }
  }
  return { status: "error", message: "Data tidak ditemukan" };
}

// ==================== PENGATURAN ============================
function getSetting() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SETTING);
  if (!sh) return { status: "error", message: "Sheet Pengaturan tidak ditemukan" };
  const data = sh.getDataRange().getValues();
  const setting = {};
  for (let i = 1; i < data.length; i++) {
    let val = data[i][1];
    if (val instanceof Date) val = jamToString(val); // jaga-jaga kalau JamMasuk/JamTelat/JamPulang ikut terkonversi Sheets
    setting[data[i][0]] = val;
  }
  return { status: "ok", data: setting };
}

function saveSetting(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SETTING);
  const data = sh.getDataRange().getValues();
  const updates = body.data;
  for (const key in updates) {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        if (key.indexOf("Jam") === 0) sh.getRange(i+1, 2).setNumberFormat("@"); // cegah auto-convert utk key JamMasuk/JamTelat/JamPulang
        sh.getRange(i+1, 2).setValue(updates[key]);
        found = true; break;
      }
    }
    if (!found) sh.appendRow([key, updates[key]]);
  }
  return { status: "ok", message: "Pengaturan disimpan" };
}

// ==================== DATA SISWA ============================
function getSiswaList() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SISWA);
  if (!sh) return { status: "error", message: "Sheet DataSiswa tidak ditemukan" };
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      list.push({
        nis:          data[i][0].toString(),
        nama:         data[i][1],
        kelas:        data[i][2],
        barcode:      data[i][3] || "",
        tanggalDaftar: data[i][4]
      });
    }
  }
  return { status: "ok", data: list };
}

function daftarSiswa(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SISWA);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === body.nis.toString()) {
      return { status: "error", message: "NIS sudah terdaftar" };
    }
  }
  const barcode = buatBarcodeSiswa(body.nis);
  sh.appendRow([body.nis, body.nama, body.kelas, barcode, getWIBDate()]);
  return { status: "ok", message: "Siswa berhasil didaftarkan", barcode: barcode };
}

function updateSiswa(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SISWA);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === body.nis.toString()) {
      sh.getRange(i+1, 2).setValue(body.nama);
      sh.getRange(i+1, 3).setValue(body.kelas);
      return { status: "ok", message: "Data siswa diperbarui" };
    }
  }
  return { status: "error", message: "Siswa tidak ditemukan" };
}

// Buat ulang barcode siswa (misal barcode hilang/rusak)
function cetakUlangBarcodeSiswa(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SISWA);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === body.nis.toString()) {
      const barcode = buatBarcodeSiswa(body.nis);
      sh.getRange(i+1, 4).setValue(barcode);
      return { status: "ok", message: "Barcode baru dibuat", barcode: barcode };
    }
  }
  return { status: "error", message: "Siswa tidak ditemukan" };
}

function hapusSiswa(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SISWA);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === body.nis.toString()) {
      sh.deleteRow(i+1);
      return { status: "ok", message: "Siswa dihapus" };
    }
  }
  return { status: "error", message: "Siswa tidak ditemukan" };
}

// Cek status absen hari ini untuk satu siswa (sudah absen masuk? sudah pulang?).
// Dipakai frontend supaya scan barcode langsung bisa otomatis menentukan
// jenis absen (Masuk atau Pulang) tanpa siswa harus pilih tombol manual dulu.
function cekStatusAbsenHariIni(params) {
  const nis = params.nis;
  if (!nis) return { status: "error", message: "Parameter nis wajib diisi" };
  const tanggal = getWIBDate();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSEN);
  if (!sh) return { status: "ok", data: { sudahAbsen: false, statusMasuk: null, sudahPulang: false } };
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (data[i][1].toString() === nis.toString() && tanggalToString(data[i][4]) === tanggal) {
      const statusMasuk = data[i][6];
      const waktuPulang = data[i][8];
      return {
        status: "ok",
        data: {
          sudahAbsen: true,
          statusMasuk: statusMasuk,
          sudahPulang: !!waktuPulang,
          waktuMasuk: data[i][5] || ""
        }
      };
    }
  }
  return { status: "ok", data: { sudahAbsen: false, statusMasuk: null, sudahPulang: false } };
}

// ==================== ABSENSI ===============================
function absenSiswa(body) {
  const now  = new Date();
  const wib  = new Date(now.getTime() + 7 * 60 * 60 * 1000);
  const pad  = n => String(n).padStart(2, "0");
  const tanggal = wib.getUTCFullYear() + "-" + pad(wib.getUTCMonth()+1) + "-" + pad(wib.getUTCDate());
  const hariIdx = wib.getUTCDay(); // 0=Minggu, 1=Senin, …

  // ── Cek hari libur ──────────────────────────────────────
  const shLibur = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LIBUR);
  if (shLibur) {
    const rows = shLibur.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      if (tanggal >= rows[i][1].toString() && tanggal <= rows[i][2].toString()) {
        return { status: "error", message: "Hari ini adalah hari libur: " + rows[i][3] };
      }
    }
  }

  // ── Ambil jadwal hari ini ────────────────────────────────
  const jadwalRes = getJadwal();
  const jadwalList = jadwalRes.data || [];
  const jadwalHariIni = jadwalList.find(j => j.hari === String(hariIdx));

  // Jika hari tidak aktif (mis. Minggu) tolak absen
  if (jadwalHariIni && jadwalHariIni.aktif === false) {
    return { status: "error", message: "Hari " + (jadwalHariIni.namaHari || "ini") + " bukan hari sekolah." };
  }

  // Gunakan jam dari jadwal hari, fallback ke setting umum
  const setting = getSetting().data;
  const jamTelat  = jadwalHariIni ? jadwalHariIni.jamTelat  : (setting.JamTelat  || "07:30");
  const jamPulang = jadwalHariIni ? jadwalHariIni.jamPulang : (setting.JamPulang || "15:00");

  // ── Verifikasi lokasi (GPS) DIHAPUS PERMANEN ─────────────
  // Absen tidak lagi memerlukan/menghitung koordinat GPS sama sekali.
  const luarTilok = false;

  const waktu = pad(wib.getUTCHours()) + ":" + pad(wib.getUTCMinutes()) + ":" + pad(wib.getUTCSeconds());

  // ── Cek sudah absen ──────────────────────────────────────
  const shAbsen = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSEN);
  hapusKolomGPSLama(shAbsen);
  const dataAbsen = shAbsen.getDataRange().getValues();
  for (let i = 1; i < dataAbsen.length; i++) {
    if (dataAbsen[i][1].toString() === body.nis.toString() &&
        tanggalToString(dataAbsen[i][4]) === tanggal) {
      return { status: "error", message: "Anda sudah absen hari ini pukul " + dataAbsen[i][5] };
    }
  }

  // ── Tentukan status ──────────────────────────────────────
  const menitAbsen = wib.getUTCHours() * 60 + wib.getUTCMinutes();
  const menitBatas = jamToMenit(jamTelat);
  const status = menitAbsen <= menitBatas ? "Hadir" : "Terlambat";

  const catatan = luarTilok ? "ABSEN DI LUAR TILOK" : "";
  const id = "ABS" + now.getTime();
  shAbsen.appendRow([id, body.nis, body.nama, body.kelas, tanggal, waktu, status, catatan]);

  return {
    status: "ok",
    message: luarTilok
      ? `Absensi berhasil! Status: ${status}. ⚠️ ABSEN DI LUAR TILOK`
      : `Absensi berhasil! Status: ${status}`,
    data: {
      tanggal, waktu, status,
      luarTilok,
      jamMasuk:  jadwalHariIni ? jadwalHariIni.jamMasuk  : (setting.JamMasuk  || "07:00"),
      jamPulang: jamPulang,
      hariNama:  jadwalHariIni ? jadwalHariIni.namaHari : ""
    }
  };
}

// ── ABSEN PULANG ─────────────────────────────────────────────
// Siswa harus sudah absen masuk (Hadir/Terlambat) hari ini sebelum bisa
// absen pulang. Verifikasi GPS sama seperti absen masuk. Hasil ditulis
// ke kolom Waktu_Pulang..Status_Pulang pada baris absen masuk yang sama.
function absenPulangSiswa(body) {
  const now  = new Date();
  const wib  = new Date(now.getTime() + 7 * 60 * 60 * 1000);
  const pad  = n => String(n).padStart(2, "0");
  const tanggal = wib.getUTCFullYear() + "-" + pad(wib.getUTCMonth()+1) + "-" + pad(wib.getUTCDate());
  const hariIdx = wib.getUTCDay();

  const shAbsen = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSEN);
  migrasiKolomPulang(shAbsen);
  hapusKolomGPSLama(shAbsen);
  const data = shAbsen.getDataRange().getValues();

  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toString() === body.nis.toString() && tanggalToString(data[i][4]) === tanggal) {
      rowIdx = i;
      break;
    }
  }
  if (rowIdx === -1) {
    return { status: "error", message: "Anda belum absen masuk hari ini. Silakan absen masuk terlebih dahulu." };
  }

  const statusMasuk = data[rowIdx][6];
  if (statusMasuk !== "Hadir" && statusMasuk !== "Terlambat") {
    return { status: "error", message: `Status kehadiran Anda hari ini adalah "${statusMasuk}", absen pulang tidak berlaku.` };
  }

  const waktuPulangSudah = data[rowIdx][8];
  if (waktuPulangSudah) {
    return { status: "error", message: "Anda sudah absen pulang hari ini pukul " + jamToString(waktuPulangSudah) };
  }

  // ── Verifikasi lokasi (GPS) DIHAPUS PERMANEN ─────────────
  const setting = getSetting().data;
  const luarTilokPulang = false;

  const waktu = pad(wib.getUTCHours()) + ":" + pad(wib.getUTCMinutes()) + ":" + pad(wib.getUTCSeconds());

  // ── Tentukan jam pulang & status (Pulang / Pulang Cepat) ─
  const jadwalRes = getJadwal();
  const jadwalHariIni = (jadwalRes.data || []).find(j => j.hari === String(hariIdx));
  const jamPulang = jadwalHariIni ? jadwalHariIni.jamPulang : (setting.JamPulang || "15:00");
  const menitPulang     = wib.getUTCHours() * 60 + wib.getUTCMinutes();
  const menitBatasPulang = jamToMenit(jamPulang);
  const statusPulang = menitPulang < menitBatasPulang ? "Pulang Cepat" : "Pulang";

  shAbsen.getRange(rowIdx + 1, 9, 1, 5).setNumberFormat("@"); // cegah auto-convert waktu
  shAbsen.getRange(rowIdx + 1, 9, 1, 5).setValues([[waktu, "", "", "", statusPulang]]);

  return {
    status: "ok",
    message: luarTilokPulang
      ? `Absen pulang berhasil! Status: ${statusPulang}. ⚠️ ABSEN DI LUAR TILOK`
      : `Absen pulang berhasil! Status: ${statusPulang}`,
    data: {
      tanggal, waktu, status: statusPulang,
      luarTilok: luarTilokPulang, jamPulang
    }
  };
}

function getAbsensi(params) {
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSEN);
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const row = {
      id: data[i][0], nis: data[i][1].toString(), nama: data[i][2],
      kelas: data[i][3], tanggal: tanggalToString(data[i][4]), waktu: data[i][5],
      status: data[i][6], catatan: data[i][7],
      waktuPulang:  data[i][8] ? jamToString(data[i][8]) : "",
      latPulang:    data[i][9]  || "",
      lngPulang:    data[i][10] || "",
      jarakPulang:  data[i][11] || "",
      statusPulang: data[i][12] || ""
    };
    if (params.tanggal && row.tanggal !== params.tanggal) continue;
    if (params.kelas   && row.kelas   !== params.kelas)   continue;
    if (params.nis     && row.nis     !== params.nis)      continue;
    list.push(row);
  }
  return { status: "ok", data: list };
}

// ── REKAP HARIAN LENGKAP (untuk Download Absen Harian) ───────
// Menggabungkan seluruh siswa terdaftar dengan data absensi pada satu
// tanggal, sehingga siswa yang belum absen sama sekali tetap muncul
// dengan status "Alpha". Dipakai untuk unduh laporan absen harian.
function getRekapHarian(params) {
  const tanggal = params.tanggal;
  if (!tanggal) return { status: "error", message: "Parameter tanggal wajib diisi" };
  const siswaList = getSiswaList().data || [];
  const absenList = getAbsensi({ tanggal: tanggal }).data || [];
  const map = {};
  absenList.forEach(a => { map[a.nis] = a; });

  const hasil = siswaList.map(s => {
    const a = map[s.nis];
    if (a) {
      return {
        nis: s.nis, nama: s.nama, kelas: s.kelas, tanggal: tanggal,
        waktuMasuk: a.waktu || "-", statusMasuk: a.status || "-", jarakMasuk: a.jarak,
        waktuPulang: a.waktuPulang || "-", statusPulang: a.statusPulang || "-", jarakPulang: a.jarakPulang || "-",
        catatan: a.catatan || ""
      };
    }
    return {
      nis: s.nis, nama: s.nama, kelas: s.kelas, tanggal: tanggal,
      waktuMasuk: "-", statusMasuk: "Alpha", jarakMasuk: "-",
      waktuPulang: "-", statusPulang: "-", jarakPulang: "-",
      catatan: ""
    };
  });
  hasil.sort((x,y) => x.kelas.localeCompare(y.kelas) || x.nama.localeCompare(y.nama));
  return { status: "ok", data: hasil };
}

// Rekap bulanan per siswa: Hadir/Terlambat/Izin/Sakit/Alpha.
// Sebelumnya hanya menghitung Hadir & Terlambat dari baris yang ADA di
// sheet, sehingga Izin/Sakit selalu 0 dan Alpha (siswa yang sama sekali
// tidak absen) tidak pernah terhitung — sekarang dihitung berdasarkan
// hari sekolah aktif (sesuai JadwalHari & HariLibur) sampai hari ini.
function getRekapBulan(params) {
  const bulan = params.bulan;
  if (!bulan) return { status: "error", message: "Parameter bulan wajib diisi" };
  const [thnStr, blnStr] = bulan.split("-");
  const thn = parseInt(thnStr, 10), bln = parseInt(blnStr, 10);
  const jumlahHari = new Date(thn, bln, 0).getDate();
  const todayStr = getWIBDate();
  const p = n => String(n).padStart(2, "0");

  const liburList  = getLibur().data  || [];
  const jadwalList = getJadwal().data || [];
  const jadwalMap  = {};
  jadwalList.forEach(j => jadwalMap[j.hari] = j);

  // Kumpulkan tanggal-tanggal sekolah aktif dalam bulan ini yang sudah lewat
  const tanggalSekolah = [];
  for (let d = 1; d <= jumlahHari; d++) {
    const tglStr = thn + "-" + p(bln) + "-" + p(d);
    if (tglStr > todayStr) break; // jangan hitung Alpha utk hari yg belum terjadi
    const hariIdx = new Date(thn, bln - 1, d).getDay();
    const jadwalHari = jadwalMap[String(hariIdx)];
    if (jadwalHari && jadwalHari.aktif === false) continue;
    const isLibur = liburList.some(l => tglStr >= l.tglMulai && tglStr <= l.tglAkhir);
    if (isLibur) continue;
    tanggalSekolah.push(tglStr);
  }

  const siswaList = getSiswaList().data || [];
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSEN);
  const data = sh.getDataRange().getValues();

  const absenMap = {}; // key: nis_tanggal -> status
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const tgl = data[i][4] ? tanggalToString(data[i][4]) : "";
    if (!tgl.startsWith(bulan)) continue;
    absenMap[data[i][1].toString() + "_" + tgl] = data[i][6];
  }

  const rekap = {};
  siswaList.forEach(s => {
    rekap[s.nis] = { nis: s.nis, nama: s.nama, kelas: s.kelas, hadir: 0, terlambat: 0, izin: 0, sakit: 0, alpha: 0 };
  });

  tanggalSekolah.forEach(tglStr => {
    siswaList.forEach(s => {
      const r = rekap[s.nis];
      if (!r) return;
      const status = absenMap[s.nis + "_" + tglStr];
      if (status === "Hadir") r.hadir++;
      else if (status === "Terlambat") r.terlambat++;
      else if (status === "Izin") r.izin++;
      else if (status === "Sakit") r.sakit++;
      else r.alpha++;
    });
  });

  return { status: "ok", data: Object.values(rekap) };
}

// Riwayat absensi harian satu siswa dalam satu bulan — dipakai untuk
// Laporan Per Siswa (PDF/Excel). Sertakan hari sekolah yang belum diabsen
// (tampil "Alpha") sampai hari ini, mengikuti JadwalHari & HariLibur.
function getRiwayatSiswa(params) {
  const nis = params.nis, bulan = params.bulan;
  if (!nis || !bulan) return { status: "error", message: "Parameter nis dan bulan wajib diisi" };

  const siswaList = getSiswaList().data || [];
  const siswa = siswaList.find(s => s.nis === nis.toString());
  if (!siswa) return { status: "error", message: "Siswa tidak ditemukan" };

  const [thnStr, blnStr] = bulan.split("-");
  const thn = parseInt(thnStr, 10), bln = parseInt(blnStr, 10);
  const jumlahHari = new Date(thn, bln, 0).getDate();
  const todayStr = getWIBDate();
  const p = n => String(n).padStart(2, "0");

  const liburList  = getLibur().data  || [];
  const jadwalList = getJadwal().data || [];
  const jadwalMap  = {};
  jadwalList.forEach(j => jadwalMap[j.hari] = j);

  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSEN);
  const data = sh.getDataRange().getValues();
  const absenMap = {}; // key: tanggal -> detail
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (data[i][1].toString() !== nis.toString()) continue;
    const tgl = data[i][4] ? tanggalToString(data[i][4]) : "";
    absenMap[tgl] = {
      waktu: data[i][5], status: data[i][6], catatan: data[i][7],
      waktuPulang: data[i][8] ? jamToString(data[i][8]) : "", statusPulang: data[i][12] || ""
    };
  }

  const hasil = [];
  for (let d = 1; d <= jumlahHari; d++) {
    const tglStr = thn + "-" + p(bln) + "-" + p(d);
    if (tglStr > todayStr) break;
    const hariIdx = new Date(thn, bln - 1, d).getDay();
    const jadwalHari = jadwalMap[String(hariIdx)];
    if (jadwalHari && jadwalHari.aktif === false) continue;
    const isLibur = liburList.some(l => tglStr >= l.tglMulai && tglStr <= l.tglAkhir);
    if (isLibur) continue;
    const a = absenMap[tglStr];
    hasil.push({
      tanggal: tglStr,
      hari: jadwalHari ? jadwalHari.namaHari : "",
      waktuMasuk: a ? (a.waktu || "-") : "-",
      statusMasuk: a ? a.status : "Alpha",
      waktuPulang: a ? (a.waktuPulang || "-") : "-",
      statusPulang: a ? (a.statusPulang || "-") : "-",
      catatan: a ? (a.catatan || "") : ""
    });
  }
  return { status: "ok", data: { siswa, riwayat: hasil } };
}

// Input keterangan Izin/Sakit/Alpha manual (dipakai modal admin & form Izin siswa).
// Catatan: sebelumnya action ini dipanggil dari frontend tapi belum ada
// implementasinya di backend, sehingga selalu gagal — sekarang sudah ditambahkan.
function absenManual(body) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSEN);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toString() === body.nis.toString() && tanggalToString(data[i][4]) === body.tanggal) {
      return { status: "error", message: `Data absensi tanggal ini sudah ada (status: ${data[i][6]})` };
    }
  }
  const id = "ABS" + new Date().getTime();
  sh.appendRow([id, body.nis, body.nama, body.kelas, body.tanggal, "-", body.status, "", "", "", body.catatan || ""]);
  return { status: "ok", message: `Keterangan "${body.status}" berhasil disimpan` };
}

function hapusAbsensi(body) {
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSEN);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === body.id.toString()) {
      sh.deleteRow(i+1);
      return { status: "ok", message: "Data absensi dihapus" };
    }
  }
  return { status: "error", message: "Data tidak ditemukan" };
}

// ==================== LOGIN ADMIN ===========================
function loginAdmin(body) {
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ADMIN);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.username && data[i][1] === body.password) {
      return { status: "ok", data: { nama: data[i][2], role: data[i][3] } };
    }
  }
  return { status: "error", message: "Username atau password salah" };
}

// ==================== UTILITIES =============================
function hitungJarak(lat1, lng1, lat2, lng2) {
  const R    = 6371000;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLng = (lng2 - lng1) * Math.PI / 180;
  const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
            Math.cos(lat1 * Math.PI/180) * Math.cos(lat2 * Math.PI/180) *
            Math.sin(dLng/2) * Math.sin(dLng/2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

function getWIBDate() {
  const wib = new Date(new Date().getTime() + 7*60*60*1000);
  const p   = n => String(n).padStart(2,"0");
  return wib.getUTCFullYear() + "-" + p(wib.getUTCMonth()+1) + "-" + p(wib.getUTCDate());
}

// ── Normalisasi nilai jam dari Sheet ─────────────────────────
// PENTING: Google Sheets otomatis mengubah string seperti "07:30" yang
// ditulis via appendRow()/setValue() menjadi tipe Waktu (Date object).
// Kalau tidak dinormalisasi, jamTelat.split(":") akan menghasilkan nilai
// acak/NaN sehingga status "Terlambat" bisa salah (kadang muncul padahal
// siswa datang tepat/lebih awal). Fungsi ini selalu mengembalikan "HH:mm".
// ── Normalisasi nilai tanggal dari Sheet ─────────────────────
// PENTING: sama seperti jam, Google Sheets kadang otomatis mengubah
// string "2026-07-17" yang ditulis via appendRow()/setValue() menjadi
// tipe Date. Kalau tidak dinormalisasi, semua perbandingan tanggal
// (mis. filter getAbsensi by tanggal) akan selalu gagal cocok sehingga
// data yang sudah absen tidak muncul di dashboard admin. Fungsi ini
// selalu mengembalikan "yyyy-MM-dd".
function tanggalToString(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone() || "Asia/Jakarta", "yyyy-MM-dd");
  }
  return val === null || val === undefined ? "" : val.toString().trim();
}

function jamToString(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone() || "Asia/Jakarta", "HH:mm");
  }
  return val === null || val === undefined ? "" : val.toString().trim();
}

// Ubah nilai jam (string "HH:mm" ATAU Date) menjadi total menit sejak 00:00
function jamToMenit(val) {
  if (val instanceof Date) {
    return val.getHours() * 60 + val.getMinutes();
  }
  const parts = val.toString().trim().split(":");
  const h = parseInt(parts[0], 10) || 0;
  const m = parseInt(parts[1], 10) || 0;
  return h * 60 + m;
}
