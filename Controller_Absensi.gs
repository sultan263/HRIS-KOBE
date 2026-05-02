/**
 * =========================================================
 * KONTROLER MODUL ABSENSI
 * Terhubung dengan Database: 15HSyULRuhrsq86TMogPfahY-KZsCuxM60k4TvQJu6Z4
 * =========================================================
 */

const DB_ABSENSI_ID = '15HSyULRuhrsq86TMogPfahY-KZsCuxM60k4TvQJu6Z4';

function getDashboardAbsensiData() {
  try {
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID);
    const sheetRaw = ss.getSheetByName('Raw_Kehadiran');
    if (!sheetRaw) return JSON.stringify({ status: 'error', message: 'Sheet Raw_Kehadiran tidak ditemukan!' });
    
    const data = sheetRaw.getDataRange().getValues();
    data.shift(); // Buang header
    
    // Ambil info hari ini
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = today.getMonth(); // Bulan dimulai dari 0
    const dd = today.getDate();
    
    let statHadir = 0, statCuti = 0, statIzinSakit = 0, statTerlambat = 0;
    let logHariIni = [];
    let logSemua = []; // Array penampung seluruh data (untuk Fallback)
    
    for (let i = 0; i < data.length; i++) {
      let row = data[i];
      let tglRow = row[0]; // Kolom A: Tanggal
      
      if (!tglRow || tglRow === "") continue; 
      
      // LOGIKA TANGGAL YANG LEBIH FLEKSIBEL
      let isToday = false;
      let rowDate = new Date(tglRow); // Paksa jadikan objek Date
      
      if (!isNaN(rowDate.getTime())) {
        // Bandingkan secara absolut (Tahun, Bulan, Tanggal) -> anti meleset
        if (rowDate.getFullYear() === yyyy && rowDate.getMonth() === mm && rowDate.getDate() === dd) {
          isToday = true;
        }
      } else {
        // Jika formatnya teks murni dan gagal di-parse, cek string secara kasar
        let tglStr = String(tglRow);
        let strToday1 = Utilities.formatDate(today, "Asia/Jakarta", "yyyy-MM-dd");
        let strToday2 = Utilities.formatDate(today, "Asia/Jakarta", "dd/MM/yyyy");
        let strToday3 = Utilities.formatDate(today, "Asia/Jakarta", "dd-MM-yyyy");
        if (tglStr.includes(strToday1) || tglStr.includes(strToday2) || tglStr.includes(strToday3)) {
          isToday = true;
        }
      }
      
      // Ambil sisa kolom
      let nrpp = row[1] || "-";
      let nama = row[2] || "-";
      let jamMasuk = row[3] || "";
      let jamKeluar = row[4] || "";
      let status = row[5] || "";
      let statusLower = String(status).toLowerCase();
      
      // Normalisasi Jam dan Keterlambatan
      let jmFormat = "-", jkFormat = "-", keterangan = "-";
      if (jamMasuk instanceof Date) {
        jmFormat = Utilities.formatDate(jamMasuk, "Asia/Jakarta", "HH:mm");
        if (jamMasuk.getHours() > 8 || (jamMasuk.getHours() === 8 && jamMasuk.getMinutes() > 0)) {
          keterangan = "Terlambat";
        }
      } else if (jamMasuk !== "") jmFormat = String(jamMasuk);
      
      if (jamKeluar instanceof Date) {
        jkFormat = Utilities.formatDate(jamKeluar, "Asia/Jakarta", "HH:mm");
      } else if (jamKeluar !== "") jkFormat = String(jamKeluar);
      
      let itemLog = {
        NRPP: nrpp, Nama: nama, JamMasuk: jmFormat, JamKeluar: jkFormat, Status: status, Keterangan: keterangan, TglAsli: String(tglRow)
      };
      
      // Kumpulkan ke keranjang semua data
      logSemua.push(itemLog);
      
      // Jika memang datanya HARI INI, masukkan ke statistik
      if (isToday) {
        logHariIni.push(itemLog);
        if (statusLower.includes('hadir')) statHadir++;
        else if (statusLower.includes('cuti')) statCuti++;
        else if (statusLower.includes('sakit') || statusLower.includes('izin')) statIzinSakit++;
        if (keterangan === "Terlambat") statTerlambat++;
      }
    }
    
    // --- MODE FALLBACK CERDAS ---
    // Jika hari ini 0 data, kita JANGAN tampilkan tabel kosong.
    // Kita paksa kembalikan 5 data terakhir dari sheet agar kamu bisa melihat isinya.
    let isFallback = false;
    if (logHariIni.length === 0 && logSemua.length > 0) {
      isFallback = true;
      logHariIni = logSemua.slice(-5).reverse(); // Ambil 5 data paling bawah, balik urutannya
    }
    
    return JSON.stringify({
      status: 'success',
      stats: { hadir: statHadir, cuti: statCuti, izin: statIzinSakit, terlambat: statTerlambat },
      logs: logHariIni,
      fallbackMode: isFallback // Penanda untuk Frontend
    });
    
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

/**
 * =========================================================
 * FUNGSI: KELOLA HARI LIBUR
 * =========================================================
 */

// 1. Mengambil daftar hari libur dari sheet Master_Libur
function getLiburData() {
  try {
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID);
    const sheet = ss.getSheetByName('Master_Libur');
    if (!sheet) return JSON.stringify({ status: 'error', message: 'Sheet Master_Libur tidak ditemukan' });
    
    const data = sheet.getDataRange().getValues();
    data.shift(); // Buang baris header
    
    let result = [];
    for (let i = 0; i < data.length; i++) {
      let tgl = data[i][0];
      if (!tgl || tgl === "") continue; // Lewati baris kosong
      
      // Format tanggal agar rapi saat ditampilkan di tabel
      let tglStr = (tgl instanceof Date) ? Utilities.formatDate(tgl, "Asia/Jakarta", "dd-MM-yyyy") : String(tgl);
      
      result.push({
        Tanggal: tglStr,
        Status: data[i][1] || '-',
        Keterangan: data[i][2] || '-'
      });
    }
    
    return JSON.stringify({ status: 'success', data: result });
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

// 2. Menyimpan hari libur baru ke sheet Master_Libur
function saveLiburData(tglInput, statusInput, ketInput) {
  try {
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID);
    const sheet = ss.getSheetByName('Master_Libur');
    if (!sheet) return JSON.stringify({ status: 'error', message: 'Sheet Master_Libur tidak ditemukan' });
    
    // Konversi input string (YYYY-MM-DD dari form HTML) ke format Date untuk Spreadsheet
    let parts = tglInput.split('-');
    let dateObj = new Date(parts[0], parts[1] - 1, parts[2]); // Tahun, Bulan (0-index), Hari
    
    // Tambahkan baris baru ke paling bawah
    sheet.appendRow([dateObj, statusInput, ketInput]);
    
    return JSON.stringify({ status: 'success' });
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

/**
 * =========================================================
 * FUNGSI: EDIT / TAMBAH ABSENSI (VERSI FLEKSIBEL)
 * =========================================================
 */

function getEditAbsensiData(pt, tglInput) {
  try {
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID);
    
    // 1. Pecah input tanggal dari UI (YYYY-MM-DD)
    let parts = tglInput.split('-');
    let tgtY = parseInt(parts[0]);
    let tgtM = parseInt(parts[1]) - 1; // Index bulan di JS mulai dari 0
    let tgtD = parseInt(parts[2]);

    // 2. AMBIL MASTER KARYAWAN
    const sheetKar = ss.getSheetByName('Master_Karyawan');
    if (!sheetKar) throw new Error("Sheet Master_Karyawan tidak ditemukan.");
    const dataKar = sheetKar.getDataRange().getValues();
    dataKar.shift(); // Buang header
    
    let listKaryawan = [];
    let ptLower = String(pt).toLowerCase(); // misal "kobe"
    
    for (let i = 0; i < dataKar.length; i++) {
      let nrppKar = String(dataKar[i][0]).trim();
      let namaKar = dataKar[i][1];
      let ptKar = String(dataKar[i][7] || "").toLowerCase(); // Kolom H
      
      if (!nrppKar) continue;
      
      // STRATEGI TOLERANSI PT: Gunakan 'includes' bukan '==='
      // Jika UI mengirim "kobe", dan di sheet tertulis "pt kobexindo", "kobe" ada di dalamnya!
      if (ptKar.includes(ptLower) || ptLower === 'all' || ptLower === '') {
        listKaryawan.push({ NRPP: nrppKar, Nama: namaKar });
      }
    }

    // Jika karyawan tidak ditemukan sama sekali, kembalikan error spesifik
    if (listKaryawan.length === 0) {
        return JSON.stringify({ status: 'error', message: `Tidak ada karyawan ditemukan untuk kata kunci PT: "${pt}" di Master_Karyawan.` });
    }

    // 3. AMBIL RAW KEHADIRAN
    const sheetRaw = ss.getSheetByName('Raw_Kehadiran');
    if (!sheetRaw) throw new Error("Sheet Raw_Kehadiran tidak ditemukan.");
    const dataRaw = sheetRaw.getDataRange().getValues();
    
    let mapAbsensi = {};
    
    for (let i = 1; i < dataRaw.length; i++) {
      let tglRow = dataRaw[i][0];
      if (!tglRow || tglRow === "") continue;
      
      // LOGIKA TANGGAL KEBAL
      let isMatch = false;
      let rowDate = new Date(tglRow);
      
      if (!isNaN(rowDate.getTime())) {
        if (rowDate.getFullYear() === tgtY && rowDate.getMonth() === tgtM && rowDate.getDate() === tgtD) {
          isMatch = true;
        }
      } else {
        // Fallback string manual
        let tglStr = String(tglRow);
        let strTgt1 = tglInput; // "YYYY-MM-DD"
        let strTgt2 = `${("0"+tgtD).slice(-2)}/${("0"+(tgtM+1)).slice(-2)}/${tgtY}`; // "DD/MM/YYYY"
        let strTgt3 = `${("0"+tgtD).slice(-2)}-${("0"+(tgtM+1)).slice(-2)}-${tgtY}`; // "DD-MM-YYYY"
        if (tglStr.includes(strTgt1) || tglStr.includes(strTgt2) || tglStr.includes(strTgt3)) {
          isMatch = true;
        }
      }

      if (isMatch) {
        let nrppRow = String(dataRaw[i][1]).trim();
        mapAbsensi[nrppRow] = {
          rowIndex: i + 1, // +1 karena array mulai dari 0 tapi sheet mulai dari 1
          JamMasuk: dataRaw[i][3] instanceof Date ? Utilities.formatDate(dataRaw[i][3], "Asia/Jakarta", "HH:mm") : (dataRaw[i][3] || ""),
          JamKeluar: dataRaw[i][4] instanceof Date ? Utilities.formatDate(dataRaw[i][4], "Asia/Jakarta", "HH:mm") : (dataRaw[i][4] || ""),
          Status: dataRaw[i][5] || ""
        };
      }
    }

    // 4. GABUNGKAN DATA
    let result = [];
    for (let k of listKaryawan) {
      let absen = mapAbsensi[k.NRPP];
      result.push({
        NRPP: k.NRPP,
        Nama: k.Nama,
        JamMasuk: absen ? absen.JamMasuk : "",
        JamKeluar: absen ? absen.JamKeluar : "",
        Status: absen ? absen.Status : "",
        RowIndex: absen ? absen.rowIndex : null
      });
    }

    return JSON.stringify({ status: 'success', data: result });
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

function saveSingleAbsensi(nrpp, nama, tglInput, jamMasuk, jamKeluar, status, rowIndex) {
  try {
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID);
    const sheetRaw = ss.getSheetByName('Raw_Kehadiran');
    
    // Normalisasi tanggal agar konsisten di spreadsheet
    let dateObj = tglInput; // Default biarkan format "YYYY-MM-DD"
    let parts = tglInput.split('-');
    if (parts.length === 3) {
      dateObj = new Date(parts[0], parts[1] - 1, parts[2]); // Convert ke objek tanggal
    }

    if (rowIndex) {
      // MODE EDIT
      sheetRaw.getRange(rowIndex, 4).setValue(jamMasuk);
      sheetRaw.getRange(rowIndex, 5).setValue(jamKeluar);
      sheetRaw.getRange(rowIndex, 6).setValue(status);
    } else {
      // MODE TAMBAH BARU
      sheetRaw.appendRow([dateObj, nrpp, nama, jamMasuk, jamKeluar, status]);
    }

    return JSON.stringify({ status: 'success' });
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}

/**
 * Mengambil daftar perusahaan unik dari Master_Karyawan Kolom H
 */
function getUniquePerusahaan() {
  try {
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID);
    const sheet = ss.getSheetByName('Master_Karyawan');
    if (!sheet) return JSON.stringify({ status: 'error', message: 'Sheet Master_Karyawan tidak ditemukan' });
    
    const data = sheet.getDataRange().getValues();
    data.shift(); // Buang header
    
    let listPT = [];
    for (let i = 0; i < data.length; i++) {
      let pt = data[i][7]; // Kolom H (index 7)
      if (pt && pt !== "" && !listPT.includes(pt)) {
        listPT.push(pt);
      }
    }
    
    // Urutkan berdasarkan abjad terlebih dahulu
    let sortedPT = listPT.sort();
    
    // [KODE BARU] Tambahkan opsi "ALL" di index ke-0 (paling atas)
    sortedPT.unshift("ALL");
    
    return JSON.stringify({ status: 'success', data: sortedPT });
  } catch (error) {
    return JSON.stringify({ status: 'error', message: error.toString() });
  }
}
/**
 * Mengambil data untuk tabel Crosstab Rekap Absensi
 * Menerima parameter dari Frontend (Js.html)
 */
function getCrosstabDataBackend(pt, start, end) {
  try {
    // Gunakan ID Sheet dari konstanta yang sudah ada di sistem Anda
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID); 
    const sheetKaryawan = ss.getSheetByName('Master_Karyawan');
    const sheetKehadiran = ss.getSheetByName('Raw_Kehadiran');

    if (!sheetKaryawan || !sheetKehadiran) throw new Error("Database Master_Karyawan atau Raw_Kehadiran tidak ditemukan.");

    // ==========================================
    // 1. SETUP ARRAY TANGGAL DINAMIS
    // ==========================================
    const startDate = new Date(start);
    const endDate = new Date(end);
    let dateArray = [];
    let currentDate = new Date(startDate);
    
    while (currentDate <= endDate) {
      // Format YYYY-MM-DD agar seragam untuk pencocokan matrix
      let dateString = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      dateArray.push(dateString);
      currentDate.setDate(currentDate.getDate() + 1);
    }

    if (dateArray.length > 31) throw new Error("Rentang tanggal terlalu besar. Maksimal 31 hari untuk Rekap Crosstab.");

    // ==========================================
    // 2. AMBIL MASTER KARYAWAN & FILTER PERUSAHAAN
    // ==========================================
    const dataKaryawan = sheetKaryawan.getDataRange().getValues();
    dataKaryawan.shift(); // Buang header baris pertama
    
    let mapKaryawan = {};
    let filteredKaryawanList = [];
    let filterPT = pt ? pt.toString().trim().toUpperCase() : "";

    // Skema Master_Karyawan: NRPP [0], Nama Karyawan [1], Gol [2], Status [3], Jabatan [4], Dept [5], Divisi [6], Perusahaan [7]
    for (let i = 0; i < dataKaryawan.length; i++) {
      let nrpp = dataKaryawan[i][0] ? dataKaryawan[i][0].toString().trim() : "";
      let nama = dataKaryawan[i][1] ? dataKaryawan[i][1].toString().trim() : "";
      let perusahaan = dataKaryawan[i][7] ? dataKaryawan[i][7].toString().trim().toUpperCase() : "";

      if (nrpp === "") continue;

      // Mock Data Mismatch Awareness: Bypass 'ALL' atau gunakan includes()
      if (filterPT === "ALL" || perusahaan.includes(filterPT)) {
        mapKaryawan[nrpp] = {
          nama: nama,
          perusahaan: perusahaan,
          kehadiran: {}, // Matrix tgl akan ditaruh di sini
          totalH: 0, totalS: 0, totalC: 0, totalA: 0
        };
        // Inisialisasi default kehadiran menjadi "-" (kosong) untuk semua tanggal di rentang
        dateArray.forEach(d => mapKaryawan[nrpp].kehadiran[d] = "-");
        filteredKaryawanList.push(nrpp);
      }
    }

    if (filteredKaryawanList.length === 0) {
       return `<tr><td colspan="${dateArray.length + 4}" class="text-center py-4 text-danger fw-bold">Tidak ada Karyawan ditemukan untuk Perusahaan tersebut.</td></tr>`;
    }

    // ==========================================
    // 3. AMBIL RAW KEHADIRAN & MAPPING MATRIX
    // ==========================================
    const dataKehadiran = sheetKehadiran.getDataRange().getValues();
    dataKehadiran.shift(); // Buang Header
    
    // Skema Raw_Kehadiran: Tanggal [0], NRPP [1], Nama [2], Jam Masuk [3], Jam Keluar [4], Status [5]
    for (let i = 0; i < dataKehadiran.length; i++) {
      let rowTgl = dataKehadiran[i][0];
      if (!rowTgl) continue;
      
      let tglRaw = new Date(rowTgl);
      if (isNaN(tglRaw.getTime())) continue; // Lewati jika format tanggal sheet rusak
      
      let tglStr = Utilities.formatDate(tglRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
      let nrpp = dataKehadiran[i][1] ? dataKehadiran[i][1].toString().trim() : "";
      let statusString = dataKehadiran[i][5] ? dataKehadiran[i][5].toString().trim().toUpperCase() : "H";
      
      // Jika karyawan ada di daftar filter DAN tanggal masuk di rentang pencarian
      if (dateArray.includes(tglStr) && mapKaryawan[nrpp]) {
        let inisialStatus = statusString.substring(0, 1); // Ambil huruf pertama (Hadir=H, Sakit=S)
        mapKaryawan[nrpp].kehadiran[tglStr] = inisialStatus;
        
        // Kalkulasi Total Summary
        if (inisialStatus === "H") mapKaryawan[nrpp].totalH++;
        else if (inisialStatus === "S") mapKaryawan[nrpp].totalS++;
        else if (inisialStatus === "C") mapKaryawan[nrpp].totalC++;
        else mapKaryawan[nrpp].totalA++; // Alpa / lainnya
      }
    }

    // ==========================================
    // 4. BANGUN HTML TABLE
    // ==========================================
    let html = "";
    
    // --- Render Header Baris 1 ---
    html += `<tr class="table-primary text-center align-middle" style="font-size: 0.85rem; border-bottom: 2px solid #dee2e6;">`;
    html += `<th rowspan="2">No</th>`;
    html += `<th rowspan="2" style="min-width: 80px;">NRPP</th>`;
    html += `<th rowspan="2" style="min-width: 150px;">Nama</th>`;
    if (filterPT === "ALL") html += `<th rowspan="2" style="min-width: 100px;">Perusahaan</th>`;
    html += `<th colspan="${dateArray.length}">Tanggal</th>`;
    html += `<th colspan="4">Total</th>`;
    html += `</tr>`;
    
    // --- Render Header Baris 2 (Angka Tanggal Saja) ---
    html += `<tr class="table-primary text-center align-middle" style="font-size: 0.8rem;">`;
    dateArray.forEach(d => {
        let tglHari = d.split("-")[2]; // Ekstrak DD dari YYYY-MM-DD
        html += `<th style="min-width:35px;">${tglHari}</th>`;
    });
    html += `<th class="text-success">H</th><th class="text-warning">S</th><th class="text-warning">C</th><th class="text-danger">A</th>`;
    html += `</tr>`;

    // --- Render Body (Looping setiap Karyawan) ---
    let no = 1;
    filteredKaryawanList.forEach(nrpp => {
      let kar = mapKaryawan[nrpp];
      html += `<tr class="text-center align-middle table-hover" style="font-size: 0.85rem;">`;
      html += `<td>${no++}</td>`;
      html += `<td>${nrpp}</td>`;
      html += `<td class="text-start text-nowrap">${kar.nama}</td>`;
      if (filterPT === "ALL") html += `<td class="text-start small">${kar.perusahaan}</td>`;
      
      // Render Matrix per Tanggal
      dateArray.forEach(d => {
         let stat = kar.kehadiran[d];
         let colorClass = stat === "H" ? "text-success fw-bold" : 
                          (stat === "S" || stat === "C") ? "text-warning fw-bold" : 
                          stat === "A" ? "text-danger fw-bold" : "text-muted";
         html += `<td class="${colorClass}">${stat}</td>`;
      });
      
      // Render Kolom Total
      html += `<td class="fw-bold bg-light">${kar.totalH}</td>`;
      html += `<td class="bg-light">${kar.totalS}</td>`;
      html += `<td class="bg-light">${kar.totalC}</td>`;
      html += `<td class="text-danger fw-bold bg-light">${kar.totalA}</td>`;
      html += `</tr>`;
    });

    return html;

  } catch (error) {
    throw new Error(error.toString());
  }
}
/**
 * Mengambil data untuk tabel Crosstab Rekap Absensi (Versi Upgrade dgn Jam Masuk/Keluar)
 * Menerima parameter dari Frontend (Js.html)
 */
function getCrosstabDataBackend(pt, start, end) {
  try {
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID); 
    const sheetKaryawan = ss.getSheetByName('Master_Karyawan');
    const sheetKehadiran = ss.getSheetByName('Raw_Kehadiran');

    if (!sheetKaryawan || !sheetKehadiran) throw new Error("Database Master_Karyawan atau Raw_Kehadiran tidak ditemukan.");

    // ==========================================
    // 1. SETUP ARRAY TANGGAL DINAMIS
    // ==========================================
    const startDate = new Date(start);
    const endDate = new Date(end);
    let dateArray = [];
    let currentDate = new Date(startDate);
    
    while (currentDate <= endDate) {
      let dateString = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      dateArray.push(dateString);
      currentDate.setDate(currentDate.getDate() + 1);
    }

    if (dateArray.length > 31) throw new Error("Rentang tanggal terlalu besar. Maksimal 31 hari untuk Rekap Crosstab.");

    // ==========================================
    // 2. AMBIL MASTER KARYAWAN & FILTER PERUSAHAAN
    // ==========================================
    const dataKaryawan = sheetKaryawan.getDataRange().getValues();
    dataKaryawan.shift(); 
    
    let mapKaryawan = {};
    let filteredKaryawanList = [];
    let filterPT = pt ? pt.toString().trim().toUpperCase() : "";

    for (let i = 0; i < dataKaryawan.length; i++) {
      let nrpp = dataKaryawan[i][0] ? dataKaryawan[i][0].toString().trim() : "";
      let nama = dataKaryawan[i][1] ? dataKaryawan[i][1].toString().trim() : "";
      let perusahaan = dataKaryawan[i][7] ? dataKaryawan[i][7].toString().trim().toUpperCase() : "";

      if (nrpp === "") continue;

      if (filterPT === "ALL" || perusahaan.includes(filterPT)) {
        mapKaryawan[nrpp] = {
          nama: nama,
          perusahaan: perusahaan,
          kehadiran: {}, 
          totalH: 0, totalS: 0, totalC: 0, totalA: 0
        };
        // [KODE BARU] Inisialisasi object untuk menyimpan status DAN jam
        dateArray.forEach(d => mapKaryawan[nrpp].kehadiran[d] = { stat: "-", masuk: "", keluar: "" });
        filteredKaryawanList.push(nrpp);
      }
    }

    if (filteredKaryawanList.length === 0) {
       return `<tr><td colspan="${dateArray.length + 4}" class="text-center py-4 text-danger fw-bold">Tidak ada Karyawan ditemukan untuk Perusahaan tersebut.</td></tr>`;
    }

    // ==========================================
    // 3. AMBIL RAW KEHADIRAN & MAPPING MATRIX
    // ==========================================
    const dataKehadiran = sheetKehadiran.getDataRange().getValues();
    dataKehadiran.shift(); 
    
    for (let i = 0; i < dataKehadiran.length; i++) {
      let rowTgl = dataKehadiran[i][0];
      if (!rowTgl) continue;
      
      let tglRaw = new Date(rowTgl);
      if (isNaN(tglRaw.getTime())) continue; 
      
      let tglStr = Utilities.formatDate(tglRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
      let nrpp = dataKehadiran[i][1] ? dataKehadiran[i][1].toString().trim() : "";
      let statusString = dataKehadiran[i][5] ? dataKehadiran[i][5].toString().trim().toUpperCase() : "H";
      
      if (dateArray.includes(tglStr) && mapKaryawan[nrpp]) {
        let inisialStatus = statusString.substring(0, 1); 
        
        // [KODE BARU] Ambil dan Format Jam Masuk / Keluar (Mengatasi Date Object dari Sheet)
        let rawMasuk = dataKehadiran[i][3];
        let rawKeluar = dataKehadiran[i][4];
        let jamMasuk = rawMasuk instanceof Date ? Utilities.formatDate(rawMasuk, Session.getScriptTimeZone(), "HH:mm") : (rawMasuk ? rawMasuk.toString().trim() : "");
        let jamKeluar = rawKeluar instanceof Date ? Utilities.formatDate(rawKeluar, Session.getScriptTimeZone(), "HH:mm") : (rawKeluar ? rawKeluar.toString().trim() : "");

        // Simpan ke dictionary
        mapKaryawan[nrpp].kehadiran[tglStr] = {
            stat: inisialStatus,
            masuk: jamMasuk,
            keluar: jamKeluar
        };
        
        // Kalkulasi Total Summary
        if (inisialStatus === "H") mapKaryawan[nrpp].totalH++;
        else if (inisialStatus === "S") mapKaryawan[nrpp].totalS++;
        else if (inisialStatus === "C") mapKaryawan[nrpp].totalC++;
        else mapKaryawan[nrpp].totalA++; 
      }
    }

    // ==========================================
    // 4. BANGUN HTML TABLE
    // ==========================================
    let html = "";
    
    html += `<tr class="table-primary text-center align-middle" style="font-size: 0.85rem; border-bottom: 2px solid #dee2e6;">`;
    html += `<th rowspan="2">No</th>`;
    html += `<th rowspan="2" style="min-width: 80px;">NRPP</th>`;
    html += `<th rowspan="2" style="min-width: 150px;">Nama</th>`;
    if (filterPT === "ALL") html += `<th rowspan="2" style="min-width: 100px;">Perusahaan</th>`;
    html += `<th colspan="${dateArray.length}">Tanggal</th>`;
    html += `<th colspan="4">Total</th>`;
    html += `</tr>`;
    
    html += `<tr class="table-primary text-center align-middle" style="font-size: 0.8rem;">`;
    dateArray.forEach(d => {
        let tglHari = d.split("-")[2]; 
        html += `<th style="min-width:60px;">${tglHari}</th>`; // Lebar kolom sedikit ditambah untuk teks waktu
    });
    html += `<th class="text-success">H</th><th class="text-warning">S</th><th class="text-warning">C</th><th class="text-danger">A</th>`;
    html += `</tr>`;

    let no = 1;
    filteredKaryawanList.forEach(nrpp => {
      let kar = mapKaryawan[nrpp];
      html += `<tr class="text-center align-middle table-hover" style="font-size: 0.85rem;">`;
      html += `<td>${no++}</td>`;
      html += `<td>${nrpp}</td>`;
      html += `<td class="text-start text-nowrap">${kar.nama}</td>`;
      if (filterPT === "ALL") html += `<td class="text-start small">${kar.perusahaan}</td>`;
      
      // Render Matrix per Tanggal
      dateArray.forEach(d => {
         let dataHariIni = kar.kehadiran[d];
         let stat = dataHariIni.stat;
         
         let colorClass = stat === "H" ? "text-success fw-bold" : 
                          (stat === "S" || stat === "C") ? "text-warning fw-bold" : 
                          stat === "A" ? "text-danger fw-bold" : "text-muted";
         
         // [KODE BARU] Sisipkan Jam Masuk & Keluar jika statusnya Hadir (H)
         let infoWaktu = "";
         if (stat === "H" && (dataHariIni.masuk || dataHariIni.keluar)) {
             let jM = dataHariIni.masuk || "--:--";
             let jK = dataHariIni.keluar || "--:--";
             infoWaktu = `<br><span style="font-size: 0.65rem; font-weight: normal; color: #6c757d; white-space: nowrap;">${jM} - ${jK}</span>`;
         }

         html += `<td class="${colorClass}">${stat}${infoWaktu}</td>`;
      });
      
      // Render Kolom Total
      html += `<td class="fw-bold bg-light">${kar.totalH}</td>`;
      html += `<td class="bg-light">${kar.totalS}</td>`;
      html += `<td class="bg-light">${kar.totalC}</td>`;
      html += `<td class="text-danger fw-bold bg-light">${kar.totalA}</td>`;
      html += `</tr>`;
    });

    return html;

  } catch (error) {
    throw new Error(error.toString());
  }
}

/**
 * Mengirim data Crosstab Rekap ke Frontend sesuai format SPA
 */
function getCrosstabRekapData(start, end, pt) {
  try {
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID);
    
    // ========================================================
    // 1. BACA DATA DATABASE
    // ========================================================
    const sheetKar = ss.getSheetByName('Master_Karyawan');
    const sheetHadir = ss.getSheetByName('Raw_Kehadiran');
    const sheetLibur = ss.getSheetByName('Master_Libur'); // Asumsi ada sheet Master_Libur
    
    if (!sheetKar || !sheetHadir) throw new Error("Sheet Master_Karyawan atau Raw_Kehadiran tidak ditemukan");
    
    const dataKar = sheetKar.getDataRange().getValues(); dataKar.shift();
    const dataHadir = sheetHadir.getDataRange().getValues(); dataHadir.shift();
    const dataLibur = sheetLibur ? sheetLibur.getDataRange().getValues() : []; 
    if(dataLibur.length > 0) dataLibur.shift();

    // ========================================================
    // 2. PETAKAN HARI LIBUR
    // ========================================================
    let mapLibur = {};
    dataLibur.forEach(row => {
      if(row[0]) {
        let tglLiburStr = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
        mapLibur[tglLiburStr] = { stat: row[1] || "L", ket: row[2] || "Libur" }; // Status [1], Keterangan [2]
      }
    });

// ========================================================
    // 3. SUSUN ARRAY DATES & HITUNG WORKDAYS
    // ========================================================
    const startDate = new Date(start);
    const endDate = new Date(end);
    let datesArray = [];
    let workdays = 0;
    
    // Array untuk mengubah angka bulan/hari menjadi teks bahasa Indonesia
    const namaBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const namaHari = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
    
    let curDate = new Date(startDate);
    while (curDate <= endDate) {
      let rawStr = Utilities.formatDate(curDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      
      // Buat format panjang: "Selasa, 28 April 2026"
      let dHari = namaHari[curDate.getDay()];
      let dTgl = curDate.getDate();
      let dBulan = namaBulan[curDate.getMonth()];
      let dTahun = curDate.getFullYear();
      let displayLengkap = `${dHari}, ${dTgl} ${dBulan} ${dTahun}`;

      let isLibur = mapLibur[rawStr] !== undefined;
      
      datesArray.push({
        raw: rawStr,
        display: displayLengkap, // <--- Perubahan di sini, menggunakan displayLengkap
        isLibur: isLibur,
        liburStat: isLibur ? mapLibur[rawStr].stat : "",
        liburKet: isLibur ? mapLibur[rawStr].ket : ""
      });
      
      if (!isLibur && curDate.getDay() !== 0) workdays++;
      curDate.setDate(curDate.getDate() + 1);
    }

    // ========================================================
    // 4. SUSUN DAFTAR KARYAWAN & STRUKTUR OBJEK
    // ========================================================
    let employeesArray = [];
    let attendanceMap = {};
    let empDict = {}; // Untuk menampung rekap sementara
    let filterPT = pt ? pt.toString().trim().toUpperCase() : "";

    dataKar.forEach(row => {
      let nrpp = row[0] ? row[0].toString().trim() : "";
      let nama = row[1] ? row[1].toString().trim() : "";
      let perusahaan = row[7] ? row[7].toString().trim().toUpperCase() : "";

      if (nrpp && (filterPT === "ALL" || perusahaan.includes(filterPT))) {
        empDict[nrpp] = {
          nrpp: nrpp, nama: nama, lateSecs: 0, H: 0, S: 0, C: 0, ST: 0, VC: 0, PG: 0
        };
        attendanceMap[nrpp] = {};
      }
    });

    // ========================================================
    // 5. ISI MATRIKS KEHADIRAN & HITUNG REKAP
    // ========================================================
    dataHadir.forEach(row => {
      let rawDate = row[0];
      if (!rawDate) return;
      
      let dateStr = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
      let nrpp = row[1] ? row[1].toString().trim() : "";
      
      // Filter jika tgl ada di rentang & karyawan valid
      if (empDict[nrpp] && datesArray.find(d => d.raw === dateStr)) {
        
        // Handling Jam (String atau Date Object dari sheet)
        let rawIn = row[3]; let rawOut = row[4];
        let jamIn = rawIn instanceof Date ? Utilities.formatDate(rawIn, Session.getScriptTimeZone(), "HH:mm") : (rawIn ? rawIn.toString().trim() : "-");
        let jamOut = rawOut instanceof Date ? Utilities.formatDate(rawOut, Session.getScriptTimeZone(), "HH:mm") : (rawOut ? rawOut.toString().trim() : "-");
        
        let statStr = row[5] ? row[5].toString().trim().toUpperCase() : "H";
        let statCode = statStr.substring(0, 2).trim(); // Ambil 1-2 huruf
        
        // Simpan ke matriks per tanggal
        attendanceMap[nrpp][dateStr] = {
          in: jamIn,
          out: jamOut,
          stat: statCode
        };

        // Tambah ke Rekap Summary
        if (statCode === "H") empDict[nrpp].H++;
        else if (statCode === "S") empDict[nrpp].S++;
        else if (statCode === "C") empDict[nrpp].C++;
        else if (statCode === "ST") empDict[nrpp].ST++;
        else if (statCode === "VC") empDict[nrpp].VC++;
        else empDict[nrpp].PG++; // PG atau Alpa

        // Hitung Keterlambatan (Late) jika Jam Masuk > 08:00
        if (jamIn !== "-" && jamIn.length >= 5) {
           let inHrs = parseInt(jamIn.substring(0,2));
           let inMins = parseInt(jamIn.substring(3,5));
           let totalInMins = (inHrs * 60) + inMins;
           let thresholdMins = 8 * 60; // 08:00
           if (totalInMins > thresholdMins) {
               empDict[nrpp].lateSecs += (totalInMins - thresholdMins) * 60;
           }
        }
      }
    });

    // ========================================================
    // 6. KONVERSI LATE SECONDS KE STRING & FINISHING
    // ========================================================
    Object.keys(empDict).forEach(k => {
      let emp = empDict[k];
      
      // Konversi detik menjadi format HH:mm:ss
      if (emp.lateSecs === 0) {
        emp.lateStr = "00:00:00";
      } else {
        let hrs = Math.floor(emp.lateSecs / 3600);
        let mins = Math.floor((emp.lateSecs % 3600) / 60);
        let secs = emp.lateSecs % 60;
        emp.lateStr = `${hrs.toString().padStart(2, '0')}:${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
      }
      
      employeesArray.push(emp);
    });

    // Mengembalikan data JSON persis seperti yang diharapkan Js.html
    return {
      dates: datesArray,
      employees: employeesArray,
      attendance: attendanceMap,
      workdays: workdays
    };

  } catch (error) {
    throw new Error(error.toString());
  }
}

/**
 * Mengambil detail kehadiran 1 karyawan spesifik untuk format Laporan/PDF
 */
function getLaporanKaryawanData(inputKaryawan, start, end, pt) {
  try {
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID); // Pastikan DB_ABSENSI_ID sudah dideklarasikan global
    const sheetKar = ss.getSheetByName('Master_Karyawan');
    const sheetHadir = ss.getSheetByName('Raw_Kehadiran');
    const sheetLibur = ss.getSheetByName('Master_Libur');

    if (!sheetKar || !sheetHadir) throw new Error("Database tidak ditemukan.");

    // =====================================
    // 1. CARI PROFIL KARYAWAN
    // =====================================
    const dataKar = sheetKar.getDataRange().getValues(); 
    dataKar.shift();
    
    let profile = null;
    let inputSearch = inputKaryawan.toString().toUpperCase().trim();
    
    for (let i = 0; i < dataKar.length; i++) {
      let nrpp = dataKar[i][0] ? dataKar[i][0].toString().toUpperCase() : "";
      let nama = dataKar[i][1] ? dataKar[i][1].toString().toUpperCase() : "";
      
      // Mock Data Mismatch Awareness: Cocokkan jika input ada di NRPP atau Nama
      if (nrpp.includes(inputSearch) || nama.includes(inputSearch)) {
        profile = {
          nrpp: dataKar[i][0], nama: dataKar[i][1], gol: dataKar[i][2],
          status: dataKar[i][3], jabatan: dataKar[i][4], dept: dataKar[i][5],
          divisi: dataKar[i][6], perusahaan: dataKar[i][7]
        };
        break; // Hentikan loop jika sudah ketemu
      }
    }

    if (!profile) throw new Error("Data Karyawan tidak ditemukan.");

    // =====================================
    // 2. PETAKAN HARI LIBUR
    // =====================================
    const dataLibur = sheetLibur ? sheetLibur.getDataRange().getValues() : [];
    if (dataLibur.length > 0) dataLibur.shift();
    let mapLibur = {};
    dataLibur.forEach(row => {
      if(row[0]) {
        let tglStr = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
        mapLibur[tglStr] = { stat: row[1] || "L", ket: row[2] || "Libur" };
      }
    });

    // =====================================
    // 3. AMBIL DATA RAW KEHADIRAN
    // =====================================
    const dataHadir = sheetHadir.getDataRange().getValues(); 
    dataHadir.shift();
    let absensiMap = {};
    
    dataHadir.forEach(row => {
      let tglRaw = row[0]; let nrppRaw = row[1];
      if (!tglRaw || !nrppRaw) return;
      
      // Pastikan NRPP cocok
      if (nrppRaw.toString().trim() === profile.nrpp.toString().trim()) {
         let tglStr = Utilities.formatDate(new Date(tglRaw), Session.getScriptTimeZone(), "yyyy-MM-dd");
         
         let jamIn = row[3] instanceof Date ? Utilities.formatDate(row[3], Session.getScriptTimeZone(), "HH:mm") : (row[3] ? row[3].toString() : "");
         let jamOut = row[4] instanceof Date ? Utilities.formatDate(row[4], Session.getScriptTimeZone(), "HH:mm") : (row[4] ? row[4].toString() : "");
         
         absensiMap[tglStr] = { in: jamIn, out: jamOut, stat: row[5] ? row[5].toString().toUpperCase() : "H" };
      }
    });

    // =====================================
    // 4. SUSUN MATRIX LAPORAN & KALKULASI
    // =====================================
    let startDate = new Date(start);
    let endDate = new Date(end);
    let hasilMatrix = [];
    let rekap = { H: 0, S: 0, C: 0, ST: 0, PG: 0, VC: 0, HariKerja: 0 };
    const namaHari = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
    
    let curDate = new Date(startDate);
    let no = 1;

    while (curDate <= endDate) {
      let tglStr = Utilities.formatDate(curDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      let dFormat = Utilities.formatDate(curDate, Session.getScriptTimeZone(), "dd-MM-yyyy");
      let dHari = namaHari[curDate.getDay()];
      
      let isLibur = mapLibur[tglStr] !== undefined;
      let att = absensiMap[tglStr];
      
      let baris = {
        no: no++, hari: dHari, tanggal: dFormat,
        jamIn: "-", jamOut: "-", telat: "-", stat: "-", durasi: "-"
      };

      if (att) {
         baris.jamIn = att.in || "-";
         baris.jamOut = att.out || "-";
         baris.stat = att.stat.substring(0, 2).trim();
         
         // Kalkulasi Keterlambatan (Batas jam 08:00)
         if (baris.jamIn !== "-") {
            let hm = baris.jamIn.split(":");
            let mntIn = parseInt(hm[0])*60 + parseInt(hm[1]);
            let limit = 8 * 60;
            if (mntIn > limit) {
               let diff = mntIn - limit;
               baris.telat = `${Math.floor(diff/60).toString().padStart(2,'0')}:${(diff%60).toString().padStart(2,'0')}`;
            }
         }
         
         // Kalkulasi Durasi Jam
         if (baris.jamIn !== "-" && baris.jamOut !== "-") {
            let inArr = baris.jamIn.split(":"); let outArr = baris.jamOut.split(":");
            let mntIn = parseInt(inArr[0])*60 + parseInt(inArr[1]);
            let mntOut = parseInt(outArr[0])*60 + parseInt(outArr[1]);
            let diff = mntOut - mntIn;
            if(diff > 0) baris.durasi = `${Math.floor(diff/60)} Jam, ${diff%60} Menit`;
         }

         // Tambah rekap
         if(baris.stat === 'H') rekap.H++;
         else if(baris.stat === 'S') rekap.S++;
         else if(baris.stat === 'C') rekap.C++;
         else if(baris.stat === 'ST') rekap.ST++;
         else if(baris.stat === 'VC') rekap.VC++;
         else rekap.PG++; // Alpa/PG

      } else if (isLibur) {
         baris.stat = mapLibur[tglStr].stat;
         baris.durasi = mapLibur[tglStr].ket;
      }

      if (!isLibur && curDate.getDay() !== 0) rekap.HariKerja++; // Asumsi hari kerja bukan minggu & libur
      
      hasilMatrix.push(baris);
      curDate.setDate(curDate.getDate() + 1);
    }

    return { status: "success", profile: profile, matrix: hasilMatrix, rekap: rekap };

  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

/**
 * Mengambil daftar karyawan untuk Auto-Suggest Datalist
 * Format output: "NRPP - NAMA KARYAWAN"
 */
function getDatalistKaryawan(pt) {
  try {
    const ss = SpreadsheetApp.openById(DB_ABSENSI_ID);
    const sheet = ss.getSheetByName('Master_Karyawan');
    if (!sheet) throw new Error("Sheet Master_Karyawan tidak ditemukan.");

    const data = sheet.getDataRange().getValues();
    data.shift(); // Buang header

    let list = [];
    let filterPT = pt ? pt.toString().trim().toUpperCase() : "";

    data.forEach(row => {
      let nrpp = row[0] ? row[0].toString().trim() : "";
      let nama = row[1] ? row[1].toString().trim() : "";
      let perusahaan = row[7] ? row[7].toString().trim().toUpperCase() : "";

      if (nrpp !== "") {
        // Jika filter "ALL" atau PT cocok, masukkan ke list
        if (filterPT === "ALL" || perusahaan.includes(filterPT) || filterPT === "") {
          list.push(`${nrpp} - ${nama}`);
        }
      }
    });

    return list;
  } catch (error) {
    throw new Error(error.toString());
  }
}

