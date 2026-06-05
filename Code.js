// Konfigurasi Header untuk setiap Modul sesuai Template SIDAK KSDAE KEMENHUT
const MODULE_HEADERS = {
  "C.01": ["ID", "Timestamp", "Tahun", "Periode", "Satuan Kerja", "Nama Ilmiah Spesies", "Jml Indv yang dijumpai", "Tanggal Perjumpaan", "Longitude (X)", "Latitude (Y)", "Kategori Perjumpaan", "Keterangan"],
  "C.02": ["ID", "Timestamp", "Tahun", "Periode", "Satuan Kerja", "Nama Lembaga Konservasi", "Betuk Lembaga Konservasi", "Alamat", "Longitude (X)", "Latitude (Y)", "Luas Areal (Ha)", "Nomor Perizinan", "Tanggal Perizinan", "Tanggal Berakhir Izin", "Keterangan"],
  "C.03": ["ID", "Timestamp", "Tahun", "Periode", "Satuan Kerja", "Nama Lembaga Konservasi", "Bentuk Lembaga Konservasi", "Nama Ilmiah Spesies", "Jantan", "Betina", "Belum Diketahui", "Keterangan"],
  "C.04": ["ID", "Timestamp", "Tahun", "Satuan Kerja", "Nama Penangkar", "Alamat", "Lokasi Unit Penangkaran", "Longitude (X)", "Latitude (Y)", "Nomor Perizinan", "Tanggal Perizinan", "Tanggal Berakhir Izin", "Luas Areal (Ha)", "Keterangan"],
  "C.05": ["ID", "Timestamp", "Tahun", "Periode", "Satuan Kerja", "Nama Penangkar", "Nama Ilmiah Spesies", "Jumlah Indukan", "Hasil Penangkaran", "Hasil Pemanfaatan Tahun ini", "Sisa Stok Hasil Pemanfaatan", "Keterangan"],
  "C.06": ["ID", "Timestamp", "Tahun", "Periode", "Satuan Kerja", "Nama Pengedar", "Kategori", "Lokasi Unit Pengedar", "Longitude (X)", "Latitude (Y)", "Nomor Perizinan", "Tanggal Perizinan", "Tanggal Berakhir Izin", "Spesies yang Diedarkan", "Keterangan"],
  "C.08": ["ID", "Timestamp", "Tahun", "Periode", "Satuan Kerja", "Nama Spesies", "Jenis", "Provinsi", "Volume Realisasi Tangkap", "Satuan Realisasi Tangkap", "Keterangan"],
  "C.09": ["ID", "Timestamp", "Tahun", "Periode", "Satuan Kerja", "Nama Penangkar", "Nama Spesies", "Jenis", "Negara", "Volume Kuota Ekspor", "Satuan Kuota Ekspor", "Volume Realisasi Ekspor", "Satuan Realisasi Ekspor", "Keterangan"],
  "C.10": ["ID", "Timestamp", "Tahun", "Periode", "Satuan Kerja", "Nama Pengedar TSL", "Nama Spesies", "Jenis", "Negara", "Volume Kuota Ekspor", "Satuan Kuota Ekspor", "Volume Realisasi Ekspor", "Satuan Realisasi Ekspor", "Keterangan"],
  "C.11": ["ID", "Timestamp", "Tahun", "Periode", "Satuan Kerja", "Nama Ilmiah Spesies", "Asal-usul", "Jml Jenis Kelamin Jantan", "Jml Jenis Kelamin Betina", "Jml Jenis Kelamin Belum Diketahui", "Keterangan"],
  "C.12": ["ID", "Timestamp", "Tahun", "Periode", "Satuan Kerja", "Akun PNBP", "Realisasi Bulan ini", "Keterangan"],
  "C.14": ["ID", "Timestamp", "Tahun", "Periode", "Satuan Kerja", "Nama Ilmiah Spesies", "Tanggal Kejadian", "Jml Indv", "Longitude (X)", "Latitude (Y)", "Korban Manusia Meninggal", "Korban Manusia Cedera", "Kerusakan Kebun/Lahan Masy (m2)", "Kerusakan Bangunan", "Korban Kambing", "Korban Sapi", "Korban Kerbau", "Korban Anjing", "Korban Babi", "Korban Unggas", "Korban Satwa Mati", "Taksiran Kerugian (Rp)", "Upaya Penanggulangan", "Keterangan"],
  "C.15": ["ID", "Timestamp", "Tahun", "Periode", "Satuan Kerja", "Nama Ilmiah Spesies", "Jml Indv", "Asal-usul", "Longitude (X)", "Latitude (Y)", "Keterangan"],
  "C.16": ["ID", "Timestamp", "Tahun", "Periode", "Satuan Kerja", "Nama Ilmiah Satwa", "Tanggal Kelahiran", "Jumlah Jantan", "Jumlah Betina", "Jumlah Belum Diketahui", "Longitude (X)", "Latitude (Y)", "Usulan Nama Satwa", "Keterangan"],
  "C.17": ["ID", "Timestamp", "Tahun", "Periode", "Satuan Kerja", "Nama Ilmiah Satwa", "Tanggal Kematian", "Jumlah Jantan", "Jumlah Betina", "Jumlah Belum Diketahui", "Perkiraan Usia Satwa", "Penyebab Kematian", "Longitude (X)", "Latitude (Y)", "Upaya Penanganan", "Keterangan"]
};

// Fungsi inisialisasi awal web app
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SIDAK KSDAE - Web Konservasi Spesies dan Genetik')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Setup otomatis database Google Sheets jika belum tersedia sheet penampung modul
function setupDatabase() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Object.keys(MODULE_HEADERS).forEach(moduleCode => {
      let sheet = ss.getSheetByName(moduleCode);
      if (!sheet) {
        sheet = ss.insertSheet(moduleCode);
        sheet.appendRow(MODULE_HEADERS[moduleCode]);
        sheet.getRange(1, 1, 1, MODULE_HEADERS[moduleCode].length)
          .setFontWeight("bold")
          .setBackground("#15803d")
          .setFontColor("#ffffff");
      }
    });
    return { success: true, message: "Database sukses divalidasi dan diinisialisasi." };
  } catch(e) {
    return { success: false, message: "Gagal inisialisasi: " + e.toString() };
  }
}

// Fungsi mendapatkan rekapitulasi data modul dari spreadsheet
function getRecentData(moduleCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(moduleCode);
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName(moduleCode);
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    if (values.length <= 1) return []; // Hanya baris header
    
    const headers = values[0];
    const records = [];
    
    // Ambil data aktif (abaikan header), urutkan terbaru di atas (kebalikan dari sheet append)
    for (let i = values.length - 1; i >= 1; i--) {
      let row = values[i];
      let obj = {};
      headers.forEach((header, colIndex) => {
        let cellVal = row[colIndex];
        if (cellVal instanceof Date) {
          // Format tanggal ramah pengguna
          obj[header] = Utilities.formatDate(cellVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          obj[header] = cellVal;
        }
      });
      records.push(obj);
    }
    return records;
  } catch(e) {
    return [];
  }
}

// Logika penyimpanan data baru atau memperbarui data lama
function submitConservationData(moduleCode, fields) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(moduleCode);
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName(moduleCode);
    }
    
    // Penentuan periode pelaporan secara modular
    let periodName = "Tahunan";
    if (fields.period_type === "Bulanan") {
      periodName = fields.bulan_val || "Januari";
    } else if (fields.period_type === "Triwulan") {
      periodName = fields.triwulan_val || "Triwulan I";
    } else if (fields.period_type === "Semester") {
      periodName = fields.semester_val || "Semester I";
    }

    const timestamp = new Date();
    const headers = MODULE_HEADERS[moduleCode];
    const isEdit = fields.row_id && fields.row_id !== "";
    const uniqueId = isEdit ? fields.row_id : (moduleCode + "-" + timestamp.getTime());
    
    // Pemetaan data secara menyeluruh dan presisi sesuai nama input formulir ke kolom Excel
    const rowValues = headers.map(header => {
      switch (header) {
        case 'ID': return uniqueId;
        case 'Timestamp': return timestamp;
        case 'Tahun': return fields.tahun || "2026";
        case 'Periode':
        case 'Bulan/Periode':
        case 'Triwulan/Periode':
        case 'Semester/Periode':
          return periodName;
        case 'Satuan Kerja': return "Balai KSDA Kalimantan Timur"; // Terkunci permanen
        
        // PETA NAMA ILMIAH SPESIES / SATWA (Mencegah tidak masuk rekap)
        case 'Nama Ilmiah Spesies':
        case 'Nama Ilmiah Satwa':
        case 'Nama Spesies':
        case 'Spesies yang Diedarkan':
          return fields.spesies_val || '';
          
        // PETA JUMLAH INDIVIDU (Mencegah tidak masuk rekap)
        case 'Jml Indv yang dijumpai':
        case 'Jumlah Jantan':
        case 'Jml Jenis Kelamin Jantan':
        case 'Jml Indv':
        case 'Jumlah Indukan':
          if (moduleCode === 'C.03' || moduleCode === 'C.11' || moduleCode === 'C.16' || moduleCode === 'C.17') {
            return fields.c03_jantan || fields.c11_jantan || fields.c16_jantan || fields.c17_jantan || '0';
          }
          return fields.jumlah_val || fields.c05_indukan || '';

        case 'Betina':
        case 'Jml Jenis Kelamin Betina':
        case 'Jumlah Betina':
          return fields.c03_betina || fields.c11_betina || fields.c16_betina || fields.c17_betina || '0';

        case 'Belum Diketahui':
        case 'Jml Jenis Kelamin Belum Diketahui':
        case 'Jumlah Belum Diketahui':
          return fields.c03_unknown || fields.c11_unknown || fields.c16_unknown || fields.c17_unknown || '0';

        // Koordinat Geografis (X/Y)
        case 'Longitude (X)': 
          return fields.c01_lng || fields.c02_lng || fields.c04_lng || fields.c06_lng || fields.c14_lng || fields.c15_lng || fields.c16_lng || fields.c17_lng || '';
        case 'Latitude (Y)': 
          return fields.c01_lat || fields.c02_lat || fields.c04_lat || fields.c06_lat || fields.c14_lat || fields.c15_lat || fields.c16_lat || fields.c17_lat || '';
          
        // C.01 Perjumpaan
        case 'Tanggal Perjumpaan': return fields.c01_tanggal || '';
        case 'Kategori Perjumpaan': return fields.c01_kategori || '';
        
        // C.02 & C.03 Lembaga Konservasi
        case 'Nama Lembaga Konservasi': return fields.c02_lembaga || fields.c03_lembaga || '';
        case 'Betuk Lembaga Konservasi': // Sesuai typo header asli template C.02
        case 'Bentuk Lembaga Konservasi': 
          return fields.c02_bentuk || fields.c03_bentuk || '';
        case 'Alamat': return fields.c02_alamat || fields.c04_alamat || '';
        case 'Luas Areal (Ha)': return fields.c02_luas || fields.c04_luas || '';
        case 'Nomor Perizinan': return fields.c02_izin_no || fields.c04_izin || fields.c06_izin || '';
        case 'Tanggal Perizinan': return fields.c02_izin_tgl || fields.c04_izin_tgl || fields.c06_izin_tgl || '';
        case 'Tanggal Berakhir Izin': return fields.c02_izin_akhir || fields.c04_izin_akhir || fields.c06_izin_akhir || '';
        
        // C.04 Penangkaran
        case 'Nama Penangkar': return fields.c04_penangkar || fields.c05_penangkar || fields.c09_penangkar || '';
        case 'Lokasi Unit Penangkaran': return fields.c04_lokasi || '';
        
        // C.05 Jenis ditangkarkan
        case 'Hasil Penangkaran': return fields.c05_hasil_breeding || '';
        case 'Hasil Pemanfaatan Tahun ini': return fields.c05_pemanfaatan || '';
        case 'Sisa Stok Hasil Pemanfaatan': return fields.c05_sisa || '';
        
        // C.06 Pengedar
        case 'Nama Pengedar': return fields.c06_pengedar || '';
        case 'Kategori': return fields.c06_kategori || '';
        case 'Lokasi Unit Pengedar': return fields.c06_lokasi || '';
        
        // C.08 Realisasi Tangkap
        case 'Jenis': return fields.c08_jenis || fields.c09_jenis || fields.c10_jenis || '';
        case 'Provinsi': return fields.c08_provinsi || '';
        case 'Volume Realisasi Tangkap': return fields.c08_volume || '';
        case 'Satuan Realisasi Tangkap': return fields.c08_satuan || '';
        
        // C.09 & C.10 Ekspor
        case 'Nama Pengedar TSL': return fields.c10_pengedar || '';
        case 'Negara': return fields.c09_negara || fields.c10_negara || '';
        case 'Volume Kuota Ekspor': return fields.c09_kuota || fields.c10_kuota || '';
        case 'Satuan Kuota Ekspor': return fields.c09_kuota_satuan || fields.c10_kuota_satuan || '';
        case 'Volume Realisasi Ekspor': return fields.c09_realisasi || fields.c10_realisasi || '';
        case 'Satuan Realisasi Ekspor': return fields.c09_realisasi_satuan || fields.c10_realisasi_satuan || '';
        
        // C.11 Rekap Sitaan
        case 'Asal-usul': return fields.c11_asal || fields.c15_asal || '';
        
        // C.12 PNBP
        case 'Akun PNBP': return fields.c12_akun || '';
        case 'Realisasi Bulan ini': return fields.c12_realisasi || '';
        
        // C.14 Interaksi Negatif
        case 'Tanggal Kejadian': return fields.c14_tanggal || '';
        case 'Korban Manusia Meninggal': return fields.c14_mati || '0';
        case 'Korban Manusia Cedera': return fields.c14_luka || '0';
        case 'Kerusakan Kebun/Lahan Masy (m2)': return fields.c14_kebun || '0';
        case 'Kerusakan Bangunan': return fields.c14_bangunan || '0';
        case 'Korban Kambing': 
          // Parsing string hewan peliharaan warga jika diinput
          return fields.c14_ternak || '0';
        case 'Korban Sapi': return '0';
        case 'Korban Kerbau': return '0';
        case 'Korban Anjing': return '0';
        case 'Korban Babi': return '0';
        case 'Korban Unggas': return '0';
        case 'Korban Satwa Mati': return fields.c14_satwa_mati || '0';
        case 'Taksiran Kerugian (Rp)': return fields.c14_rugi || '0';
        case 'Upaya Penanggulangan': return fields.c14_upaya || '';
        
        // C.16 Kelahiran
        case 'Tanggal Kelahiran': return fields.c16_tanggal || '';
        case 'Usulan Nama Satwa': return fields.c16_nama_usul || '';
        
        // C.17 Kematian
        case 'Tanggal Kematian': return fields.c17_tanggal || '';
        case 'Perkiraan Usia Satwa': return fields.c17_usia || '';
        case 'Penyebab Kematian': return fields.c17_penyebab || '';
        case 'Upaya Penanganan': return fields.c17_upaya || '';

        case 'Keterangan': return fields.keterangan_val || '';
        default: return fields[header] || '';
      }
    });

    if (isEdit) {
      // Perbarui baris yang ada di spreadsheet
      const data = sheet.getDataRange().getValues();
      let rowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === uniqueId) {
          rowIndex = i + 1;
          break;
        }
      }
      if (rowIndex !== -1) {
        sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
        return { success: true, message: "Laporan Modul " + moduleCode + " sukses diperbarui!" };
      } else {
        throw new Error("ID data '" + uniqueId + "' tidak ditemukan di database.");
      }
    } else {
      // Tambah baris baru
      sheet.appendRow(rowValues);
      return { success: true, message: "Sukses! Laporan Modul " + moduleCode + " terekam di database." };
    }
  } catch(e) {
    return { success: false, message: "Aktivitas penyimpanan gagal: " + e.toString() };
  }
}

// Fungsi utama menghapus data berdasarkan ID Modul (MEMPERBAIKI ERROR TIDAK BISA MENGHAPUS)
function deleteConservationData(moduleCode, id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(moduleCode);
    if (!sheet) {
      throw new Error("Sheet modul '" + moduleCode + "' tidak ditemukan.");
    }
    
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    // Cari baris yang memiliki ID yang sesuai pada kolom ke-1 (indeks 0)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        rowIndex = i + 1; // Konversi ke baris indeks 1-based di Google Sheets
        break;
      }
    }
    
    if (rowIndex !== -1) {
      sheet.deleteRow(rowIndex);
      return { success: true, message: "Data dengan ID " + id + " sukses dihapus dari database." };
    } else {
      return { success: false, message: "Gagal menghapus: Data tidak ditemukan." };
    }
  } catch(e) {
    return { success: false, message: "Gagal menghapus data: " + e.toString() };
  }
}