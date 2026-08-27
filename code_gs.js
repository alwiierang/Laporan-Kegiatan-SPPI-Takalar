// KONFIGURASI KUNCI - Sesuai dengan data Anda
var SPREADSHEET_ID = '1ZPgwss0Sc2gMB07te3q47NSUjA4l3gQcha-bYauXllo';
var FOLDER_ID = '1LVCWeBKVMNekgrezrCrJ7kYePCaax_3-';
var EMAIL_TUJUAN = 'alwiierang@gmail.com';

/**
 * Fungsi doGet: Menangani saat Web App diakses / di-embed
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Aplikasi Pelaporan Briefing SPPG Takalar Aktif')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // Penting untuk Embed Canva
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Fungsi utama untuk memproses laporan dari frontend HTML
 */
function processForm(formData) {
  var lock = LockService.getScriptLock();
  // Menunggu hingga 10 detik agar antrean pengiriman tidak bentrok
  lock.waitLock(10000); 

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var folder = DriveApp.getFolderById(FOLDER_ID);
    
    var cleanName = formData.nama ? formData.nama.replace(/[^a-zA-Z0-9]/g, '_') : 'Relawan';
    var fileName = 'Bukti_Briefing_' + cleanName + '_' + new Date().getTime() + '.jpg';
    var fileUrl = 'Tidak Ada Foto';
    var photoBlob = null;

    // 1. Proses Upload Foto jika ada
    if (formData.fileData) {
      var contentType = formData.mimeType || 'image/jpeg';
      var decodedData = Utilities.base64Decode(formData.fileData);
      
      // Validasi Ukuran (Maksimal ~10MB)
      if (decodedData.length > 10 * 1024 * 1024) {
        throw new Error("Ukuran foto terlalu besar. Maksimal 10MB.");
      }
      
      photoBlob = Utilities.newBlob(decodedData, contentType, fileName);
      var file = folder.createFile(photoBlob);
      // Buat file publik agar bisa dilihat dari link
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }

    // 2. Kirim Email Notifikasi
    try {
      var subject = '[Laporan Briefing SPPG] ' + (formData.sppg || 'SPPG') + ' - ' + (formData.nama || 'Relawan');
      var body = "Detail Laporan Briefing:\n\n" +
                 "• Tanggal Briefing: " + formData.tanggal + " " + formData.waktu + "\n" +
                 "• Email Pelapor: " + formData.email + "\n" +
                 "• Nama Pelapor: " + formData.nama + "\n" +
                 "• Lokasi SPPG: " + formData.sppg + "\n" +
                 "• Status Operasional: " + formData.status + "\n" +
                 "• Divisi Briefing: " + formData.divisi + "\n" +
                 "• Catatan: " + (formData.catatan || '-') + "\n" +
                 "• Link Google Drive Foto: " + fileUrl + "\n\n" +
                 "Bukti foto briefing juga turut dilampirkan pada email ini (jika ada).";
      
      var mailOptions = {};
      if (photoBlob) {
        mailOptions.attachments = [photoBlob];
      }
      GmailApp.sendEmail(EMAIL_TUJUAN, subject, body, mailOptions);
    } catch (mailErr) {
      Logger.log("Gagal Mengirim Email: " + mailErr.toString());
    }

    // 3. Menyiapkan Sheet berdasarkan format Tanggal (DD/MM/YY)
    var targetDate = new Date(formData.tanggal);
    if (isNaN(targetDate.getTime())) {
      targetDate = new Date(); // Fallback ke hari ini jika input salah
    }
    
    // Format tanggal ke zona waktu WITA (Asia/Makassar)
    var dateString = Utilities.formatDate(targetDate, "Asia/Makassar", "dd/MM/yy");
    var timestamp = Utilities.formatDate(new Date(), "Asia/Makassar", "yyyy-MM-dd HH:mm:ss");

    var sheet = ss.getSheetByName(dateString);
    if (!sheet) {
      // Buat sheet baru jika belum ada
      sheet = ss.insertSheet(dateString);
      sheet.appendRow([
        "Timestamp_WITA", 
        "Tanggal_Briefing",
        "Waktu_Briefing",
        "Email_Address", 
        "Nama_Lengkap", 
        "Nama_SPPG", 
        "Status_Operasional", 
        "Divisi_yang_di_Briefing",
        "Catatan",
        "Upload_Bukti_Briefing"
      ]);
      sheet.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#f3f4f6");
    }

    // 4. Masukkan data ke dalam Spreadsheet
    sheet.appendRow([
      timestamp,
      formData.tanggal,
      formData.waktu,
      formData.email,
      formData.nama,
      formData.sppg,
      formData.status,
      formData.divisi,
      formData.catatan,
      fileUrl
    ]);

    return { success: true, message: 'Laporan berhasil diarsipkan!' };

  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function tesKoneksiSistem() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
  var folder = DriveApp.getFolderById(FOLDER_ID);
  Logger.log("Koneksi Berhasil! Spreadsheet: " + ss.getName() + " | Folder Drive: " + folder.getName());
}