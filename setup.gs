var SHEET_NAME = 'BPM_Tasks';

/**
 * Inisialisasi Database Non-Destruktif (Safe Migrate)
 * Membuat sheet jika belum ada dan mengisi data dummy awal tanpa merusak data yang ada.
 */
function setupDatabase() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    var headers = ['ID_Task', 'Task_Name', 'Description', 'Category', 'Status', 'Order_Index', 'Created_At'];
    
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#151b26').setFontColor('#00ffaa');
      sheet.setFrozenRows(1);
      
      var now = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
      var dummyData = [
        ['BPM-0001', 'Mapping SOP Customer Onboarding', 'Membuat alur proses dari registrasi hingga aktivasi akun client', 'Process Mapping', 'Active', 1, now],
        ['BPM-0002', 'Audit Sistem Core Banking v2', 'Menganalisis bottleneck transaksi pada jam sibuk', 'System Audit', 'Active', 2, now],
        ['BPM-0003', 'Review SLA Client Enterprise X', 'Verifikasi kesesuaian response time dengan kontrak layanan', 'Client Review', 'Completed', 3, now],
        ['BPM-0004', 'Internal BPM Weekly Sync', 'Review KPI bulanan dan optimasi alur kerja internal', 'Internal BPM', 'Active', 4, now]
      ];
      
      sheet.getRange(2, 1, dummyData.length, headers.length).setValues(dummyData);
      SpreadsheetApp.flush();
      return { success: true, message: 'Database BPM_Tasks berhasil dibuat beserta data awal!' };
    } else {
      var existingHeadersRange = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1);
      var existingHeaders = existingHeadersRange.getDisplayValues()[0];
      
      if (existingHeaders.length === 0 || existingHeaders[0] === '') {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#151b26').setFontColor('#00ffaa');
        sheet.setFrozenRows(1);
      }
      SpreadsheetApp.flush();
      return { success: true, message: 'Database BPM_Tasks sudah ada. Struktur dipastikan aman (Safe Migrate).' };
    }
  } catch (error) {
    Logger.log('Error setupDatabase: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}
