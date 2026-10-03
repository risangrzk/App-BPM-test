var SHEET_NAME = 'BPM_Tasks';

/**
 * Melayani antarmuka web HTML
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('BPM Task Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper untuk mengambil sheet database BPM
 */
function getDbSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName(SHEET_NAME);
  }
  return sheet;
}

/**
 * Mengambil daftar tugas dengan Server-Side Filtering & Pagination (Anti-Lag)
 */
function getTasksBackend(params) {
  try {
    params = params || {};
    var statusFilter = params.status || 'All'; // 'All', 'Active', 'Completed'
    var categoryFilter = params.category || 'All';
    var searchKeyword = (params.search || '').toString().toLowerCase().trim();
    var page = parseInt(params.page || 1, 10);
    var limit = parseInt(params.limit || 50, 10);

    var sheet = getDbSheet();
    var lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { success: true, tasks: [], totalCount: 0, categories: [] };
    }

    var range = sheet.getRange(2, 1, lastRow - 1, 7);
    var values = range.getDisplayValues(); // String safe formatting

    var filteredTasks = [];
    var categoriesMap = {};

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var taskId = row[0];
      var taskName = row[1];
      var description = row[2];
      var category = row[3];
      var status = row[4];
      var orderIndex = parseInt(row[5] || (i + 1), 10);
      var createdAt = row[6];

      if (category) {
        categoriesMap[category] = true;
      }

      // Filter Status
      if (statusFilter !== 'All' && status !== statusFilter) {
        continue;
      }

      // Filter Kategori
      if (categoryFilter !== 'All' && category !== categoryFilter) {
        continue;
      }

      // Filter Pencarian
      if (searchKeyword !== '') {
        var matchName = taskName.toLowerCase().indexOf(searchKeyword) !== -1;
        var matchDesc = description.toLowerCase().indexOf(searchKeyword) !== -1;
        var matchId = taskId.toLowerCase().indexOf(searchKeyword) !== -1;
        if (!matchName && !matchDesc && !matchId) {
          continue;
        }
      }

      filteredTasks.push({
        id: taskId,
        name: taskName,
        description: description,
        category: category,
        status: status,
        orderIndex: orderIndex,
        createdAt: createdAt
      });
    }

    filteredTasks.sort(function(a, b) {
      return a.orderIndex - b.orderIndex;
    });

    // Pagination
    var totalCount = filteredTasks.length;
    var startIndex = (page - 1) * limit;
    var paginatedTasks = filteredTasks.slice(startIndex, startIndex + limit);

    var categoriesList = Object.keys(categoriesMap);

    return {
      success: true,
      tasks: paginatedTasks,
      totalCount: totalCount,
      categories: categoriesList
    };
  } catch (error) {
    Logger.log('Error getTasksBackend: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Membuat Tugas Baru dengan ID Sekuensial BPM-xxxx
 */
function addTaskBackend(taskData) {
  try {
    var sheet = getDbSheet();
    var lastRow = sheet.getLastRow();
    var values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 6).getDisplayValues() : [];

    // Logika ID Sekuensial (BPM-0001)
    var maxIdNum = 0;
    var maxOrder = 0;

    for (var i = 0; i < values.length; i++) {
      var curId = values[i][0];
      var curOrder = parseInt(values[i][5] || 0, 10);
      if (curOrder > maxOrder) {
        maxOrder = curOrder;
      }
      if (curId && curId.indexOf('BPM-') === 0) {
        var numPart = parseInt(curId.replace('BPM-', ''), 10);
        if (!isNaN(numPart) && numPart > maxIdNum) {
          maxIdNum = numPart;
        }
      }
    }

    var nextNum = maxIdNum + 1;
    var formattedNum = ('000' + nextNum).slice(-4);
    var newId = 'BPM-' + formattedNum;

    var createdAt = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var newOrderIndex = maxOrder + 1;

    var newRow = [
      newId,
      String(taskData.name || '').trim(),
      String(taskData.description || '').trim(),
      String(taskData.category || 'General').trim(),
      'Active',
      newOrderIndex,
      createdAt
    ];

    sheet.appendRow(newRow);
    SpreadsheetApp.flush(); // Pastikan sinkronisasi instan

    return {
      success: true,
      task: {
        id: newId,
        name: newRow[1],
        description: newRow[2],
        category: newRow[3],
        status: newRow[4],
        orderIndex: newOrderIndex,
        createdAt: createdAt
      }
    };
  } catch (error) {
    Logger.log('Error addTaskBackend: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Mengubah data Tugas eksisting
 */
function updateTaskBackend(taskData) {
  try {
    var sheet = getDbSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: 'Tugas tidak ditemukan' };

    var range = sheet.getRange(2, 1, lastRow - 1, 1);
    var ids = range.getDisplayValues();
    var rowIndex = -1;

    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] === taskData.id) {
        rowIndex = i + 2;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: 'ID Tugas tidak ditemukan' };
    }

    sheet.getRange(rowIndex, 2).setValue(String(taskData.name).trim());
    sheet.getRange(rowIndex, 3).setValue(String(taskData.description).trim());
    sheet.getRange(rowIndex, 4).setValue(String(taskData.category).trim());

    SpreadsheetApp.flush();

    return { success: true };
  } catch (error) {
    Logger.log('Error updateTaskBackend: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Mengubah status Tugas (Active <-> Completed)
 */
function toggleTaskStatusBackend(taskId, newStatus) {
  try {
    var sheet = getDbSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: 'Data kosong' };

    var range = sheet.getRange(2, 1, lastRow - 1, 1);
    var ids = range.getDisplayValues();
    var rowIndex = -1;

    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] === taskId) {
        rowIndex = i + 2;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: 'ID Tugas tidak ditemukan' };
    }

    sheet.getRange(rowIndex, 5).setValue(newStatus);
    SpreadsheetApp.flush();

    return { success: true };
  } catch (error) {
    Logger.log('Error toggleTaskStatusBackend: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Menghapus tugas dari sheet
 */
function deleteTaskBackend(taskId) {
  try {
    var sheet = getDbSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: 'Data kosong' };

    var range = sheet.getRange(2, 1, lastRow - 1, 1);
    var ids = range.getDisplayValues();
    var rowIndex = -1;

    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] === taskId) {
        rowIndex = i + 2;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: 'ID Tugas tidak ditemukan' };
    }

    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();

    return { success: true };
  } catch (error) {
    Logger.log('Error deleteTaskBackend: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Menyimpan urutan baru hasil Drag-and-Drop (Batch Processing)
 */
function reorderTasksBackend(orderedIds) {
  try {
    if (!orderedIds || !orderedIds.length) return { success: true };

    var sheet = getDbSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true };

    var idsRange = sheet.getRange(2, 1, lastRow - 1, 1);
    var ids = idsRange.getDisplayValues();

    var orderMap = {};
    for (var k = 0; k < orderedIds.length; k++) {
      orderMap[orderedIds[k]] = k + 1;
    }

    var orderColumnValues = [];
    for (var i = 0; i < ids.length; i++) {
      var id = ids[i][0];
      if (orderMap.hasOwnProperty(id)) {
        orderColumnValues.push([orderMap[id]]);
      } else {
        orderColumnValues.push([999 + i]);
      }
    }

    sheet.getRange(2, 6, orderColumnValues.length, 1).setValues(orderColumnValues);
    SpreadsheetApp.flush();

    return { success: true };
  } catch (error) {
    Logger.log('Error reorderTasksBackend: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}
