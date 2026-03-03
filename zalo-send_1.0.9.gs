// ==========================================
// ⚙️ CẤU HÌNH HỆ THỐNG
// ==========================================
const SYSTEM_CONFIG = {
  SHEETS: {
    LEADS: "lead-mkt",
    QUEUE: "Zalo_Queue",
    RVA_CONFIG: "RVA_Config",
    LOG: "System_Log",
    MEMORY: "System_Memory",
    LEAD_HISTORY: "Lead_History"
  },
  LEADS_CONFIG: {
    START_ROW: 5,
    COL_STT: 1,
    COL_LEAD_ID: 2,
    COL_NAME: 3,
    COL_PHONE: 4,
    COL_NEED: 5,
    COL_PROJECT: 6,
    COL_SOURCE: 7,
    RVA_START_COL: 8
  },
  QUEUE_CONFIG: {
    COL_TIMESTAMP: 1,
    COL_RVA_ID: 2,
    COL_RVA_NAME: 3,
    COL_ZALO_PHONE: 4,
    COL_ZALO_ID: 5,
    COL_LEAD_ID: 6,
    COL_LEAD_NAME: 7,
    COL_LEAD_PHONE: 8,
    COL_NEED: 9,
    COL_PROJECT: 10,
    COL_MESSAGE: 11,
    COL_STATUS: 12,
    COL_ORIGINAL_ROW: 13,
    COL_COLUMN_INDEX: 14,
    COL_ERROR: 15,
    COL_SENT_TIME: 16,
    COL_QUEUE_ROW: 17
  },
  N8N_CONFIG: {
    WEBHOOK_URL: 'https://n8n.rever.io.vn/webhook/zalo-trigger',
    RESULT_COL: 18,
    QUEUE_ROW_COL: 17
  }
};

// ==========================================
// 🗂️ QUẢN LÝ BỘ NHỚ TRẠNG THÁI
// ==========================================
function getSystemMemory() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let memorySheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MEMORY);
    
    if (!memorySheet) {
      memorySheet = ss.insertSheet(SYSTEM_CONFIG.SHEETS.MEMORY);
      memorySheet.appendRow(["Key", "Value", "Description", "Last Updated"]);
      memorySheet.getRange("1:1").setFontWeight("bold");
      memorySheet.getRange("1:1").setBackground("#667eea");
      memorySheet.getRange("1:1").setFontColor("#ffffff");
      
      memorySheet.appendRow(["lastProcessedLeadId", "", "Lead ID cuối cùng đã xử lý vào Queue", new Date()]);
      memorySheet.appendRow(["nextRvaIndex", 0, "Index của RVA tiếp theo (0-based)", new Date()]);
      memorySheet.appendRow(["lastQueueRow", 1, "Row cuối cùng trong Queue (bao gồm header)", new Date()]);
      
      memorySheet.setColumnWidth(1, 200);
      memorySheet.setColumnWidth(2, 150);
      memorySheet.setColumnWidth(3, 300);
      memorySheet.setColumnWidth(4, 200);
      
      logSystem("Đã tạo sheet Memory mới với giá trị mặc định");
      
      return {
        lastProcessedLeadId: "",
        nextRvaIndex: 0,
        lastQueueRow: 1
      };
    }
    
    const data = memorySheet.getRange(2, 1, 3, 2).getValues();
    
    const memory = {
      lastProcessedLeadId: String(data[0][1] || "").trim(),
      nextRvaIndex: parseInt(data[1][1]) || 0,
      lastQueueRow: parseInt(data[2][1]) || 1
    };
    
    logSystem("Đọc Memory: lastLeadId=" + memory.lastProcessedLeadId + 
             ", nextRVA=" + memory.nextRvaIndex + 
             ", lastQueueRow=" + memory.lastQueueRow);
    
    return memory;
    
  } catch (error) {
    logSystem("Lỗi getSystemMemory: " + error.toString(), "ERROR");
    return {
      lastProcessedLeadId: "",
      nextRvaIndex: 0,
      lastQueueRow: 1
    };
  }
}

function updateSystemMemory(memory) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let memorySheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MEMORY);
    
    if (!memorySheet) {
      getSystemMemory();
      memorySheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.MEMORY);
    }
    
    const timestamp = new Date();
    
    if (memory.hasOwnProperty('lastProcessedLeadId')) {
      memorySheet.getRange(2, 2).setValue(memory.lastProcessedLeadId);
      memorySheet.getRange(2, 4).setValue(timestamp);
    }
    
    if (memory.hasOwnProperty('nextRvaIndex')) {
      memorySheet.getRange(3, 2).setValue(memory.nextRvaIndex);
      memorySheet.getRange(3, 4).setValue(timestamp);
    }
    
    if (memory.hasOwnProperty('lastQueueRow')) {
      memorySheet.getRange(4, 2).setValue(memory.lastQueueRow);
      memorySheet.getRange(4, 4).setValue(timestamp);
    }
    
    logSystem("Cập nhật Memory: " + JSON.stringify(memory));
    
  } catch (error) {
    logSystem("Lỗi updateSystemMemory: " + error.toString(), "ERROR");
  }
}

function resetSystemMemory() {
  try {
    updateSystemMemory({
      lastProcessedLeadId: "",
      nextRvaIndex: 0,
      lastQueueRow: 1
    });
    
    logSystem("Đã reset System Memory về mặc định", "SUCCESS");
    
    return {
      success: true,
      message: "✅ Đã reset bộ nhớ hệ thống thành công!"
    };
    
  } catch (error) {
    logSystem("Lỗi resetSystemMemory: " + error.toString(), "ERROR");
    return {
      success: false,
      message: "❌ Lỗi: " + error.toString()
    };
  }
}

// ==========================================
// 🔖 TẠO LEAD ID TỰ ĐỘNG
// ==========================================
function generateLeadId(rowIndex) {
  const idNumber = String(rowIndex).padStart(2, '0');
  return 'le' + idNumber;
}

function autoGenerateLeadIds() {
  try {
    logSystem("Bắt đầu tạo Lead ID tự động");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    
    if (!leadSheet) {
      return {
        success: false,
        message: '❌ Không tìm thấy sheet "' + SYSTEM_CONFIG.SHEETS.LEADS + '"'
      };
    }
    
    const lastRow = leadSheet.getLastRow();
    const startRow = SYSTEM_CONFIG.LEADS_CONFIG.START_ROW;
    
    if (lastRow < startRow) {
      return {
        success: false,
        message: "⚠️ Không có dữ liệu lead"
      };
    }
    
    const nameData = leadSheet.getRange(startRow, SYSTEM_CONFIG.LEADS_CONFIG.COL_NAME, lastRow - startRow + 1, 1).getValues();
    const leadIdData = leadSheet.getRange(startRow, SYSTEM_CONFIG.LEADS_CONFIG.COL_LEAD_ID, lastRow - startRow + 1, 1).getValues();
    
    let createdCount = 0;
    let skippedCount = 0;
    
    let maxNumber = 0;
    leadIdData.forEach(function(row) {
      const leadId = String(row[0]).trim();
      if (leadId.startsWith('le')) {
        const num = parseInt(leadId.substring(2));
        if (!isNaN(num) && num > maxNumber) {
          maxNumber = num;
        }
      }
    });
    
    let currentNumber = maxNumber;
    
    for (let i = 0; i < nameData.length; i++) {
      const fullName = nameData[i][0];
      const leadId = String(leadIdData[i][0]).trim();
      
      if (fullName && fullName !== "") {
        if (leadId === "" || leadId === "null" || leadId === "undefined") {
          currentNumber++;
          const newLeadId = 'le' + String(currentNumber).padStart(2, '0');
          leadSheet.getRange(startRow + i, SYSTEM_CONFIG.LEADS_CONFIG.COL_LEAD_ID).setValue(newLeadId);
          createdCount++;
        } else {
          skippedCount++;
        }
      }
    }
    
    logSystem('✅ Đã tạo ' + createdCount + ' Lead ID mới', "SUCCESS");
    
    return {
      success: true,
      message: '✅ Tạo Lead ID hoàn tất!\n\n' +
              '📝 Đã tạo: ' + createdCount + ' ID mới\n' +
              '✓ Đã có sẵn: ' + skippedCount + ' ID\n' +
              '➡️ ID tiếp theo sẽ là: le' + String(currentNumber + 1).padStart(2, '0')
    };
    
  } catch (error) {
    logSystem('❌ Lỗi tạo Lead ID: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: '❌ Lỗi: ' + error.toString()
    };
  }
}

// ==========================================
// 📊 API: LẤY CẤU HÌNH RVA
// ==========================================
function getRVAConfig() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.RVA_CONFIG);
    
    if (!configSheet) {
      return [
        {
          id: "RV001",
          name: "RVA 1",
          phone: "0901234567",
          zaloid: "1243438522397465634",
          columnIndex: 8,
          active: true,
          note: ""
        },
        {
          id: "RV002",
          name: "RVA 2",
          phone: "0902345678",
          zaloid: "3837551771715024930",
          columnIndex: 11,
          active: true,
          note: ""
        },
        {
          id: "RV003",
          name: "RVA 3",
          phone: "0903456789",
          zaloid: "1504484729431570818",
          columnIndex: 14,
          active: true,
          note: ""
        }
      ];
    }
    
    const lastRow = configSheet.getLastRow();
    if (lastRow <= 1) return [];
    
    const data = configSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    
    return data.map(row => ({
      id: row[0] || "",
      name: row[1] || "",
      phone: row[2] || "",
      zaloid: row[3] || "",
      columnIndex: Number(row[4]) || 8,
      active: row[5] === true || row[5] === "TRUE" || row[5] === "true",
      note: row[6] || ""
    })).filter(rva => rva.id !== "");
    
  } catch (error) {
    Logger.log("❌ Lỗi getRVAConfig: " + error.toString());
    return [];
  }
}

// ==========================================
// 🔄 ĐỒNG BỘ CỘT RVA VÀO LEAD-MKT (CHỈ CHẠY THỦ CÔNG)
// ==========================================
function syncRVAColumnsToLeadSheet() {
  try {
    logSystem("🔄 [MANUAL] Bắt đầu đồng bộ cột RVA vào lead-mkt");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    const rvaConfig = getRVAConfig();
    
    if (!leadSheet) {
      return {
        success: false,
        message: '❌ Không tìm thấy sheet "' + SYSTEM_CONFIG.SHEETS.LEADS + '"'
      };
    }
    
    if (rvaConfig.length === 0) {
      return {
        success: false,
        message: '⚠️ Không có RVA nào trong RVA_Config!\n\nVui lòng cấu hình RVA trước.'
      };
    }
    
    const lastCol = leadSheet.getLastColumn();
    const headers = leadSheet.getRange(4, 1, 1, lastCol).getValues()[0];
    
    const rvaStartCol = SYSTEM_CONFIG.LEADS_CONFIG.RVA_START_COL;
    
    const expectedRVAColumns = [];
    rvaConfig.forEach(function(rva) {
      expectedRVAColumns.push(rva.name + " - Status");
      expectedRVAColumns.push(rva.name + " - Assigned Time");
      expectedRVAColumns.push(rva.name + " - Care Status");
    });
    
    logSystem('📋 Cần có ' + expectedRVAColumns.length + ' cột RVA (' + rvaConfig.length + ' RVA x 3 cột)');
    
    const currentRVAColumns = headers.slice(rvaStartCol - 1);
    
    let addedCount = 0;
    let updatedCount = 0;
    let currentCol = rvaStartCol + currentRVAColumns.length;
    
    expectedRVAColumns.forEach(function(colName, index) {
      const existingIndex = currentRVAColumns.indexOf(colName);
      
      if (existingIndex === -1) {
        leadSheet.insertColumnAfter(currentCol - 1);
        leadSheet.getRange(4, currentCol).setValue(colName);
        
        leadSheet.getRange(4, currentCol)
          .setFontWeight("bold")
          .setBackground("#f3f3f3")
          .setHorizontalAlignment("center")
          .setBorder(true, true, true, true, false, false);
        
        if (colName.includes("Status")) {
          leadSheet.setColumnWidth(currentCol, 120);
        } else if (colName.includes("Time")) {
          leadSheet.setColumnWidth(currentCol, 160);
        } else if (colName.includes("Care")) {
          leadSheet.setColumnWidth(currentCol, 150);
        }
        
        logSystem('➕ Thêm cột mới: ' + colName + ' (Cột ' + currentCol + ')');
        currentCol++;
        addedCount++;
      } else {
        const existingCol = rvaStartCol + existingIndex;
        leadSheet.getRange(4, existingCol)
          .setFontWeight("bold")
          .setBackground("#f3f3f3")
          .setHorizontalAlignment("center");
        
        updatedCount++;
      }
    });
    
    updateRVAColumnIndexes();
    
    logSystem('✅ Đồng bộ RVA columns hoàn tất: Thêm ' + addedCount + ' cột, Cập nhật ' + updatedCount + ' cột', "SUCCESS");
    
    return {
      success: true,
      message: '✅ Đồng bộ cột RVA hoàn tất!\n\n' +
              '📊 THỐNG KÊ:\n' +
              '• Tổng RVA: ' + rvaConfig.length + '\n' +
              '• Tổng cột cần có: ' + expectedRVAColumns.length + '\n' +
              '• Đã thêm: ' + addedCount + ' cột\n' +
              '• Đã có sẵn: ' + updatedCount + ' cột\n\n' +
              '📋 MỖI RVA CÓ 3 CỘT:\n' +
              '  1. Status (trạng thái gửi)\n' +
              '  2. Assigned Time (thời gian giao lead)\n' +
              '  3. Care Status (trạng thái chăm sóc)\n\n' +
              '✓ Column Index đã được cập nhật trong RVA_Config'
    };
    
  } catch (error) {
    logSystem('❌ Lỗi syncRVAColumnsToLeadSheet: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: '❌ Lỗi: ' + error.toString()
    };
  }
}

function updateRVAColumnIndexes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    const configSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.RVA_CONFIG);
    
    if (!leadSheet || !configSheet) {
      logSystem('⚠️ Không tìm thấy sheet cần thiết để cập nhật Column Index', "WARNING");
      return false;
    }
    
    const lastCol = leadSheet.getLastColumn();
    const headers = leadSheet.getRange(4, 1, 1, lastCol).getValues()[0];
    
    const rvaConfig = getRVAConfig();
    
    let updatedCount = 0;
    
    for (let i = 0; i < rvaConfig.length; i++) {
      const rva = rvaConfig[i];
      const statusColName = rva.name + " - Status";
      
      const colIndex = headers.indexOf(statusColName) + 1;
      
      if (colIndex > 0) {
        configSheet.getRange(i + 2, 5).setValue(colIndex);
        logSystem('✓ Cập nhật Column Index cho ' + rva.name + ': Cột ' + colIndex);
        updatedCount++;
      } else {
        logSystem('⚠️ Không tìm thấy cột Status cho ' + rva.name, "WARNING");
      }
    }
    
    logSystem('✅ Đã cập nhật Column Index cho ' + updatedCount + '/' + rvaConfig.length + ' RVA');
    return true;
    
  } catch (error) {
    logSystem('❌ Lỗi updateRVAColumnIndexes: ' + error.toString(), "ERROR");
    return false;
  }
}

// ==========================================
// 🔍 KIỂM TRA CẤU TRÚC CỘT RVA
// ==========================================
function checkRVAColumnsStatus() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    const rvaConfig = getRVAConfig();
    
    if (!leadSheet) {
      return {
        success: false,
        message: '❌ Không tìm thấy sheet "lead-mkt"'
      };
    }
    
    const lastCol = leadSheet.getLastColumn();
    const headers = leadSheet.getRange(4, 1, 1, lastCol).getValues()[0];
    const rvaStartCol = SYSTEM_CONFIG.LEADS_CONFIG.RVA_START_COL;
    
    let report = '📊 BÁO CÁO CẤU TRÚC CỘT RVA\n';
    report += '==========================================\n\n';
    
    report += '📋 TỔNG QUAN:\n';
    report += '• Tổng RVA trong config: ' + rvaConfig.length + '\n';
    report += '• Tổng cột trong sheet: ' + lastCol + '\n';
    report += '• Cột bắt đầu RVA: ' + rvaStartCol + '\n\n';
    
    report += '👥 CHI TIẾT TỪNG RVA:\n';
    
    rvaConfig.forEach(function(rva, index) {
      report += '\n' + (index + 1) + '. ' + rva.name + ' (ID: ' + rva.id + ')\n';
      report += '   Active: ' + (rva.active ? 'Có ✓' : 'Không ✗') + '\n';
      
      const statusCol = rva.name + " - Status";
      const timeCol = rva.name + " - Assigned Time";
      const careCol = rva.name + " - Care Status";
      
      const hasStatus = headers.indexOf(statusCol) !== -1;
      const hasTime = headers.indexOf(timeCol) !== -1;
      const hasCare = headers.indexOf(careCol) !== -1;
      
      report += '   Cột Status: ' + (hasStatus ? 'Có ✓' : 'Thiếu ✗') + '\n';
      report += '   Cột Time: ' + (hasTime ? 'Có ✓' : 'Thiếu ✗') + '\n';
      report += '   Cột Care: ' + (hasCare ? 'Có ✓' : 'Thiếu ✗') + '\n';
      
      if (hasStatus) {
        const colIndex = headers.indexOf(statusCol) + 1;
        report += '   → Column Index: ' + colIndex + '\n';
        
        if (colIndex !== rva.columnIndex) {
          report += '   ⚠️ Chú ý: Column Index trong config (' + rva.columnIndex + ') khác với thực tế (' + colIndex + ')\n';
        }
      }
    });
    
    report += '\n==========================================';
    
    logSystem('✅ Đã tạo báo cáo cấu trúc cột RVA');
    
    return {
      success: true,
      message: report
    };
    
  } catch (error) {
    return {
      success: false,
      message: '❌ Lỗi: ' + error.toString()
    };
  }
}

// ==========================================
// ⏰ CẬP NHẬT THỜI GIAN & TRẠNG THÁI (CẢI TIẾN)
// ==========================================
function updateLeadAssignment(leadId, rvaId, status, careStatus) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    const rvaConfig = getRVAConfig();
    
    const rva = rvaConfig.find(function(r) { return r.id === rvaId; });
    if (!rva) {
      logSystem('❌ Không tìm thấy RVA: ' + rvaId, "ERROR");
      return false;
    }
    
    const lastRow = leadSheet.getLastRow();
    const startRow = SYSTEM_CONFIG.LEADS_CONFIG.START_ROW;
    const leadIdData = leadSheet.getRange(startRow, SYSTEM_CONFIG.LEADS_CONFIG.COL_LEAD_ID, 
                                          lastRow - startRow + 1, 1).getValues();
    
    let leadRow = -1;
    for (let i = 0; i < leadIdData.length; i++) {
      if (String(leadIdData[i][0]).trim() === String(leadId).trim()) {
        leadRow = startRow + i;
        break;
      }
    }
    
    if (leadRow === -1) {
      logSystem('❌ Không tìm thấy lead: ' + leadId, "ERROR");
      return false;
    }
    
    const statusCol = rva.columnIndex;
    const timeCol = statusCol + 1;
    const careStatusCol = statusCol + 2;
    
    // ✅ KIỂM TRA STATUS HIỆN TẠI
    const currentStatus = leadSheet.getRange(leadRow, statusCol).getValue();
    
    // ✅ NẾU ĐÃ LÀ SUCCESS/ERROR RỒI → KHÔNG CẬP NHẬT LẠI
    if (currentStatus === "Success" || currentStatus === "Error") {
      logSystem('⏭️ Lead ' + leadId + ' đã có status: ' + currentStatus + ', bỏ qua cập nhật', "INFO");
      return false;
    }
    
    // ✅ CẬP NHẬT STATUS MỚI
    if (status) {
      leadSheet.getRange(leadRow, statusCol).setValue(status);
    }
    
    // ✅ CHỈ CẬP NHẬT THỜI GIAN LẦN ĐẦU (KHI CHƯA CÓ)
    const currentTime = leadSheet.getRange(leadRow, timeCol).getValue();
    if (!currentTime || currentTime === "") {
      leadSheet.getRange(leadRow, timeCol).setValue(new Date());
    }
    
    // ✅ CẬP NHẬT CARE STATUS
    if (careStatus) {
      leadSheet.getRange(leadRow, careStatusCol).setValue(careStatus);
    }
    
    logSystem('✅ Cập nhật lead ' + leadId + ' cho RVA ' + rva.name + ': ' + status);
    return true;
    
  } catch (error) {
    logSystem('❌ Lỗi updateLeadAssignment: ' + error.toString(), "ERROR");
    return false;
  }
}

// ==========================================
// 🔄 CHUYỂN LEAD GIỮA CÁC RVA
// ==========================================
function logLeadTransfer(leadId, fromRvaId, toRvaId, reason, previousStatus, previousCareStatus) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let historySheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEAD_HISTORY);
    
    if (!historySheet) {
      historySheet = ss.insertSheet(SYSTEM_CONFIG.SHEETS.LEAD_HISTORY);
      historySheet.appendRow([
        "Timestamp",
        "Lead ID",
        "From RVA",
        "To RVA",
        "Reason",
        "Previous Status",
        "Previous Care Status",
        "Transfer By"
      ]);
      
      historySheet.getRange("1:1").setFontWeight("bold");
      historySheet.getRange("1:1").setBackground("#667eea");
      historySheet.getRange("1:1").setFontColor("#ffffff");
    }
    
    historySheet.appendRow([
      new Date(),
      leadId,
      fromRvaId,
      toRvaId,
      reason || "Chuyển lead",
      previousStatus || "",
      previousCareStatus || "",
      Session.getActiveUser().getEmail()
    ]);
    
    logSystem('📝 Đã log chuyển lead: ' + leadId + ' từ ' + fromRvaId + ' → ' + toRvaId);
    
  } catch (error) {
    logSystem('❌ Lỗi logLeadTransfer: ' + error.toString(), "ERROR");
  }
}

function transferLeadToNewRVA(leadId, fromRvaId, toRvaId, reason) {
  try {
    logSystem('🔄 Bắt đầu chuyển lead ' + leadId + ' từ ' + fromRvaId + ' → ' + toRvaId);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    let queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    const rvaConfig = getRVAConfig();
    
    const fromRVA = rvaConfig.find(function(r) { return r.id === fromRvaId; });
    const toRVA = rvaConfig.find(function(r) { return r.id === toRvaId; });
    
    if (!fromRVA || !toRVA) {
      return {
        success: false,
        message: '❌ Không tìm thấy RVA'
      };
    }
    
    const lastRow = leadSheet.getLastRow();
    const startRow = SYSTEM_CONFIG.LEADS_CONFIG.START_ROW;
    const data = leadSheet.getRange(startRow, 1, lastRow - startRow + 1, 
                                    leadSheet.getLastColumn()).getValues();
    
    let leadRow = -1;
    let leadData = null;
    
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][SYSTEM_CONFIG.LEADS_CONFIG.COL_LEAD_ID - 1]).trim() === String(leadId).trim()) {
        leadRow = startRow + i;
        leadData = data[i];
        break;
      }
    }
    
    if (leadRow === -1) {
      return {
        success: false,
        message: '❌ Không tìm thấy lead: ' + leadId
      };
    }
    
    const oldStatusCol = fromRVA.columnIndex;
    const oldTimeCol = oldStatusCol + 1;
    const oldCareStatusCol = oldStatusCol + 2;
    
    const oldStatus = leadSheet.getRange(leadRow, oldStatusCol).getValue();
    const oldTime = leadSheet.getRange(leadRow, oldTimeCol).getValue();
    const oldCareStatus = leadSheet.getRange(leadRow, oldCareStatusCol).getValue();
    
    logLeadTransfer(leadId, fromRvaId, toRvaId, reason, oldStatus, oldCareStatus);
    
    leadSheet.getRange(leadRow, oldStatusCol).setValue("");
    leadSheet.getRange(leadRow, oldTimeCol).setValue("");
    leadSheet.getRange(leadRow, oldCareStatusCol).setValue("");
    
    const leadName = leadData[SYSTEM_CONFIG.LEADS_CONFIG.COL_NAME - 1];
    const leadPhone = leadData[SYSTEM_CONFIG.LEADS_CONFIG.COL_PHONE - 1];
    const leadNeed = leadData[SYSTEM_CONFIG.LEADS_CONFIG.COL_NEED - 1];
    const leadProject = leadData[SYSTEM_CONFIG.LEADS_CONFIG.COL_PROJECT - 1];
    
    // ✅ TẠO TIN NHẮN CONTEXT ĐẦY ĐỦ (FORMAT ĐẸP HƠN)
    const contextMessage = '🔄 LEAD CHUYỂN GIAO\n' +
                          '\n' +
                          '📋 THÔNG TIN LEAD:\n' +
                          'Lead ID: ' + leadId + '\n' +
                          'Tên: ' + leadName + '\n' +
                          'SĐT: ' + leadPhone + '\n' +
                          'Nhu cầu: ' + leadNeed + '\n' +
                          'Dự án: ' + leadProject + '\n' +
                          '\n' +
                          '📊 LỊCH SỬ CHĂM SÓC:\n' +
                          'RVA trước: ' + fromRVA.name + '\n' +
                          'Thời gian nhận: ' + (oldTime ? Utilities.formatDate(new Date(oldTime), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : "N/A") + '\n' +
                          'Trạng thái: ' + (oldStatus || "N/A") + '\n' +
                          'Chăm sóc: ' + (oldCareStatus || "Chưa có") + '\n' +
                          '\n' +
                          '💡 LÝ DO CHUYỂN:\n' +
                          reason + '\n' +
                          '\n' +
                          '👉 Vui lòng tiếp tục chăm sóc lead này!';
    
    logSystem('📝 Tin nhắn context đã tạo:\n' + contextMessage);
    
    // ✅ TẠO QUEUE SHEET NẾU CHƯA CÓ
    if (!queueSheet) {
      queueSheet = ss.insertSheet(SYSTEM_CONFIG.SHEETS.QUEUE);
      queueSheet.appendRow([
        "Timestamp",
        "RVA ID",
        "RVA Name",
        "Zalo Phone",
        "Zalo ID",
        "Lead ID",
        "Lead Name",
        "Lead Phone",
        "Need",
        "Project",
        "Message",        // K
        "Status",         // L
        "Original Row",
        "Column Index",
        "Error",
        "Sent Time",
        "Queue Row"
      ]);
      
      const headerRange = queueSheet.getRange("1:1");
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#667eea");
      headerRange.setFontColor("#ffffff");
      
      logSystem("✅ Đã tạo sheet Queue mới");
    }
    
    const queueRowNumber = queueSheet.getLastRow() + 1;
    
    // ✅ THÊM VÀO QUEUE VỚI TIN NHẮN CONTEXT
    queueSheet.appendRow([
      new Date(),                    // A: Timestamp
      toRVA.id,                      // B: RVA ID
      toRVA.name,                    // C: RVA Name
      toRVA.phone,                   // D: Zalo Phone
      toRVA.zaloid,                  // E: Zalo ID
      leadId,                        // F: Lead ID
      leadName,                      // G: Lead Name
      leadPhone,                     // H: Lead Phone
      leadNeed,                      // I: Need
      leadProject,                   // J: Project
      contextMessage,                // K: Message ✅ TIN NHẮN CONTEXT
      "Pending",                     // L: Status
      leadRow,                       // M: Original Row
      toRVA.columnIndex,             // N: Column Index
      "",                            // O: Error
      "",                            // P: Sent Time
      queueRowNumber                 // Q: Queue Row
    ]);
    
    logSystem('✅ Đã thêm lead vào Queue với tin nhắn context (Row ' + queueRowNumber + ')');
    
    // ✅ CẬP NHẬT STATUS MỚI
    updateLeadAssignment(leadId, toRvaId, "Pending", "Chuyển giao - Chưa liên hệ");
    
    logSystem('✅ Chuyển lead thành công: ' + leadId + ' → ' + toRVA.name, "SUCCESS");
    
    // ✅ TÙY CHỌN: TRIGGER N8N NGAY LẬP TỨC
    const autoTrigger = true; // Đặt true nếu muốn tự động gửi ngay
    
    if (autoTrigger) {
      logSystem('🚀 Đang trigger n8n để gửi tin nhắn chuyển giao...');
      
      const n8nWebhookUrl = SYSTEM_CONFIG.N8N_CONFIG.WEBHOOK_URL;
      
      if (n8nWebhookUrl && n8nWebhookUrl.indexOf('http') === 0) {
        try {
          const payload = {
            trigger: 'lead_transfer',
            timestamp: new Date().toISOString(),
            leadId: leadId,
            fromRVA: fromRVA.name,
            toRVA: toRVA.name,
            reason: reason,
            queueRow: queueRowNumber,
            message: contextMessage
          };
          
          const options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
          };
          
          const response = UrlFetchApp.fetch(n8nWebhookUrl, options);
          const responseCode = response.getResponseCode();
          
          if (responseCode === 200) {
            logSystem('✅ n8n đã nhận webhook chuyển lead', "SUCCESS");
          } else {
            logSystem('⚠️ n8n trả về code: ' + responseCode, "WARNING");
          }
          
        } catch (webhookError) {
          logSystem('⚠️ Không thể trigger n8n: ' + webhookError.toString(), "WARNING");
        }
      }
    }
    
    return {
      success: true,
      message: '✅ Chuyển lead thành công!\n\n' +
              '📋 Lead: ' + leadId + '\n' +
              '📤 Từ: ' + fromRVA.name + '\n' +
              '📥 Sang: ' + toRVA.name + '\n\n' +
              '💬 Tin nhắn đã được thêm vào Queue (Row ' + queueRowNumber + '):\n' +
              '━━━━━━━━━━━━━━━━━━━━\n' +
              contextMessage + '\n' +
              '━━━━━━━━━━━━━━━━━━━━\n\n' +
              (autoTrigger ? '🚀 n8n đã được trigger tự động\n' : '⏳ Chờ n8n đọc queue hoặc chạy Giai đoạn 3')
    };
    
  } catch (error) {
    logSystem('❌ Lỗi transferLeadToNewRVA: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: '❌ Lỗi: ' + error.toString()
    };
  }
}

// ==========================================
// 📊 API: LẤY THỐNG KÊ HỆ THỐNG
// ==========================================
function getSystemStats() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    const queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    
    let leadStats = {
      total: 0,
      pending: 0,
      sent: 0,
      error: 0
    };
    
    if (leadSheet && leadSheet.getLastRow() >= SYSTEM_CONFIG.LEADS_CONFIG.START_ROW) {
      const lastRow = leadSheet.getLastRow();
      const startRow = SYSTEM_CONFIG.LEADS_CONFIG.START_ROW;
      
      const data = leadSheet.getRange(startRow, SYSTEM_CONFIG.LEADS_CONFIG.COL_NAME, 
                                      lastRow - startRow + 1, 
                                      leadSheet.getLastColumn() - SYSTEM_CONFIG.LEADS_CONFIG.COL_NAME + 1).getValues();
      
      data.forEach(row => {
        const fullName = row[0];
        if (fullName && fullName !== "") {
          leadStats.total++;
          
          let hasSuccess = false;
          let hasError = false;
          let hasPending = false;
          
          for (let i = SYSTEM_CONFIG.LEADS_CONFIG.RVA_START_COL - SYSTEM_CONFIG.LEADS_CONFIG.COL_NAME; i < row.length; i += 3) {
            const status = String(row[i]).trim().toLowerCase();
            if (status === "success") hasSuccess = true;
            if (status === "error") hasError = true;
            if (status === "pending") hasPending = true;
          }
          
          if (hasSuccess) {
            leadStats.sent++;
          } else if (hasError) {
            leadStats.error++;
          } else if (hasPending) {
            leadStats.pending++;
          } else {
          leadStats.pending++;
          }
        }
      });
    }
    
    let queueStats = {
      total: 0,
      pending: 0,
      sent: 0,
      error: 0
    };
    
    if (queueSheet && queueSheet.getLastRow() > 1) {
      const statusCol = SYSTEM_CONFIG.QUEUE_CONFIG.COL_STATUS;
      const data = queueSheet.getRange(2, statusCol, queueSheet.getLastRow() - 1, 1).getValues();
      
      queueStats.total = data.length;
      data.forEach(row => {
        const status = String(row[0]).trim();
        if (status === "Pending") queueStats.pending++;
        else if (status === "Sent" || status.includes("Sent")) queueStats.sent++;
        else if (status === "Error") queueStats.error++;
      });
    }
    
    const rvaConfig = getRVAConfig();
    const rvaStats = {
      total: rvaConfig.length,
      active: rvaConfig.filter(rva => rva.active).length,
      inactive: rvaConfig.filter(rva => !rva.active).length
    };
    
    const memory = getSystemMemory();
    const memoryStats = {
      lastProcessedLeadId: memory.lastProcessedLeadId || "Chưa có",
      nextRvaIndex: memory.nextRvaIndex,
      lastQueueRow: memory.lastQueueRow
    };
    
    const logSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LOG);
    let recentLogs = [];
    
    if (logSheet && logSheet.getLastRow() > 1) {
      const lastRow = logSheet.getLastRow();
      const startRow = Math.max(2, lastRow - 9);
      const logs = logSheet.getRange(startRow, 1, lastRow - startRow + 1, 3).getValues();
      
      recentLogs = logs.reverse().map(log => {
        try {
          return {
            timestamp: Utilities.formatDate(new Date(log[0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
            level: log[1] || "INFO",
            message: log[2] || ""
          };
        } catch (e) {
          return {
            timestamp: "N/A",
            level: "INFO",
            message: String(log[2] || "")
          };
        }
      });
    }
    
    return {
      success: true,
      data: {
        lead: leadStats,
        queue: queueStats,
        rva: rvaStats,
        memory: memoryStats,
        logs: recentLogs,
        timestamp: new Date().toISOString()
      }
    };
    
  } catch (error) {
    Logger.log("❌ Error in getSystemStats: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ==========================================
// 📝 LOGGING SYSTEM
// ==========================================
function logSystem(message, level) {
  level = level || "INFO";
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LOG);
    
    if (!logSheet) {
      logSheet = ss.insertSheet(SYSTEM_CONFIG.SHEETS.LOG);
      logSheet.appendRow(["Timestamp", "Level", "Message"]);
      logSheet.getRange("1:1").setFontWeight("bold");
    }
    
    logSheet.appendRow([new Date(), level, message]);
    
    if (logSheet.getLastRow() > 1001) {
      logSheet.deleteRows(2, logSheet.getLastRow() - 1001);
    }
    
    Logger.log('[' + level + '] ' + message);
    
  } catch (error) {
    Logger.log("❌ Lỗi logging: " + error.toString());
  }
}

// ==========================================
// 🎯 CONTROL PANEL
// ==========================================
function showControlPanel() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const stats = getSystemStats();
    
    if (!stats.success) {
      ui.alert('❌ Lỗi', 'Không thể lấy thông tin hệ thống:\n' + stats.error, ui.ButtonSet.OK);
      return;
    }
    
    const data = stats.data;
    
    const message = 
      '📊 THỐNG KÊ HỆ THỐNG\n' +
      '==========================================\n\n' +
      '📋 LEAD:\n' +
      '   • Tổng số: ' + data.lead.total + '\n' +
      '   • Chờ gửi: ' + data.lead.pending + '\n' +
      '   • Đã gửi: ' + data.lead.sent + '\n\n' +
      '📝 HÀNG ĐỢI:\n' +
      '   • Tổng số: ' + data.queue.total + '\n' +
      '   • Pending: ' + data.queue.pending + '\n' +
      '   • Đã gửi: ' + data.queue.sent + '\n' +
      '   • Lỗi: ' + data.queue.error + '\n\n' +
      '👥 RVA:\n' +
      '   • Tổng số: ' + data.rva.total + '\n' +
      '   • Hoạt động: ' + data.rva.active + '\n' +
      '   • Không hoạt động: ' + data.rva.inactive + '\n\n' +
      '🗂️ BỘ NHỚ:\n' +
      '   • Lead cuối: ' + data.memory.lastProcessedLeadId + '\n' +
      '   • RVA tiếp: #' + (data.memory.nextRvaIndex + 1) + '\n' +
      '   • Queue row: ' + data.memory.lastQueueRow + '\n\n' +
      '⏰ Cập nhật: ' + new Date(data.timestamp).toLocaleString('vi-VN') + '\n\n' +
      '==========================================';
    
    ui.alert('🎯 Zalo Lead Distribution - Control Panel', message, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('❌ Lỗi', 'Lỗi trong showControlPanel:\n' + error.toString(), ui.ButtonSet.OK);
    Logger.log("❌ Error in showControlPanel: " + error.toString());
  }
}

// ==========================================
// 🎯 GIAI ĐOẠN 1: PHÂN TÍCH LEAD
// ==========================================
function runStage1FromUI() {
  try {
    logSystem("🔄 Bắt đầu Giai đoạn 1");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    
    if (!leadSheet) {
      return {
        success: false,
        message: '❌ Không tìm thấy sheet "' + SYSTEM_CONFIG.SHEETS.LEADS + '"'
      };
    }
    
    const lastRow = leadSheet.getLastRow();
    const startRow = SYSTEM_CONFIG.LEADS_CONFIG.START_ROW;
    
    if (lastRow < startRow) {
      return {
        success: false,
        message: "⚠️ Không có dữ liệu lead để phân tích"
      };
    }
    
    const data = leadSheet.getRange(startRow, 1, lastRow - startRow + 1, SYSTEM_CONFIG.LEADS_CONFIG.COL_LEAD_ID).getValues();
    
    let total = 0;
    let missingLeadId = 0;
    
    data.forEach(function(row) {
      const fullName = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_NAME - 1];
      const leadId = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_LEAD_ID - 1];
      
      if (fullName && fullName !== "") {
        total++;
        
        if (!leadId || leadId === "") {
          missingLeadId++;
        }
      }
    });
    
    const queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    const sentLeadIds = new Set();
    const errorLeadIds = new Set();
    
    if (queueSheet && queueSheet.getLastRow() > 1) {
      const queueLastRow = queueSheet.getLastRow();
      const queueData = queueSheet.getRange(2, 1, queueLastRow - 1, SYSTEM_CONFIG.QUEUE_CONFIG.COL_STATUS).getValues();
      
      queueData.forEach(row => {
        const leadId = String(row[SYSTEM_CONFIG.QUEUE_CONFIG.COL_LEAD_ID - 1]).trim();
        const status = String(row[SYSTEM_CONFIG.QUEUE_CONFIG.COL_STATUS - 1]).trim().toLowerCase();
        
        if (leadId) {
          if (status === "success" || status === "sent" || status.includes("success")) {
            sentLeadIds.add(leadId);
          } else if (status === "error" || status.includes("error")) {
            errorLeadIds.add(leadId);
          }
        }
      });
    }
    
    const sent = sentLeadIds.size;
    const error = errorLeadIds.size;
    const pending = total - sent - error;
    
    const rvaConfig = getRVAConfig();
    const activeRVA = rvaConfig.filter(function(r) { return r.active; }).length;
    
    const memory = getSystemMemory();
    
    logSystem('✅ Giai đoạn 1 hoàn tất: ' + total + ' lead, ' + pending + ' chờ gửi, ' + sent + ' đã gửi', "SUCCESS");
    
    let warningMsg = '';
    if (missingLeadId > 0) {
      warningMsg = '\n\n⚠️ CÓ ' + missingLeadId + ' LEAD CHƯA CÓ ID!\nVui lòng chạy "Tạo Lead ID" trước.';
    }
    
    return {
      success: true,
      message: '✅ Phân tích hoàn tất!\n\n' +
              '📊 Tổng lead: ' + total + '\n' +
              '⏳ Lead chờ gửi: ' + pending + '\n' +
              '✓ Đã gửi: ' + sent + '\n' +
              '❌ Lỗi: ' + error + '\n' +
              '👥 RVA hoạt động: ' + activeRVA + '/' + rvaConfig.length + '\n\n' +
              '🗂️ BỘ NHỚ:\n' +
              'Lead cuối xử lý: ' + (memory.lastProcessedLeadId || "Chưa có") + '\n' +
              'RVA tiếp theo: #' + (memory.nextRvaIndex + 1) +
              warningMsg
    };
    
  } catch (error) {
    logSystem('❌ Lỗi Giai đoạn 1: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: '❌ Lỗi: ' + error.toString()
    };
  }
}

// ==========================================
// 🎯 GIAI ĐOẠN 2: TẠO HÀNG ĐỢI
// ==========================================
function runStage2FromUI() {
  try {
    logSystem("🔄 Bắt đầu Giai đoạn 2 - Tạo hàng đợi (Phiên bản cải tiến với Memory)");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    
    if (!leadSheet) {
      return {
        success: false,
        message: '❌ Không tìm thấy sheet "' + SYSTEM_CONFIG.SHEETS.LEADS + '"'
      };
    }
    
    const rvaConfig = getRVAConfig();
    const activeRVAs = rvaConfig.filter(function(rva) { return rva.active; });
    
    if (activeRVAs.length === 0) {
      return {
        success: false,
        message: "❌ Không có RVA nào đang hoạt động!\n\nVui lòng cấu hình RVA trong sheet 'RVA_Config'"
      };
    }
    
    logSystem('✅ Tìm thấy ' + activeRVAs.length + ' RVA đang hoạt động');
    
    const memory = getSystemMemory();
    logSystem('🗂️ Memory hiện tại: lastLeadId=' + memory.lastProcessedLeadId + 
             ', nextRVA=' + memory.nextRvaIndex);
    
    const lastRow = leadSheet.getLastRow();
    const startRow = SYSTEM_CONFIG.LEADS_CONFIG.START_ROW;
    
    if (lastRow < startRow) {
      return {
        success: false,
        message: "⚠️ Không có dữ liệu lead để xử lý"
      };
    }
    
    let queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    
    const data = leadSheet.getRange(startRow, 1, lastRow - startRow + 1, SYSTEM_CONFIG.LEADS_CONFIG.COL_PROJECT).getValues();
    
    const processedLeadIds = new Set();
    
    if (queueSheet && queueSheet.getLastRow() > 1) {
      const queueData = queueSheet.getRange(2, SYSTEM_CONFIG.QUEUE_CONFIG.COL_LEAD_ID, queueSheet.getLastRow() - 1, 1).getValues();
      queueData.forEach(function(row) {
        const leadId = String(row[0]).trim();
        if (leadId) {
          processedLeadIds.add(leadId);
        }
      });
      logSystem('📊 Đã tìm thấy ' + processedLeadIds.size + ' Lead ID trong Queue');
    }
    
    const pendingLeads = [];
    let skippedDuplicate = 0;
    let skippedProcessed = 0;
    let missingLeadId = 0;
    let foundLastProcessed = memory.lastProcessedLeadId === "";
    
    data.forEach(function(row, index) {
      const leadId = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_LEAD_ID - 1];
      const fullName = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_NAME - 1];
      const phone = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_PHONE - 1];
      const need = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_NEED - 1];
      const project = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_PROJECT - 1];
      
      if (fullName && fullName !== "") {
        if (!leadId || leadId === "") {
          missingLeadId++;
          logSystem('⚠️ CẢNH BÁO: Lead "' + fullName + '" chưa có ID', "WARNING");
          return;
        }
        
        if (!foundLastProcessed) {
          if (String(leadId).trim() === memory.lastProcessedLeadId) {
            foundLastProcessed = true;
            logSystem('✅ Đã tìm thấy lead cuối xử lý: ' + leadId);
          }
          skippedProcessed++;
          return;
        }
        
        if (processedLeadIds.has(String(leadId).trim())) {
          skippedDuplicate++;
          logSystem('⏭️ Bỏ qua Lead ID đã có trong Queue: ' + leadId, "INFO");
          return;
        }
        
        pendingLeads.push({
          rowIndex: startRow + index,
          leadId: leadId,
          fullName: fullName,
          phone: phone || "",
          need: need || "",
          project: project || ""
        });
      }
    });
    
    if (missingLeadId > 0) {
      return {
        success: false,
        message: "⚠️ CÓ " + missingLeadId + " LEAD CHƯA CÓ ID!\n\n" +
                "Vui lòng chạy 'Tạo Lead ID' trước khi tạo Queue."
      };
    }
    
    if (pendingLeads.length === 0) {
      let msg = "⚠️ Không có lead mới nào cần gửi!\n\n";
      if (skippedProcessed > 0) {
        msg += "Đã bỏ qua " + skippedProcessed + " lead đã xử lý (dựa vào Memory).\n";
      }
      if (skippedDuplicate > 0) {
        msg += "Đã bỏ qua " + skippedDuplicate + " lead trùng trong Queue.\n";
      }
      return {
        success: false,
        message: msg
      };
    }
    
    logSystem('📊 Tìm thấy ' + pendingLeads.length + ' lead mới (bỏ qua ' + skippedProcessed + ' lead đã xử lý, ' + skippedDuplicate + ' lead trùng)');
    
    if (!queueSheet) {
      queueSheet = ss.insertSheet(SYSTEM_CONFIG.SHEETS.QUEUE);
      queueSheet.appendRow([
        "Timestamp",
        "RVA ID",
        "RVA Name",
        "Zalo Phone",
        "Zalo ID",
        "Lead ID",
        "Lead Name",
        "Lead Phone",
        "Need",
        "Project",
        "Message",
        "Status",
        "Original Row",
        "Column Index",
        "Error",
        "Sent Time",
        "Queue Row"
      ]);
      
      const headerRange = queueSheet.getRange("1:1");
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#667eea");
      headerRange.setFontColor("#ffffff");
      
      queueSheet.setColumnWidth(1, 150);
      queueSheet.setColumnWidth(6, 100);
      queueSheet.setColumnWidth(7, 150);
      queueSheet.setColumnWidth(11, 300);
      queueSheet.setColumnWidth(12, 100);
      
      logSystem("✅ Đã tạo sheet Queue mới");
    }
    
    let queuedCount = 0;
    let rvaIndex = memory.nextRvaIndex;
    let lastProcessedLeadId = "";
    
    logSystem('📦 Bắt đầu phân chia ' + pendingLeads.length + ' lead cho ' + activeRVAs.length + ' RVA (bắt đầu từ RVA #' + (rvaIndex + 1) + ')');
    
    pendingLeads.forEach(function(lead, index) {
      const currentRVA = activeRVAs[rvaIndex];
      
      const message ='Lead ID: ' + lead.leadId + '\n' +
                     'Tên: ' + lead.fullName + '\n' +
                     'SĐT: ' + lead.phone + '\n' +
                     'Nhu cầu: ' + lead.need + '\n' +
                     'Dự án: ' + lead.project;
      
      const queueRowNumber = queueSheet.getLastRow() + 1;
      
      queueSheet.appendRow([
        new Date(),
        currentRVA.id,
        currentRVA.name,
        currentRVA.phone,
        currentRVA.zaloid,
        lead.leadId,
        lead.fullName,
        lead.phone,
        lead.need,
        lead.project,
        message,
        "Pending",
        lead.rowIndex,
        currentRVA.columnIndex,
        "",
        "",
        queueRowNumber
      ]);
      
      queuedCount++;
      lastProcessedLeadId = lead.leadId;
      
      rvaIndex = (rvaIndex + 1) % activeRVAs.length;
      
      if ((index + 1) % 50 === 0) {
        logSystem('📊 Đã xử lý: ' + (index + 1) + '/' + pendingLeads.length + ' lead');
      }
    });
    
    updateSystemMemory({
      lastProcessedLeadId: lastProcessedLeadId,
      nextRvaIndex: rvaIndex,
      lastQueueRow: queueSheet.getLastRow()
    });
    
    logSystem('🗂️ Đã cập nhật Memory: lastLeadId=' + lastProcessedLeadId + ', nextRVA=' + rvaIndex);
    
    try {
      const statusRange = queueSheet.getRange("L:L");
      
      const pendingRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Pending")
        .setBackground("#FFF3CD")
        .setFontColor("#856404")
        .setRanges([statusRange])
        .build();
      
      const sentRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("Sent")
        .setBackground("#D4EDDA")
        .setFontColor("#155724")
        .setRanges([statusRange])
        .build();
      
      const errorRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("Error")
        .setBackground("#F8D7DA")
        .setFontColor("#721c24")
        .setRanges([statusRange])
        .build();
      
      queueSheet.setConditionalFormatRules([pendingRule, sentRule, errorRule]);
    } catch (e) {
      logSystem('⚠️ Không thể set conditional formatting: ' + e.toString(), "WARNING");
    }
    
    logSystem('✅ Giai đoạn 2 hoàn tất - Đã thêm ' + queuedCount + ' lead vào queue', "SUCCESS");
    
    logSystem('🔄 Bắt đầu tự động đồng bộ Status về lead-mkt...', "INFO");
    const syncResult = syncStatusToLeadSheetV2();
    
    let resultMsg = '✅ Tạo hàng đợi thành công!\n\n' +
                   '📦 Đã thêm: ' + queuedCount + ' lead mới\n' +
                   '👥 Phân cho: ' + activeRVAs.length + ' RVA\n';
    
    if (skippedProcessed > 0) {
      resultMsg += '⏭️ Bỏ qua (Memory): ' + skippedProcessed + ' lead\n';
    }
    if (skippedDuplicate > 0) {
      resultMsg += '⏭️ Bỏ qua (Trùng): ' + skippedDuplicate + ' lead\n';
    }
    
    resultMsg += '\n📌 RVA tiếp theo: #' + (rvaIndex + 1) + '\n';
    resultMsg += '🔖 Lead cuối: ' + lastProcessedLeadId + '\n';
    
    if (syncResult.success) {
      resultMsg += '\n✅ Đã tự động cập nhật Status "Pending"\nvào sheet lead-mkt\n';
    }
    
    resultMsg += '\n🚀 Tiếp theo: Chạy Giai đoạn 3 hoặc để n8n tự động xử lý';
    
    return {
      success: true,
      message: resultMsg
    };
    
  } catch (error) {
    logSystem('❌ Lỗi Giai đoạn 2: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: '❌ Lỗi: ' + error.toString()
    };
  }
}

// ==========================================
// 🎯 GIAI ĐOẠN 3: KÍCH HOẠT GỬI
// ==========================================
function runStage3FromUI() {
  try {
    logSystem("🚀 Bắt đầu Giai đoạn 3");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    
    if (!queueSheet) {
      logSystem("❌ Sheet Queue không tồn tại", "ERROR");
      return {
        success: false,
        message: "❌ Sheet 'Zalo_Queue' không tồn tại!\n\nVui lòng chạy Giai đoạn 2 trước."
      };
    }
    
    const lastRow = queueSheet.getLastRow();
    
    if (lastRow <= 1) {
      logSystem("⚠️ Queue sheet trống", "WARNING");
      return {
        success: false,
        message: "⚠️ Hàng đợi trống!\n\nVui lòng chạy Giai đoạn 2 để tạo queue."
      };
    }
    
    logSystem('📊 Queue có ' + (lastRow - 1) + ' dòng dữ liệu');
    
    const allData = queueSheet.getDataRange().getValues();
    const headers = allData[0];
    const statusCol = SYSTEM_CONFIG.QUEUE_CONFIG.COL_STATUS;
    const statusIndex = statusCol - 1;
    
    let pendingCount = 0;
    let sentCount = 0;
    let errorCount = 0;
    const pendingRows = [];
    
    for (let i = 1; i < allData.length; i++) {
      const status = String(allData[i][statusIndex]).trim();
      
      if (status === "Pending") {
        pendingCount++;
        const rowData = {};
        headers.forEach(function(header, index) {
          rowData[header] = allData[i][index];
        });
        
        pendingRows.push({
          rowNumber: i + 1,
          rowData: rowData
        });
      } else if (status === "Sent" || status.includes("Sent")) {
        sentCount++;
      } else if (status === "Error") {
        errorCount++;
      }
    }
    
    logSystem('📊 Thống kê: Pending=' + pendingCount + ', Sent=' + sentCount + ', Error=' + errorCount);
    
    if (pendingCount === 0) {
      return {
        success: false,
        message: '⚠️ Không có lead nào cần gửi!\n\n' +
                '✓ Sent: ' + sentCount + '\n' +
                '❌ Error: ' + errorCount
      };
    }
    
    logSystem('📤 Đang gửi ' + pendingCount + ' lead đến n8n...');
    
    const n8nWebhookUrl = SYSTEM_CONFIG.N8N_CONFIG.WEBHOOK_URL;
    
    if (!n8nWebhookUrl || n8nWebhookUrl === '' || n8nWebhookUrl.indexOf('http') !== 0) {
      return {
        success: false,
        message: '❌ Chưa cấu hình n8n Webhook URL!\n\n' +
                'Vui lòng cập nhật:\n' +
                'SYSTEM_CONFIG.N8N_CONFIG.WEBHOOK_URL'
      };
    }
    
    const payload = {
      trigger: 'apps_script_stage3',
      timestamp: new Date().toISOString(),
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      sheetName: SYSTEM_CONFIG.SHEETS.QUEUE,
      pendingCount: pendingCount,
      pendingRows: pendingRows
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    try {
      const response = UrlFetchApp.fetch(n8nWebhookUrl, options);
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();
      
      logSystem('📥 n8n Response Code: ' + responseCode);
      
      if (responseCode === 200) {
        logSystem('✅ Webhook gửi thành công', "SUCCESS");
        
        return {
          success: true,
          message: '✅ Đã kích hoạt n8n thành công!\n\n' +
                  '📊 ' + pendingCount + ' lead đang được xử lý\n' +
                  '🤖 n8n đang gửi tin nhắn Zalo\n\n' +
                  '📋 Theo dõi tiến độ tại sheet "Zalo_Queue"\n' +
                  'Cột K (Status) sẽ tự động cập nhật thành "Success"\n' +
                  'khi n8n gửi tin nhắn thành công'
        };
        
      } else {
        logSystem('❌ n8n trả về lỗi: ' + responseCode, "ERROR");
        
        return {
          success: false,
          message: '❌ Lỗi khi gọi n8n!\n\n' +
                  'Response Code: ' + responseCode + '\n' +
                  'Message: ' + responseBody
        };
      }
      
    } catch (webhookError) {
      logSystem('❌ Lỗi kết nối n8n: ' + webhookError.toString(), "ERROR");
      
      return {
        success: false,
        message: '❌ Không thể kết nối đến n8n!\n\n' +
                'Error: ' + webhookError.toString()
      };
    }
    
  } catch (error) {
    logSystem('❌ Lỗi Giai đoạn 3: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: '❌ Lỗi: ' + error.toString()
    };
  }
}

// ==========================================
// 🔄 ĐỒNG BỘ STATUS VỀ LEAD-MKT (v3 - CHỈ CẬP NHẬT KHI CẦN)
// ==========================================
function syncStatusToLeadSheetV2() {
  try {
    logSystem("🔄 Bắt đầu đồng bộ Status & Time từ Queue về lead-mkt (v3 - Smart Update)");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    const queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    
    if (!leadSheet || !queueSheet || queueSheet.getLastRow() <= 1) {
      return {
        success: false,
        message: "⚠️ Queue trống hoặc sheet không tồn tại"
      };
    }
    
    const lastRow = queueSheet.getLastRow();
    const queueData = queueSheet.getRange(2, 1, lastRow - 1, SYSTEM_CONFIG.QUEUE_CONFIG.COL_QUEUE_ROW).getValues();
    
    let updatedCount = 0;
    let skippedCount = 0;
    let errorCount = 0;
    
    // ✅ TẠO MAP ĐỂ THEO DÕI LEAD ĐÃ CẬP NHẬT
    const updatedLeads = new Set();
    
    queueData.forEach(function(row, index) {
      const leadId = String(row[SYSTEM_CONFIG.QUEUE_CONFIG.COL_LEAD_ID - 1]).trim();
      const rvaId = String(row[SYSTEM_CONFIG.QUEUE_CONFIG.COL_RVA_ID - 1]).trim();
      const status = String(row[SYSTEM_CONFIG.QUEUE_CONFIG.COL_STATUS - 1]).trim();
      const originalRow = row[SYSTEM_CONFIG.QUEUE_CONFIG.COL_ORIGINAL_ROW - 1];
      const columnIndex = row[SYSTEM_CONFIG.QUEUE_CONFIG.COL_COLUMN_INDEX - 1];
      
      if (!leadId || !rvaId || !status) {
        return; // Bỏ qua dòng không đủ thông tin
      }
      
      // ✅ CHỈ XỬ LÝ STATUS ĐÃ THAY ĐỔI (Success hoặc Error)
      // BỎ QUA Pending vì chưa có kết quả
      if (status.toLowerCase() === "pending") {
        return;
      }
      
      // ✅ TẠO KEY DUY NHẤT: leadId + rvaId + columnIndex
      const uniqueKey = leadId + "_" + rvaId + "_" + columnIndex;
      
      // ✅ KIỂM TRA ĐÃ CẬP NHẬT TRONG LẦN CHẠY NÀY CHƯA
      if (updatedLeads.has(uniqueKey)) {
        skippedCount++;
        return; // Đã cập nhật rồi, bỏ qua
      }
      
      // ✅ KIỂM TRA LEAD SHEET CÓ STATUS NÀY CHƯA
      try {
        const currentStatus = leadSheet.getRange(originalRow, columnIndex).getValue();
        
        // Nếu status trong lead-mkt đã là Success/Error rồi → Bỏ qua
        if (currentStatus === "Success" || currentStatus === "Error") {
          skippedCount++;
          logSystem('⏭️ Bỏ qua lead ' + leadId + ' (đã có status: ' + currentStatus + ')', "INFO");
          return;
        }
        
      } catch (e) {
        logSystem('⚠️ Không thể đọc status hiện tại của lead ' + leadId, "WARNING");
      }
      
      // ✅ CẬP NHẬT STATUS MỚI
      let statusToWrite = "";
      let careStatus = "";
      
      if (status.toLowerCase() === "success" || status.toLowerCase() === "sent") {
        statusToWrite = "Success";
        careStatus = "Đã nhận lead";
      } else if (status.toLowerCase() === "error") {
        statusToWrite = "Error";
        careStatus = "Lỗi gửi";
      }
      
      if (statusToWrite) {
        const result = updateLeadAssignment(leadId, rvaId, statusToWrite, careStatus);
        if (result) {
          updatedCount++;
          updatedLeads.add(uniqueKey);
          logSystem('✅ Cập nhật ' + leadId + ' → ' + statusToWrite, "INFO");
        } else {
          errorCount++;
        }
      }
    });
    
    logSystem('✅ Đồng bộ v3 hoàn tất: Cập nhật=' + updatedCount + ', Bỏ qua=' + skippedCount + ', Lỗi=' + errorCount, "SUCCESS");
    
    return {
      success: true,
      message: '✅ Đồng bộ hoàn tất!\n\n' +
              '📊 ĐÃ CẬP NHẬT: ' + updatedCount + ' lead\n' +
              '⏭️ BỎ QUA: ' + skippedCount + ' lead (đã có status)\n' +
              (errorCount > 0 ? '❌ LỖI: ' + errorCount + ' lead\n' : '') +
              '\n💡 Chỉ cập nhật lead có status Success/Error mới'
    };
    
  } catch (error) {
    logSystem('❌ Lỗi syncStatusToLeadSheetV2: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: '❌ Lỗi: ' + error.toString()
    };
  }
}

// ==========================================
// 🧹 DỌN DẸP QUEUE
// ==========================================
function cleanupQueueFromUI() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    
    if (!queueSheet) {
      return {
        success: false,
        message: "⚠️ Sheet Queue không tồn tại"
      };
    }
    
    if (queueSheet.getLastRow() > 1) {
      queueSheet.deleteRows(2, queueSheet.getLastRow() - 1);
    }
    
    resetSystemMemory();
    
    logSystem("✅ Đã dọn dẹp queue và reset memory", "SUCCESS");
    
    return {
      success: true,
      message: "✅ Đã dọn dẹp queue và reset bộ nhớ thành công!"
    };
    
  } catch (error) {
    logSystem('❌ Lỗi cleanup: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: '❌ Lỗi: ' + error.toString()
    };
  }
}

// ==========================================
// 📋 MENU CHÍNH
// ==========================================
function showMainMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '🎯 ZALO LEAD DISTRIBUTION SYSTEM (v3.0 - CRM Edition)',
    'Chọn OK để mở menu điều khiển',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response == ui.Button.OK) {
    showActionMenu();
  }
}

function showActionMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    '🎯 CHỌN HÀNH ĐỘNG',
    'Nhập số tương ứng:\n\n' +
    '1. 📊 Xem thống kê hệ thống\n' +
    '2. 🔖 Tạo Lead ID tự động\n' +
    '3. 1️⃣ Chạy Giai đoạn 1 (Phân tích Lead)\n' +
    '4. 2️⃣ Chạy Giai đoạn 2 (Tạo Queue)\n' +
    '5. 3️⃣ Chạy Giai đoạn 3 (Kích hoạt gửi)\n' +
    '6. 🔄 Đồng bộ Status về lead-mkt\n' +
    '7. 👥 Quản lý RVA\n' +
    '8. 🧹 Dọn dẹp Queue\n' +
    '9. 📝 Xem Log\n' +
    '10. ♻️ Reset Memory\n' +
    '11. ✨ Đồng bộ cột RVA (CHỈ CHẠY KHI CẦN)\n' +
    '12. 🔄 Chuyển Lead giữa RVA\n' +
    '13. 📊 Xem lịch sử chuyển Lead\n' +
    '14. 🔍 Kiểm tra cấu trúc cột RVA\n\n' +
    '0. ❌ Thoát\n\n' +
    'Nhập lựa chọn:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const choice = response.getResponseText().trim();
    
    switch(choice) {
      case '1':
        showControlPanel();
        showActionMenu();
        break;
      case '2':
        generateLeadIdMenu();
        break;
      case '3':
        runStage1Menu();
        break;
      case '4':
        runStage2Menu();
        break;
      case '5':
        runStage3Menu();
        break;
      case '6':
        syncStatusMenu();
        break;
      case '7':
        manageRVAMenu();
        break;
      case '8':
        cleanupQueueMenu();
        break;
      case '9':
        viewLogMenu();
        break;
      case '10':
        resetMemoryMenu();
        break;
      case '11':
        syncRVAColumnsMenu();
        break;
      case '12':
        transferLeadMenu();
        break;
      case '13':
        viewLeadHistoryMenu();
        break;
      case '14':
        checkRVAColumnsMenu();
        break;
      case '0':
        return;
      default:
        ui.alert('⚠️ Cảnh báo', 'Lựa chọn không hợp lệ!', ui.ButtonSet.OK);
        showActionMenu();
    }
  }
}

// ==========================================
// 📋 CÁC HÀM MENU
// ==========================================
function generateLeadIdMenu() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '🔖 TẠO LEAD ID',
    'Tự động tạo Lead ID cho các lead chưa có ID?\n\n' +
    'Định dạng: le01, le02, le03...',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm == ui.Button.YES) {
    const result = autoGenerateLeadIds();
    ui.alert(result.success ? '✅ Thành công' : '❌ Lỗi', result.message, ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function runStage1Menu() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '📊 GIAI ĐOẠN 1',
    'Phân tích lead từ sheet. Tiếp tục?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm == ui.Button.YES) {
    const result = runStage1FromUI();
    ui.alert(result.success ? '✅ Thành công' : '❌ Lỗi', result.message, ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function runStage2Menu() {
  const ui = SpreadsheetApp.getUi();
  const result = runStage2FromUI();
  ui.alert(result.success ? '✅ Thành công' : '❌ Lỗi', result.message, ui.ButtonSet.OK);
  showActionMenu();
}

function runStage3Menu() {
  const ui = SpreadsheetApp.getUi();
  const result = runStage3FromUI();
  ui.alert(result.success ? '✅ Thành công' : '❌ Lỗi', result.message, ui.ButtonSet.OK);
  showActionMenu();
}

function syncStatusMenu() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '🔄 ĐỒNG BỘ STATUS',
    'Cập nhật status từ Queue về các cột RVA trong lead-mkt?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm == ui.Button.YES) {
    const result = syncStatusToLeadSheetV2();
    ui.alert(result.success ? '✅ Thành công' : '❌ Lỗi', result.message, ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function manageRVAMenu() {
  const ui = SpreadsheetApp.getUi();
  const config = getRVAConfig();
  
  let msg = '👥 DANH SÁCH RVA:\n\n';
  config.forEach(function(rva, i) {
    msg += (i+1) + '. ' + rva.name + ' - ' + (rva.active ? '✅ Hoạt động' : '❌ Không hoạt động') + '\n';
  });
  
  ui.alert('👥 RVA', msg, ui.ButtonSet.OK);
  showActionMenu();
}

function cleanupQueueMenu() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '🧹 XÓA QUEUE & RESET MEMORY',
    'Xác nhận xóa Queue và reset bộ nhớ?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm == ui.Button.YES) {
    const result = cleanupQueueFromUI();
    ui.alert(result.success ? '✅ Thành công' : '❌ Lỗi', result.message, ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function resetMemoryMenu() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '♻️ RESET MEMORY',
    'Reset bộ nhớ về trạng thái ban đầu?\n\n' +
    'Hệ thống sẽ quét lại từ đầu trong lần chạy tiếp theo.',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm == ui.Button.YES) {
    const result = resetSystemMemory();
    ui.alert(result.success ? '✅ Thành công' : '❌ Lỗi', result.message, ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function syncRVAColumnsMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const previewResponse = ui.alert(
    '✨ ĐỒNG BỘ CỘT RVA VÀO LEAD-MKT',
    'Chức năng này sẽ:\n\n' +
    '1️⃣ Đọc danh sách RVA từ sheet "RVA_Config"\n' +
    '2️⃣ Tự động thêm các cột vào sheet "lead-mkt":\n' +
    '   • [Tên RVA] - Status\n' +
    '   • [Tên RVA] - Assigned Time\n' +
    '   • [Tên RVA] - Care Status\n' +
    '3️⃣ Cập nhật Column Index trong RVA_Config\n\n' +
    '⚠️ LƯU Ý:\n' +
    '• Chỉ THÊM cột mới, không xóa cột cũ\n' +
    '• Các cột đã có sẽ được giữ nguyên\n' +
    '• Nên backup sheet trước khi chạy\n\n' +
    'Bạn có muốn tiếp tục không?',
    ui.ButtonSet.YES_NO
  );
  
  if (previewResponse == ui.Button.YES) {
    const result = syncRVAColumnsToLeadSheet();
    ui.alert(
      result.success ? '✅ Đồng bộ thành công!' : '❌ Có lỗi xảy ra',
      result.message,
      ui.ButtonSet.OK
    );
  }
  
  showActionMenu();
}

function transferLeadMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const leadIdResponse = ui.prompt(
    '🔄 CHUYỂN LEAD - Bước 1',
    'Nhập Lead ID cần chuyển (ví dụ: le01):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (leadIdResponse.getSelectedButton() != ui.Button.OK) {
    showActionMenu();
    return;
  }
  
  const leadId = leadIdResponse.getResponseText().trim();
  
  const fromRvaResponse = ui.prompt(
    '🔄 CHUYỂN LEAD - Bước 2',
    'Nhập RVA ID hiện tại (ví dụ: RV001):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (fromRvaResponse.getSelectedButton() != ui.Button.OK) {
    showActionMenu();
    return;
  }
  
  const fromRvaId = fromRvaResponse.getResponseText().trim();
  
  const toRvaResponse = ui.prompt(
    '🔄 CHUYỂN LEAD - Bước 3',
    'Nhập RVA ID mới (ví dụ: RV002):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (toRvaResponse.getSelectedButton() != ui.Button.OK) {
    showActionMenu();
    return;
  }
  
  const toRvaId = toRvaResponse.getResponseText().trim();
  
  const reasonResponse = ui.prompt(
    '🔄 CHUYỂN LEAD - Bước 4',
    'Nhập lý do chuyển:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (reasonResponse.getSelectedButton() != ui.Button.OK) {
    showActionMenu();
    return;
  }
  
  const reason = reasonResponse.getResponseText().trim();
  
  const confirm = ui.alert(
    '⚠️ XÁC NHẬN CHUYỂN LEAD',
    'Lead: ' + leadId + '\n' +
    'Từ: ' + fromRvaId + '\n' +
    'Sang: ' + toRvaId + '\n' +
    'Lý do: ' + reason + '\n\n' +
    'Xác nhận chuyển?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm == ui.Button.YES) {
    const result = transferLeadToNewRVA(leadId, fromRvaId, toRvaId, reason);
    ui.alert(result.success ? '✅ Thành công' : '❌ Lỗi', result.message, ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function viewLeadHistoryMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEAD_HISTORY);
  
  if (!historySheet || historySheet.getLastRow() <= 1) {
    ui.alert('📊 Lịch sử', 'Chưa có lịch sử chuyển lead', ui.ButtonSet.OK);
  } else {
    ss.setActiveSheet(historySheet);
    ui.alert('📊 Lịch sử', 'Sheet lịch sử đã mở', ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function checkRVAColumnsMenu() {
  const ui = SpreadsheetApp.getUi();
  const result = checkRVAColumnsStatus();
  ui.alert(result.success ? '📊 Báo cáo' : '❌ Lỗi', result.message, ui.ButtonSet.OK);
  showActionMenu();
}

function viewLogMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LOG);
  
  if (!logSheet || logSheet.getLastRow() <= 1) {
    ui.alert('📝 Log', 'Chưa có log', ui.ButtonSet.OK);
  } else {
    ss.setActiveSheet(logSheet);
    ui.alert('📝 Log', 'Sheet log đã mở', ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

// ==========================================
// 🚀 ON OPEN - TẠO MENU
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('⚡ Zalo CRM v3')
    .addItem('🎨 Control Panel', 'showMainMenu')
    .addItem('📈 Thống kê', 'showControlPanel')
    .addSeparator()
    .addSubMenu(ui.createMenu('📋 Quản lý Lead')
      .addItem('🔖 Tạo Lead ID', 'generateLeadIdMenu')
      .addItem('1️⃣ Giai đoạn 1 - Phân tích', 'runStage1Menu')
      .addItem('2️⃣ Giai đoạn 2 - Tạo Queue', 'runStage2Menu')
      .addItem('3️⃣ Giai đoạn 3 - Kích hoạt', 'runStage3Menu'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Cấu hình & Đồng bộ')
      .addItem('✨ Đồng bộ cột RVA', 'syncRVAColumnsMenu')
      .addItem('🔄 Đồng bộ Status', 'syncStatusMenu')
      .addItem('🔍 Kiểm tra cột RVA', 'checkRVAColumnsMenu')
      .addItem('👥 Quản lý RVA', 'manageRVAMenu'))
    .addSubMenu(ui.createMenu('🔀 Chuyển & Lịch sử')
      .addItem('🔀 Chuyển Lead', 'transferLeadMenu')
      .addItem('📊 Xem lịch sử', 'viewLeadHistoryMenu'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🛠️ Công cụ')
      .addItem('♻️ Reset Memory', 'resetMemoryMenu')
      .addItem('🧹 Dọn dẹp Queue', 'cleanupQueueMenu')
      .addItem('📝 Xem Log', 'viewLogMenu'))
    .addToUi();
  
  logSystem("✅ Menu khởi tạo - Version 3.0 CRM Edition with Manual RVA Sync");
}
