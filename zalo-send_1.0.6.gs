// ==========================================
// CAU HINH HE THONG
// ==========================================
const SYSTEM_CONFIG = {
  SHEETS: {
    LEADS: "lead-mkt",
    QUEUE: "Zalo_Queue",
    RVA_CONFIG: "RVA_Config",
    LOG: "System_Log"
  },
  LEADS_CONFIG: {
    START_ROW: 5,
    COL_STT: 1,
    COL_LEAD_ID: 2,        // THÊM MỚI: Cột Lead ID
    COL_NAME: 3,           // Cập nhật từ 2 -> 3
    COL_PHONE: 4,          // Cập nhật từ 3 -> 4
    COL_NEED: 5,           // Cập nhật từ 4 -> 5
    COL_PROJECT: 6,        // Cập nhật từ 5 -> 6
    COL_RVA1: 8,           // Cập nhật từ 7 -> 8
    COL_RVA2: 9,           // Cập nhật từ 8 -> 9
    COL_RVA3: 10           // Cập nhật từ 9 -> 10
  },
  QUEUE_CONFIG: {
    COL_TIMESTAMP: 1,
    COL_RVA_ID: 2,
    COL_RVA_NAME: 3,
    COL_ZALO_PHONE: 4,
    COL_ZALO_ID: 5,
    COL_LEAD_ID: 6,        // THÊM MỚI: Cột Lead ID trong Queue
    COL_LEAD_NAME: 7,      // Cập nhật từ 6 -> 7
    COL_LEAD_PHONE: 8,     // Cập nhật từ 7 -> 8
    COL_NEED: 9,           // Cập nhật từ 8 -> 9
    COL_PROJECT: 10,       // Cập nhật từ 9 -> 10
    COL_MESSAGE: 11,       // Cập nhật từ 10 -> 11
    COL_STATUS: 12,        // Cập nhật từ 11 -> 12
    COL_ORIGINAL_ROW: 13,  // Cập nhật từ 12 -> 13
    COL_COLUMN_INDEX: 14,  // Cập nhật từ 13 -> 14
    COL_ERROR: 15,         // Cập nhật từ 14 -> 15
    COL_SENT_TIME: 16,     // Cập nhật từ 15 -> 16
    COL_QUEUE_ROW: 17      // THÊM MỚI: Số thứ tự row trong Queue
  },
  N8N_CONFIG: {
    WEBHOOK_URL: 'https://n8n.rever.io.vn/webhook/zalo-trigger',
    RESULT_COL: 18,        // Cập nhật từ 18 -> 18 (Message ID từ n8n)
    QUEUE_ROW_COL: 17      // THÊM MỚI - Cột Queue Row
  }
};

// ==========================================
// HÀM TẠO LEAD ID TỰ ĐỘNG
// ==========================================
function generateLeadId(rowIndex) {
  // Tạo ID dạng le01, le02, le03...
  const idNumber = String(rowIndex).padStart(2, '0');
  return 'le' + idNumber;
}

// ==========================================
// HÀM TẠO LEAD ID CHO TẤT CẢ LEAD CHƯA CÓ ID
// ==========================================
function autoGenerateLeadIds() {
  try {
    logSystem("Bắt đầu tạo Lead ID tự động");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    
    if (!leadSheet) {
      return {
        success: false,
        message: 'Không tìm thấy sheet "' + SYSTEM_CONFIG.SHEETS.LEADS + '"'
      };
    }
    
    const lastRow = leadSheet.getLastRow();
    const startRow = SYSTEM_CONFIG.LEADS_CONFIG.START_ROW;
    
    if (lastRow < startRow) {
      return {
        success: false,
        message: "Không có dữ liệu lead"
      };
    }
    
    // Đọc dữ liệu cột Name và Lead ID
    const nameData = leadSheet.getRange(startRow, SYSTEM_CONFIG.LEADS_CONFIG.COL_NAME, lastRow - startRow + 1, 1).getValues();
    const leadIdData = leadSheet.getRange(startRow, SYSTEM_CONFIG.LEADS_CONFIG.COL_LEAD_ID, lastRow - startRow + 1, 1).getValues();
    
    let createdCount = 0;
    let skippedCount = 0;
    
    // Tìm số thứ tự lớn nhất hiện có
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
    
    // Tạo Lead ID cho các dòng chưa có
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
    
    logSystem('Đã tạo ' + createdCount + ' Lead ID mới', "SUCCESS");
    
    return {
      success: true,
      message: 'Tạo Lead ID hoàn tất!\n\n' +
              'Đã tạo: ' + createdCount + ' ID mới\n' +
              'Đã có sẵn: ' + skippedCount + ' ID\n' +
              'ID tiếp theo sẽ là: le' + String(currentNumber + 1).padStart(2, '0')
    };
    
  } catch (error) {
    logSystem('Lỗi tạo Lead ID: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: 'Lỗi: ' + error.toString()
    };
  }
}

// ==========================================
// API: LẤY THỐNG KÊ HỆ THỐNG
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
    
    // Đếm lead từ sheet lead-mkt và kiểm tra status từ các cột RVA
    if (leadSheet && leadSheet.getLastRow() >= SYSTEM_CONFIG.LEADS_CONFIG.START_ROW) {
      const lastRow = leadSheet.getLastRow();
      const startRow = SYSTEM_CONFIG.LEADS_CONFIG.START_ROW;
      
      // Đọc từ cột Name đến cột RVA3
      const data = leadSheet.getRange(startRow, SYSTEM_CONFIG.LEADS_CONFIG.COL_NAME, 
                                      lastRow - startRow + 1, 
                                      SYSTEM_CONFIG.LEADS_CONFIG.COL_RVA3 - SYSTEM_CONFIG.LEADS_CONFIG.COL_NAME + 1).getValues();
      
      data.forEach(row => {
        const fullName = row[0]; // COL_NAME
        if (fullName && fullName !== "") {
          leadStats.total++;
          
          // Các cột RVA: RVA1, RVA2, RVA3
          const rva1 = String(row[SYSTEM_CONFIG.LEADS_CONFIG.COL_RVA1 - SYSTEM_CONFIG.LEADS_CONFIG.COL_NAME]).trim().toLowerCase();
          const rva2 = String(row[SYSTEM_CONFIG.LEADS_CONFIG.COL_RVA2 - SYSTEM_CONFIG.LEADS_CONFIG.COL_NAME]).trim().toLowerCase();
          const rva3 = String(row[SYSTEM_CONFIG.LEADS_CONFIG.COL_RVA3 - SYSTEM_CONFIG.LEADS_CONFIG.COL_NAME]).trim().toLowerCase();
          
          // Kiểm tra nếu có bất kỳ RVA nào có status
          const hasSuccess = rva1 === "success" || rva2 === "success" || rva3 === "success";
          const hasError = rva1 === "error" || rva2 === "error" || rva3 === "error";
          const hasPending = rva1 === "pending" || rva2 === "pending" || rva3 === "pending";
          const isEmpty = rva1 === "" && rva2 === "" && rva3 === "";
          
          if (hasSuccess) {
            leadStats.sent++;
          } else if (hasError) {
            leadStats.error++;
          } else if (hasPending) {
            leadStats.pending++;
          } else if (isEmpty) {
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
      
      Logger.log('Đọc Queue Status (cột L): Total=' + queueStats.total + ', Pending=' + queueStats.pending + ', Sent=' + queueStats.sent);
    }
    
    const rvaConfig = getRVAConfig();
    const rvaStats = {
      total: rvaConfig.length,
      active: rvaConfig.filter(rva => rva.active).length,
      inactive: rvaConfig.filter(rva => !rva.active).length
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
        logs: recentLogs,
        timestamp: new Date().toISOString()
      }
    };
    
  } catch (error) {
    Logger.log("Error in getSystemStats: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ==========================================
// API: LẤY CẤU HÌNH RVA
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
          columnIndex: 9,
          active: true,
          note: ""
        },
        {
          id: "RV003",
          name: "RVA 3",
          phone: "0903456789",
          zaloid: "1504484729431570818",
          columnIndex: 10,
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
    Logger.log("Lỗi getRVAConfig: " + error.toString());
    return [];
  }
}

// ==========================================
// LOGGING SYSTEM
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
    Logger.log("Lỗi logging: " + error.toString());
  }
}

// ==========================================
// CONTROL PANEL
// ==========================================
function showControlPanel() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const stats = getSystemStats();
    
    if (!stats.success) {
      ui.alert('Lỗi', 'Không thể lấy thông tin hệ thống:\n' + stats.error, ui.ButtonSet.OK);
      return;
    }
    
    const data = stats.data;
    
    const message = 
      'THỐNG KÊ HỆ THỐNG\n' +
      '==========================================\n\n' +
      'LEAD:\n' +
      '   Tổng số: ' + data.lead.total + '\n' +
      '   Chờ gửi: ' + data.lead.pending + '\n' +
      '   Đã gửi: ' + data.lead.sent + '\n\n' +
      'HÀNG ĐỢI:\n' +
      '   Tổng số: ' + data.queue.total + '\n' +
      '   Pending: ' + data.queue.pending + '\n' +
      '   Đã gửi: ' + data.queue.sent + '\n' +
      '   Lỗi: ' + data.queue.error + '\n\n' +
      'RVA:\n' +
      '   Tổng số: ' + data.rva.total + '\n' +
      '   Hoạt động: ' + data.rva.active + '\n' +
      '   Không hoạt động: ' + data.rva.inactive + '\n\n' +
      'Cập nhật: ' + new Date(data.timestamp).toLocaleString('vi-VN') + '\n\n' +
      '==========================================';
    
    ui.alert('Zalo Lead Distribution - Control Panel', message, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Lỗi', 'Lỗi trong showControlPanel:\n' + error.toString(), ui.ButtonSet.OK);
    Logger.log("Error in showControlPanel: " + error.toString());
  }
}

// ==========================================
// GIAI ĐOẠN 1: PHÂN TÍCH LEAD
// ==========================================
function runStage1FromUI() {
  try {
    logSystem("Bắt đầu Giai đoạn 1");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    
    if (!leadSheet) {
      return {
        success: false,
        message: 'Không tìm thấy sheet "' + SYSTEM_CONFIG.SHEETS.LEADS + '"'
      };
    }
    
    const lastRow = leadSheet.getLastRow();
    const startRow = SYSTEM_CONFIG.LEADS_CONFIG.START_ROW;
    
    if (lastRow < startRow) {
      return {
        success: false,
        message: "Không có dữ liệu lead để phân tích"
      };
    }
    
    const data = leadSheet.getRange(startRow, 1, lastRow - startRow + 1, SYSTEM_CONFIG.LEADS_CONFIG.COL_LEAD_ID).getValues();
    
    let total = 0;
    let missingLeadId = 0;
    
    // Đếm tổng lead và lead thiếu ID
    data.forEach(function(row) {
      const fullName = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_NAME - 1];
      const leadId = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_LEAD_ID - 1];
      
      if (fullName && fullName !== "") {
        total++;
        
        // Kiểm tra Lead ID
        if (!leadId || leadId === "") {
          missingLeadId++;
        }
      }
    });
    
    // Kiểm tra Queue để đếm lead đã gửi và lỗi
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
    
    logSystem('Giai đoạn 1 hoàn tất: ' + total + ' lead, ' + pending + ' chờ gửi, ' + sent + ' đã gửi', "SUCCESS");
    
    let warningMsg = '';
    if (missingLeadId > 0) {
      warningMsg = '\n\n⚠️ CÓ ' + missingLeadId + ' LEAD CHƯA CÓ ID!\nVui lòng chạy "Tạo Lead ID" trước.';
    }
    
    return {
      success: true,
      message: 'Phân tích hoàn tất!\n\n' +
              'Tổng lead: ' + total + '\n' +
              'Lead chờ gửi: ' + pending + '\n' +
              'Đã gửi: ' + sent + '\n' +
              'Lỗi: ' + error + '\n' +
              'RVA hoạt động: ' + activeRVA + '/' + rvaConfig.length +
              warningMsg
    };
    
  } catch (error) {
    logSystem('Lỗi Giai đoạn 1: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: 'Lỗi: ' + error.toString()
    };
  }
}

// ==========================================
// GIAI ĐOẠN 2: TẠO HÀNG ĐỢI
// ==========================================
function runStage2FromUI() {
  try {
    logSystem("Bắt đầu Giai đoạn 2 - Tạo hàng đợi");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    
    if (!leadSheet) {
      return {
        success: false,
        message: 'Không tìm thấy sheet "' + SYSTEM_CONFIG.SHEETS.LEADS + '"'
      };
    }
    
    const rvaConfig = getRVAConfig();
    const activeRVAs = rvaConfig.filter(function(rva) { return rva.active; });
    
    if (activeRVAs.length === 0) {
      return {
        success: false,
        message: "Không có RVA nào đang hoạt động!\n\nVui lòng cấu hình RVA trong sheet 'RVA_Config'"
      };
    }
    
    logSystem('Tìm thấy ' + activeRVAs.length + ' RVA đang hoạt động');
    
    const lastRow = leadSheet.getLastRow();
    const startRow = SYSTEM_CONFIG.LEADS_CONFIG.START_ROW;
    
    if (lastRow < startRow) {
      return {
        success: false,
        message: "Không có dữ liệu lead để xử lý"
      };
    }
    
    // Kiểm tra hoặc tạo Queue Sheet
    let queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    
    const data = leadSheet.getRange(startRow, 1, lastRow - startRow + 1, SYSTEM_CONFIG.LEADS_CONFIG.COL_PROJECT).getValues();
    
    // Lấy danh sách Lead ID đã gửi từ Queue (bao gồm cả Success và Error)
    const processedLeadIds = new Set();
    
    if (queueSheet && queueSheet.getLastRow() > 1) {
      const queueData = queueSheet.getRange(2, SYSTEM_CONFIG.QUEUE_CONFIG.COL_LEAD_ID, queueSheet.getLastRow() - 1, 1).getValues();
      queueData.forEach(function(row) {
        const leadId = String(row[0]).trim();
        if (leadId) {
          processedLeadIds.add(leadId);
        }
      });
      logSystem('Đã tìm thấy ' + processedLeadIds.size + ' Lead ID trong Queue (bao gồm cả Pending)');
    }
    
    const pendingLeads = [];
    let skippedDuplicate = 0;
    let missingLeadId = 0;
    
    data.forEach(function(row, index) {
      const leadId = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_LEAD_ID - 1];
      const fullName = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_NAME - 1];
      const phone = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_PHONE - 1];
      const need = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_NEED - 1];
      const project = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_PROJECT - 1];
      
      if (fullName && fullName !== "") {
        // Kiểm tra Lead ID
        if (!leadId || leadId === "") {
          missingLeadId++;
          logSystem('CẢNH BÁO: Lead "' + fullName + '" chưa có ID', "WARNING");
          return; // Bỏ qua lead này
        }
        
        // Kiểm tra trùng lặp - nếu Lead ID đã có trong Queue (dù Pending, Success hay Error)
        if (processedLeadIds.has(String(leadId).trim())) {
          skippedDuplicate++;
          logSystem('Bỏ qua Lead ID đã xử lý: ' + leadId, "INFO");
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
      let msg = "Không có lead nào cần gửi!\n\n";
      if (skippedDuplicate > 0) {
        msg += "Đã bỏ qua " + skippedDuplicate + " lead trùng lặp.";
      } else {
        msg += "Tất cả lead đã được xử lý.";
      }
      return {
        success: false,
        message: msg
      };
    }
    
    logSystem('Tìm thấy ' + pendingLeads.length + ' lead chờ gửi (bỏ qua ' + skippedDuplicate + ' lead trùng)');
    
    if (!queueSheet) {
      queueSheet = ss.insertSheet(SYSTEM_CONFIG.SHEETS.QUEUE);
      queueSheet.appendRow([
        "Timestamp",
        "RVA ID",
        "RVA Name",
        "Zalo Phone",
        "Zalo ID",
        "Lead ID",         // THÊM MỚI
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
        "Queue Row"        // THÊM MỚI - Số thứ tự row trong Queue
      ]);
      
      const headerRange = queueSheet.getRange("1:1");
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#667eea");
      headerRange.setFontColor("#ffffff");
      
      queueSheet.setColumnWidth(1, 150);
      queueSheet.setColumnWidth(6, 100);  // Lead ID
      queueSheet.setColumnWidth(7, 150);  // Lead Name
      queueSheet.setColumnWidth(11, 300); // Message
      queueSheet.setColumnWidth(12, 100); // Status
      
      logSystem("Đã tạo sheet Queue mới");
    }
    
    let queuedCount = 0;
    let rvaIndex = 0;
    
    logSystem('Bắt đầu phân chia ' + pendingLeads.length + ' lead cho ' + activeRVAs.length + ' RVA');
    
    pendingLeads.forEach(function(lead, index) {
      const currentRVA = activeRVAs[rvaIndex];
      
      // Format tin nhắn mới với Lead ID
      const message = 'LEAD ID: ' + lead.leadId + '\n' +
                     'Tên: ' + lead.fullName + '\n' +
                     'SĐT: ' + lead.phone + '\n' +
                     'Nhu cầu: ' + lead.need + '\n' +
                     'Dự án: ' + lead.project;
      
      // Tính Queue Row number (row 1 là header, data bắt đầu từ row 2)
      const queueRowNumber = queueSheet.getLastRow() + 1;
      
      queueSheet.appendRow([
        new Date(),
        currentRVA.id,
        currentRVA.name,
        currentRVA.phone,
        currentRVA.zaloid,
        lead.leadId,        // THÊM MỚI
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
        queueRowNumber      // THÊM MỚI - Queue Row Number
      ]);
      
      queuedCount++;
      rvaIndex = (rvaIndex + 1) % activeRVAs.length;
      
      if ((index + 1) % 50 === 0) {
        logSystem('Đã xử lý: ' + (index + 1) + '/' + pendingLeads.length + ' lead');
      }
    });
    
    try {
      const statusRange = queueSheet.getRange("L:L"); // Cập nhật từ K:K -> L:L
      
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
      logSystem('Không thể set conditional formatting: ' + e.toString(), "WARNING");
    }
    
    logSystem('Giai đoạn 2 hoàn tất - Đã thêm ' + queuedCount + ' lead vào queue', "SUCCESS");
    
    // Tự động đồng bộ Status về lead-mkt
    logSystem('Bắt đầu tự động đồng bộ Status về lead-mkt...', "INFO");
    const syncResult = syncStatusToLeadSheet();
    
    let resultMsg = 'Tạo hàng đợi thành công!\n\n' +
                   'Đã thêm: ' + queuedCount + ' lead\n' +
                   'Phân cho: ' + activeRVAs.length + ' RVA\n';
    
    if (skippedDuplicate > 0) {
      resultMsg += 'Bỏ qua: ' + skippedDuplicate + ' lead trùng\n';
    }
    
    if (syncResult.success) {
      resultMsg += '\n✅ Đã tự động cập nhật Status "Pending"\nvào sheet lead-mkt\n';
    }
    
    resultMsg += '\nTiếp theo: Chạy Giai đoạn 3 hoặc để n8n tự động xử lý\n\n' +
                'Kiểm tra sheet "' + SYSTEM_CONFIG.SHEETS.QUEUE + '" để xem chi tiết';
    
    return {
      success: true,
      message: resultMsg
    };
    
  } catch (error) {
    logSystem('Lỗi Giai đoạn 2: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: 'Lỗi: ' + error.toString()
    };
  }
}

// ==========================================
// GIAI ĐOẠN 3: KÍCH HOẠT GỬI
// ==========================================
function runStage3FromUI() {
  try {
    logSystem("Bắt đầu Giai đoạn 3");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    
    if (!queueSheet) {
      logSystem("Sheet Queue không tồn tại", "ERROR");
      return {
        success: false,
        message: "Sheet 'Zalo_Queue' không tồn tại!\n\nVui lòng chạy Giai đoạn 2 trước."
      };
    }
    
    const lastRow = queueSheet.getLastRow();
    
    if (lastRow <= 1) {
      logSystem("Queue sheet trống", "WARNING");
      return {
        success: false,
        message: "Hàng đợi trống!\n\nVui lòng chạy Giai đoạn 2 để tạo queue."
      };
    }
    
    logSystem('Queue có ' + (lastRow - 1) + ' dòng dữ liệu');
    
    const allData = queueSheet.getDataRange().getValues();
    const headers = allData[0];
    const statusCol = SYSTEM_CONFIG.QUEUE_CONFIG.COL_STATUS;
    const statusIndex = statusCol - 1;
    
    logSystem('Sample status (5 dòng đầu):');
    for (let i = 1; i < Math.min(6, allData.length); i++) {
      logSystem('  Row ' + (i + 1) + ': "' + allData[i][statusIndex] + '"');
    }
    
    let pendingCount = 0;
    let sentCount = 0;
    let errorCount = 0;
    let emptyCount = 0;
    const pendingRows = [];
    
    for (let i = 1; i < allData.length; i++) {
      const status = String(allData[i][statusIndex]).trim();
      
      if (status === "" || status === "null" || status === "undefined") {
        emptyCount++;
      } else if (status === "Pending") {
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
    
    logSystem('Thống kê: Pending=' + pendingCount + ', Sent=' + sentCount + ', Error=' + errorCount + ', Empty=' + emptyCount);
    
    if (pendingCount === 0) {
      if (emptyCount > 0) {
        return {
          success: false,
          message: 'Không có lead "Pending" nào!\n\n' +
                  'Có ' + emptyCount + ' dòng Status rỗng.\n' +
                  'Hãy kiểm tra sheet Queue.'
        };
      }
      
      return {
        success: false,
        message: 'Không có lead nào cần gửi!\n\n' +
                'Sent: ' + sentCount + '\n' +
                'Error: ' + errorCount
      };
    }
    
    logSystem('Đang gửi ' + pendingCount + ' lead đến n8n...');
    
    const n8nWebhookUrl = SYSTEM_CONFIG.N8N_CONFIG.WEBHOOK_URL;
    
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
      
      logSystem('n8n Response Code: ' + responseCode);
      logSystem('n8n Response: ' + responseBody);
      
      if (responseCode === 200) {
        logSystem('Webhook gửi thành công', "SUCCESS");
        
        return {
          success: true,
          message: 'Đã kích hoạt n8n thành công!\n\n' +
                  pendingCount + ' lead đang được xử lý\n' +
                  'n8n đang gửi tin nhắn Zalo\n\n' +
                  'Theo dõi tiến độ tại sheet "Zalo_Queue"\n' +
                  'Cột L (Status) sẽ tự động cập nhật thành "Success"\n' +
                  'khi n8n gửi tin nhắn thành công'
        };
        
      } else {
        logSystem('n8n trả về lỗi: ' + responseCode, "ERROR");
        
        return {
          success: false,
          message: 'Lỗi khi gọi n8n!\n\n' +
                  'Response Code: ' + responseCode + '\n' +
                  'Message: ' + responseBody + '\n\n' +
                  'Vui lòng kiểm tra:\n' +
                  '1. n8n workflow đang Active\n' +
                  '2. Webhook URL đúng\n' +
                  '3. Xem logs n8n để biết chi tiết'
        };
      }
      
    } catch (webhookError) {
      logSystem('Lỗi kết nối n8n: ' + webhookError.toString(), "ERROR");
      
      return {
        success: false,
        message: 'Không thể kết nối đến n8n!\n\n' +
                'Error: ' + webhookError.toString() + '\n\n' +
                'Vui lòng kiểm tra:\n' +
                '1. n8n server đang chạy\n' +
                '2. Webhook URL đúng\n' +
                '3. Network/firewall settings'
      };
    }
    
  } catch (error) {
    logSystem('Lỗi Giai đoạn 3: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: 'Lỗi: ' + error.toString()
    };
  }
}

// ==========================================
// ĐỒNG BỘ STATUS TỪ QUEUE VỀ LEAD-MKT
// ==========================================
function syncStatusToLeadSheet() {
  try {
    logSystem("Bắt đầu đồng bộ Status từ Queue về lead-mkt");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    const queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    
    if (!leadSheet) {
      return {
        success: false,
        message: 'Không tìm thấy sheet "' + SYSTEM_CONFIG.SHEETS.LEADS + '"'
      };
    }
    
    if (!queueSheet || queueSheet.getLastRow() <= 1) {
      return {
        success: false,
        message: "Queue trống, không có gì để đồng bộ"
      };
    }
    
    const lastRow = queueSheet.getLastRow();
    const queueData = queueSheet.getRange(2, 1, lastRow - 1, SYSTEM_CONFIG.QUEUE_CONFIG.COL_COLUMN_INDEX).getValues();
    
    let updatedCount = 0;
    let pendingCount = 0;
    let successCount = 0;
    let errorCount = 0;
    
    // Duyệt qua từng dòng trong Queue
    queueData.forEach(function(row) {
      const originalRow = row[SYSTEM_CONFIG.QUEUE_CONFIG.COL_ORIGINAL_ROW - 1];
      const columnIndex = row[SYSTEM_CONFIG.QUEUE_CONFIG.COL_COLUMN_INDEX - 1];
      const status = String(row[SYSTEM_CONFIG.QUEUE_CONFIG.COL_STATUS - 1]).trim();
      
      if (originalRow && columnIndex && status) {
        let statusToWrite = "";
        
        if (status.toLowerCase() === "pending") {
          statusToWrite = "Pending";
          pendingCount++;
        } else if (status.toLowerCase() === "success" || status.toLowerCase() === "sent") {
          statusToWrite = "Success";
          successCount++;
        } else if (status.toLowerCase() === "error" || status.toLowerCase().includes("error")) {
          statusToWrite = "Error";
          errorCount++;
        }
        
        if (statusToWrite) {
          try {
            leadSheet.getRange(originalRow, columnIndex).setValue(statusToWrite);
            updatedCount++;
          } catch (e) {
            logSystem('Lỗi cập nhật row ' + originalRow + ', col ' + columnIndex + ': ' + e.toString(), "ERROR");
          }
        }
      }
    });
    
    logSystem('Đồng bộ hoàn tất: ' + updatedCount + ' ô được cập nhật', "SUCCESS");
    
    return {
      success: true,
      message: 'Đồng bộ hoàn tất!\n\n' +
              'Đã cập nhật: ' + updatedCount + ' ô\n' +
              '  - Pending: ' + pendingCount + '\n' +
              '  - Success: ' + successCount + '\n' +
              '  - Error: ' + errorCount + '\n\n' +
              'Kiểm tra sheet "' + SYSTEM_CONFIG.SHEETS.LEADS + '"\n' +
              'các cột RVA (H, I, J)'
    };
    
  } catch (error) {
    logSystem('Lỗi đồng bộ Status: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: 'Lỗi: ' + error.toString()
    };
  }
}

// ==========================================
// TỰ ĐỘNG ĐỒNG BỘ SAU KHI TẠO QUEUE
// ==========================================
function autoSyncAfterQueue() {
  // Hàm này sẽ tự động gọi syncStatusToLeadSheet sau khi tạo Queue
  try {
    logSystem("Tự động đồng bộ Status sau khi tạo Queue", "INFO");
    const result = syncStatusToLeadSheet();
    if (result.success) {
      logSystem("Tự động đồng bộ thành công", "SUCCESS");
    }
    return result;
  } catch (error) {
    logSystem('Lỗi tự động đồng bộ: ' + error.toString(), "ERROR");
  }
}

// ==========================================
// DỌN DẸP QUEUE
// ==========================================
function cleanupQueueFromUI() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    
    if (!queueSheet) {
      return {
        success: false,
        message: "Sheet Queue không tồn tại"
      };
    }
    
    if (queueSheet.getLastRow() > 1) {
      queueSheet.deleteRows(2, queueSheet.getLastRow() - 1);
    }
    
    logSystem("Đã dọn dẹp queue", "SUCCESS");
    
    return {
      success: true,
      message: "Đã dọn dẹp queue thành công!"
    };
    
  } catch (error) {
    logSystem('Lỗi cleanup: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: 'Lỗi: ' + error.toString()
    };
  }
}

// ==========================================
// MENU
// ==========================================
function showMainMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'ZALO LEAD DISTRIBUTION SYSTEM',
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
    'CHỌN HÀNH ĐỘNG',
    'Nhập số tương ứng:\n\n' +
    '1. Xem thống kê hệ thống\n' +
    '2. Tạo Lead ID tự động\n' +
    '3. Chạy Giai đoạn 1 (Phân tích Lead)\n' +
    '4. Chạy Giai đoạn 2 (Tạo Queue)\n' +
    '5. Chạy Giai đoạn 3 (Kích hoạt gửi)\n' +
    '6. Đồng bộ Status về lead-mkt\n' +
    '7. Quản lý RVA\n' +
    '8. Dọn dẹp Queue\n' +
    '9. Xem Log\n\n' +
    '0. Thoát\n\n' +
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
      case '0':
        return;
      default:
        ui.alert('Cảnh báo', 'Lựa chọn không hợp lệ!', ui.ButtonSet.OK);
        showActionMenu();
    }
  }
}

function generateLeadIdMenu() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'TẠO LEAD ID',
    'Tự động tạo Lead ID cho các lead chưa có ID?\n\n' +
    'Định dạng: le01, le02, le03...',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm == ui.Button.YES) {
    const result = autoGenerateLeadIds();
    ui.alert(result.success ? 'Thành công' : 'Lỗi', result.message, ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function runStage1Menu() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'GIAI ĐOẠN 1',
    'Phân tích lead từ sheet. Tiếp tục?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm == ui.Button.YES) {
    const result = runStage1FromUI();
    ui.alert(result.success ? 'Thành công' : 'Lỗi', result.message, ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function runStage2Menu() {
  const ui = SpreadsheetApp.getUi();
  const result = runStage2FromUI();
  ui.alert(result.success ? 'Thành công' : 'Lỗi', result.message, ui.ButtonSet.OK);
  showActionMenu();
}

function runStage3Menu() {
  const ui = SpreadsheetApp.getUi();
  const result = runStage3FromUI();
  ui.alert(result.success ? 'Thành công' : 'Lỗi', result.message, ui.ButtonSet.OK);
  showActionMenu();
}

function syncStatusMenu() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'ĐỒNG BỘ STATUS',
    'Cập nhật status từ Queue về các cột RVA trong lead-mkt?\n\n' +
    'Pending → Pending\n' +
    'Success → Success\n' +
    'Error → Error',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm == ui.Button.YES) {
    const result = syncStatusToLeadSheet();
    ui.alert(result.success ? 'Thành công' : 'Lỗi', result.message, ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function manageRVAMenu() {
  const ui = SpreadsheetApp.getUi();
  const config = getRVAConfig();
  
  let msg = 'DANH SÁCH RVA:\n\n';
  config.forEach(function(rva, i) {
    msg += (i+1) + '. ' + rva.name + ' - ' + (rva.active ? 'Hoạt động' : 'Không hoạt động') + '\n';
  });
  
  ui.alert('RVA', msg, ui.ButtonSet.OK);
  showActionMenu();
}

function cleanupQueueMenu() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('Xóa Queue?', 'Xác nhận xóa?', ui.ButtonSet.YES_NO);
  
  if (confirm == ui.Button.YES) {
    const result = cleanupQueueFromUI();
    ui.alert(result.success ? 'Thành công' : 'Lỗi', result.message, ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function viewLogMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LOG);
  
  if (!logSheet || logSheet.getLastRow() <= 1) {
    ui.alert('Log', 'Chưa có log', ui.ButtonSet.OK);
  } else {
    ss.setActiveSheet(logSheet);
    ui.alert('Log', 'Sheet log đã mở', ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('⚡ Zalo System')
    .addItem('📊 Control Panel', 'showMainMenu')
    .addItem('📈 Thống kê', 'showControlPanel')
    .addSeparator()
    .addItem('🔖 Tạo Lead ID', 'generateLeadIdMenu')
    .addItem('1️⃣ Giai đoạn 1', 'runStage1Menu')
    .addItem('2️⃣ Giai đoạn 2', 'runStage2Menu')
    .addItem('3️⃣ Giai đoạn 3', 'runStage3Menu')
    .addSeparator()
    .addItem('🔄 Đồng bộ Status', 'syncStatusMenu')
    .addItem('🗑️ Dọn dẹp Queue', 'cleanupQueueMenu')
    .addToUi();
  
  logSystem("Menu khởi tạo - Có đồng bộ Status");
}
