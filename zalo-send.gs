// ==========================================
// ⚙️ CẤU HÌNH HỆ THỐNG
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
    COL_STT: 1,        // A
    COL_NAME: 2,       // B
    COL_PHONE: 3,      // C
    COL_NEED: 4,       // D
    COL_PROJECT: 5,    // E
    COL_RVA1: 7,       // G
    COL_RVA2: 8,       // H
    COL_RVA3: 9        // I
  },
  QUEUE_CONFIG: {
    COL_TIMESTAMP: 1,      // A
    COL_RVA_ID: 2,         // B
    COL_RVA_NAME: 3,       // C
    COL_ZALO_PHONE: 4,     // D
    COL_ZALO_ID: 5,        // E
    COL_LEAD_NAME: 6,      // F
    COL_LEAD_PHONE: 7,     // G
    COL_NEED: 8,           // H
    COL_PROJECT: 9,        // I
    COL_MESSAGE: 10,       // J
    COL_STATUS: 11,        // K 
    COL_ORIGINAL_ROW: 12,  // L
    COL_COLUMN_INDEX: 13,  // M
    COL_ERROR: 14,         // N
    COL_SENT_TIME: 15      // O
  }
};

// ==========================================
// 📊 API: LẤY THỐNG KÊ HỆ THỐNG
// ==========================================
function getSystemStats() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Thống kê Lead
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    let leadStats = {
      total: 0,
      pending: 0,
      sent: 0,
      error: 0
    };
    
    if (leadSheet && leadSheet.getLastRow() >= SYSTEM_CONFIG.LEADS_CONFIG.START_ROW) {
      const lastRow = leadSheet.getLastRow();
      const startRow = SYSTEM_CONFIG.LEADS_CONFIG.START_ROW;
      const data = leadSheet.getRange(startRow, 1, lastRow - startRow + 1, 9).getValues();
      
      data.forEach(row => {
        const fullName = row[1]; // Cột B
        if (fullName && fullName !== "") {
          leadStats.total++;
          
          const rva1 = row[6]; // Cột G
          const rva2 = row[7]; // Cột H
          const rva3 = row[8]; // Cột I
          
          if (rva1 === "" && rva2 === "" && rva3 === "") {
            leadStats.pending++;
          } else {
            leadStats.sent++;
          }
        }
      });
    }
    
    // 2. Thống kê Queue (ĐỌC CỘT K - 11)
    const queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    let queueStats = {
      total: 0,
      pending: 0,
      sent: 0,
      error: 0
    };
    
    if (queueSheet && queueSheet.getLastRow() > 1) {
      const statusCol = SYSTEM_CONFIG.QUEUE_CONFIG.COL_STATUS; // Cột 11 (K)
      const data = queueSheet.getRange(2, statusCol, queueSheet.getLastRow() - 1, 1).getValues();
      
      queueStats.total = data.length;
      data.forEach(row => {
        const status = String(row[0]).trim();
        if (status === "Pending") queueStats.pending++;
        else if (status === "Sent ✓" || status === "Sent") queueStats.sent++;
        else if (status === "Error") queueStats.error++;
      });
      
      Logger.log(`✅ Đọc Queue Status (cột K): Total=${queueStats.total}, Pending=${queueStats.pending}, Sent=${queueStats.sent}`);
    }
    
    // 3. Thống kê RVA
    const rvaConfig = getRVAConfig();
    const rvaStats = {
      total: rvaConfig.length,
      active: rvaConfig.filter(rva => rva.active).length,
      inactive: rvaConfig.filter(rva => !rva.active).length
    };
    
    // 4. Log gần nhất
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
    Logger.log("❌ Error in getSystemStats: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ==========================================
// 📋 API: LẤY CẤU HÌNH RVA
// ==========================================
function getRVAConfig() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.RVA_CONFIG);
    
    if (!configSheet) {
      // Config mặc định
      return [
        {
          id: "RV001",
          name: "RVA 1",
          phone: "0901234567",
          zaloid: "1243438522397465634",
          columnIndex: 7,
          active: true,
          note: ""
        },
        {
          id: "RV002",
          name: "RVA 2",
          phone: "0902345678",
          zaloid: "3837551771715024930",
          columnIndex: 8,
          active: true,
          note: ""
        },
        {
          id: "RV003",
          name: "RVA 3",
          phone: "0903456789",
          zaloid: "1504484729431570818",
          columnIndex: 9,
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
      columnIndex: Number(row[4]) || 7,
      active: row[5] === true || row[5] === "TRUE" || row[5] === "true",
      note: row[6] || ""
    })).filter(rva => rva.id !== "");
    
  } catch (error) {
    Logger.log("❌ Lỗi getRVAConfig: " + error.toString());
    return [];
  }
}

// ==========================================
// 📝 LOGGING SYSTEM
// ==========================================
function logSystem(message, level = "INFO") {
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
    
    Logger.log(`[${level}] ${message}`);
    
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
    
    const message = `
📊 THỐNG KÊ HỆ THỐNG
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📋 LEAD:
   • Tổng số: ${data.lead.total}
   • Chờ gửi: ${data.lead.pending}
   • Đã gửi: ${data.lead.sent}

📝 HÀNG ĐỢI:
   • Tổng số: ${data.queue.total}
   • Pending: ${data.queue.pending}
   • Đã gửi: ${data.queue.sent}
   • Lỗi: ${data.queue.error}

👥 RVA:
   • Tổng số: ${data.rva.total}
   • Hoạt động: ${data.rva.active}
   • Không hoạt động: ${data.rva.inactive}

⏰ Cập nhật: ${new Date(data.timestamp).toLocaleString('vi-VN')}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    `.trim();
    
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
        message: `❌ Không tìm thấy sheet "${SYSTEM_CONFIG.SHEETS.LEADS}"`
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
    
    const data = leadSheet.getRange(startRow, 1, lastRow - startRow + 1, 9).getValues();
    
    let total = 0;
    let pending = 0;
    let sent = 0;
    
    data.forEach(row => {
      const fullName = row[1];
      if (fullName && fullName !== "") {
        total++;
        
        const rva1 = row[6];
        const rva2 = row[7];
        const rva3 = row[8];
        
        if (rva1 === "" && rva2 === "" && rva3 === "") {
          pending++;
        } else {
          sent++;
        }
      }
    });
    
    const rvaConfig = getRVAConfig();
    const activeRVA = rvaConfig.filter(r => r.active).length;
    
    logSystem(`✅ Giai đoạn 1 hoàn tất: ${total} lead, ${pending} chờ gửi`, "SUCCESS");
    
    return {
      success: true,
      message: `✅ Phân tích hoàn tất!\n\n` +
              `📊 Tổng lead: ${total}\n` +
              `⏳ Lead chờ gửi: ${pending}\n` +
              `✓ Đã gửi: ${sent}\n` +
              `👥 RVA hoạt động: ${activeRVA}/${rvaConfig.length}`
    };
    
  } catch (error) {
    logSystem(`❌ Lỗi Giai đoạn 1: ${error.toString()}`, "ERROR");
    return {
      success: false,
      message: `❌ Lỗi: ${error.toString()}`
    };
  }
}

// ==========================================
// 🎯 GIAI ĐOẠN 2: TẠO HÀNG ĐỢI
// ==========================================
function runStage2FromUI() {
  try {
    logSystem("🔄 Bắt đầu Giai đoạn 2 - Tạo hàng đợi");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    
    if (!leadSheet) {
      return {
        success: false,
        message: `❌ Không tìm thấy sheet "${SYSTEM_CONFIG.SHEETS.LEADS}"`
      };
    }
    
    // Lấy cấu hình RVA
    const rvaConfig = getRVAConfig();
    const activeRVAs = rvaConfig.filter(rva => rva.active);
    
    if (activeRVAs.length === 0) {
      return {
        success: false,
        message: "❌ Không có RVA nào đang hoạt động!\n\nVui lòng cấu hình RVA trong sheet 'RVA_Config'"
      };
    }
    
    logSystem(`✅ Tìm thấy ${activeRVAs.length} RVA đang hoạt động`);
    
    // Đọc dữ liệu lead
    const lastRow = leadSheet.getLastRow();
    const startRow = SYSTEM_CONFIG.LEADS_CONFIG.START_ROW;
    
    if (lastRow < startRow) {
      return {
        success: false,
        message: "⚠️ Không có dữ liệu lead để xử lý"
      };
    }
    
    const data = leadSheet.getRange(startRow, 1, lastRow - startRow + 1, 9).getValues();
    
    // Lọc lead chưa gửi
    const pendingLeads = [];
    
    data.forEach((row, index) => {
      const fullName = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_NAME - 1];
      const phone = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_PHONE - 1];
      const need = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_NEED - 1];
      const project = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_PROJECT - 1];
      const rva1 = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_RVA1 - 1];
      const rva2 = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_RVA2 - 1];
      const rva3 = row[SYSTEM_CONFIG.LEADS_CONFIG.COL_RVA3 - 1];
      
      if (fullName && fullName !== "" && rva1 === "" && rva2 === "" && rva3 === "") {
        pendingLeads.push({
          rowIndex: startRow + index,
          fullName: fullName,
          phone: phone || "",
          need: need || "",
          project: project || ""
        });
      }
    });
    
    if (pendingLeads.length === 0) {
      return {
        success: false,
        message: "⚠️ Không có lead nào cần gửi!\n\nTất cả lead đã được xử lý."
      };
    }
    
    logSystem(`📊 Tìm thấy ${pendingLeads.length} lead chờ gửi`);
    
    // Tạo hoặc lấy Queue Sheet
    let queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    
    if (!queueSheet) {
      queueSheet = ss.insertSheet(SYSTEM_CONFIG.SHEETS.QUEUE);
      queueSheet.appendRow([
        "Timestamp",      // A
        "RVA ID",         // B
        "RVA Name",       // C
        "Zalo Phone",     // D
        "Zalo ID",        // E
        "Lead Name",      // F
        "Lead Phone",     // G
        "Need",           // H
        "Project",        // I
        "Message",        // J
        "Status",         // K ✅
        "Original Row",   // L
        "Column Index",   // M
        "Error",          // N
        "Sent Time"       // O
      ]);
      
      // Format header
      const headerRange = queueSheet.getRange("1:1");
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#667eea");
      headerRange.setFontColor("#ffffff");
      
      // Set column widths
      queueSheet.setColumnWidth(1, 150);   // Timestamp
      queueSheet.setColumnWidth(6, 150);   // Lead Name
      queueSheet.setColumnWidth(10, 300);  // Message
      queueSheet.setColumnWidth(11, 100);  // Status
      
      logSystem("✅ Đã tạo sheet Queue mới");
    }
    
    // Phân chia lead (Round-robin)
    let queuedCount = 0;
    let rvaIndex = 0;
    
    logSystem(`📦 Bắt đầu phân chia ${pendingLeads.length} lead cho ${activeRVAs.length} RVA`);
    
    pendingLeads.forEach((lead, index) => {
      const currentRVA = activeRVAs[rvaIndex];
      
      const message = `📋 LEAD MỚI
👤 Tên: ${lead.fullName}
📞 SĐT: ${lead.phone}
💼 Nhu cầu: ${lead.need}
🏢 Dự án: ${lead.project}`;
      
      // Thêm vào queue (ĐÚNG THỨ TỰ CỘT)
      queueSheet.appendRow([
        new Date(),                 // A: Timestamp
        currentRVA.id,              // B: RVA ID
        currentRVA.name,            // C: RVA Name
        currentRVA.phone,           // D: Zalo Phone
        currentRVA.zaloid,          // E: Zalo ID
        lead.fullName,              // F: Lead Name
        lead.phone,                 // G: Lead Phone
        lead.need,                  // H: Need
        lead.project,               // I: Project
        message,                    // J: Message
        "Pending",                  // K: Status ✅
        lead.rowIndex,              // L: Original Row
        currentRVA.columnIndex,     // M: Column Index
        "",                         // N: Error
        ""                          // O: Sent Time
      ]);
      
      queuedCount++;
      rvaIndex = (rvaIndex + 1) % activeRVAs.length;
      
      if ((index + 1) % 50 === 0) {
        logSystem(`📊 Đã xử lý: ${index + 1}/${pendingLeads.length} lead`);
      }
    });
    
    // Format Status column
    try {
      const statusRange = queueSheet.getRange("K:K");
      
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
      logSystem(`⚠️ Không thể set conditional formatting: ${e.toString()}`, "WARNING");
    }
    
    logSystem(`✅ Giai đoạn 2 hoàn tất - Đã thêm ${queuedCount} lead vào queue`, "SUCCESS");
    
    return {
      success: true,
      message: `✅ Tạo hàng đợi thành công!\n\n` +
              `📦 Đã thêm: ${queuedCount} lead\n` +
              `👥 Phân cho: ${activeRVAs.length} RVA\n\n` +
              `➡️ Tiếp theo: Chạy Giai đoạn 3 hoặc để n8n tự động xử lý\n\n` +
              `📋 Kiểm tra sheet "${SYSTEM_CONFIG.SHEETS.QUEUE}" để xem chi tiết`
    };
    
  } catch (error) {
    logSystem(`❌ Lỗi Giai đoạn 2: ${error.toString()}`, "ERROR");
    return {
      success: false,
      message: `❌ Lỗi: ${error.toString()}`
    };
  }
}

// ==========================================
// 🎯 GIAI ĐOẠN 3: KÍCH HOẠT GỬI
// ==========================================
function runStage3FromUI() {
  try {
    logSystem("🔄 Bắt đầu Giai đoạn 3");
    
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
    
    logSystem(`📊 Queue có ${lastRow - 1} dòng dữ liệu`);
    
    // Đọc toàn bộ dữ liệu từ queue
    const allData = queueSheet.getDataRange().getValues();
    const headers = allData[0];
    const statusCol = SYSTEM_CONFIG.QUEUE_CONFIG.COL_STATUS;
    const statusIndex = statusCol - 1; // Convert to 0-based index
    
    // Log sample
    logSystem(`🔍 Sample status (5 dòng đầu):`);
    for (let i = 1; i < Math.min(6, allData.length); i++) {
      logSystem(`  Row ${i + 1}: "${allData[i][statusIndex]}"`);
    }
    
    // Lọc và đếm status
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
        // Thêm vào danh sách pending với row number và full data
        const rowData = {};
        headers.forEach((header, index) => {
          rowData[header] = allData[i][index];
        });
        
        pendingRows.push({
          rowNumber: i + 1, // +1 vì row trong sheet bắt đầu từ 1
          rowData: rowData
        });
      } else if (status === "Sent ✓" || status === "Sent") {
        sentCount++;
      } else if (status === "Error") {
        errorCount++;
      }
    }
    
    logSystem(`📊 Thống kê: Pending=${pendingCount}, Sent=${sentCount}, Error=${errorCount}, Empty=${emptyCount}`);
    
    if (pendingCount === 0) {
      if (emptyCount > 0) {
        return {
          success: false,
          message: `⚠️ Không có lead 'Pending' nào!\n\n` +
                  `Có ${emptyCount} dòng Status rỗng.\n` +
                  `Hãy kiểm tra sheet Queue.`
        };
      }
      
      return {
        success: false,
        message: `⚠️ Không có lead nào cần gửi!\n\n` +
                `• Sent: ${sentCount}\n` +
                `• Error: ${errorCount}`
      };
    }
    
    // ==========================================
    // GỬI REQUEST ĐẾN N8N WEBHOOK
    // ==========================================
    
    logSystem(`🚀 Đang gửi ${pendingCount} lead đến n8n...`);
    
    const n8nWebhookUrl = 'https://n8n.rever.io.vn/webhook/zalo-trigger'; 
    
    // Chuẩn bị payload
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
      
      logSystem(`📡 n8n Response Code: ${responseCode}`);
      logSystem(`📡 n8n Response: ${responseBody}`);
      
      if (responseCode === 200) {
        logSystem(`✅ Webhook gửi thành công`, "SUCCESS");
        
        return {
          success: true,
          message: `✅ Đã kích hoạt n8n thành công!\n\n` +
                  `📊 ${pendingCount} lead đang được xử lý\n` +
                  `🤖 n8n đang gửi tin nhắn Zalo\n\n` +
                  `📋 Theo dõi tiến độ tại sheet "Zalo_Queue"\n` +
                  `Cột K (Status) sẽ chuyển từ "Pending" → "Sent ✓"\n` +
                  `Cột P (Message ID) sẽ hiển thị ID tin nhắn`
        };
        
      } else {
        logSystem(`❌ n8n trả về lỗi: ${responseCode}`, "ERROR");
        
        return {
          success: false,
          message: `❌ Lỗi khi gọi n8n!\n\n` +
                  `Response Code: ${responseCode}\n` +
                  `Message: ${responseBody}\n\n` +
                  `Vui lòng kiểm tra:\n` +
                  `1. n8n workflow đang Active\n` +
                  `2. Webhook URL đúng\n` +
                  `3. Xem logs n8n để biết chi tiết`
        };
      }
      
    } catch (webhookError) {
      logSystem(`❌ Lỗi kết nối n8n: ${webhookError.toString()}`, "ERROR");
      
      return {
        success: false,
        message: `❌ Không thể kết nối đến n8n!\n\n` +
                `Error: ${webhookError.toString()}\n\n` +
                `Vui lòng kiểm tra:\n` +
                `1. n8n server đang chạy\n` +
                `2. Webhook URL đúng\n` +
                `3. Network/firewall settings`
      };
    }
    
  } catch (error) {
    logSystem(`❌ Lỗi Giai đoạn 3: ${error.toString()}`, "ERROR");
    return {
      success: false,
      message: `❌ Lỗi: ${error.toString()}`
    };
  }
}


// ==========================================
// HÀM KIỂM TRA N8N ONLINE (OPTIONAL)
// ==========================================

function checkN8nHealth() {
  const healthCheckUrl = 'YOUR_N8N_WEBHOOK_URL_HERE/health'; // ← Thay đổi URL
  
  try {
    const response = UrlFetchApp.fetch(healthCheckUrl, {
      method: 'get',
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      logSystem('✅ n8n is online');
      return true;
    } else {
      logSystem(`⚠️ n8n health check failed: ${response.getResponseCode()}`);
      return false;
    }
  } catch (error) {
    logSystem(`❌ n8n health check error: ${error.toString()}`);
    return false;
  }
}


// ==========================================
// HÀM ENHANCED VỚI HEALTH CHECK
// ==========================================

function runStage3FromUIEnhanced() {
  try {
    logSystem("🔄 Bắt đầu Giai đoạn 3");
    
    // Kiểm tra n8n online trước
    if (!checkN8nHealth()) {
      return {
        success: false,
        message: `❌ n8n đang offline!\n\n` +
                `Vui lòng:\n` +
                `1. Kiểm tra n8n server đang chạy\n` +
                `2. Workflow đã Active\n` +
                `3. Thử lại sau`
      };
    }
    
    // Tiếp tục với logic cũ...
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    
    // ... (phần code giống như trên)
    
  } catch (error) {
    logSystem(`❌ Lỗi Giai đoạn 3: ${error.toString()}`, "ERROR");
    return {
      success: false,
      message: `❌ Lỗi: ${error.toString()}`
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
    
    logSystem("✅ Đã dọn dẹp queue", "SUCCESS");
    
    return {
      success: true,
      message: "✅ Đã dọn dẹp queue thành công!"
    };
    
  } catch (error) {
    logSystem(`❌ Lỗi cleanup: ${error.toString()}`, "ERROR");
    return {
      success: false,
      message: `❌ Lỗi: ${error.toString()}`
    };
  }
}

// ==========================================
// 📋 MENU
// ==========================================
function showMainMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '🎯 ZALO LEAD DISTRIBUTION SYSTEM',
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
    `Nhập số tương ứng:

1. Xem thống kê hệ thống
2. Chạy Giai đoạn 1 (Phân tích Lead)
3. Chạy Giai đoạn 2 (Tạo Queue)
4. Chạy Giai đoạn 3 (Kích hoạt gửi)
5. Quản lý RVA
6. Dọn dẹp Queue
7. Xem Log

0. Thoát

Nhập lựa chọn:`,
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
        runStage1Menu();
        break;
      case '3':
        runStage2Menu();
        break;
      case '4':
        runStage3Menu();
        break;
      case '5':
        manageRVAMenu();
        break;
      case '6':
        cleanupQueueMenu();
        break;
      case '7':
        viewLogMenu();
        break;
      case '0':
        return;
      default:
        ui.alert('⚠️', 'Lựa chọn không hợp lệ!', ui.ButtonSet.OK);
        showActionMenu();
    }
  }
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
    ui.alert(result.success ? '✅' : '❌', result.message, ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function runStage2Menu() {
  const ui = SpreadsheetApp.getUi();
  const result = runStage2FromUI();
  ui.alert(result.success ? '✅' : '❌', result.message, ui.ButtonSet.OK);
  showActionMenu();
}

function runStage3Menu() {
  const ui = SpreadsheetApp.getUi();
  const result = runStage3FromUI();
  ui.alert(result.success ? '✅' : '❌', result.message, ui.ButtonSet.OK);
  showActionMenu();
}

function manageRVAMenu() {
  const ui = SpreadsheetApp.getUi();
  const config = getRVAConfig();
  
  let msg = '👥 DANH SÁCH RVA:\n\n';
  config.forEach((rva, i) => {
    msg += `${i+1}. ${rva.name} - ${rva.active ? '✅' : '❌'}\n`;
  });
  
  ui.alert('👥 RVA', msg, ui.ButtonSet.OK);
  showActionMenu();
}

function cleanupQueueMenu() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('🧹 Xóa Queue?', 'Xác nhận xóa?', ui.ButtonSet.YES_NO);
  
  if (confirm == ui.Button.YES) {
    const result = cleanupQueueFromUI();
    ui.alert(result.success ? '✅' : '❌', result.message, ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function viewLogMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LOG);
  
  if (!logSheet || logSheet.getLastRow() <= 1) {
    ui.alert('📋 Log', 'Chưa có log', ui.ButtonSet.OK);
  } else {
    ss.setActiveSheet(logSheet);
    ui.alert('📋 Log', 'Sheet log đã mở', ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🎯 Zalo System')
    .addItem('🎨 Control Panel', 'showMainMenu')
    .addItem('📊 Thống kê', 'showControlPanel')
    .addSeparator()
    .addItem('1️⃣ Giai đoạn 1', 'runStage1Menu')
    .addItem('2️⃣ Giai đoạn 2', 'runStage2Menu')
    .addItem('3️⃣ Giai đoạn 3', 'runStage3Menu')
    .addToUi();
  
  logSystem("✅ Menu khởi tạo");
}
