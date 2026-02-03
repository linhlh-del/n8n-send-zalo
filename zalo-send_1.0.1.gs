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
    COL_NAME: 2,
    COL_PHONE: 3,
    COL_NEED: 4,
    COL_PROJECT: 5,
    COL_RVA1: 7,
    COL_RVA2: 8,
    COL_RVA3: 9
  },
  QUEUE_CONFIG: {
    COL_TIMESTAMP: 1,
    COL_RVA_ID: 2,
    COL_RVA_NAME: 3,
    COL_ZALO_PHONE: 4,
    COL_ZALO_ID: 5,
    COL_LEAD_NAME: 6,
    COL_LEAD_PHONE: 7,
    COL_NEED: 8,
    COL_PROJECT: 9,
    COL_MESSAGE: 10,
    COL_STATUS: 11,
    COL_ORIGINAL_ROW: 12,
    COL_COLUMN_INDEX: 13,
    COL_ERROR: 14,
    COL_SENT_TIME: 15
  },
  N8N_CONFIG: {
    WEBHOOK_URL: 'https://n8n.rever.io.vn/webhook/zalo-trigger',
    RESULT_COL: 17
  }
};

// ==========================================
// API: LAY THONG KE HE THONG
// ==========================================
function getSystemStats() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
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
        const fullName = row[1];
        if (fullName && fullName !== "") {
          leadStats.total++;
          
          const rva1 = row[6];
          const rva2 = row[7];
          const rva3 = row[8];
          
          if (rva1 === "" && rva2 === "" && rva3 === "") {
            leadStats.pending++;
          } else {
            leadStats.sent++;
          }
        }
      });
    }
    
    const queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
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
      
      Logger.log('Đọc Queue Status (cột K): Total=' + queueStats.total + ', Pending=' + queueStats.pending + ', Sent=' + queueStats.sent);
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
// API: LAY CAU HINH RVA
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
    Logger.log("Loi getRVAConfig: " + error.toString());
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
    Logger.log("Loi logging: " + error.toString());
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
      ui.alert('Loi', 'Khong the lay thong tin he thong:\n' + stats.error, ui.ButtonSet.OK);
      return;
    }
    
    const data = stats.data;
    
    const message = 
      'THONG KE HE THONG\n' +
      '==========================================\n\n' +
      'LEAD:\n' +
      '   Tong so: ' + data.lead.total + '\n' +
      '   Cho gui: ' + data.lead.pending + '\n' +
      '   Da gui: ' + data.lead.sent + '\n\n' +
      'HANG DOI:\n' +
      '   Tong so: ' + data.queue.total + '\n' +
      '   Pending: ' + data.queue.pending + '\n' +
      '   Da gui: ' + data.queue.sent + '\n' +
      '   Loi: ' + data.queue.error + '\n\n' +
      'RVA:\n' +
      '   Tong so: ' + data.rva.total + '\n' +
      '   Hoat dong: ' + data.rva.active + '\n' +
      '   Khong hoat dong: ' + data.rva.inactive + '\n\n' +
      'Cap nhat: ' + new Date(data.timestamp).toLocaleString('vi-VN') + '\n\n' +
      '==========================================';
    
    ui.alert('Zalo Lead Distribution - Control Panel', message, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Loi', 'Loi trong showControlPanel:\n' + error.toString(), ui.ButtonSet.OK);
    Logger.log("Error in showControlPanel: " + error.toString());
  }
}

// ==========================================
// GIAI DOAN 1: PHAN TICH LEAD
// ==========================================
function runStage1FromUI() {
  try {
    logSystem("Bat dau Giai doan 1");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    
    if (!leadSheet) {
      return {
        success: false,
        message: 'Khong tim thay sheet "' + SYSTEM_CONFIG.SHEETS.LEADS + '"'
      };
    }
    
    const lastRow = leadSheet.getLastRow();
    const startRow = SYSTEM_CONFIG.LEADS_CONFIG.START_ROW;
    
    if (lastRow < startRow) {
      return {
        success: false,
        message: "Khong co du lieu lead de phan tich"
      };
    }
    
    const data = leadSheet.getRange(startRow, 1, lastRow - startRow + 1, 9).getValues();
    
    let total = 0;
    let pending = 0;
    let sent = 0;
    
    data.forEach(function(row) {
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
    const activeRVA = rvaConfig.filter(function(r) { return r.active; }).length;
    
    logSystem('Giai doan 1 hoan tat: ' + total + ' lead, ' + pending + ' cho gui', "SUCCESS");
    
    return {
      success: true,
      message: 'Phan tich hoan tat!\n\n' +
              'Tong lead: ' + total + '\n' +
              'Lead cho gui: ' + pending + '\n' +
              'Da gui: ' + sent + '\n' +
              'RVA hoat dong: ' + activeRVA + '/' + rvaConfig.length
    };
    
  } catch (error) {
    logSystem('Loi Giai doan 1: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: 'Loi: ' + error.toString()
    };
  }
}

// ==========================================
// GIAI DOAN 2: TAO HANG DOI
// ==========================================
function runStage2FromUI() {
  try {
    logSystem("Bat dau Giai doan 2 - Tao hang doi");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leadSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LEADS);
    
    if (!leadSheet) {
      return {
        success: false,
        message: 'Khong tim thay sheet "' + SYSTEM_CONFIG.SHEETS.LEADS + '"'
      };
    }
    
    const rvaConfig = getRVAConfig();
    const activeRVAs = rvaConfig.filter(function(rva) { return rva.active; });
    
    if (activeRVAs.length === 0) {
      return {
        success: false,
        message: "Khong co RVA nao dang hoat dong!\n\nVui long cau hinh RVA trong sheet 'RVA_Config'"
      };
    }
    
    logSystem('Tim thay ' + activeRVAs.length + ' RVA dang hoat dong');
    
    const lastRow = leadSheet.getLastRow();
    const startRow = SYSTEM_CONFIG.LEADS_CONFIG.START_ROW;
    
    if (lastRow < startRow) {
      return {
        success: false,
        message: "Khong co du lieu lead de xu ly"
      };
    }
    
    const data = leadSheet.getRange(startRow, 1, lastRow - startRow + 1, 9).getValues();
    
    const pendingLeads = [];
    
    data.forEach(function(row, index) {
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
        message: "Khong co lead nao can gui!\n\nTat ca lead da duoc xu ly."
      };
    }
    
    logSystem('Tim thay ' + pendingLeads.length + ' lead cho gui');
    
    let queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    
    if (!queueSheet) {
      queueSheet = ss.insertSheet(SYSTEM_CONFIG.SHEETS.QUEUE);
      queueSheet.appendRow([
        "Timestamp",
        "RVA ID",
        "RVA Name",
        "Zalo Phone",
        "Zalo ID",
        "Lead Name",
        "Lead Phone",
        "Need",
        "Project",
        "Message",
        "Status",
        "Original Row",
        "Column Index",
        "Error",
        "Sent Time"
      ]);
      
      const headerRange = queueSheet.getRange("1:1");
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#667eea");
      headerRange.setFontColor("#ffffff");
      
      queueSheet.setColumnWidth(1, 150);
      queueSheet.setColumnWidth(6, 150);
      queueSheet.setColumnWidth(10, 300);
      queueSheet.setColumnWidth(11, 100);
      
      logSystem("Da tao sheet Queue moi");
    }
    
    let queuedCount = 0;
    let rvaIndex = 0;
    
    logSystem('Bat dau phan chia ' + pendingLeads.length + ' lead cho ' + activeRVAs.length + ' RVA');
    
    pendingLeads.forEach(function(lead, index) {
      const currentRVA = activeRVAs[rvaIndex];
      
      const message = 'LEAD MOI\n' +
                     'Ten: ' + lead.fullName + '\n' +
                     'SDT: ' + lead.phone + '\n' +
                     'Nhu cau: ' + lead.need + '\n' +
                     'Du an: ' + lead.project;
      
      queueSheet.appendRow([
        new Date(),
        currentRVA.id,
        currentRVA.name,
        currentRVA.phone,
        currentRVA.zaloid,
        lead.fullName,
        lead.phone,
        lead.need,
        lead.project,
        message,
        "Pending",
        lead.rowIndex,
        currentRVA.columnIndex,
        "",
        ""
      ]);
      
      queuedCount++;
      rvaIndex = (rvaIndex + 1) % activeRVAs.length;
      
      if ((index + 1) % 50 === 0) {
        logSystem('Da xu ly: ' + (index + 1) + '/' + pendingLeads.length + ' lead');
      }
    });
    
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
      logSystem('Khong the set conditional formatting: ' + e.toString(), "WARNING");
    }
    
    logSystem('Giai doan 2 hoan tat - Da them ' + queuedCount + ' lead vao queue', "SUCCESS");
    
    return {
      success: true,
      message: 'Tao hang doi thanh cong!\n\n' +
              'Da them: ' + queuedCount + ' lead\n' +
              'Phan cho: ' + activeRVAs.length + ' RVA\n\n' +
              'Tiep theo: Chay Giai doan 3 hoac de n8n tu dong xu ly\n\n' +
              'Kiem tra sheet "' + SYSTEM_CONFIG.SHEETS.QUEUE + '" de xem chi tiet'
    };
    
  } catch (error) {
    logSystem('Loi Giai doan 2: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: 'Loi: ' + error.toString()
    };
  }
}

// ==========================================
// GIAI DOAN 3: KICH HOAT GUI
// ==========================================
function runStage3FromUI() {
  try {
    logSystem("Bat dau Giai doan 3");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    
    if (!queueSheet) {
      logSystem("Sheet Queue khong ton tai", "ERROR");
      return {
        success: false,
        message: "Sheet 'Zalo_Queue' khong ton tai!\n\nVui long chay Giai doan 2 truoc."
      };
    }
    
    const lastRow = queueSheet.getLastRow();
    
    if (lastRow <= 1) {
      logSystem("Queue sheet trong", "WARNING");
      return {
        success: false,
        message: "Hang doi trong!\n\nVui long chay Giai doan 2 de tao queue."
      };
    }
    
    logSystem('Queue co ' + (lastRow - 1) + ' dong du lieu');
    
    const allData = queueSheet.getDataRange().getValues();
    const headers = allData[0];
    const statusCol = SYSTEM_CONFIG.QUEUE_CONFIG.COL_STATUS;
    const statusIndex = statusCol - 1;
    
    logSystem('Sample status (5 dong dau):');
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
    
    logSystem('Thong ke: Pending=' + pendingCount + ', Sent=' + sentCount + ', Error=' + errorCount + ', Empty=' + emptyCount);
    
    if (pendingCount === 0) {
      if (emptyCount > 0) {
        return {
          success: false,
          message: 'Khong co lead "Pending" nao!\n\n' +
                  'Co ' + emptyCount + ' dong Status rong.\n' +
                  'Hay kiem tra sheet Queue.'
        };
      }
      
      return {
        success: false,
        message: 'Khong co lead nao can gui!\n\n' +
                'Sent: ' + sentCount + '\n' +
                'Error: ' + errorCount
      };
    }
    
    logSystem('Dang gui ' + pendingCount + ' lead den n8n...');
    
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
        logSystem('Webhook gui thanh cong', "SUCCESS");
        
        return {
          success: true,
          message: 'Da kich hoat n8n thanh cong!\n\n' +
                  pendingCount + ' lead dang duoc xu ly\n' +
                  'n8n dang gui tin nhan Zalo\n\n' +
                  'Theo doi tien do tai sheet "Zalo_Queue"\n' +
                  'Cot K (Status) se chuyen tu "Pending" sang "Sent"\n' +
                  'Cot P (Message ID) se hien thi ID tin nhan'
        };
        
      } else {
        logSystem('n8n tra ve loi: ' + responseCode, "ERROR");
        
        return {
          success: false,
          message: 'Loi khi goi n8n!\n\n' +
                  'Response Code: ' + responseCode + '\n' +
                  'Message: ' + responseBody + '\n\n' +
                  'Vui long kiem tra:\n' +
                  '1. n8n workflow dang Active\n' +
                  '2. Webhook URL dung\n' +
                  '3. Xem logs n8n de biet chi tiet'
        };
      }
      
    } catch (webhookError) {
      logSystem('Loi ket noi n8n: ' + webhookError.toString(), "ERROR");
      
      return {
        success: false,
        message: 'Khong the ket noi den n8n!\n\n' +
                'Error: ' + webhookError.toString() + '\n\n' +
                'Vui long kiem tra:\n' +
                '1. n8n server dang chay\n' +
                '2. Webhook URL dung\n' +
                '3. Network/firewall settings'
      };
    }
    
  } catch (error) {
    logSystem('Loi Giai doan 3: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: 'Loi: ' + error.toString()
    };
  }
}

// ==========================================
// GIAI DOAN 4: DONG BO STATUS TU N8N
// ==========================================
function syncStatusFromN8n() {
  try {
    logSystem("Bat dau dong bo Status tu cot Q sang cot K");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    
    if (!queueSheet) {
      return {
        success: false,
        message: "Sheet Queue khong ton tai"
      };
    }
    
    const lastRow = queueSheet.getLastRow();
    
    if (lastRow <= 1) {
      return {
        success: false,
        message: "Queue trong"
      };
    }
    
    const statusCol = SYSTEM_CONFIG.QUEUE_CONFIG.COL_STATUS;
    const n8nResultCol = SYSTEM_CONFIG.N8N_CONFIG.RESULT_COL;
    
    const statusData = queueSheet.getRange(2, statusCol, lastRow - 1, 1).getValues();
    const n8nResultData = queueSheet.getRange(2, n8nResultCol, lastRow - 1, 1).getValues();
    
    let updatedCount = 0;
    let skippedCount = 0;
    const updatesToApply = [];
    
    for (let i = 0; i < statusData.length; i++) {
      const currentStatus = String(statusData[i][0]).trim();
      const n8nResult = String(n8nResultData[i][0]).trim();
      
      if (n8nResult.toLowerCase().includes("success") && currentStatus === "Pending") {
        updatesToApply.push({
          row: i + 2,
          oldStatus: currentStatus,
          newStatus: "Sent"
        });
        updatedCount++;
      } else {
        skippedCount++;
      }
    }
    
    if (updatesToApply.length > 0) {
      logSystem('Dang cap nhat ' + updatesToApply.length + ' dong...');
      
      updatesToApply.forEach(function(update) {
        queueSheet.getRange(update.row, statusCol).setValue(update.newStatus);
      });
      
      logSystem('Da cap nhat ' + updatedCount + ' dong tu Pending sang Sent', "SUCCESS");
    }
    
    return {
      success: true,
      message: 'Dong bo hoan tat!\n\n' +
              'Da cap nhat: ' + updatedCount + ' dong\n' +
              'Bo qua: ' + skippedCount + ' dong\n\n' +
              'Kiem tra sheet "Zalo_Queue" cot K'
    };
    
  } catch (error) {
    logSystem('Loi dong bo Status: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: 'Loi: ' + error.toString()
    };
  }
}

// ==========================================
// HAM TU DONG DONG BO (NANG CAO)
// ==========================================
function autoSyncStatusFromN8n() {
  try {
    logSystem("Auto-sync: Bat dau kiem tra va dong bo");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    
    if (!queueSheet || queueSheet.getLastRow() <= 1) {
      logSystem("Auto-sync: Queue trong, bo qua");
      return;
    }
    
    const lastRow = queueSheet.getLastRow();
    const statusCol = SYSTEM_CONFIG.QUEUE_CONFIG.COL_STATUS;
    const n8nResultCol = SYSTEM_CONFIG.N8N_CONFIG.RESULT_COL;
    
    const statusData = queueSheet.getRange(2, statusCol, lastRow - 1, 1).getValues();
    const n8nResultData = queueSheet.getRange(2, n8nResultCol, lastRow - 1, 1).getValues();
    
    let updatedCount = 0;
    
    for (let i = 0; i < statusData.length; i++) {
      const currentStatus = String(statusData[i][0]).trim();
      const n8nResult = String(n8nResultData[i][0]).trim();
      
      if (n8nResult.toLowerCase().includes("success") && currentStatus === "Pending") {
        queueSheet.getRange(i + 2, statusCol).setValue("Sent");
        updatedCount++;
      }
    }
    
    if (updatedCount > 0) {
      logSystem('Auto-sync: Da cap nhat ' + updatedCount + ' dong', "SUCCESS");
    } else {
      logSystem("Auto-sync: Khong co gi can cap nhat");
    }
    
  } catch (error) {
    logSystem('Auto-sync error: ' + error.toString(), "ERROR");
  }
}

// ==========================================
// DON DEP QUEUE
// ==========================================
function cleanupQueueFromUI() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const queueSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.QUEUE);
    
    if (!queueSheet) {
      return {
        success: false,
        message: "Sheet Queue khong ton tai"
      };
    }
    
    if (queueSheet.getLastRow() > 1) {
      queueSheet.deleteRows(2, queueSheet.getLastRow() - 1);
    }
    
    logSystem("Da don dep queue", "SUCCESS");
    
    return {
      success: true,
      message: "Da don dep queue thanh cong!"
    };
    
  } catch (error) {
    logSystem('Loi cleanup: ' + error.toString(), "ERROR");
    return {
      success: false,
      message: 'Loi: ' + error.toString()
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
    'Chon OK de mo menu dieu khien',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response == ui.Button.OK) {
    showActionMenu();
  }
}

function showActionMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'CHON HANH DONG',
    'Nhap so tuong ung:\n\n' +
    '1. Xem thong ke he thong\n' +
    '2. Chay Giai doan 1 (Phan tich Lead)\n' +
    '3. Chay Giai doan 2 (Tao Queue)\n' +
    '4. Chay Giai doan 3 (Kich hoat gui)\n' +
    '5. Dong bo Status (Giai doan 4)\n' +
    '6. Quan ly RVA\n' +
    '7. Don dep Queue\n' +
    '8. Xem Log\n\n' +
    '0. Thoat\n\n' +
    'Nhap lua chon:',
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
        syncStatusMenu();
        break;
      case '6':
        manageRVAMenu();
        break;
      case '7':
        cleanupQueueMenu();
        break;
      case '8':
        viewLogMenu();
        break;
      case '0':
        return;
      default:
        ui.alert('Canh bao', 'Lua chon khong hop le!', ui.ButtonSet.OK);
        showActionMenu();
    }
  }
}

function runStage1Menu() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'GIAI DOAN 1',
    'Phan tich lead tu sheet. Tiep tuc?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm == ui.Button.YES) {
    const result = runStage1FromUI();
    ui.alert(result.success ? 'Thanh cong' : 'Loi', result.message, ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function runStage2Menu() {
  const ui = SpreadsheetApp.getUi();
  const result = runStage2FromUI();
  ui.alert(result.success ? 'Thanh cong' : 'Loi', result.message, ui.ButtonSet.OK);
  showActionMenu();
}

function runStage3Menu() {
  const ui = SpreadsheetApp.getUi();
  const result = runStage3FromUI();
  ui.alert(result.success ? 'Thanh cong' : 'Loi', result.message, ui.ButtonSet.OK);
  showActionMenu();
}

function syncStatusMenu() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'DONG BO STATUS',
    'Cap nhat Status tu ket qua n8n?\n\nPending sang Sent (neu cot Q co Success)',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm == ui.Button.YES) {
    const result = syncStatusFromN8n();
    ui.alert(result.success ? 'Thanh cong' : 'Loi', result.message, ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function manageRVAMenu() {
  const ui = SpreadsheetApp.getUi();
  const config = getRVAConfig();
  
  let msg = 'DANH SACH RVA:\n\n';
  config.forEach(function(rva, i) {
    msg += (i+1) + '. ' + rva.name + ' - ' + (rva.active ? 'Hoat dong' : 'Khong hoat dong') + '\n';
  });
  
  ui.alert('RVA', msg, ui.ButtonSet.OK);
  showActionMenu();
}

function cleanupQueueMenu() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('Xoa Queue?', 'Xac nhan xoa?', ui.ButtonSet.YES_NO);
  
  if (confirm == ui.Button.YES) {
    const result = cleanupQueueFromUI();
    ui.alert(result.success ? 'Thanh cong' : 'Loi', result.message, ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function viewLogMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.LOG);
  
  if (!logSheet || logSheet.getLastRow() <= 1) {
    ui.alert('Log', 'Chua co log', ui.ButtonSet.OK);
  } else {
    ss.setActiveSheet(logSheet);
    ui.alert('Log', 'Sheet log da mo', ui.ButtonSet.OK);
  }
  
  showActionMenu();
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Zalo System')
    .addItem('Control Panel', 'showMainMenu')
    .addItem('Thong ke', 'showControlPanel')
    .addSeparator()
    .addItem('Giai doan 1', 'runStage1Menu')
    .addItem('Giai doan 2', 'runStage2Menu')
    .addItem('Giai doan 3', 'runStage3Menu')
    .addItem('Dong bo Status', 'syncStatusMenu')
    .addSeparator()
    .addItem('Don dep Queue', 'cleanupQueueMenu')
    .addToUi();
  
  logSystem("Menu khoi tao");
}
