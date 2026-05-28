// ตั้งค่าความปลอดภัยรหัสผ่านผู้จัดการ (Admin)
const ADMIN_PASSWORD = "pd51234";

function doGet(e) {
  return ContentService.createTextOutput("ระบบ Mixing Backend ออนไลน์พร้อมทำงาน");
}

function updateAndFetchDetails(qrCodeMessage, operatorName, batchNumber) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. ตรวจสอบแผ่นงานสูตรผลิต (Formula)
    const formulaSheet = ss.getSheetByName("Formula");
    if (!formulaSheet) {
      return { success: false, message: "❌ ไม่พบแผ่นงานชื่อ 'Formula' ใน Google Sheets" };
    }
    
    const formulaData = formulaSheet.getDataRange().getValues();
    let matchedFormulaRow = null;
    let rowIndexInFormula = -1;
    
    // ค้นหา QR Code ในคอลัมน์แรก (Index 0) ของหน้า Formula
    for (let i = 1; i < formulaData.length; i++) {
      if (formulaData[i][0].toString().trim() === qrCodeMessage.trim()) {
        matchedFormulaRow = formulaData[i];
        rowIndexInFormula = i + 1; // ล็อกแถวจริงบน Sheet
        break;
      }
    }
    
    if (!matchedFormulaRow) {
      logToSheet(qrCodeMessage, operatorName, batchNumber, "⚠️ ปฏิเสธ: ไม่พบรหัสสารนี้ในสูตรผลิต");
      return { success: false, message: "❌ สารเคมีนี้ไม่อยู่ในสูตรการผลิต (ไม่พบคิวอาร์ในหน้า Formula)" };
    }
    
    // โครงสร้างคอลัมน์หน้า Formula: [0] QR, [1] ชื่อสาร, [2] ขนาดถุง, [3] สแกนแล้ว, [4] ลิมิตสูตร
    const substanceName = matchedFormulaRow[1] || "ไม่ระบุชื่อสาร";
    const targetLimit = parseInt(matchedFormulaRow[4]) || 0;
    
    // 2. ตรวจสอบแผ่นงานบันทึกประวัติ (Mixing_Logs) เพื่อตรวจสอบจำนวนการสแกนใน Batch นี้
    const logsSheet = ss.getSheetByName("Mixing_Logs") || ss.insertSheet("Mixing_Logs");
    if (logsSheet.getLastRow() === 0) {
      logsSheet.appendRow(["Timestamp", "DateGroup", "Operator", "Batch", "QRCode", "SubstanceName", "TargetQtyFromFormula", "Status"]);
    }
    
    const logsData = logsSheet.getDataRange().getValues();
    let currentScannedInBatch = 0;
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    // นับเฉพาะรายการที่ "บันทึกสำเร็จ" ใน Batch เดียวกัน และเป็นสารตัวเดียวกัน
    for (let j = 1; j < logsData.length; j++) {
      const logDateGroup = logsData[j][1] ? logsData[j][1].toString().trim() : "";
      const logBatch = logsData[j][3] ? logsData[j][3].toString().trim() : "";
      const logSubstance = logsData[j][5] ? logsData[j][5].toString().trim() : "";
      const logStatus = logsData[j][7] ? logsData[j][7].toString().trim() : "";
      
      if (logDateGroup === todayStr && logBatch === batchNumber.trim() && logSubstance === substanceName && logStatus === "บันทึกสำเร็จ") {
        currentScannedInBatch++;
      }
    }
    
    // 3. ตรวจสอบโควตา (Validation)
    if (currentScannedInBatch >= targetLimit) {
      logToSheet(qrCodeMessage, operatorName, batchNumber, "⚠️ ปฏิเสธ: ปริมาณสารเกินกำหนดกลุ่มสูตรแล้ว", substanceName, targetLimit);
      return { 
        success: false, 
        message: "🚫 ปฏิเสธการบันทึก! สาร " + substanceName + " ถูกสแกนครบตามจำนวนจำกัด (" + targetLimit + " ถุง) ใน BATCH นี้แล้ว หากต้องการสแกนเพิ่มโปรดติดต่อแอดมินลบประวัติเดิม" 
      };
    }
    
    // 4. บันทึกผลสำเร็จลงตารางประวัติ และอัปเดตยอดสะสมหน้าสูตร
    logToSheet(qrCodeMessage, operatorName, batchNumber, "บันทึกสำเร็จ", substanceName, targetLimit);
    
    const newScannedCount = currentScannedInBatch + 1;
    formulaSheet.getRange(rowIndexInFormula, 4).setValue(newScannedCount); // อัปเดตคอลัมน์ "สแกนแล้ว" ในหน้า Formula
    
    return {
      success: true,
      message: "🎉 บันทึกส่วนผสม " + substanceName + " เรียบร้อย! (ถุงที่ " + newScannedCount + "/" + targetLimit + ")",
      serverScannedCount: newScannedCount,
      rowData: matchedFormulaRow
    };
    
  } catch (error) {
    return { success: false, message: "💥 เกิดข้อผิดพลาดบนระบบคลาวด์: " + error.toString() };
  }
}

// ฟังก์ชันภายในช่วยบันทึกข้อมูลลง Sheet ประวัติ
function logToSheet(qrCode, operator, batch, status, substanceName = "-", targetLimit = "-") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logsSheet = ss.getSheetByName("Mixing_Logs") || ss.insertSheet("Mixing_Logs");
  
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const dateGroup = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  logsSheet.appendRow([
    timestamp,
    dateGroup,
    operator,
    batch,
    qrCode,
    substanceName,
    targetLimit,
    status
  ]);
}

function getLogsData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Mixing_Logs");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const logs = [];
    
    for (let i = data.length - 1; i >= 1; i--) {
      logs.push({
        timestamp: data[i][0] ? Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : "",
        dateGroup: data[i][1] ? data[i][1].toString() : "",
        operator: data[i][2] || "",
        batch: data[i][3] || "",
        qrCode: data[i][4] || "",
        substanceName: data[i][5] || "",
        targetQtyFromFormula: data[i][6] || "",
        status: data[i][7] || ""
      });
    }
    return logs;
  } catch (e) {
    return [];
  }
}

function deleteLogFromServer(timestamp, qrCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Mixing_Logs");
    if (!sheet) return { success: false, message: "ไม่พบหน้าต่างบันทึกข้อมูล" };
    
    const data = sheet.getDataRange().getValues();
    let deletedCount = 0;
    
    // วนลูปจากล่างขึ้นบนเพื่อป้องกันดัชนีเลื่อนเวลาลบแถว
    for (let i = data.length - 1; i >= 1; i--) {
      const rowTime = data[i][0] ? Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : "";
      const rowQr = data[i][4] || "";
      
      if (rowTime === timestamp && rowQr === qrCode) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    return { success: true, message: "ลบรายการประวัติสำเร็จ จำนวน " + deletedCount + " แถว" };
  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() };
  }
}
