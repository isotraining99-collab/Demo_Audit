var SHEET_ID = "1zUeF35TzE3l6pm45cMKclPGx0XGHNvaEE6xhAcPT5sw"; 

function doGet() { return HtmlService.createHtmlOutputFromFile('index'); }

// ==========================================
// 🌟 ระบบของ 02 Audit Result (ระบบเดิม)
// ==========================================
function getAuditData() {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var data = sheet.getDataRange().getDisplayValues(); 
    if (data.length <= 1) return []; data.shift(); return data;
  } catch (e) { throw new Error(e.message); }
}

function createNewCustomerFolders(year, customerName, realName) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var rawDate = new Date();
    // 🌟 แก้บั๊ก 00:00:00 โดยบังคับแปลงเป็น Text ตรงๆ
    var timestampStr = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var userEmail = Session.getActiveUser().getEmail() || "Unknown User"; 
    var auditId = "AD-" + Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyMMdd-HHmmss");
    var docTypes = ["1. Plant Tour information", "2. Audit Time Table", "3. Audit Preparation Material", "4. Audit Report (Customer)", "5. Audit Report (Internal)"];
    var dataToWrite = [];
    for (var i = 0; i < docTypes.length; i++) { dataToWrite.push([timestampStr, year, customerName, auditId, docTypes[i], "", "", userEmail, realName, ""]); }
    sheet.getRange(sheet.getLastRow() + 1, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
    SpreadsheetApp.flush(); return "Success: Audit Session [" + auditId + "] has been created.";
  } catch (e) { throw new Error(e.message); }
}

function uploadFileToDrive(base64Data, fileName, year, customerName, auditId, docType, folderId, detail) {
  try {
    var folder = DriveApp.getFolderById(folderId);
    var splitBase = base64Data.split(','); var type = splitBase[0].split(';')[0].replace('data:', '');
    var blob = Utilities.newBlob(Utilities.base64Decode(splitBase[1]), type, fileName);
    var fileUrl = folder.createFile(blob).getUrl();
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var timestampStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var userEmail = Session.getActiveUser().getEmail() || "Unknown User";
    var realName = ""; var emptyRowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] == auditId && data[i][4] == docType) {
        realName = data[i][8]; if (data[i][5] == "" && emptyRowIndex == -1) emptyRowIndex = i + 1; 
      }
    }
    if (emptyRowIndex != -1) {
      sheet.getRange(emptyRowIndex, 1).setValue(timestampStr); sheet.getRange(emptyRowIndex, 6).setValue(fileName);  
      sheet.getRange(emptyRowIndex, 7).setValue(fileUrl); sheet.getRange(emptyRowIndex, 8).setValue(userEmail); sheet.getRange(emptyRowIndex, 10).setValue(detail); 
    } else {
      sheet.appendRow([timestampStr, year, customerName, auditId, docType, fileName, fileUrl, userEmail, realName, detail]);
    }
    SpreadsheetApp.flush(); return "Success: File uploaded and details saved.";
  } catch (e) { throw new Error(e.message); }
}

function editFileDetail(auditId, docType, fileLink, newDetail) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) { if (data[i][3] == auditId && data[i][4] == docType && data[i][6] == fileLink) { sheet.getRange(i + 1, 10).setValue(newDetail); SpreadsheetApp.flush(); return "Success: Detail updated."; } }
    throw new Error("File not found.");
  } catch(e) { throw new Error(e.message); }
}

function removeFileFromSheet(auditId, docType, fileLink) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; var data = sheet.getDataRange().getValues();
    var matchCount = 0; var targetRow = -1;
    for (var i = 1; i < data.length; i++) { if (data[i][3] == auditId && data[i][4] == docType) { matchCount++; if (data[i][6] == fileLink) targetRow = i + 1; } }
    if (targetRow != -1) {
      var fLink = sheet.getRange(targetRow, 7).getValue();
      if (fLink) { var fileIdMatch = fLink.match(/[-\w]{25,}/); if (fileIdMatch) { try { DriveApp.getFileById(fileIdMatch[0]).setTrashed(true); } catch(e) { Logger.log("Drive Delete Error"); } } }
      if (matchCount > 1) { sheet.deleteRow(targetRow); } else { sheet.getRange(targetRow, 6).clearContent(); sheet.getRange(targetRow, 7).clearContent(); sheet.getRange(targetRow, 10).clearContent(); }
      SpreadsheetApp.flush(); return "Success: File removed from system and Drive.";
    }
  } catch(e) { throw new Error(e.message); }
}

function deleteAuditSession(auditId) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) { if (data[i][3] == auditId) sheet.deleteRow(i + 1); }
    SpreadsheetApp.flush(); return "Success: Session deleted.";
  } catch(e) { throw new Error(e.message); }
}

function renameCustomerInSheet(oldName, newName) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) { if (data[i][2] == oldName) sheet.getRange(i + 1, 3).setValue(newName); }
    SpreadsheetApp.flush(); return "Success: Customer name updated.";
  } catch(e) { throw new Error(e.message); }
}

// ==========================================
// 🌟 ระบบของ 01 Custom Audit (สร้างเอกสารอิสระ 1 Session = 1 Doc)
// ==========================================

function getCustomAuditData() {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("01_CustomAudit");
    if(!sheet) return [];
    var data = sheet.getDataRange().getDisplayValues(); 
    if (data.length <= 1) return []; data.shift(); return data;
  } catch (e) { throw new Error(e.message); }
}

function createNewCustomSession(year, customerName, realName, docTitle) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("01_CustomAudit");
    if(!sheet) throw new Error("Sheet '01_CustomAudit' not found!");
    
    var rawDate = new Date();
    var timestampStr = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var userEmail = Session.getActiveUser().getEmail() || "Unknown User"; 
    var auditId = "CA-" + Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyMMdd-HHmmss");
    
    // 🌟 สร้างแค่ 1 บรรทัดต่อ 1 เอกสาร (ไม่มีการยัด 5 หัวข้อแล้ว)
    sheet.appendRow([timestampStr, year, customerName, auditId, docTitle, "", userEmail, realName, "Draft"]); 
    
    SpreadsheetApp.flush(); return "Success: Document Session [" + auditId + "] created.";
  } catch (e) { throw new Error(e.message); }
}

function saveCustomDocContent(auditId, htmlContent) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("01_CustomAudit");
    var data = sheet.getDataRange().getValues();
    var timestampStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var userEmail = Session.getActiveUser().getEmail() || "Unknown User";
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] == auditId) { // เช็คจากรหัส ID
        var targetRow = i + 1;
        sheet.getRange(targetRow, 1).setValue(timestampStr); 
        sheet.getRange(targetRow, 6).setValue(htmlContent); 
        sheet.getRange(targetRow, 7).setValue(userEmail); 
        sheet.getRange(targetRow, 9).setValue("Completed"); 
        SpreadsheetApp.flush();
        return "Success: Document saved successfully!";
      }
    }
    throw new Error("Document not found");
  } catch(e) { throw new Error(e.message); }
}

function deleteCustomSession(auditId) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("01_CustomAudit");
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) { if (data[i][3] == auditId) sheet.deleteRow(i + 1); }
    SpreadsheetApp.flush(); return "Success: Document deleted.";
  } catch(e) { throw new Error(e.message); }
}

// ==========================================
// 🌟 ฟังก์ชันพิเศษสำหรับรับไฟล์จากหน้า Word (TinyMCE)
// ==========================================
function uploadMediaToDrive(base64Data, fileName, isImage) {
  try {
    var folderId = "1cX4x5WnHrKUg0zNWHH-8qjQ3d1iCWbwQ"; // ใช้โฟลเดอร์หลักของเรา
    var folder = DriveApp.getFolderById(folderId);
    
    // แปลงไฟล์จาก Base64 กลับเป็นไฟล์ปกติ
    var splitBase = base64Data.split(',');
    var type = splitBase[0].split(';')[0].replace('data:', '');
    var blob = Utilities.newBlob(Utilities.base64Decode(splitBase[1]), type, fileName);
    var file = folder.createFile(blob);
    
    // 🌟 เปิดสิทธิ์ให้ "ทุกคนที่มีลิงก์" ดูได้ (เพื่อให้รูปโชว์บนหน้าเว็บได้ทันที)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    if (isImage) {
      // ถ้าเป็นรูปภาพ ส่ง Direct Link กลับไปให้โชว์ภาพ
      return "https://drive.google.com/uc?export=view&id=" + file.getId();
    } else {
      // ถ้าเป็นไฟล์เอกสาร ส่งลิงก์ธรรมดากลับไปให้คลิกโหลด
      return file.getUrl();
    }
  } catch (e) {
    throw new Error(e.message);
  }
}
