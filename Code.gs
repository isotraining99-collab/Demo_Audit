var SHEET_ID = "1zUeF35TzE3l6pm45cMKclPGx0XGHNvaEE6xhAcPT5sw"; 

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

function getAuditData() {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var data = sheet.getDataRange().getDisplayValues(); 
    if (data.length <= 1) return []; 
    data.shift(); 
    return data;
  } catch (e) { throw new Error(e.message); }
}

// 🌟 สร้างโฟลเดอร์ลูกค้า (พร้อม Gen รหัส Audit_ID)
function createNewCustomerFolders(year, customerName) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var timestamp = new Date();
    var userEmail = Session.getActiveUser().getEmail() || "Unknown User"; 
    
    // สร้าง Audit ID อัตโนมัติ (รูปแบบ AD-ปีเดือนวัน-เวลา)
    var auditId = "AD-" + Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyMMdd-HHmmss");
    
    var docTypes = [
      "1. Plant Tour information", "2. Audit Time Table", "3. Audit Preparation Material",
      "4. Audit Report (Customer)", "5. Audit Report (Internal)"
    ];
    var dataToWrite = [];
    for (var i = 0; i < docTypes.length; i++) {
      // เรียงใหม่: Timestamp, Year, Customer, Audit_ID, Doc_Type, File_Name, File_Link, Editor
      dataToWrite.push([timestamp, year, customerName, auditId, docTypes[i], "", "", userEmail]);
    }
    sheet.getRange(sheet.getLastRow() + 1, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
    SpreadsheetApp.flush(); 
    return "สร้างรอบ Audit ใหม่รหัส [" + auditId + "] สำเร็จ!";
  } catch (e) { throw new Error(e.message); }
}

function uploadFileToDrive(base64Data, fileName, year, customerName, auditId, docType, folderId) {
  try {
    var folder = DriveApp.getFolderById(folderId);
    var splitBase = base64Data.split(',');
    var type = splitBase[0].split(';')[0].replace('data:', '');
    var blob = Utilities.newBlob(Utilities.base64Decode(splitBase[1]), type, fileName);
    var fileUrl = folder.createFile(blob).getUrl();

    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var timestamp = new Date();
    var userEmail = Session.getActiveUser().getEmail() || "Unknown User";
    
    for (var i = 1; i < data.length; i++) {
      // เช็คให้ตรงทั้ง ปี, ลูกค้า, รหัสรอบ Audit, และประเภทเอกสาร
      if (data[i][1] == year && data[i][2] == customerName && data[i][3] == auditId && data[i][4] == docType) {
        var targetRow = i + 1; 
        sheet.getRange(targetRow, 1).setValue(timestamp); 
        sheet.getRange(targetRow, 6).setValue(fileName);  
        sheet.getRange(targetRow, 7).setValue(fileUrl);   
        sheet.getRange(targetRow, 8).setValue(userEmail); 
        SpreadsheetApp.flush(); 
        return "อัปโหลดไฟล์สำเร็จ";
      }
    }
    throw new Error("หาบรรทัดใน Sheet ไม่เจอ");
  } catch (e) { throw new Error(e.message); }
}

function removeFileFromSheet(auditId, docType) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] == auditId && data[i][4] == docType) {
        sheet.getRange(i + 1, 6).clearContent(); 
        sheet.getRange(i + 1, 7).clearContent(); 
        SpreadsheetApp.flush();
        return "ถอดไฟล์ออกเรียบร้อย";
      }
    }
  } catch(e) { throw new Error(e.message); }
}

// 🌟 เปลี่ยนชื่อลูกค้า (เปลี่ยนทุกรอบประวัติที่มีชื่อนี้)
function renameCustomerInSheet(oldName, newName) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] == oldName) {
        sheet.getRange(i + 1, 3).setValue(newName); 
      }
    }
    SpreadsheetApp.flush();
    return "เปลี่ยนชื่อสำเร็จ";
  } catch(e) { throw new Error(e.message); }
}

// 🌟 ลบประวัติ Audit (เลือกลบเฉพาะบางรอบได้ โดยอ้างอิงจาก ID)
function deleteAuditSession(auditId) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][3] == auditId) {
        sheet.deleteRow(i + 1);
      }
    }
    SpreadsheetApp.flush();
    return "ลบรอบ Audit รหัส " + auditId + " ออกเรียบร้อย";
  } catch(e) { throw new Error(e.message); }
}
