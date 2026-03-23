function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('OT Planning System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ฟังก์ชันดึง Sheet อัตโนมัติ ไม่ต้องใส่ ID ให้วุ่นวาย
function getDB() {
  return SpreadsheetApp.getActiveSpreadsheet(); 
}

function registerUser(data) {
  try {
    const ss = getDB();
    let sheet = ss.getSheetByName('Users');
    
    // ถ้ายังไม่มีชีต ให้สร้างใหม่และเขียนหัวคอลัมน์ 8 ช่องให้เป๊ะ
    if (!sheet) {
      sheet = ss.insertSheet('Users');
      sheet.appendRow(['EmpID', 'Name', 'Position', 'Phone', 'Email', 'Password', 'Timestamp', 'Order']);
      sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#f3f3f3");
    }
    
    const dataRange = sheet.getDataRange().getValues();
    let nextOrder = dataRange.length;
    
    for (let i = 1; i < dataRange.length; i++) {
      if (String(dataRange[i][0]) === String(data.empId)) {
        return { success: false, message: 'รหัสพนักงานนี้มีการลงทะเบียนแล้ว!' };
      }
    }
    
    // เรียงข้อมูลที่จะลง Sheet: 0=ID, 1=ชื่อ, 2=ตำแหน่ง, 3=เบอร์, 4=อีเมล, 5=รหัสผ่าน, 6=เวลา, 7=ลำดับ
    sheet.appendRow([
      String(data.empId), 
      data.name, 
      data.position, 
      data.phone, 
      data.email, 
      data.password, 
      new Date(), 
      nextOrder
    ]);
    
    return { success: true, message: 'สมัครสมาชิกสำเร็จ!' };
  } catch (error) { return { success: false, message: error.message }; }
}

function loginUser(empId, password) {
  try {
    const sheet = getDB().getSheetByName('Users');
    if (!sheet) return { success: false, message: 'ระบบยังไม่มีฐานข้อมูลผู้ใช้' };
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(empId) && String(data[i][5]) === String(password)) {
        return { success: true, userData: { empId: String(data[i][0]), name: data[i][1], position: data[i][2] } };
      }
    }
    return { success: false, message: 'รหัสพนักงาน หรือ รหัสผ่านไม่ถูกต้อง' };
  } catch (error) { return { success: false, message: error.message }; }
}

function getUsersList() {
  try {
    const sheet = getDB().getSheetByName('Users');
    if (!sheet) return { success: true, data: [] };
    
    const data = sheet.getDataRange().getValues();
    let users = [];
    for (let i = 1; i < data.length; i++) {
      users.push({ 
        id: String(data[i][0]), 
        name: data[i][1], 
        pos: data[i][2], 
        order: data[i][7] || i 
      });
    }
    users.sort((a, b) => a.order - b.order);
    return { success: true, data: users };
  } catch (e) { return { success: false, message: e.message }; }
}

function getMonthlyOT(monthKey) {
  try {
    const ss = getDB();
    const sheetName = 'OT_' + monthKey;
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['EmpID', 'Status', 'OT_Data_JSON', 'LastUpdate']);
      sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#d9ead3");
      return { success: true, data: {} };
    }
    
    const data = sheet.getDataRange().getValues();
    let otRecords = {};
    for (let i = 1; i < data.length; i++) {
      otRecords[String(data[i][0])] = { 
        status: data[i][1] || 'Pending', 
        otData: JSON.parse(data[i][2] || '{}') 
      };
    }
    return { success: true, data: otRecords };
  } catch (e) { return { success: false, message: e.message }; }
}

function saveMonthlyOT(monthKey, empId, otDataObj, statusStr) {
  try {
    const ss = getDB();
    const sheetName = 'OT_' + monthKey;
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['EmpID', 'Status', 'OT_Data_JSON', 'LastUpdate']);
    }
    
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(empId)) {
        if(statusStr) sheet.getRange(i + 1, 2).setValue(statusStr);
        if(otDataObj) sheet.getRange(i + 1, 3).setValue(JSON.stringify(otDataObj));
        sheet.getRange(i + 1, 4).setValue(new Date());
        found = true; break;
      }
    }
    if (!found) sheet.appendRow([String(empId), statusStr || 'Pending', JSON.stringify(otDataObj || {}), new Date()]);
    
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateUsersInfo(usersArray) {
  try {
    const sheet = getDB().getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      let empId = String(data[i][0]);
      let userObj = usersArray.find(u => String(u.id) === empId);
      if (userObj) {
        sheet.getRange(i + 1, 2).setValue(userObj.name);
        sheet.getRange(i + 1, 3).setValue(userObj.pos);
        sheet.getRange(i + 1, 8).setValue(userObj.order);
      }
    }
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}
