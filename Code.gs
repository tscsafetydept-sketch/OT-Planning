// ==========================================
// ส่วนของการตั้งค่า (Configuration)
// ==========================================
const USERS_SHEET = "Users";
const OT_SHEET_PREFIX = "OT_";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('OT Planning System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// 1. ฟังก์ชันจัดการฐานข้อมูลและผู้ใช้งาน
// ==========================================
function getDb() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function setupSheets() {
  const db = getDb();
  if (!db.getSheetByName(USERS_SHEET)) {
    const ws = db.insertSheet(USERS_SHEET);
    ws.appendRow(["EmpID", "Name", "Position", "Phone", "Email", "Password", "Order"]);
  }
}

function registerUser(data) {
  setupSheets();
  const ws = getDb().getSheetByName(USERS_SHEET);
  const dataArr = ws.getDataRange().getValues();
  
  // เช็คว่ารหัสพนักงานซ้ำหรือไม่
  for (let i = 1; i < dataArr.length; i++) {
    if (String(dataArr[i][0]) === String(data.empId)) {
      return { success: false, message: "รหัสพนักงานนี้มีในระบบแล้ว" };
    }
  }
  
  const order = dataArr.length;
  ws.appendRow([data.empId, data.name, data.position, data.phone, data.email, data.password, order]);
  return { success: true, message: "ลงทะเบียนสำเร็จเรียบร้อย" };
}

function loginUser(id, pass) {
  const ws = getDb().getSheetByName(USERS_SHEET);
  if(!ws) return { success: false, message: "ไม่พบฐานข้อมูลผู้ใช้งาน" };
  const data = ws.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id) && String(data[i][5]) === String(pass)) {
      return { 
        success: true, 
        userData: { empId: data[i][0], name: data[i][1], position: data[i][2] } 
      };
    }
  }
  return { success: false, message: "รหัสพนักงานหรือรหัสผ่านไม่ถูกต้อง" };
}

function getUsersList() {
  const ws = getDb().getSheetByName(USERS_SHEET);
  if(!ws) return { success: true, data: [] };
  const data = ws.getDataRange().getValues();
  let users = [];
  
  for (let i = 1; i < data.length; i++) {
    users.push({
      id: String(data[i][0]),
      name: data[i][1],
      pos: data[i][2], // ตำแหน่ง (คอลัมน์ C)
      order: data[i][6] || i // การจัดเรียง (คอลัมน์ G)
    });
  }
  return { success: true, data: users };
}

function updateUsersInfo(usersList) {
  const ws = getDb().getSheetByName(USERS_SHEET);
  if(!ws) return { success: false, message: "ไม่พบฐานข้อมูล" };
  const data = ws.getDataRange().getValues();
  
  for(let j=0; j < usersList.length; j++){
     for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(usersList[j].id)) {
           ws.getRange(i + 1, 7).setValue(usersList[j].order); // อัปเดตลำดับ
           break;
        }
     }
  }
  return { success: true, message: "อัปเดตลำดับเรียบร้อย" };
}

// ==========================================
// 2. ฟังก์ชันจัดการข้อมูล OT รายเดือน
// ==========================================
function getMonthlyOT(monthVal) {
  const sheetName = OT_SHEET_PREFIX + monthVal;
  const ws = getDb().getSheetByName(sheetName);
  if (!ws) return { success: true, data: {} };
  
  const data = ws.getDataRange().getValues();
  let result = {};
  
  for (let i = 1; i < data.length; i++) {
    const empId = String(data[i][0]);
    const status = data[i][1];
    const otDataStr = data[i][2];
    let otData = {};
    try { otData = JSON.parse(otDataStr); } catch(e) {}
    result[empId] = { status: status, otData: otData };
  }
  return { success: true, data: result };
}

function saveMonthlyOT(monthVal, empId, otData, status) {
  const sheetName = OT_SHEET_PREFIX + monthVal;
  let ws = getDb().getSheetByName(sheetName);
  
  if (!ws) {
    ws = getDb().insertSheet(sheetName);
    ws.appendRow(["EmpID", "Status", "OT_JSON"]);
  }
  
  const data = ws.getDataRange().getValues();
  let found = false;
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(empId)) {
      ws.getRange(i + 1, 2).setValue(status);
      ws.getRange(i + 1, 3).setValue(JSON.stringify(otData));
      found = true;
      break;
    }
  }
  
  if (!found) {
    ws.appendRow([empId, status, JSON.stringify(otData)]);
  }
  return { success: true, message: "บันทึกข้อมูลสำเร็จ" };
}
