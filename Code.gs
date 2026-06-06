// ==========================================
// การตั้งค่า
// ==========================================
// แก้ไข YOUR_SPREADSHEET_ID ให้ตรงกับ ID ของ Google Sheets ที่คุณสร้าง
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet() ? SpreadsheetApp.getActiveSpreadsheet().getId() : "YOUR_SPREADSHEET_ID";

// ฟังก์ชันช่วยดึงข้อมูลชีต ถ้าไม่มีให้สร้างใหม่พร้อมหัวตาราง (Headers)
function getOrCreateSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    
    // 1. ถ้าไม่พบแท็บ ให้สร้างใหม่
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    // 2. ตรวจสอบว่ามีข้อมูลหรือไม่ (ถ้าชีตว่างเปล่าจะเขียนหัวตารางให้ทันที)
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      let headers = [];
      if (sheetName === 'Users') {
        headers = ['id', 'name', 'username', 'password'];
      } else if (sheetName === 'Students') {
        headers = ['id', 'studentId', 'name', 'class', 'totalScore'];
      } else if (sheetName === 'Behaviors') {
        headers = ['id', 'studentId', 'date', 'type', 'details', 'scoreChange', 'teacherName'];
      }
      
      if (headers.length > 0) {
        sheet.appendRow(headers);
        // จัดรูปแบบหัวตารางให้เป็นตัวหนา
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        
        // หากเป็นชีต Users ให้สร้างบัญชี admin/admin ให้โดยอัตโนมัติหากยังไม่มีข้อมูล
        if (sheetName === 'Users') {
          sheet.appendRow([Utilities.getUuid(), 'ผู้ดูแลระบบ', 'admin', 'admin']);
        }
      }
    }
    return sheet;
  } catch (e) {
    throw new Error('ไม่สามารถเข้าถึงหรือสร้างชีต "' + sheetName + '" ได้: ' + e.message);
  }
}

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('ระบบบันทึกและติดตามพฤติกรรมนักเรียน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// ระบบ Login
// ==========================================
function login(username, password) {
  try {
    const sheet = getOrCreateSheet('Users');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      // Columns: id(0), name(1), username(2), password(3)
      if (String(data[i][2]) === String(username) && String(data[i][3]) === String(password)) {
        return { 
          success: true, 
          user: { 
            id: data[i][0], 
            name: data[i][1], 
            username: data[i][2] 
          } 
        };
      }
    }
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  } catch(e) {
    return { success: false, message: 'ข้อผิดพลาด: ' + e.message };
  }
}

// ==========================================
// Dashboard
// ==========================================
function getDashboardData() {
  try {
    const sheetBehaviors = getOrCreateSheet('Behaviors');
    const sheetStudents = getOrCreateSheet('Students');
    
    const behaviorsData = sheetBehaviors.getDataRange().getValues();
    const studentsData = sheetStudents.getDataRange().getValues();
    
    let totalStudents = studentsData.length > 1 ? studentsData.length - 1 : 0;
    
    let stats = { light: 0, medium: 0, severe: 0, good: 0 };
    let recentLogs = [];
    
    // Last 7 days setup
    const today = new Date();
    const last7Days = [];
    for(let i=6; i>=0; i--) {
      let d = new Date(today);
      d.setDate(d.getDate() - i);
      last7Days.push({
         dateStr: Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM"),
         dateObj: d,
         good: 0,
         bad: 0
      });
    }

    // Process Behaviors
    // Columns: id(0), studentId(1), date(2), type(3), details(4), scoreChange(5), teacherName(6)
    for (let i = behaviorsData.length - 1; i >= 1; i--) {
      const row = behaviorsData[i];
      const type = row[3];
      const date = new Date(row[2]);
      const scoreChange = Number(row[5]);
      
      if (type === 'ผิดขั้นเบา') stats.light++;
      else if (type === 'ผิดขั้นปานกลาง') stats.medium++;
      else if (type === 'ผิดขั้นร้ายแรง') stats.severe++;
      else if (type === 'ทำดี') stats.good++;
      
      if (recentLogs.length < 5) {
        // Find student name
        let sName = 'ไม่ทราบชื่อ';
        for (let j = 1; j < studentsData.length; j++) {
          if (studentsData[j][1] == row[1]) {
            sName = studentsData[j][2];
            break;
          }
        }
        
        recentLogs.push({
          id: row[0],
          studentId: row[1],
          studentName: sName,
          date: Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
          type: type,
          details: row[4],
          scoreChange: scoreChange,
          teacherName: row[6]
        });
      }
      
      // Chart data population
      for(let j=0; j<last7Days.length; j++) {
        let d1 = last7Days[j].dateObj;
        if (d1.getDate() === date.getDate() && d1.getMonth() === date.getMonth() && d1.getFullYear() === date.getFullYear()) {
          if (scoreChange > 0) last7Days[j].good++;
          else if (scoreChange < 0) last7Days[j].bad++;
        }
      }
    }
    
    // Top Negative Students
    let studentScores = [];
    for(let i=1; i<studentsData.length; i++) {
      studentScores.push({
        studentId: studentsData[i][1],
        name: studentsData[i][2],
        class: studentsData[i][3],
        totalScore: Number(studentsData[i][4]) || 0
      });
    }
    studentScores.sort((a, b) => a.totalScore - b.totalScore);
    const topNegativeStudents = studentScores.slice(0, 5);

    return {
      success: true,
      totalStudents: totalStudents,
      stats: stats,
      recentLogs: recentLogs,
      chartData: {
        labels: last7Days.map(d => d.dateStr),
        goodScores: last7Days.map(d => d.good),
        badScores: last7Days.map(d => d.bad)
      },
      topNegativeStudents: topNegativeStudents
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// Students Management
// ==========================================
function getStudents() {
  try {
    const sheet = getOrCreateSheet('Students');
    const data = sheet.getDataRange().getValues();
    let result = [];
    for (let i = 1; i < data.length; i++) {
      result.push({
        id: data[i][0],
        studentId: data[i][1],
        name: data[i][2],
        class: data[i][3],
        totalScore: data[i][4]
      });
    }
    return result;
  } catch(e) {
    return [];
  }
}

function addStudent(student) {
  try {
    const sheet = getOrCreateSheet('Students');
    const newId = Utilities.getUuid();
    // Start with 100 points
    sheet.appendRow([newId, student.studentId, student.name, student.class, 100]); 
    return { success: true, message: 'เพิ่มข้อมูลนักเรียนสำเร็จ' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateStudent(student) {
  try {
    const sheet = getOrCreateSheet('Students');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === student.id) {
        sheet.getRange(i + 1, 2, 1, 3).setValues([[student.studentId, student.name, student.class]]);
        return { success: true, message: 'แก้ไขข้อมูลสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลนักเรียน' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deleteStudent(id) {
  try {
    const sheet = getOrCreateSheet('Students');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบข้อมูลสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลนักเรียน' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// Behaviors Management
// ==========================================
function addBehavior(behavior) {
  try {
    const sheet = getOrCreateSheet('Behaviors');
    const newId = Utilities.getUuid();
    const date = new Date();
    
    sheet.appendRow([
      newId, 
      behavior.studentId, 
      date, 
      behavior.type, 
      behavior.details, 
      behavior.scoreChange, 
      behavior.teacherName
    ]);
    
    // Update student total score
    const sheetStudents = getOrCreateSheet('Students');
    const studentsData = sheetStudents.getDataRange().getValues();
    for (let i = 1; i < studentsData.length; i++) {
      if (String(studentsData[i][1]) === String(behavior.studentId)) {
        let currentScore = Number(studentsData[i][4]) || 0;
        let newScore = currentScore + Number(behavior.scoreChange);
        sheetStudents.getRange(i + 1, 5).setValue(newScore);
        break;
      }
    }
    
    return { success: true, message: 'บันทึกพฤติกรรมและอัปเดตคะแนนสำเร็จ' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getStudentReport(studentId) {
  try {
    const sheetBehaviors = getOrCreateSheet('Behaviors');
    const behaviorsData = sheetBehaviors.getDataRange().getValues();
    
    let history = [];
    for (let i = behaviorsData.length - 1; i >= 1; i--) {
      if (String(behaviorsData[i][1]) === String(studentId)) {
        history.push({
          date: Utilities.formatDate(new Date(behaviorsData[i][2]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
          type: behaviorsData[i][3],
          details: behaviorsData[i][4],
          scoreChange: behaviorsData[i][5],
          teacherName: behaviorsData[i][6]
        });
      }
    }
    
    const sheetStudents = getOrCreateSheet('Students');
    const studentsData = sheetStudents.getDataRange().getValues();
    let studentInfo = null;
    for (let i = 1; i < studentsData.length; i++) {
      if (String(studentsData[i][1]) === String(studentId)) {
        studentInfo = {
          studentId: studentsData[i][1],
          name: studentsData[i][2],
          class: studentsData[i][3],
          totalScore: studentsData[i][4]
        };
        break;
      }
    }
    
    if(!studentInfo) return { success: false, message: 'ไม่พบข้อมูลนักเรียนรหัสนี้' };
    
    return {
      success: true,
      student: studentInfo,
      history: history
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// Users Management (Teachers)
// ==========================================
function getUsers() {
  try {
    const sheet = getOrCreateSheet('Users');
    const data = sheet.getDataRange().getValues();
    let result = [];
    for (let i = 1; i < data.length; i++) {
      result.push({
        id: data[i][0],
        name: data[i][1],
        username: data[i][2]
      });
    }
    return result;
  } catch(e) {
    return [];
  }
}

function addUser(user) {
  try {
    const sheet = getOrCreateSheet('Users');
    // Check duplicate username
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]) === String(user.username)) {
        return { success: false, message: 'Username นี้ถูกใช้งานแล้ว' };
      }
    }
    
    const newId = Utilities.getUuid();
    sheet.appendRow([newId, user.name, user.username, user.password]);
    return { success: true, message: 'เพิ่มผู้ใช้งานสำเร็จ' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deleteUser(id) {
  try {
    const sheet = getOrCreateSheet('Users');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบผู้ใช้งานสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบผู้ใช้งาน' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}
