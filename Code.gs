// === Code.gs (Backend API) ===

function doPost(e) {
  let response = { status: 'error', message: 'Unknown error', data: null };
  
  try {
    let requestData = JSON.parse(e.postData.contents);
    let action = requestData.action;
    let payload = requestData.payload || {};

    // --- ROUTER: Map frontend actions to backend functions ---
    if (action === 'getUsers') {
      response.data = getUsers();
    } else if (action === 'getTodayMenu') {
      response.data = getTodayMenu();
    } else if (action === 'submitBooking') {
      response.data = submitBooking(payload.userName, payload.mealType);
    } else if (action === 'getMyTodayBookings') {
      response.data = getMyTodayBookings(payload.userName);
    } else if (action === 'getMyMonthlySummary') {
      response.data = getMyMonthlySummary(payload.userName);
    } else if (action === 'submitVisitor') {
      response.data = submitVisitor(payload.requestedBy, payload.description, payload.count, payload.mealType);
    } else if (action === 'getDashboardStats') {
      response.data = getDashboardStats();
    } else if (action === 'getAdminDashboardData') {
      response.data = getAdminDashboardData();
    } else if (action === 'getAllStaff') {
      response.data = getAllStaff();
    } else if (action === 'updateStaff') {
      response.data = updateStaff(payload.id, payload.name, payload.email, payload.dept);
    } else if (action === 'deleteStaff') {
      response.data = deleteStaff(payload.id);
    } else if (action === 'updateDailyMenu') {
      response.data = updateDailyMenu(payload.date, payload.bfast, payload.lunch, payload.dinner);
    } else if (action === 'addEmployee') {
      response.data = addEmployee(payload.id, payload.name, payload.email, payload.dept);
    } else {
      throw new Error("Action not found.");
    }
    
    response.status = 'success';
  } catch (error) {
    response.message = error.toString();
  }

  // Return standard JSON bypassing standard CORS restrictions
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// Fallback for simple pings
function doGet(e) {
  return ContentService.createTextOutput("Darul Iman API is running.").setMimeType(ContentService.MimeType.TEXT);
}

// ==========================================
// CORE SPREADSHEET LOGIC (Unchanged)
// ==========================================

function getUsers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => row[1]).filter(String); 
}

function submitBooking(userName, mealType) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(3000); } catch (e) { return "System busy. Please try again."; }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bookings");
    const data = sheet.getDataRange().getValues();
    const timestamp = new Date();
    const formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    for (let i = 1; i < data.length; i++) {
      let rowDate = data[i][1];
      if (rowDate instanceof Date) rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (rowDate === formattedDate && data[i][2] === userName && data[i][3] === mealType && data[i][4] === "Booked") {
        return `Notice: You have already booked ${mealType} for today!`;
      }
    }

    sheet.appendRow([timestamp, formattedDate, userName, mealType, "Booked"]);
    return `Success! ${mealType} booked for ${userName}.`;
  } catch (error) { return "Error saving booking: " + error.toString(); } finally { lock.releaseLock(); }
}

function getMyTodayBookings(userName) {
  if (!userName) return [];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bookings");
  const data = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  let myMeals = [];
  for (let i = data.length - 1; i > 0; i--) {
    let rowDate = data[i][1];
    if (rowDate instanceof Date) rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (rowDate === today && data[i][2] === userName && data[i][4] === "Booked") myMeals.push(data[i][3]); 
  }
  return myMeals; 
}

function submitVisitor(requestedBy, description, count, mealType) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(3000); } catch (e) { return "System busy."; }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Visitors");
    const timestamp = new Date();
    const formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.appendRow([timestamp, formattedDate, description, count, mealType, requestedBy]);
    return `Success! ${count} ${mealType}(s) booked for ${description}.`;
  } catch (error) { return "Error: " + error.toString(); } finally { lock.releaseLock(); }
}

function getDashboardStats() {
  const bookSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bookings");
  const visSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Visitors");
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  let stats = { Breakfast: { staff: 0, visitor: 0, total: 0 }, Lunch: { staff: 0, visitor: 0, total: 0 }, Dinner: { staff: 0, visitor: 0, total: 0 } };

  const bookData = bookSheet.getDataRange().getValues();
  for (let i = 1; i < bookData.length; i++) {
    let rowDate = bookData[i][1];
    if (rowDate instanceof Date) rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (rowDate === today && bookData[i][4] === "Booked") {
      let meal = bookData[i][3];
      if (stats[meal] !== undefined) { stats[meal].staff++; stats[meal].total++; }
    }
  }

  const visData = visSheet.getDataRange().getValues();
  for (let i = 1; i < visData.length; i++) {
    let rowDate = visData[i][1];
    if (rowDate instanceof Date) rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (rowDate === today) {
      let count = Number(visData[i][3]) || 0;
      let meal = visData[i][4];
      if (stats[meal] !== undefined) { stats[meal].visitor += count; stats[meal].total += count; }
    }
  }
  return stats;
}

function getTodayMenu() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");
  if (!sheet) return { Breakfast: "No menu sheet", Lunch: "No menu sheet", Dinner: "No menu sheet" };
  const data = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  let menu = { Breakfast: "Pending...", Lunch: "Pending...", Dinner: "Pending..." };
  
  for (let i = 1; i < data.length; i++) {
    let rowDate = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "yyyy-MM-dd") : data[i][0];
    if (rowDate === today) {
      let type = data[i][1]; let desc = data[i][2]; 
      if (menu[type] !== undefined) menu[type] = desc;
    }
  }
  return menu;
}

function addEmployee(id, name, email, department) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users").appendRow([id, name, email, department, "User"]);
  return `Successfully added ${name}!`;
}

function updateDailyMenu(date, bfast, lunch, dinner) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");
  const data = sheet.getDataRange().getValues();
  const formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const types = ["Breakfast", "Lunch", "Dinner"], descs = [bfast, lunch, dinner];

  for (let i = 0; i < types.length; i++) {
    let type = types[i], desc = descs[i], found = false;
    for (let r = 1; r < data.length; r++) {
      let rowDate = data[r][0] instanceof Date ? Utilities.formatDate(data[r][0], Session.getScriptTimeZone(), "yyyy-MM-dd") : data[r][0];
      if (rowDate === formattedDate && data[r][1] === type) { sheet.getRange(r + 1, 3).setValue(desc); found = true; break; }
    }
    if (!found) sheet.appendRow([formattedDate, type, desc, "", "Active"]);
  }
  return `Menu updated successfully!`;
}

function getAdminDashboardData() {
  const stats = getDashboardStats(), menu = getTodayMenu();
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bookings").getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  let todayBookings = [];
  
  for (let i = data.length - 1; i > 0; i--) {
    let rowDate = data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), "yyyy-MM-dd") : data[i][1];
    if (rowDate === today && data[i][4] === "Booked") {
      let timeStr = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "hh:mm a") : "";
      todayBookings.push({ time: timeStr, name: data[i][2], meal: data[i][3] });
    }
  }
  return { stats: stats, menu: menu, bookings: todayBookings };
}

function getAllStaff() {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users").getDataRange().getValues();
  let staffList = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][1]) staffList.push({ id: data[i][0] || "-", name: data[i][1], email: data[i][2] || "-", dept: data[i][3] || "-" });
  }
  return staffList;
}

function updateStaff(id, name, email, dept) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) { 
      sheet.getRange(i + 1, 2).setValue(name); sheet.getRange(i + 1, 3).setValue(email); sheet.getRange(i + 1, 4).setValue(dept);
      return `Updated ${name}!`;
    }
  }
  return `Error finding ID.`;
}

function deleteStaff(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (data[i][0] == id) { sheet.deleteRow(i + 1); return `Deleted ${id}.`; } }
  return `Error finding ID.`;
}

function getMyMonthlySummary(userName) {
  if (!userName) return { totals: { Breakfast: 0, Lunch: 0, Dinner: 0, Total: 0 }, history: [] };
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bookings").getDataRange().getValues();
  const now = new Date(), currentMonth = now.getMonth(), currentYear = now.getFullYear();
  let summary = { totals: { Breakfast: 0, Lunch: 0, Dinner: 0, Total: 0 }, history: [] };

  for (let i = data.length - 1; i > 0; i--) {
    let rawDateObj = data[i][1] instanceof Date ? data[i][1] : new Date(data[i][1]);
    if (!isNaN(rawDateObj) && data[i][2] === userName && data[i][4] === "Booked") {
      if (rawDateObj.getMonth() === currentMonth && rawDateObj.getFullYear() === currentYear) {
        let mealType = data[i][3];
        if (summary.totals[mealType] !== undefined) { summary.totals[mealType]++; summary.totals.Total++; }
        if (summary.history.length < 10) summary.history.push({ date: Utilities.formatDate(rawDateObj, Session.getScriptTimeZone(), "MMM dd"), meal: mealType });
      }
    }
  }
  return summary;
}