/**
 * 1. WEB APP ROUTER
 * Handles page navigation via ?page= parameter
 */
function doGet(e) {
  try {
    const page = (e && e.parameter.page) || 'index';

    // Routing Logic
    if (page === 'attendance') return render('Attendance', 'Attendance Tracker');
    if (page === 'list') return render('ListPage', 'Prefect List');
    // if (page === 'cc') return render('CC', 'Circle Calculator');
    // if (page === 'pos') return render('CustomerPOS', 'Product Menu');
    // if (page === 'seller') return render('SellerPage', 'Seller Dashboard');
    if (page === 'excuse') return render('ExcuseForm', 'Meeting Excuse Form');
    if (page === 'admine') return render('AdminExcuseView', 'Admin: View Excuses');
    // if (page === 'posv') return render('CostomerPOSV', 'Bake Sale - Visual Menu');
    //if (page === 'team')     return render('TeamDirectory', 'Prefect Team');
    if (page === 'duty') return render('Duty', 'Duty Checker');
    if (page === 'sc') return render('SC_Profiles', 'SC Directory');
    if (page === 'scedit') return render('SC_Edit', 'Manage Council');
    if (page === 'checkin') return render('CheckIn', 'Daily Attendance');
    if (page === 'ranking') return render('Ranking', 'Rank')
    if (page === 'adminrank') return render('adminrank', 'Admin Ranking')
    //if (page === 'ai') return render('MiniGemini', 'Mini Gemini AI');
    if (page === 'pages') return render('PagesFile', 'Website Directory');
    if (page === 'broadcast') return render('Broadcast', 'SC Announcement Center');
  
    // Default Portal (Login Page)
    return render('Index', 'Portal Login');
  } catch (err) {
    return HtmlService.createHtmlOutput("<h1>Portal Error</h1><p>" + err.toString() + "</p>");
  }
}

function render(file, title) {
  return HtmlService.createTemplateFromFile(file).evaluate()
      .setTitle(title)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 2. AUTHENTICATION & ADMIN USER CONTROL
 */
function handleAuth(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users") || ss.insertSheet("Users");
  const values = sheet.getDataRange().getValues();
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, data.password)
                .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');

  if (data.action === "signup") {
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === data.username) return { success: false, msg: "Username taken" };
    }
    // [Username, PasswordHash, Email, ?, Name, ?, Role]
    sheet.appendRow([data.username, hash, data.email, "", data.name, "", "user"]);
    return { success: true };
  }

  if (data.action === "login") {
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === data.username && values[i][1] === hash) {
        return { success: true, name: values[i][4], role: values[i][6] };
      }
    }
    return { success: false, msg: "Invalid Credentials" };
  }
}

function getAllUsers() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users").getDataRange().getValues();
}

function adminUpdateUser(rowIdx, d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  sheet.getRange(rowIdx + 1, 1).setValue(d.username);
  sheet.getRange(rowIdx + 1, 5).setValue(d.name);
  sheet.getRange(rowIdx + 1, 7).setValue(d.role);
  return "User Updated";
}

function adminDeleteUser(rowIdx) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users").deleteRow(rowIdx + 1);
  return "User Deleted";
}

/**
 * 3. SELLER POS & INVENTORY LOGIC (฿ Baht)
 */
function registerItem(name, price, qty) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Inventory") || ss.insertSheet("Inventory");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === name.toLowerCase()) {
      sheet.getRange(i + 1, 2).setValue(price);
      sheet.getRange(i + 1, 3).setValue(qty);
      return "Item Updated!";
    }
  }
  sheet.appendRow([name, price, qty]);
  return "New Item Registered!";
}

function reduceStock(name, amount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName("Inventory");
  const histSheet = ss.getSheetByName("History") || ss.insertSheet("History");
  const data = invSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      const newQty = Number(data[i][2]) - amount;
      if (newQty < 0) return "Error: No Stock!";
      invSheet.getRange(i + 1, 3).setValue(newQty);
      histSheet.appendRow([new Date(), name, "-" + amount, newQty]);
      return "Success! Remaining: " + newQty;
    }
  }
}

function getCustomerStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Inventory"); 
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // Skip the header row
  
  // Filter out empty rows and map to Objects
  return rows.filter(row => row[0] && row[0].toString().trim() !== "").map(row => {
    return {
      name: String(row[0]),
      price: isNaN(parseFloat(row[1])) ? 0 : parseFloat(row[1]),
      qty: isNaN(parseInt(row[2])) ? 0 : parseInt(row[2])
    };
  });
}
function loadStock() {
    document.getElementById('stockList').innerHTML = "Refreshing list...";
    google.script.run.withSuccessHandler(data => {
      if (!data || data.length === 0) {
        document.getElementById('stockList').innerHTML = "No items found in inventory.";
        return;
      }
      
      // Map data using Object properties (.name, .qty, .price)
      document.getElementById('stockList').innerHTML = data.map(item => `
        <div class="stock-item">
          <div class="item-details">
            <b>${item.name}</b><br>
            <small>Stock: ${item.qty} | ฿${item.price}</small>
          </div>
          <button class="reduce-btn" onclick="handleReduce('${item.name}')">SELL 1</button>
        </div>
      `).join('');
    }).getCustomerStock();
  }

function getHistory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("History");
  if (!sheet) return [];
  return sheet.getDataRange().getValues().reverse().slice(0, 25);
}

/**
 * 4. ATTENDANCE LOGIC
 */
function getAttendanceList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("attendance list");
  const today = new Date(); 
  
  let targetDate = new Date(today);
  const dayOfWeek = today.getDay(); 
  let daysToTuesday = (2 - dayOfWeek + 7) % 7;
  targetDate.setDate(today.getDate() + daysToTuesday);

  const dateStr = "DATE: " + Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "d MMMM").toUpperCase();
  let data = sheet.getDataRange().getValues();
  let headers = data[0];
  let currentCol = headers.indexOf(dateStr) + 1;

  if (currentCol <= 0) {
    currentCol = headers.length + 1;
    sheet.getRange(1, currentCol).setValue(dateStr).setBackground("#ffff00").setFontWeight("bold");
    if (sheet.getLastRow() > 1) sheet.getRange(2, currentCol, sheet.getLastRow() - 1, 1).setValue("A");
    data = sheet.getDataRange().getValues(); 
  }

  let students = [], groups = [], currentGroup = "General";
  for (let i = 0; i < data.length; i++) {
    let name = String(data[i][0]).trim();
    if (!name || name === "ATTENDANCE" || (name.startsWith("DATE:") && !name.toLowerCase().includes("and"))) continue;
    if (name.toLowerCase().includes(" and ")) {
      currentGroup = name; groups.push(name);
      students.push({ rowIdx: i + 1, name: name, type: "Header", group: currentGroup });
    } else {
      students.push({ rowIdx: i + 1, name: name, status: data[i][currentCol - 1] || "A", type: "Student", group: currentGroup });
    }
  }
  return { students, date: dateStr.replace("DATE: ", ""), availableGroups: groups, colIndex: currentCol };
}

function saveAttendance(rowIdx, status, colIdx) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("attendance list");
  sheet.getRange(rowIdx, colIdx).setValue(status);
  return "Saved";
}

function submitExcuse(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Excuses") || ss.insertSheet("Excuses");
  
  // If new sheet, add headers
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Name", "Class", "Guardian Name", "Reason for Absence"]);
  }
  
  sheet.appendRow([
    new Date(), 
    data.name, 
    data.class, 
    data.guardian, 
    data.reason
  ]);
  
  return "Your excuse has been submitted successfully.";
}
function getExcusesData(user, pass) {
  // Security Check
  if (user === "admin" && pass === "sisbpu") {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Ensure this matches your Tab Name exactly (case sensitive)
    const sheet = ss.getSheetByName("Excuses"); 
    
    if (!sheet) return "Error: Sheet 'Excuses' not found";
    
    const data = sheet.getDataRange().getValues();
    
    // If sheet only has headers or is empty
    if (data.length < 2) return []; 
    
    // Convert Dates to strings so they don't break during transfer to HTML
    const formattedData = data.map(row => {
      return [
        row[0] instanceof Date ? row[0].toLocaleDateString() : row[0], // Date
        row[1], // Name
        row[2], // Class
        row[3], // Guardian
        row[4]  // Reason
      ];
    });
    
    return formattedData;
  } else {
    return "Unauthorized Access";
  }
}
function getSheetDatas() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventory").getDataRange().getValues(); }
function getSheetData() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("list").getDataRange().getValues(); }
/**
 * Saves excuse to Sheet and emails Tonkla
 */
/**
 * Sends an email based on the selected SC leader
 */
/**
 * Sends a formal excuse email to the selected SC leader
 */
/**
 * Sends a formal excuse for a Prefects Meeting
 */
function sendSCNotification(leaderName, studentName, studentClass, guardianName, reason) {
  let recipientEmail = "";

  // Custom Logic for leader emails
  switch (leaderName) {
    case "Jena": recipientEmail = "sisbpusc@proton.me, sisbpusc1@proton.me"; break;
    case "Millin": recipientEmail = "sisbpusc@proton.me, sisbpusc1@proton.me"; break;
    case "Bena": recipientEmail = "sisbpusc@proton.me, sisbpusc1@proton.me"; break;
    case "Zizi": recipientEmail = "sisbpusc@proton.me, sisbpusc1@proton.me"; break;
    case "Tonkla": recipientEmail = "sisbpusc@proton.me, sisbpusc1@proton.me"; break;
    case "zenzen": recipientEmail = "sisbpusc@proton.me, sisbpusc1@proton.me"; break;
    default: recipientEmail = "tonkla.pcstudentt@sisbschool.com";
  }

  const subject = `PREFECTS MEETING EXCUSE: ${studentName} (${studentClass})`;
  
  const htmlBody = `
    <div style="font-family: 'Courier New', Courier, monospace; border: 4px solid #000; padding: 20px; max-width: 600px;">
      <h2 style="text-transform: uppercase; background: #000; color: #fff; padding: 10px; text-align: center;">Official Excuse Form</h2>
      <p><strong>To Leader:</strong> ${leaderName}</p>
      <hr>
      <p><strong>Prefect Name:</strong> ${studentName}</p>
      <p><strong>Class:</strong> ${studentClass}</p>
      <p><strong>Guardian Name:</strong> ${guardianName}</p>
      <p><strong>Reason for Absence:</strong></p>
      <div style="border: 1px solid #000; padding: 10px; background: #f0f0f0;">
        ${reason}
      </div>
      <p style="font-size: 10px; margin-top: 15px;">SISB Student Council Prefect Division</p>
    </div>
  `;

  MailApp.sendEmail({
    to: recipientEmail,
    subject: subject,
    htmlBody: htmlBody
  });

  return "Excuse for Prefects Meeting sent to " + leaderName;
}

function getMemberDuty(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("duty"); 
  if (!sheet) return { found: false };
  
  const data = sheet.getDataRange().getValues();
  const backgroundColors = sheet.getDataRange().getBackgrounds(); // To find the green headers
  
  let assignments = [];
  let firstName = "";

  for (let row = 0; row < data.length; row++) {
    for (let col = 1; col < data[row].length; col++) {
      let cellValue = data[row][col].toString();
      
      if (cellValue.toLowerCase().includes(name.toLowerCase())) {
        if (!firstName) firstName = cellValue.split(/[ -|]/)[0];

        // 1. Find the Location (Search UP for the Green Header)
        let location = "General Duty";
        for (let r = row; r >= 0; r--) {
          // Check for the bright green color (#66ff66 or similar)
          if (backgroundColors[r][col-1] === "#66ff66" || backgroundColors[r][col] === "#66ff66" || data[r][col].toString().includes("DUTY")) {
            location = data[r][col].toString().split('\n')[0]; // Gets "PLAYGROUND DUTY"
            break;
          }
        }

        // 2. Find the Day (Column A)
        let day = data[row][0] || "Scheduled Day";

        // 3. Find the Time Slot (The closest Header Row)
        // Usually, the nearest header is a few rows above the duty section
        let timeSlot = "Check Schedule";
        for (let r = row; r >= 0; r--) {
           if (data[r][col] && (data[r][col].toString().includes("WEEK") || data[r][col].toString().includes("SNACK"))) {
             timeSlot = data[r][col].toString();
             break;
           }
        }

        assignments.push({
          location: location,
          day: day,
          time: timeSlot,
          status: cellValue.includes("Excused") ? "EXCUSED" : "ASSIGNED"
        });
      }
    }
  }

  if (assignments.length > 0) {
    return { found: true, displayName: firstName, assignments: assignments };
  }
  return { found: false };
}
function getSCMembers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Members");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Only headers exist
  
  data.shift(); // Remove headers
  return data.map(row => ({
    name: row[0],
    role: row[1],
    grade: row[2],
    bio: row[3],
    img: row[4] || 'https://via.placeholder.com/100'
  }));
}

/**
 * ADD DATA: Save a new member to the "Members" sheet
 */
function addSCMember(member) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Members");
  
  if (!sheet) {
    sheet = ss.insertSheet("Members");
    sheet.appendRow(["Name", "Role", "Grade", "Bio", "ImageURL"]);
  }
  
  sheet.appendRow([member.name, member.role, member.grade, member.bio, member.img]);
  return "Added " + member.name + " to the Council Sheet!";
}
/**
 * Verification for the dedicated AI Page
 */
function checkSCAccess(name) {
  // Add all SC nicknames here
  const scRoster = ["Jena", "Millin", "Bena", "Zizi", "Tonkla", "zenzen",];
  
  const cleanName = name.trim().toLowerCase();
  const isFound = scRoster.some(member => member.toLowerCase() === cleanName);
  
  return { allowed: isFound, user: name };
}

/**
 * The Mini Gemini Engine
 */
function callMiniGemini(prompt, name) {
  const apiKey = "AAIzaSyD9azFN4kWF0GgjsHtEYnv8kd-gz41tUqM"; // Replace with your actual Gemini API Key
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=$AIzaSyD9azFN4kWF0GgjsHtEYnv8kd-gz41tUqM`;

  const context = `You are Mini Gemini, the assistant for SISB Student Council. 
    You are helping ${name}. Since students cannot use phones during school hours, 
    provide clear, actionable advice they can remember or write down for tomorrow. 
    Keep responses professional but friendly.`;

  const payload = {
    "contents": [{
      "parts": [{ "text": context + "\n\nUser Question: " + prompt }]
    }]
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    return result.candidates[0].content.parts[0].text;
  } catch (e) {
    return "Mini Gemini is currently resting. Please check your API key or connection.";
  }
}
/**
 * GET DATA: Finds everyone on duty for a SPECIFIC date
 */
/**
 * Scans the complex grid from the "PREFECTS DUTIES" sheet.
 * Handles grouped headers like Quad Duty, Stairway Duty, and Assembly.
 */
/**
 * Scans the Duty sheet based on the Term 2 PDF/Sheet layout.
 */
function getDutyList(week, day) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DutyData");
  const data = sheet.getDataRange().getValues();
  
  // Filter row[0] (Week) and row[1] (Day)
  return data.slice(1).filter(row => {
    return row[0].toString() == week && row[1].toLowerCase() == day.toLowerCase();
  }).map(row => ({
    week: row[0],
    day: row[1],
    location: row[2],
    name: row[3],
    shift: row[4]
  }));
}

/**
 * Logs attendance to the AttendanceLog sheet
 */
function markAttendance(name, status, location, shift, week, comment) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("AttendanceLog") || ss.insertSheet("AttendanceLog");
  
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(["Date", "Week", "Name", "Location", "Shift", "Status", "Comment", "Timestamp"]);
  }

  const finalStatus = (status === "PRESENT") ? "HERE" : "ABSENT";
  
  logSheet.appendRow([
    new Date().toLocaleDateString(), 
    "Week " + week,
    name, 
    location, 
    shift, 
    finalStatus, 
    comment || "No comment", // Save the comment here
    new Date()
  ]);

  if (status === "ABSENT") {
    MailApp.sendEmail({
      to: "sisbpusc@proton.me",
      subject: "URGENT: Prefect Missing from Duty",
      body: "Prefect: " + name + 
            "\nLocation: " + location + 
            "\nShift: " + shift + 
            "\nReason/Comment: " + (comment || "None provided") +
            "\nTime of check: " + new Date().toLocaleString()
    });
  }

  return "Recorded: " + name;
}
// 2. LEADERBOARD LOGIC
function getPublicLeaderboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PrefectScores");
  
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  // Skip headers, grab Name (Col A) and Total Score (Col D)
  return data.slice(1).map(row => ({
    name: row[0],
    score: row[3] || 0
  })).sort((a, b) => b.score - a.score);
}
// 3. ADMIN LOGIC
const ADMIN_PASS = "sisbpusc";
function verifyAdmin(password) {
  return password === ADMIN_PASS;
}

// Fetch list of prefects currently in the scoring system
function getAdminUserList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PrefectScores");
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1).map(row => ({ name: row[0], score: row[3] }));
}

// Add a new student to the "Good Prefect" list
function addNewPrefect(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PrefectScores") || ss.insertSheet("PrefectScores");
  if (sheet.getLastRow() === 0) sheet.appendRow(["Name", "Base", "Extra", "Total"]);
  
  const nextRow = sheet.getLastRow() + 1;
  sheet.appendRow([name, 0, 0, `=B${nextRow}+C${nextRow}`]);
  return "Success: " + name + " added.";
}

// Adjust points (+5 or -5 etc)
function updateScore(name, points) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PrefectScores");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      sheet.getRange(i + 1, 3).setValue((data[i][2] || 0) + parseInt(points));
      return "Score Updated for " + name;
    }
  }
}
/**
 * 1. BROADCAST ROUTER
 * Add this line to your existing doGet(e) function:
 * if (page === 'broadcast') return render('Broadcast', 'SC Announcement Center');
 */

/**
 * Updates the current live announcement AND logs it to a history sheet.
 * Automatically creates the sheets and headers if they don't exist.
 */
function updateAnnouncement(msg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Handle SETTINGS Sheet (The Live Message)
  let settingsSheet = ss.getSheetByName("Settings");
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("Settings");
    settingsSheet.getRange("A1:B1").setValues([["Key", "Value"]]).setFontWeight("bold");
    settingsSheet.getRange("A2").setValue("CurrentAnnouncement");
  }
  settingsSheet.getRange("B2").setValue(msg);
  
  // 2. Handle BROADCASTLOG Sheet (The History)
  let logSheet = ss.getSheetByName("BroadcastLog");
  if (!logSheet) {
    logSheet = ss.insertSheet("BroadcastLog");
    logSheet.appendRow(["Timestamp", "Announcement Message", "Status"]);
    logSheet.getRange("1:1").setFontWeight("bold").setBackground("#f3f3f3");
  }
  
  const timestamp = new Date();
  const status = (msg === "" || msg === null) ? "CLEARED" : "LIVE";
  logSheet.appendRow([timestamp, msg || "[Announcement Cleared]", status]);
  
  return "Broadcast Updated and Logged!";
}

/**
 * Fetches the current live announcement for the Index page
 */
function getAnnouncement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  if (!sheet) return ""; 
  return sheet.getRange("B2").getValue() || "";
}

/**
 * Fetches the last 5 announcements for the Admin UI History
 */
function getAnnouncementHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("BroadcastLog");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  // Get last 5 rows, reverse them so newest is first
  return data.slice(1).reverse().slice(0, 5).map(row => ({
    time: Utilities.formatDate(row[0], ss.getSpreadsheetTimeZone(), "MMM d, HH:mm"),
    msg: row[1],
    status: row[2]
  }));
}
