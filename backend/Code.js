/**
 * ACC Substitute Request Form & Availability Management System
 * Backend Google Apps Script (GAS)
 */

// ==========================================
// 1. ENTRY POINTS (HTTP / EVENT HANDLERS)
// ==========================================

/**
 * Main Web App POST handler for form submissions.
 * Handles both "Substitute Addition" and "Substitute Request" forms.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    // Wait up to 10 seconds for exclusive lock to prevent concurrency race conditions
    const success = lock.tryLock(10000);
    if (!success) {
      return ContentService.createTextOutput("Server busy, please try again in a moment.")
        .setMimeType(ContentService.MimeType.TEXT);
    }

    const params = e.parameter || {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet1 = ss.getSheetByName("Sub Requests");
    const sheet2 = ss.getSheetByName("Sub Availability");

    // Call getHeaderMap EXACTLY ONCE per sheet at entry point (stores _lastColumn)
    const hdr1 = getHeaderMap(sheet1);
    const hdr2 = getHeaderMap(sheet2);

    const isAddition = params.formType === "addition";

    if (isAddition) {
      return handleSubAddition(sheet2, hdr2, params);
    } else {
      return handleSubRequest(sheet1, sheet2, hdr1, hdr2, params);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Web App GET handler for substitute acceptance links.
 */
function doGet(e) {
  const lock = LockService.getScriptLock();
  try {
    // Wait up to 10 seconds for exclusive lock to prevent double confirmations
    const success = lock.tryLock(10000);
    if (!success) {
      return createConfirmationHtml("Server Busy", "The server is currently busy. Please refresh the page in a few seconds.", false);
    }

    const params = e.parameter || {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Sub Requests");
    const subsSheet = ss.getSheetByName("Sub Availability");

    // Call getHeaderMap EXACTLY ONCE per sheet at entry point
    const hdr = getHeaderMap(sheet);
    const subsHdr = getHeaderMap(subsSheet);

    // Store sheet boundaries at top of function
    const sheet1LastRow = sheet.getLastRow();
    const subsLastRow = subsSheet.getLastRow();

    const rowIndex = parseInt(params.rowIndex, 10);
    const substituteName = (params.name || "").trim();

    if (!rowIndex || !substituteName) {
      return createConfirmationHtml("Missing Parameters", "Required request parameters were not provided.", false);
    }

    const actualName = (sheet.getRange(rowIndex, hdr["Available Substitutes"] + 1).getValue() || "").toString().trim();
    if (actualName.toLowerCase() !== substituteName.toLowerCase()) {
      return createConfirmationHtml("Verification Failed", "Could not verify substitute name at this request row.", false);
    }

    // Find the request row above the current substitute row
    const requestRow = findParentRequestRow(sheet, hdr, rowIndex);

    // Check if any substitute for this request is already confirmed
    if (isRequestAlreadyConfirmed(sheet, hdr, requestRow, sheet1LastRow)) {
      return createConfirmationHtml("Spot Filled", "This substitute spot has already been filled. Thank you!", false);
    }

    const teacherEmail = sheet.getRange(requestRow, hdr["Requester Email"] + 1).getValue();
    const teacherName = sheet.getRange(requestRow, hdr["Instructor Name"] + 1).getValue();
    const className = sheet.getRange(requestRow, hdr["Class Name"] + 1).getValue();
    const classDays = sheet.getRange(requestRow, hdr["Meeting Days"] + 1).getValue();
    const classTimes = sheet.getRange(requestRow, hdr["Meeting Times"] + 1).getValue();
    const classTime = `${classDays} ${classTimes}`;
    const classCode = sheet.getRange(requestRow, hdr["Class Code"] + 1).getValue();

    const rawDate = sheet.getRange(requestRow, hdr["Date Sub is Needed"] + 1).getDisplayValue();
    const formattedDate = formatDisplayDate(rawDate);

    const locationInfo = formatLocation(
      sheet.getRange(requestRow, hdr["Campus/Location"] + 1).getValue(),
      sheet.getRange(requestRow, hdr["Building & Room #"] + 1).getValue()
    );

    const timeStr = formatTimeRange(classTime);

    // Notify confirming substitute and requesting teacher using stored subsLastRow
    const subsData = subsLastRow > 1 ? subsSheet.getRange(1, 1, subsLastRow, subsHdr._lastColumn).getValues() : [];
    const confirmedSub = subsData.find(row => (row[subsHdr["Name"]] || "").toString().trim().toLowerCase() === substituteName.toLowerCase());

    if (confirmedSub) {
      const email = confirmedSub[subsHdr["Email"]];
      const phone = confirmedSub[subsHdr["Phone Number"]];
      const message = `<b>This is an automated email.</b><br><br>You have been confirmed for the substitute assignment for ${className}${locationInfo.campus}${locationInfo.roomWithIn}, on ${formattedDate}, subbing for ${teacherName}, which meets from ${timeStr}.`;

      sheet.getRange(rowIndex, hdr["Sub Confirmed"] + 1).setValue(substituteName);
      sheet.getRange(rowIndex, hdr["Sub Email"] + 1).setValue(email);

      // Set cell to the right of the checkbox on the substitute's row to "Confirmed"
      const statusColRightOfCheckbox = hdr["Available Substitutes"] + 3;
      sheet.getRange(rowIndex, statusColRightOfCheckbox).setValue("Confirmed");

      MailApp.sendEmail({ to: email, subject: "Substitute Confirmation", htmlBody: message });

      const teacherMsg = `<b>This is an automated email.</b><br><br>The substitute for your class, ${classCode}, has been confirmed.<br><br>Name: ${substituteName}<br>Email: ${email}<br>Phone Number: ${phone}`;
      MailApp.sendEmail({ to: teacherEmail, subject: "Substitute Confirmed", htmlBody: teacherMsg });
    }

    // Notify other matched substitutes that the spot has been filled
    try {
      notifyOtherSubstitutesSpotFilled(sheet, subsSheet, hdr, subsHdr, requestRow, substituteName, className, formattedDate, subsData, sheet1LastRow);
    } catch (err) {
      return createConfirmationHtml("Internal Error", "An internal error occurred: " + err, false);
    }

    return createConfirmationHtml(
      "Confirmation Recorded!",
      `Thank you ${substituteName}, your confirmation has been recorded.\n\nDejanira will reach out to you with further instructions. Please watch out for her email.`,
      true
    );
  } finally {
    lock.releaseLock();
  }
}

function createConfirmationHtml(title, message, isSuccess = true) {
  const icon = isSuccess ? "&#10004;" : "&#9888;";
  const iconBg = isSuccess ? "#e8f5e9" : "#ffebee";
  const iconColor = isSuccess ? "#2e7d32" : "#c62828";
  const headerBg = isSuccess ? "#003366" : "#8b0000";

  const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${title} | ACC Substitute System</title>
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      background-color: #f4f6f9;
      margin: 0;
      padding: 20px;
      display: flex;
      justify-content: center;
      align-items: center;
      min-height: 100vh;
    }
    .card {
      background: #ffffff;
      border-radius: 12px;
      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
      max-width: 480px;
      width: 100%;
      overflow: hidden;
      text-align: center;
    }
    .card-header {
      background-color: ${headerBg};
      color: #ffffff;
      padding: 24px 20px;
      font-size: 20px;
      font-weight: 700;
      letter-spacing: 0.5px;
    }
    .card-body {
      padding: 32px 24px;
    }
    .icon-circle {
      width: 64px;
      height: 64px;
      border-radius: 50%;
      background-color: ${iconBg};
      color: ${iconColor};
      font-size: 32px;
      line-height: 64px;
      margin: 0 auto 20px auto;
    }
    .card-title {
      font-size: 22px;
      color: #1a1a1a;
      margin: 0 0 12px 0;
      font-weight: 600;
    }
    .card-message {
      font-size: 15px;
      color: #555555;
      line-height: 1.6;
      margin: 0;
      white-space: pre-line;
    }
    .card-footer {
      background-color: #fafbfc;
      padding: 16px 24px;
      border-top: 1px solid #eeeeee;
      font-size: 13px;
      color: #888888;
    }
  </style>
</head>
<body>
  <div class="card">
    <div class="card-header">Austin Community College</div>
    <div class="card-body">
      <div class="icon-circle">${icon}</div>
      <h1 class="card-title">${title}</h1>
      <p class="card-message">${message}</p>
    </div>
    <div class="card-footer">ACC Substitute Request & Availability System</div>
  </div>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html).setTitle(title);
}

/**
 * OnEdit trigger handler for sending opportunity emails when checkboxes are clicked.
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "Sub Requests") return;

  const ss = e.source;
  const subsSheet = ss.getSheetByName("Sub Availability");
  const hdr = getHeaderMap(sheet);
  const subsHdr = getHeaderMap(subsSheet);

  const editedColumn = e.range.getColumn();
  const row = e.range.getRow();

  const subNamesCol = hdr["Available Substitutes"];
  const checkboxCol = subNamesCol + 2; // Column immediately following substitute name

  if (editedColumn !== checkboxCol) return;

  const checked = e.range.getValue();
  if (!checked) return;

  const name = (sheet.getRange(row, subNamesCol + 1).getValue() || "").toString().trim();
  const isRequestRow = (sheet.getRange(row, hdr["Program Type"] + 1).getValue() || "").toString().trim() === "";
  if (!isRequestRow || !name) return;

  const requestRow = findParentRequestRow(sheet, hdr, row);

  const teacherName = sheet.getRange(requestRow, hdr["Instructor Name"] + 1).getValue();
  const className = sheet.getRange(requestRow, hdr["Class Name"] + 1).getValue();
  const classDays = sheet.getRange(requestRow, hdr["Meeting Days"] + 1).getValue();
  const classTimes = sheet.getRange(requestRow, hdr["Meeting Times"] + 1).getValue();
  const classTime = `${classDays} ${classTimes}`;

  const rawDate = sheet.getRange(requestRow, hdr["Date Sub is Needed"] + 1).getDisplayValue();
  const formattedDate = formatDisplayDate(rawDate);

  const locationInfo = formatLocation(
    sheet.getRange(requestRow, hdr["Campus/Location"] + 1).getValue(),
    sheet.getRange(requestRow, hdr["Building & Room #"] + 1).getValue()
  );

  const timeStr = formatTimeRange(classTime);

  const subsLastRow = subsSheet.getLastRow();
  const subsData = subsLastRow > 1 ? subsSheet.getRange(1, 1, subsLastRow, subsHdr._lastColumn).getValues() : [];

  for (let i = 1; i < subsData.length; i++) {
    if ((subsData[i][subsHdr["Name"]] || "").toString().trim().toLowerCase() === name.toLowerCase()) {
      const email = subsData[i][subsHdr["Email"]];
      const encodedName = encodeURIComponent(name);
      const webAppUrl = PropertiesService.getScriptProperties().getProperty("WEB_APP_URL");
      const link = `${webAppUrl}?rowIndex=${row}&name=${encodedName}`;

      const msg = `<b>This is an automated email.</b><br><br>You have a sub opportunity for ${className}${locationInfo.campusPrefix}${locationInfo.roomPrefix} on ${formattedDate} from ${timeStr}, subbing for ${teacherName}. Please click this <a href="${link}" target="_blank">link</a> to accept.<br><br>If you can't accept this opportunity, please ignore this email.`;
      const subject = "Substitute Assignment Opportunity";

      MailApp.sendEmail({ to: email, subject: subject, htmlBody: msg });
      break;
    }
  }
}

// ==========================================
// 2. FORM SUBMISSION LOGIC
// ==========================================

function handleSubAddition(sheet2, hdr2, params) {
  const accEmail = (params.accEmail || "").toLowerCase().trim();
  const name = params.name || "";
  const phone = (params.phone || "").replace(/\D/g, "");
  const contactMethod = params.contactMethod || "";
  const daysAvailable = params.daysAvailable || "";
  const timesAvailable = params.timesAvailable || "";
  const availabilityMode = params.availabilityMode || "";
  const campuses = params.campuses || "";
  const classesCanTeach = params.classesCanTeach || "";

  if (!accEmail) {
    return ContentService.createTextOutput("Missing ACC email").setMimeType(ContentService.MimeType.TEXT);
  }

  const lastRow = sheet2.getLastRow();
  const lastCol = hdr2._lastColumn;
  const timestamp = getTimestampCST();

  const newRow = [];
  for (let i = 0; i < lastCol; i++) {
    switch (Object.keys(hdr2)[i]) {
      case "Timestamp": newRow.push(timestamp); break;
      case "ACC Email": newRow.push(accEmail); break;
      case "Name": newRow.push(name); break;
      case "Email": newRow.push(accEmail); break;
      case "Phone Number": newRow.push(phone); break;
      case "Preferred Method of Contact": newRow.push(contactMethod); break;
      case "Days Availiable to Sub": newRow.push(daysAvailable); break;
      case "Times Available to Sub": newRow.push(timesAvailable); break;
      case "In-person or Online": newRow.push(availabilityMode); break;
      case "Campuses Open to Sub": newRow.push(campuses); break;
      case "Classes Open to Sub": newRow.push(classesCanTeach); break;
      default: newRow.push(""); break;
    }
  }

  let found = false;
  if (lastRow > 1) {
    const emailColIndex = (hdr2["ACC Email"] !== undefined ? hdr2["ACC Email"] : 1) + 1;
    const emails = sheet2.getRange(2, emailColIndex, lastRow - 1, 1).getValues();
    for (let r = 0; r < emails.length; r++) {
      if ((emails[r][0] || "").toString().toLowerCase().trim() === accEmail) {
        sheet2.getRange(r + 2, 1, 1, newRow.length).setValues([newRow]);
        found = true;
        break;
      }
    }
  }

  if (!found) {
    sheet2.appendRow(newRow);
  }

  return ContentService.createTextOutput(found ? "Information updated" : "Substitute added successfully")
    .setMimeType(ContentService.MimeType.TEXT);
}

function handleSubRequest(sheet1, sheet2, hdr1, hdr2, params) {
  const program = params.program || "";
  const className = params.class || "";
  const instructor = params.instructor || "";
  const meetingTime = params.meetingTime || "";
  const dateNeeded = params.dateNeeded || "";
  const classType = params.classType || "";
  const location = params.classLocation || "";
  const building = params.buildingRoom || "";
  const teacherEmail = params.teacherEmail || "";
  const classCode = params.classCode || "";
  const note = params.note || "";

  const match = meetingTime.match(/^(.+?)\s+(\d{1,2}:\d{2}[AP]M.*)$/);
  const meetingDays = match ? match[1] : "";
  const meetingTimes = match ? match[2] : "";

  const dateObj = new Date(`${dateNeeded}T00:00:00`);
  const weekdayStr = Utilities.formatDate(dateObj, "UTC", "E");
  const formattedDate = `${weekdayStr} ${Utilities.formatDate(dateObj, "UTC", "MM-dd-yy")}`;

  const lastCol = hdr1._lastColumn;
  const timestamp = getTimestampCST();

  const newRow = [];
  for (let i = 0; i < lastCol; i++) {
    switch (Object.keys(hdr1)[i]) {
      case "Timestamp": newRow.push(timestamp); break;
      case "Program Type": newRow.push(program); break;
      case "Class Name": newRow.push(className); break;
      case "Instructor Name": newRow.push(instructor); break;
      case "Meeting Days": newRow.push(meetingDays); break;
      case "Meeting Times": newRow.push(meetingTimes); break;
      case "Date Sub is Needed": newRow.push(formattedDate); break;
      case "In-Person/Online": newRow.push(classType); break;
      case "Campus/Location": newRow.push(location); break;
      case "Building & Room #": newRow.push(building); break;
      case "Requester Email": newRow.push(teacherEmail); break;
      case "Class Code": newRow.push(classCode); break;
      case "Additional Notes": newRow.push(note); break;
      default: newRow.push(""); break;
    }
  }

  // Explicit row append recalculates getLastRow ONCE
  sheet1.appendRow(newRow);
  const newRowIndex = sheet1.getLastRow();

  updateAvailableSubstitutes(sheet1, sheet2, hdr1, hdr2, newRowIndex, newRow);

  return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);
}

// ==========================================
// 3. SUBSTITUTE MATCHING & COMPUTATION
// ==========================================

function updateAvailableSubstitutes(sheet1, sheet2, hdr1, hdr2, rowIndex, inMemoryRowValues = null) {
  const today = new Date(new Date().setUTCHours(0, 0, 0, 0));
  const subsLastRow = sheet2.getLastRow();
  const subsLastCol = hdr2._lastColumn;

  if (subsLastRow <= 1) return; // No substitutes to process

  const subs = sheet2.getRange(1, 1, subsLastRow, subsLastCol).getValues();

  // Use inMemoryRowValues if provided directly, otherwise fallback to single row fetch
  const requestRowValues = inMemoryRowValues || sheet1.getRange(rowIndex, 1, 1, hdr1._lastColumn).getValues()[0];

  const className = (requestRowValues[hdr1["Class Name"]] || "").toString().toLowerCase().trim();
  const classDays = requestRowValues[hdr1["Meeting Days"]] || "";
  const classTimes = requestRowValues[hdr1["Meeting Times"]] || "";
  const classTimeStr = `${classDays} ${classTimes}`;
  const cellValue = requestRowValues[hdr1["Date Sub is Needed"]];
  const dateNeededStr = (typeof cellValue === "string") ? cellValue.replace(/^\w+\s/, "") : cellValue;
  const dateNeeded = (dateNeededStr instanceof Date) ? dateNeededStr : new Date(new Date(dateNeededStr).setUTCHours(0, 0, 0, 0));
  const mode = (requestRowValues[hdr1["In-Person/Online"]] || "").toString().toLowerCase().trim();
  const location = (requestRowValues[hdr1["Campus/Location"]] || "").toString().toLowerCase().trim();

  if (isNaN(dateNeeded.getTime()) || dateNeeded < today) {
    sheet1.getRange(rowIndex, hdr1["Available Substitutes"] + 1).setValue("date parsing failed");
    return;
  }

  const weekday = getWeekdayAbbreviation(dateNeeded);
  const classTimeRange = parseClassTimeRange(classTimeStr);
  if (!classTimeRange) {
    sheet1.getRange(rowIndex, hdr1["Available Substitutes"] + 1).setValue("class time parsing failed");
    return;
  }

  const matches = subs.slice(1).filter(sub => {
    const name = sub[hdr2["Name"]];
    const subTimeJson = sub[hdr2["Times Available to Sub"]] || "";
    const subMode = (sub[hdr2["In-person or Online"]] || "").toString().toLowerCase().trim();
    const subCampuses = (sub[hdr2["Campuses Open to Sub"]] || "").toString().toLowerCase();
    const subClasses = (sub[hdr2["Classes Open to Sub"]] || "").toString().toLowerCase();

    const modeMatches = subMode === "both" || subMode === mode;
    const campusMatches = mode !== "in-person" || subCampuses.includes(location);
    const classMatches = subClasses.includes(className);
    const timeMatches = isSubAvailableThatDay(subTimeJson, weekday, classTimeRange, sheet1, hdr1);

    return name && modeMatches && campusMatches && classMatches && timeMatches;
  }).map(sub => sub[hdr2["Name"]]);

  insertSubstituteRows(sheet1, hdr1, rowIndex, matches);
}

function isSubAvailableThatDay(timeJson, weekday, classTimeRange, sheet1, hdr1) {
  try {
    const availability = JSON.parse(timeJson);
    const keyAliases = {
      "Tue": ["Tue", "Tues"],
      "Thu": ["Thu", "Thurs"]
    };
    const possibleKeys = keyAliases[weekday] || [weekday];
    let window = null;
    for (const k of possibleKeys) {
      if (availability[k]) {
        window = availability[k];
        break;
      }
    }
    if (!window) return false;
    const [startStr, endStr] = window.split("to").map(s => s.trim());
    const subStart = parseTimeString(startStr);
    const subEnd = parseTimeString(endStr);
    return classTimeRange.start >= subStart && classTimeRange.end <= subEnd;
  } catch (e) {
    if (sheet1 && hdr1) {
      sheet1.getRange(2, hdr1["Available Substitutes"] + 1).setValue("JSON parse error");
    }
    return false;
  }
}

function insertSubstituteRows(sheet1, hdr1, rowIndex, names) {
  if (!names || names.length === 0) return;
  const subCol = hdr1["Available Substitutes"] + 1;

  sheet1.insertRowsAfter(rowIndex, names.length);

  // Build 2D array in memory: [Name, false (checkbox value)]
  const values = names.map(name => [name, false]);

  // Bulk update substitute names and checkboxes in 2 single API calls instead of a loop
  const targetRange = sheet1.getRange(rowIndex + 1, subCol, names.length, 2);
  targetRange.setValues(values);
  sheet1.getRange(rowIndex + 1, subCol + 1, names.length, 1).insertCheckboxes();
}

function setupCheckboxesForRow(sheet1, hdr1, rowIndex) {
  const subsColIndex = hdr1["Available Substitutes"];
  if (subsColIndex === undefined) {
    throw new Error("Header 'Available Substitutes' not found.");
  }

  const namesCell = sheet1.getRange(rowIndex, subsColIndex + 1);
  const names = namesCell.getValue().toString().split("\n").filter(Boolean);

  const checkboxRange = sheet1.getRange(rowIndex, subsColIndex + 2, 1, names.length);
  checkboxRange.insertCheckboxes();
  checkboxRange.setValues([Array(names.length).fill(false)]);
}

// ==========================================
// 4. HELPER UTILITIES
// ==========================================

function getHeaderMap(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((name, i) => {
    map[name.toString().trim()] = i;
  });
  Object.defineProperty(map, "_lastColumn", {
    value: lastCol,
    enumerable: false,
    configurable: true,
    writable: true
  });
  return map;
}

function getTimestampCST() {
  return Utilities.formatDate(new Date(), "America/Chicago", "MM/dd/yyyy hh:mm a");
}

function getWeekdayAbbreviation(date) {
  const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  return days[date.getUTCDay()];
}

function parseTimeString(str) {
  if (!str) return null;
  if (/^\d+$/.test(str.trim())) {
    const hours = parseInt(str.trim(), 10);
    return hours * 60;
  }
  if (!/[APM]/i.test(str)) {
    const [hours, minutes] = str.split(":").map(Number);
    return hours * 60 + minutes;
  }
  str = str.toUpperCase().replace(/(\d{1,2}:\d{2})(AM|PM)/, "$1 $2").trim();
  const time = str.replace(/[^0-9:]/g, "").trim();
  const [hours, minutes] = time.split(":").map(Number);
  if (str.includes("PM") && hours < 12) return (hours + 12) * 60 + minutes;
  return hours * 60 + minutes;
}

function parseClassTimeRange(timeStr) {
  const match = (timeStr || "").match(/(\d{1,2}:\d{2}[APMapm]{2})\s+(\d{1,2}:\d{2}[APMapm]{2})/);
  if (!match) return null;
  return {
    start: parseTimeString(match[1]),
    end: parseTimeString(match[2])
  };
}

function findParentRequestRow(sheet, hdr, currentRow) {
  let requestRow = currentRow;
  while (requestRow > 2 && (sheet.getRange(requestRow - 1, hdr["Program Type"] + 1).getValue() || "").toString().trim() === "") {
    requestRow--;
  }
  return requestRow - 1;
}

function isRequestAlreadyConfirmed(sheet, hdr, requestRow, sheet1LastRow) {
  let i = requestRow + 1;
  const lastRow = sheet1LastRow || sheet.getLastRow();
  while (i <= lastRow && (sheet.getRange(i, hdr["Program Type"] + 1).getValue() || "").toString().trim() === "") {
    const status = (sheet.getRange(i, hdr["Class Code"] + 1).getValue() || "").toString().toLowerCase().trim();
    if (status === "confirmed") {
      return true;
    }
    i++;
  }
  return false;
}

function formatLocation(rawCampus, rawRoom) {
  let campus = rawCampus || "";
  let room = rawRoom || "";

  let campusPrefix = "";
  let roomPrefix = "";
  let roomWithIn = "";

  if (room !== "") {
    roomWithIn = " in " + room;
    room = " " + room;
  }

  if (campus === "" || campus === "DIL") {
    campusPrefix = " (online)";
    campus = " online";
    roomPrefix = "";
    roomWithIn = "";
  } else {
    campusPrefix = " at " + campus;
    campus = " at " + campus;
    roomPrefix = " in " + room.trim();
  }

  return { campus, room, campusPrefix, roomPrefix, roomWithIn };
}

function formatTimeRange(classTime) {
  const timeMatch = (classTime || "").match(/\d{1,2}:\d{2}\s*[APMapm]{2}/g);
  return (timeMatch && timeMatch.length === 2) ? `${timeMatch[0]} to ${timeMatch[1]}` : classTime;
}

function formatDisplayDate(rawDate) {
  const parts = (rawDate || "").replace(",", "").split(/[\s\-]+/);
  if (parts.length >= 4) {
    const [weekday, month, day, year] = parts;
    return `${weekday} ${month}/${day}/${year}`;
  }
  return rawDate;
}

function notifyOtherSubstitutesSpotFilled(sheet, subsSheet, hdr, subsHdr, requestRow, confirmingSubName, className, formattedDate, subsData, sheet1LastRow) {
  let i = requestRow + 1;
  const lastRow = sheet1LastRow || sheet.getLastRow();

  while (i <= lastRow && (sheet.getRange(i, hdr["Program Type"] + 1).getValue() || "").toString().trim() === "") {
    const name = (sheet.getRange(i, hdr["Available Substitutes"] + 1).getValue() || "").toString().trim();
    const isChecked = sheet.getRange(i, hdr["Available Substitutes"] + 2).getValue() === true;

    if (name && name.toLowerCase() !== confirmingSubName.toLowerCase() && isChecked) {
      const matchRow = subsData.find(r => (r[subsHdr["Name"]] || "").toString().trim().toLowerCase() === name.toLowerCase());
      if (matchRow) {
        const subEmail = matchRow[subsHdr["Email"]];
        const msg = `<b>This is an automated email.</b><br><br>The substitute spot for ${className} on ${formattedDate} has already been filled by another teacher. We will reach out to you if there are other substitute assignments coming up.<br><br>Thank you.`;
        MailApp.sendEmail({ to: subEmail, subject: "Substitute Spot Filled", htmlBody: msg });
      }
    }
    i++;
  }
}