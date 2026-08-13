function getHeaderMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((name, i) => {
    map[name.trim()] = i;
  });
  return map;
}

function doPost(e) {
  const params = e.parameter;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const isAddition = params.formType === "addition";
  const sheet1 = ss.getSheetByName("Sub Requests");
  const sheet2 = ss.getSheetByName("Sub Availability");

  const hdr1 = getHeaderMap(sheet1);
  const hdr2 = getHeaderMap(sheet2);

  if (isAddition) {
    const accEmail = (params.accEmail || "").toLowerCase().trim();
    const name = params.name || "";
    const email = accEmail;
    const phone = params.phone.replace(/\D/g, "") || "";
    const contactMethod = params.contactMethod || "";
    const daysAvailable = params.daysAvailable || "";
    const timesAvailable = params.timesAvailable || "";
    const availabilityMode = params.availabilityMode || "";
    const campuses = params.campuses || "";
    const classesCanTeach = params.classesCanTeach || "";

    if (!accEmail) {
      return ContentService.createTextOutput("Missing ACC email").setMimeType(ContentService.MimeType.TEXT);
    }

    const data = sheet2.getDataRange().getValues();

    // Row builder for Sub Availability
    const newRow = [];
    let j = 0;
    for (let i = 0; i < data[0].length; i++) {
      switch (Object.keys(hdr2)[i]) {
        case "Timestamp":
          var timeCST = Utilities.formatDate(
            new Date(),
            "America/Chicago",
            "MM/dd/yyyy hh:mm a" // e.g., 09/28/2025 01:30 PM
          ); 
          newRow.push(timeCST);
          break;
        case "ACC Email": newRow.push(accEmail); break;
        case "Name": newRow.push(name); break;
        case "Email": newRow.push(email); break;
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

    for (let i = 1; i < data.length; i++) {
      if ((data[i][hdr2["ACC Email"]] || "").toLowerCase().trim() === accEmail) {
        sheet2.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);
        found = true;
        break;
      }
    }

    if (!found) {
      sheet2.appendRow(newRow);
    }

    return ContentService.createTextOutput(found ? "Information updated" : "Substitute added successfully").setMimeType(ContentService.MimeType.TEXT);
  } else {
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

    const [placeholder, meetingDays, meetingTimes] = meetingTime.match(/^(.+?)\s+(\d{1,2}:\d{2}[AP]M.*)$/);

    const dateObj = new Date(`${dateNeeded}T00:00:00`);
    const weekdayStr = Utilities.formatDate(dateObj, "UTC", "E");
    const formattedDate = `${weekdayStr} ${Utilities.formatDate(dateObj, "UTC", "MM-dd-yy")}`;

    const data = sheet1.getDataRange().getValues();

    // Row builder for Sub Requests
    const newRow = [];
    let j = 0;
    for (let i = 0; i < data[0].length; i++) {
      switch (Object.keys(hdr1)[i]) {
        case "Timestamp": 
          var timeCST = Utilities.formatDate(
            new Date(),
            "America/Chicago",
            "MM/dd/yyyy hh:mm a" // e.g., 09/28/2025 01:30 PM
          ); 
          newRow.push(timeCST);
          break;
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

    sheet1.appendRow(newRow);

    updateAvailableSubstitutes(sheet1, sheet2, sheet1.getLastRow());
    return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);
  }
}

function setupCheckboxesForRow(sheet1, rowIndex) {
  const hdr1 = getHeaderMap(sheet1);
  const subsColIndex = hdr1["Available Substitutes"];
  if (subsColIndex === undefined) {
    throw new Error("Header 'Available Substitutes' not found.");
  }

  const namesCell = sheet1.getRange(rowIndex, subsColIndex + 1);
  const names = namesCell.getValue().split("\n").filter(Boolean);

  const checkboxRange = sheet1.getRange(rowIndex, subsColIndex + 2, 1, names.length);
  checkboxRange.insertCheckboxes();
  checkboxRange.setValues([Array(names.length).fill(false)]);
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();

  if (sheet.getName() !== "Sub Requests") return;

  const hdr = getHeaderMap(sheet);
  const editedColumn = e.range.getColumn();
  const row = e.range.getRow();

  // Ensure the edited column is the checkbox column
  const subNamesCol = hdr["Available Substitutes"];
  const checkboxStartCol = subNamesCol + 1;
  if (editedColumn != checkboxStartCol + 1) return;

  const checked = e.range.getValue();
  if (!checked) return;

  const name = (sheet.getRange(row, subNamesCol + 1).getValue() || "").trim();
  const isRequestRow = (sheet.getRange(row, hdr["Program Type"] + 1).getValue() || "").toString().trim() === "";
  if (!isRequestRow || !name) return;

  // Find the request row
  let requestRow = row;
  while (requestRow > 2 && sheet.getRange(requestRow - 1, hdr["Program Type"] + 1).getValue() === "") {
    requestRow--;
  }
  requestRow--;

  const teacherName = sheet.getRange(requestRow, hdr["Instructor Name"] + 1).getValue();
  const className = sheet.getRange(requestRow, hdr["Class Name"] + 1).getValue();

  const classDays = sheet.getRange(requestRow, hdr["Meeting Days"] + 1).getValue();
  const classTimes = sheet.getRange(requestRow, hdr["Meeting Times"] + 1).getValue();
  const classTime = `${classDays} ${classTimes}`;

  const rawDate = sheet.getRange(requestRow, hdr["Date Sub is Needed"] + 1).getDisplayValue();
  const [weekday, month, day, year] = rawDate.replace(",", "").split(/[\s\-]+/);
  let campus = sheet.getRange(requestRow, hdr["Campus/Location"] + 1).getValue();
  let room = sheet.getRange(requestRow, hdr["Building & Room #"] + 1).getValue();

  if (campus === "" || campus === "DIL") {
    campus = "(online)";
    room = "";
  } else {
    campus = " at " + campus;
    room = " in " + room;
  }

  const timeMatch = (classTime || "").match(/\d{1,2}:\d{2}\s*[APMapm]{2}/g);
  const timeStr = (timeMatch && timeMatch.length === 2) ? `${timeMatch[0]} to ${timeMatch[1]}` : classTime;

  const subsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sub Availability");
  const subsData = subsSheet.getDataRange().getValues();
  const subsHdr = getHeaderMap(subsSheet);

  for (let i = 1; i < subsData.length; i++) {
    if ((subsData[i][subsHdr["Name"]] || "").trim().toLowerCase() === name.toLowerCase()) {
      const email = subsData[i][subsHdr["Email"]];

      const encodedName = encodeURIComponent(name);
      const webAppUrl = PropertiesService.getScriptProperties().getProperty("WEB_APP_URL");
      const link = `${webAppUrl}?rowIndex=${row}&name=${encodedName}`;

      const msg = `You have a sub opportunity for ${className}${campus}${room} on ${weekday} ${month}/${day}/${year} from ${timeStr}, subbing for ${teacherName}. Please click this <a href="${link}" target="_blank">link</a> to accept.\nIf you can't accept this opportunity, please ignore this email.`;
      const subject = "Substitute Assignment Opportunity";

      MailApp.sendEmail({ to: email, subject: subject, htmlBody: msg });
      break;
    }
  }
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sub Requests");
  const subsSheet = ss.getSheetByName("Sub Availability");

  const hdr = getHeaderMap(sheet);
  const subsHdr = getHeaderMap(subsSheet);

  const rowIndex = parseInt(e.parameter.rowIndex);
  const substituteName = (e.parameter.name || "").trim();

  if (!rowIndex || !substituteName) {
    return HtmlService.createHtmlOutput("Missing parameters.");
  }

  const actualName = (sheet.getRange(rowIndex, hdr["Available Substitutes"] + 1).getValue() || "").trim();
  if (actualName.toLowerCase() !== substituteName.toLowerCase()) {
    return HtmlService.createHtmlOutput("Could not verify substitute name at this row.");
  }

  // Determine the request row above the current sub row
  let requestRow = rowIndex;
  while (requestRow > 1 && sheet.getRange(requestRow - 1, hdr["Program Type"] + 1).getValue() === "") {
    requestRow--;
  }
  requestRow--;

  // Check if any substitute for this request is already confirmed
  let i = requestRow + 1;
  while (i <= sheet.getLastRow() && sheet.getRange(i, hdr["Program Type"] + 1).getValue() === "") {
    const status = (sheet.getRange(i, hdr["Class Code"] + 1).getValue() || "").toString().toLowerCase();
    if (status === "confirmed") {
      return HtmlService.createHtmlOutput("This substitute spot has already been filled.");
    }
    i++;
  }

  // Mark this substitute as confirmed
  sheet.getRange(rowIndex, hdr["Class Code"] + 1).setValue("Confirmed");

  // Get request details
  const classCode = sheet.getRange(requestRow, hdr["Class Code"] + 1).getValue();
  const className = sheet.getRange(requestRow, hdr["Class Name"] + 1).getValue();
  const teacherName = sheet.getRange(requestRow, hdr["Instructor Name"] + 1).getValue();

  const classDays = sheet.getRange(requestRow, hdr["Meeting Days"] + 1).getValue();
  const classTimes = sheet.getRange(requestRow, hdr["Meeting Times"] + 1).getValue();
  const classTime = `${classDays} ${classTimes}`;

  const rawDate = sheet.getRange(requestRow, hdr["Date Sub is Needed"] + 1).getDisplayValue();
  const [weekday, month, day, year] = rawDate.replace(",", "").split(/[\s\-]+/);
  const formattedDate = `${weekday} ${month}/${day}/${year}`;
  let campus = sheet.getRange(requestRow, hdr["Campus/Location"] + 1).getValue();
  let room = sheet.getRange(requestRow, hdr["Building & Room #"] + 1).getValue();
  const teacherEmail = sheet.getRange(requestRow, hdr["Requester Email"] + 1).getValue();
  if (room !== "") {
    room = " " + room;
  }
  if (campus === "" || campus === "DIL") {
    campus = "online";
    room = "";
  } else {
    campus = " at " + campus;
    room = " in" + room;
  }

  const timeMatch = (classTime || "").match(/\d{1,2}:\d{2}\s*[APMapm]{2}/g);
  const timeStr = (timeMatch && timeMatch.length === 2) ? `${timeMatch[0]} to ${timeMatch[1]}` : classTime;

  // Notify confirming substitute
  const subsData = subsSheet.getDataRange().getValues();
  const confirmed = subsData.find(row => (row[subsHdr["Name"]] || "").trim().toLowerCase() === substituteName.toLowerCase());
  if (confirmed) {
    const email = confirmed[subsHdr["Email"]];
    const phone = confirmed[subsHdr["Phone Number"]];
    const message = `You have been confirmed for the substitute assignment for ${className}${campus}${room}, on ${formattedDate}, subbing for ${teacherName}, which meets from ${timeStr}.`;

    sheet.getRange(rowIndex, hdr["Sub Confirmed"] + 1).setValue(substituteName);
    sheet.getRange(rowIndex, hdr["Sub Email"] + 1).setValue(email);

    MailApp.sendEmail(email, "Substitute Confirmation", message);
    // Notify teacher
    const teacherMsg = `The substitute for your class, ${classCode}, has been confirmed.\n\nName: ${substituteName}\nEmail: ${email}\nPhone Number: ${phone}`;
    MailApp.sendEmail(teacherEmail, "Substitute Confirmed", teacherMsg);
  }

  // Notify other substitutes that the spot is filled
  try {
    i = requestRow + 1;
    while (i <= sheet.getLastRow() && sheet.getRange(i, hdr["Program Type"] + 1).getValue() === "") {
      const name = (sheet.getRange(i, hdr["Available Substitutes"] + 1).getValue() || "").trim();
      if (name && name.toLowerCase() !== substituteName.toLowerCase() && sheet.getRange(i, hdr["Available Substitutes"] + 2).getValue() === true) {
        const row = subsData.find(r => (r[subsHdr["Name"]] || "").trim().toLowerCase() === name.toLowerCase());
        if (row) {
          const newEmail = row[subsHdr["Email"]];
          const msg = `The substitute spot for ${className} on ${formattedDate} has already been filled by another teacher. We will reach out to you if there are other substitute assignments coming up.\nThank you.`;
          MailApp.sendEmail(newEmail, "Substitute Spot Filled", msg);
        }
      }
      i++;
    }
  } catch (err) {
    return HtmlService.createHtmlOutput("An internal error occurred: " + err);
  }

  return HtmlService.createHtmlOutput(`Thank you ${substituteName}, your confirmation has been recorded. Dejanira will reach out to you with further instructions. Please watch out for her email.\nThank you.`);
}

function getWeekdayAbbreviation(date) {
  const days = ["Sun", "Mon", "Tues", "Wed", "Thurs", "Fri", "Sat"];
  return days[date.getUTCDay()];
}

function parseTimeString(str) {
  if (!str) return 1;
  str = str.trim();
  if (!/AM|PM/i.test(str)) {
    const [hours, minutes] = str.split(":").map(Number);
    if (isNaN(hours) || isNaN(minutes)) return 3;
    return hours * 60 + minutes;
  }
  str = str.toUpperCase().replace(/(\d{1,2}:\d{2})(AM|PM)/, "$1 $2").trim();
  let time = str.replace(/[^0-9:]/g, '').trim();
  const [hours, minutes] = time.split(":").map(Number);
  if (str.includes("PM") && hours < 12) return (hours + 12) * 60 + minutes;
  return hours * 60 + minutes;
}

function parseClassTimeRange(timeStr) {
  const match = timeStr.match(/(\d{1,2}:\d{2}[APMapm]{2})\s+(\d{1,2}:\d{2}[APMapm]{2})/);
  if (!match) return null;
  return {
    start: parseTimeString(match[1]),
    end: parseTimeString(match[2])
  };
}

function isSubAvailableThatDay(timeJson, weekday, classTimeRange) {
  try {
    const availability = JSON.parse(timeJson);
    const window = availability[weekday];
    if (!window) return false;
    const [startStr, endStr] = window.split("to").map(s => s.trim());
    const subStart = parseTimeString(startStr);
    const subEnd = parseTimeString(endStr);
    const classStart = classTimeRange.start;
    const classEnd = classTimeRange.end;
    return classStart >= subStart && classEnd <= subEnd;
  } catch (e) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sub Requests");
    const hdr = getHeaderMap(sheet);
    sheet.getRange(2, hdr["Available Substitutes"] + 1).setValue("JSON parse error");
    return false;
  }
}

function insertSubstituteRows(sheet1, rowIndex, names) {
  if (names.length === 0) return;
  const hdr = getHeaderMap(sheet1);
  const subCol = hdr["Available Substitutes"] + 1;
  const confirmCol = subCol + 1;
  sheet1.insertRowsAfter(rowIndex, names.length);
  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    const targetRow = rowIndex + 1 + i;
    sheet1.getRange(targetRow, subCol).setValue(name);
    sheet1.getRange(targetRow, confirmCol).insertCheckboxes();
  }
}

function updateAvailableSubstitutes(sheet1, sheet2, rowIndex) {
  const hdr1 = getHeaderMap(sheet1);
  const hdr2 = getHeaderMap(sheet2);

  const today = new Date(new Date().setUTCHours(0, 0, 0, 0));
  const subs = sheet2.getDataRange().getValues();
  const i = rowIndex - 1;
  const requests = sheet1.getDataRange().getValues();

  const className = (requests[i][hdr1["Class Name"]] || "").toLowerCase().trim();
  const classDays = requests[i][hdr1["Meeting Days"]] || "";
  const classTimes = requests[i][hdr1["Meeting Times"]] || "";
  const classTimeStr = `${classDays} ${classTimes}`;
  const cellValue = sheet1.getRange(i + 1, hdr1["Date Sub is Needed"] + 1).getValue();
  const dateNeededStr = (typeof cellValue === "string") ? cellValue.replace(/^\w+\s/, "") : cellValue;
  const dateNeeded = (dateNeededStr instanceof Date) ? dateNeededStr : new Date(new Date(dateNeededStr).setUTCHours(0, 0, 0, 0));
  const mode = (requests[i][hdr1["In-Person/Online"]] || "").toLowerCase().trim();
  const location = (requests[i][hdr1["Campus/Location"]] || "").toLowerCase().trim();

  if (isNaN(dateNeeded.getTime()) || dateNeeded < today) {
    sheet1.getRange(i + 1, hdr1["Available Substitutes"] + 1).setValue("date parsing failed");
    return;
  }

  const weekday = getWeekdayAbbreviation(dateNeeded);
  const classTimeRange = parseClassTimeRange(classTimeStr);
  if (!classTimeRange) {
    sheet1.getRange(i + 1, hdr1["Available Substitutes"] + 1).setValue("class time parsing failed");
    return;
  }

  const matches = subs.slice(1).filter(sub => {
    const name = sub[hdr2["Name"]];
    const subTimeJson = sub[hdr2["Times Available to Sub"]] || "";
    const subMode = (sub[hdr2["In-person or Online"]] || "").toLowerCase().trim();
    const subCampuses = (sub[hdr2["Campuses Open to Sub"]] || "").toLowerCase();
    const subClasses = (sub[hdr2["Classes Open to Sub"]] || "").toLowerCase();

    const modeMatches = subMode === "both" || subMode === mode;
    const campusMatches = mode !== "in-person" || subCampuses.includes(location);
    const classMatches = subClasses.includes(className);
    const timeMatches = isSubAvailableThatDay(subTimeJson, weekday, classTimeRange);
    return name && modeMatches && campusMatches && classMatches && timeMatches;
  }).map(sub => sub[hdr2["Name"]]);

  insertSubstituteRows(sheet1, i + 1, matches);
}