/**
 * ============================================================================
 * FAST-NUCES CFD — PROJECT GROUP & PARTNER PORTAL BACKEND (Google Apps Script)
 * ============================================================================
 * Features:
 * 1. Auto-creates & formats 5 sheets: Courses, Students, Groups, Requests, Backups
 * 2. Uses LockService to prevent race conditions when many students submit at once
 * 3. Automatically saves timestamped JSON snapshots to the 'Backups' sheet
 * 4. Includes installDailyBackupTrigger() for scheduled nightly backups
 * ============================================================================
 */

const SHEET_NAMES = {
  COURSES: "Courses",
  STUDENTS: "Students",
  GROUPS: "Groups",
  REQUESTS: "Requests",
  BACKUPS: "Backups"
};

/**
 * Run this function ONCE from the Apps Script editor (or it runs automatically)
 * to create and style all required tabs in your Google Spreadsheet.
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const schemas = [
    {
      name: SHEET_NAMES.COURSES,
      headers: ["Course ID", "Course Code", "Course Name", "Max Members", "Custom Fields (JSON)", "Status", "Created At"],
      color: "#1e293b"
    },
    {
      name: SHEET_NAMES.STUDENTS,
      headers: ["Roll No", "Full Name", "University Email", "WhatsApp", "PIN", "Registered At"],
      color: "#312e81"
    },
    {
      name: SHEET_NAMES.GROUPS,
      headers: ["Group ID", "Course ID", "Sort Order", "Group Name", "Leader Roll No", "Members (JSON)", "Custom Data (JSON)", "Open For Partner", "Updated At"],
      color: "#064e3b"
    },
    {
      name: SHEET_NAMES.REQUESTS,
      headers: ["Request ID", "Group ID", "Course ID", "Sender Roll No", "Sender Name", "Sender Email", "Sender WhatsApp", "Note", "Status", "Created At"],
      color: "#78350f"
    },
    {
      name: SHEET_NAMES.BACKUPS,
      headers: ["Backup ID", "Timestamp", "Trigger Event", "Total Courses", "Total Students", "Total Groups", "Full Snapshot (JSON)"],
      color: "#4c1d95"
    }
  ];

  schemas.forEach(schema => {
    let sheet = ss.getSheetByName(schema.name);
    if (!sheet) {
      sheet = ss.insertSheet(schema.name);
    }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(schema.headers);
      const headerRange = sheet.getRange(1, 1, 1, schema.headers.length);
      headerRange
        .setBackground(schema.color)
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setHorizontalAlignment("left");
      sheet.setFrozenRows(1);
    }
  });

  // Remove default blank 'Sheet1' if other tabs exist
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1 && defaultSheet.getLastRow() === 0) {
    ss.deleteSheet(defaultSheet);
  }
}

/**
 * Optional: Run this once in Apps Script editor to enable automatic nightly backups at 1:00 AM.
 */
function installDailyBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "scheduledDailyBackup") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("scheduledDailyBackup")
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
}

function scheduledDailyBackup() {
  setupSheets();
  createBackupSnapshot_("Scheduled Daily Auto-Backup");
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function safeParseJson_(str, fallback) {
  try {
    return str ? JSON.parse(str) : fallback;
  } catch (e) {
    return fallback;
  }
}

function readAllData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheets();

  // 1. Courses
  const cRows = ss.getSheetByName(SHEET_NAMES.COURSES).getDataRange().getValues();
  const courses = [];
  for (let i = 1; i < cRows.length; i++) {
    const r = cRows[i];
    if (!r[0]) continue;
    courses.push({
      courseId: String(r[0]),
      courseCode: String(r[1]),
      courseName: String(r[2]),
      maxMembers: Number(r[3]) || 2,
      customFields: safeParseJson_(r[4], []),
      status: String(r[5] || "OPEN")
    });
  }

  // 2. Students
  const sRows = ss.getSheetByName(SHEET_NAMES.STUDENTS).getDataRange().getValues();
  const students = [];
  for (let i = 1; i < sRows.length; i++) {
    const r = sRows[i];
    if (!r[0]) continue;
    students.push({
      rollNo: String(r[0]).toUpperCase(),
      name: String(r[1]),
      email: String(r[2]),
      whatsapp: String(r[3] || ""),
      pin: String(r[4])
    });
  }

  // 3. Groups
  const gRows = ss.getSheetByName(SHEET_NAMES.GROUPS).getDataRange().getValues();
  const groups = [];
  for (let i = 1; i < gRows.length; i++) {
    const r = gRows[i];
    if (!r[0]) continue;
    groups.push({
      groupId: String(r[0]),
      courseId: String(r[1]),
      slotNo: Number(r[2]) || i,
      groupName: String(r[3]),
      leaderRollNo: String(r[4]).toUpperCase(),
      members: safeParseJson_(r[5], []),
      customData: safeParseJson_(r[6], {}),
      lookingForPartner: r[7] === true || String(r[7]).toUpperCase() === "TRUE"
    });
  }

  // 4. Requests
  const reqRows = ss.getSheetByName(SHEET_NAMES.REQUESTS).getDataRange().getValues();
  const requests = [];
  for (let i = 1; i < reqRows.length; i++) {
    const r = reqRows[i];
    if (!r[0]) continue;
    requests.push({
      requestId: String(r[0]),
      groupId: String(r[1]),
      courseId: String(r[2]),
      senderRollNo: String(r[3]).toUpperCase(),
      senderName: String(r[4]),
      senderEmail: String(r[5]),
      senderWhatsApp: String(r[6] || ""),
      note: String(r[7] || ""),
      status: String(r[8] || "PENDING")
    });
  }

  const backupSheet = ss.getSheetByName(SHEET_NAMES.BACKUPS);
  const backupsCount = Math.max(0, backupSheet.getLastRow() - 1);

  return { ok: true, courses, students, groups, requests, backupsCount };
}

function createBackupSnapshot_(reason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupSheet = ss.getSheetByName(SHEET_NAMES.BACKUPS);
  if (!backupSheet) return null;

  const snapshot = readAllData_();
  // Strip PINs from backup JSON for extra privacy
  const cleanStudents = snapshot.students.map(s => ({
    rollNo: s.rollNo,
    name: s.name,
    email: s.email,
    whatsapp: s.whatsapp
  }));

  const payload = JSON.stringify({
    exportedAt: new Date().toISOString(),
    courses: snapshot.courses,
    students: cleanStudents,
    groups: snapshot.groups,
    requests: snapshot.requests
  });

  const backupId = "BKP_" + Date.now();
  const ts = Utilities.formatDate(new Date(), "Asia/Karachi", "yyyy-MM-dd HH:mm:ss");

  backupSheet.appendRow([
    backupId,
    ts,
    reason || "Auto Snapshot",
    snapshot.courses.length,
    snapshot.students.length,
    snapshot.groups.length,
    payload
  ]);

  // Keep the latest 30 backups to keep spreadsheet fast
  const totalDataRows = backupSheet.getLastRow() - 1;
  if (totalDataRows > 30) {
    backupSheet.deleteRows(2, totalDataRows - 30);
  }

  return { backupId, timestamp: ts };
}

function doGet(e) {
  try {
    return jsonResponse_(readAllData_());
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    setupSheets();

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nowStr = Utilities.formatDate(new Date(), "Asia/Karachi", "yyyy-MM-dd HH:mm:ss");

    // 1. REGISTER STUDENT
    if (action === "registerStudent") {
      const sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
      const rows = sheet.getDataRange().getValues();
      const rollNo = String(payload.rollNo || "").trim().toUpperCase();
      const email = String(payload.email || "").trim().toLowerCase();
      const name = String(payload.name || "").trim();
      const whatsapp = String(payload.whatsapp || "").trim();
      const pin = String(payload.pin || "").trim();

      if (!email.endsWith("@cfd.nu.edu.pk")) {
        return jsonResponse_({ ok: false, error: "Please use your official @cfd.nu.edu.pk email." });
      }
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).toUpperCase() === rollNo) {
          return jsonResponse_({ ok: false, error: "Roll Number is already registered. Please sign in." });
        }
        if (String(rows[i][2]).toLowerCase() === email) {
          return jsonResponse_({ ok: false, error: "This university email is already registered." });
        }
      }

      sheet.appendRow([rollNo, name, email, whatsapp, pin, nowStr]);
      return jsonResponse_({
        ok: true,
        student: { rollNo, name, email, whatsapp }
      });
    }

    // 2. LOGIN STUDENT
    if (action === "loginStudent") {
      const sheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
      const rows = sheet.getDataRange().getValues();
      const rollNo = String(payload.rollNo || "").trim().toUpperCase();
      const pin = String(payload.pin || "").trim();

      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).toUpperCase() === rollNo && String(rows[i][4]) === pin) {
          return jsonResponse_({
            ok: true,
            student: {
              rollNo: String(rows[i][0]).toUpperCase(),
              name: String(rows[i][1]),
              email: String(rows[i][2]),
              whatsapp: String(rows[i][3] || "")
            }
          });
        }
      }
      return jsonResponse_({ ok: false, error: "Invalid Roll Number or PIN." });
    }

    // 3. SAVE COURSE
    if (action === "saveCourse") {
      const sheet = ss.getSheetByName(SHEET_NAMES.COURSES);
      const rows = sheet.getDataRange().getValues();
      const courseId = String(payload.courseId);
      const rowData = [
        courseId,
        String(payload.courseCode || "").toUpperCase(),
        String(payload.courseName || ""),
        Number(payload.maxMembers) || 2,
        JSON.stringify(payload.customFields || []),
        String(payload.status || "OPEN"),
        nowStr
      ];

      let updated = false;
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === courseId) {
          sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
          updated = true;
          break;
        }
      }
      if (!updated) sheet.appendRow(rowData);

      createBackupSnapshot_("Course Saved: " + payload.courseCode);
      return jsonResponse_({ ok: true });
    }

    // 4. TOGGLE COURSE STATUS
    if (action === "toggleCourseStatus") {
      const sheet = ss.getSheetByName(SHEET_NAMES.COURSES);
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(payload.courseId)) {
          sheet.getRange(i + 1, 6).setValue(String(payload.status));
          break;
        }
      }
      return jsonResponse_({ ok: true });
    }

    // 5. DELETE COURSE
    if (action === "deleteCourse") {
      const cId = String(payload.courseId);
      deleteRowsByColMatch_(ss.getSheetByName(SHEET_NAMES.COURSES), 0, cId);
      deleteRowsByColMatch_(ss.getSheetByName(SHEET_NAMES.GROUPS), 1, cId);
      deleteRowsByColMatch_(ss.getSheetByName(SHEET_NAMES.REQUESTS), 2, cId);
      createBackupSnapshot_("Course Deleted: " + cId);
      return jsonResponse_({ ok: true });
    }

    // 6. CREATE GROUP
    if (action === "createGroup") {
      const sheet = ss.getSheetByName(SHEET_NAMES.GROUPS);
      const rows = sheet.getDataRange().getValues();
      const courseId = String(payload.courseId);
      const incomingMembers = payload.members || [];
      const incomingRolls = incomingMembers.map(m => String(m.rollNo).toUpperCase());

      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][1]) === courseId) {
          const existingMembers = safeParseJson_(rows[i][5], []);
          for (let m of existingMembers) {
            if (incomingRolls.indexOf(String(m.rollNo).toUpperCase()) !== -1) {
              return jsonResponse_({ ok: false, error: m.rollNo + " is already registered in a group for this course." });
            }
          }
        }
      }

      sheet.appendRow([
        String(payload.groupId),
        courseId,
        Number(payload.slotNo) || Date.now(),
        String(payload.groupName || ""),
        String(payload.leaderRollNo || "").toUpperCase(),
        JSON.stringify(incomingMembers),
        JSON.stringify(payload.customData || {}),
        Boolean(payload.lookingForPartner),
        nowStr
      ]);

      createBackupSnapshot_("Group Registered: " + payload.groupName);
      return jsonResponse_({ ok: true });
    }

    // 7. UPDATE GROUP
    if (action === "updateGroup") {
      const sheet = ss.getSheetByName(SHEET_NAMES.GROUPS);
      const rows = sheet.getDataRange().getValues();
      const groupId = String(payload.groupId);

      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === groupId) {
          sheet.getRange(i + 1, 4).setValue(String(payload.groupName || ""));
          sheet.getRange(i + 1, 7).setValue(JSON.stringify(payload.customData || {}));
          sheet.getRange(i + 1, 8).setValue(Boolean(payload.lookingForPartner));
          sheet.getRange(i + 1, 9).setValue(nowStr);
          break;
        }
      }
      return jsonResponse_({ ok: true });
    }

    // 8. DELETE GROUP
    if (action === "deleteGroup") {
      const gId = String(payload.groupId);
      deleteRowsByColMatch_(ss.getSheetByName(SHEET_NAMES.GROUPS), 0, gId);
      deleteRowsByColMatch_(ss.getSheetByName(SHEET_NAMES.REQUESTS), 1, gId);
      createBackupSnapshot_("Group Deleted: " + gId);
      return jsonResponse_({ ok: true });
    }

    // 9. SEND PARTNER REQUEST
    if (action === "sendRequest") {
      const sheet = ss.getSheetByName(SHEET_NAMES.REQUESTS);
      const rows = sheet.getDataRange().getValues();
      const groupId = String(payload.groupId);
      const senderRoll = String(payload.senderRollNo || "").toUpperCase();

      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][1]) === groupId && String(rows[i][3]).toUpperCase() === senderRoll && String(rows[i][8]) === "PENDING") {
          return jsonResponse_({ ok: false, error: "You already have a pending request for this group." });
        }
      }

      sheet.appendRow([
        String(payload.requestId),
        groupId,
        String(payload.courseId),
        senderRoll,
        String(payload.senderName || ""),
        String(payload.senderEmail || ""),
        String(payload.senderWhatsApp || ""),
        String(payload.note || ""),
        "PENDING",
        nowStr
      ]);
      return jsonResponse_({ ok: true });
    }

    // 10. RESPOND TO PARTNER REQUEST
    if (action === "respondRequest") {
      const reqSheet = ss.getSheetByName(SHEET_NAMES.REQUESTS);
      const reqRows = reqSheet.getDataRange().getValues();
      const requestId = String(payload.requestId);
      const decision = String(payload.decision);

      let targetRow = -1;
      let reqData = null;
      for (let i = 1; i < reqRows.length; i++) {
        if (String(reqRows[i][0]) === requestId) {
          targetRow = i + 1;
          reqData = reqRows[i];
          break;
        }
      }
      if (!reqData) return jsonResponse_({ ok: false, error: "Request not found." });

      if (decision === "REJECT") {
        reqSheet.getRange(targetRow, 9).setValue("REJECTED");
        return jsonResponse_({ ok: true });
      }

      // ACCEPT
      const groupId = String(reqData[1]);
      const courseId = String(reqData[2]);
      const senderRoll = String(reqData[3]).toUpperCase();

      // Lookup maxMembers for course
      const cRows = ss.getSheetByName(SHEET_NAMES.COURSES).getDataRange().getValues();
      let maxMembers = 2;
      for (let c = 1; c < cRows.length; c++) {
        if (String(cRows[c][0]) === courseId) {
          maxMembers = Number(cRows[c][3]) || 2;
          break;
        }
      }

      const grpSheet = ss.getSheetByName(SHEET_NAMES.GROUPS);
      const grpRows = grpSheet.getDataRange().getValues();

      // Ensure sender isn't already in a group in this course
      for (let g = 1; g < grpRows.length; g++) {
        if (String(grpRows[g][1]) === courseId) {
          const mList = safeParseJson_(grpRows[g][5], []);
          if (mList.some(m => String(m.rollNo).toUpperCase() === senderRoll)) {
            reqSheet.getRange(targetRow, 9).setValue("CANCELLED");
            return jsonResponse_({ ok: false, error: "This student has already joined another group." });
          }
        }
      }

      for (let g = 1; g < grpRows.length; g++) {
        if (String(grpRows[g][0]) === groupId) {
          const members = safeParseJson_(grpRows[g][5], []);
          if (members.length >= maxMembers) {
            return jsonResponse_({ ok: false, error: "Your group is already full." });
          }
          members.push({
            rollNo: senderRoll,
            name: String(reqData[4]),
            email: String(reqData[5]),
            whatsapp: String(reqData[6] || "")
          });
          const stillOpen = members.length < maxMembers;
          grpSheet.getRange(g + 1, 6).setValue(JSON.stringify(members));
          grpSheet.getRange(g + 1, 8).setValue(stillOpen);
          grpSheet.getRange(g + 1, 9).setValue(nowStr);
          break;
        }
      }

      reqSheet.getRange(targetRow, 9).setValue("ACCEPTED");
      createBackupSnapshot_("Partner Accepted: " + senderRoll);
      return jsonResponse_({ ok: true });
    }

    // 11. MANUAL CLOUD BACKUP SNAPSHOT
    if (action === "createBackup") {
      const bkp = createBackupSnapshot_("Manual Admin Backup");
      return jsonResponse_({ ok: true, backup: bkp });
    }

    return jsonResponse_({ ok: false, error: "Unknown action: " + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

function deleteRowsByColMatch_(sheet, colIndex, matchValue) {
  if (!sheet) return;
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][colIndex]) === String(matchValue)) {
      sheet.deleteRow(i + 1);
    }
  }
}
