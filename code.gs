// === CONFIGURATION ===
const SPREADSHEET_ID = "1FotLFASWuFinDnvpyLTsyO51OpJeKWtuG31VFje3Oik"; // Only the ID
const SHEET_NAMES = {
  adherence: "Adherence Tracker",
  database: "Data Base",
  schedule: "Schedules",
  logs: "Logs",
  otherCodes: "Other Codes",
  leaveRequests: "Leave Requests", 
  coaching_OLD: "Coaching", // Renamed old sheet
  coachingSessions: "CoachingSessions", // NEW
  coachingScores: "CoachingScores" // NEW
};
// --- Break Time Configuration (in seconds) ---
const PLANNED_BREAK_SECONDS = 15 * 60; // 15 minutes
const PLANNED_LUNCH_SECONDS = 30 * 60; // 30 minutes

// --- Shift Cutoff Hour (e.g., 7 = 7 AM) ---
const SHIFT_CUTOFF_HOUR = 7; 

// ================= WEB APP ENTRY =================
function doGet() {
  // This is the correct function. It serves the HTML file directly.
  return HtmlService.createHtmlOutputFromFile('index') 
    .setTitle('Konecta Adherence Portal (KAP)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// ================= WEB APP APIs (UPDATED) =================

// === Web App API for Punching ===
function webPunch(action, targetUserName, adminTimestamp) { 
  try {
    const puncherEmail = Session.getActiveUser().getEmail().toLowerCase();
    const result = punch(action, targetUserName, puncherEmail, adminTimestamp); 
    return result;
  } catch (err) {
    return "Error: " + err.message;
  }
}

// === Web App API for Schedule Range ===
function webSubmitScheduleRange(userEmail, userName, startDateStr, endDateStr, startTime, endTime, leaveType) {
  try {
    const puncherEmail = Session.getActiveUser().getEmail().toLowerCase();
    const result = submitScheduleRange(puncherEmail, userEmail, userName, startDateStr, endDateStr, startTime, endTime, leaveType);
    return result;
  } catch (err) {
    return "Error: " + err.message;
  }
}

// === Web App APIs for Leave Requests ===
function webSubmitLeaveRequest(requestObject, targetUserEmail) { // Now accepts optional target user
  try {
    const submitterEmail = Session.getActiveUser().getEmail().toLowerCase();
    return submitLeaveRequest(submitterEmail, requestObject, targetUserEmail);
  } catch (err) {
    return "Error: " + err.message;
  }
}

function webGetMyRequests_V2() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getMyRequests(userEmail); 
  } catch (err) {
    Logger.log("Error in webGetMyRequests_V2: " + err.message);
    throw new Error(err.message); 
  }
}

function webGetPendingRequests_V2(filter) {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getPendingRequests(userEmail, filter);
  } catch (err) {
    Logger.log("Error in webGetPendingRequests_V2: " + err.message);
    throw new Error(err.message); 
  }
}

function webApproveDenyRequest(requestID, newStatus) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return approveDenyRequest(adminEmail, requestID, newStatus);
  } catch (err) {
    return "Error: " + err.message;
  }
}

// === Web App API for History ===
function webGetAdherenceRange(userNames, startDateStr, endDateStr) {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getAdherenceRange(userEmail, userNames, startDateStr, endDateStr);
  } catch (err) {
    return { error: "Error: " + err.message };
  }
}

// === Web App API for My Schedule ===
function webGetMySchedule() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getMySchedule(userEmail);
  } catch (err) {
    return { error: "Error: " + err.message };
  }
}

// === Web App API for Admin Tools ===
function webAdjustLeaveBalance(userEmail, leaveType, amount, reason) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return adjustLeaveBalance(adminEmail, userEmail, leaveType, amount, reason);
  } catch (err) {
    return "Error: " + err.message;
  }
}

function webImportScheduleCSV(csvData) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return importScheduleCSV(adminEmail, csvData);
  } catch (err) {
    return "Error: " + err.message;
  }
}

// === Web App API for Dashboard ===
function webGetDashboardData(userEmails, date) { 
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getDashboardData(adminEmail, userEmails, date);
  } catch (err) {
    Logger.log("webGetDashboardData Error: " + err.message);
    throw new Error(err.message);
  }
}

// --- MODIFIED: "My Team" Functions ---
function webSaveMyTeam(userEmails) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return saveMyTeam(adminEmail, userEmails);
  } catch (err) {
    return "Error: " + err.message;
  }
}

function webGetMyTeam() {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return getMyTeam(adminEmail);
  } catch (err) {
    return "Error: " + err.message;
  }
}

// --- Web App API for Reporting Line ---
function webUpdateReportingLine(userEmail, newSupervisorEmail) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    return updateReportingLine(adminEmail, userEmail, newSupervisorEmail);
  } catch (err) {
    return "Error: " + err.message;
  }
}

// ==========================================================
// === NEW/REPLACED COACHING FUNCTIONS (START) ===
// ==========================================================

/**
 * (REPLACED FOR DYNAMIC TEMPLATES - PHASE 4)
 * Saves a new coaching session and its detailed scores from any template.
 */
function webSubmitCoaching(sessionObject) {
  try {
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    const sessionSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingSessions);
    const scoreSheet = getOrCreateSheet(ss, "CoachingScores_V2"); // Use the NEW scores sheet

    const coachEmail = Session.getActiveUser().getEmail().toLowerCase();
    const coachName = userData.emailToName[coachEmail] || coachEmail;

    // 1. Validation
    if (!sessionObject.agentEmail || !sessionObject.sessionDate || !sessionObject.templateID) {
      throw new Error("Agent, Session Date, and TemplateID are required.");
    }

    const agentName = userData.emailToName[sessionObject.agentEmail.toLowerCase()];
    if (!agentName) {
      throw new Error(`Could not find agent with email ${sessionObject.agentEmail}.`);
    }

    const sessionID = `CS-${new Date().getTime()}`; // Simple unique ID
    const sessionDate = new Date(sessionObject.sessionDate + 'T00:00:00');
    const followUpDate = sessionObject.followUpDate ? new Date(sessionObject.followUpDate + 'T00:00:00') : null;
    const followUpStatus = followUpDate ? "Pending" : "";

    // 2. Log the main session
    sessionSheet.appendRow([
      sessionID,
      sessionObject.agentEmail,
      agentName,
      coachEmail,
      coachName,
      sessionDate,
      sessionObject.weekNumber,
      sessionObject.overallScore, // Pass calculated score from client
      sessionObject.followUpComment,
      new Date(), // Timestamp of submission
      followUpDate || "", 
      followUpStatus,
      sessionObject.templateID // NEW: Save the TemplateID
    ]);

    // 3. Log the individual scores to V2 sheet
    const scoresToLog = [];
    if (sessionObject.scores && Array.isArray(sessionObject.scores)) {
      sessionObject.scores.forEach(score => {
        scoresToLog.push([
          sessionID,
          score.itemID,
          score.scoreValue,
          score.commentText,
          sessionObject.templateID
        ]);
      });
    }

    if (scoresToLog.length > 0) {
      scoreSheet.getRange(scoreSheet.getLastRow() + 1, 1, scoresToLog.length, 5).setValues(scoresToLog);
    }

    return `Coaching session for ${agentName} saved successfully.`;

  } catch (err) {
    Logger.log("webSubmitCoaching Error: " + err.message);
    return "Error: " + err.message;
  }
}

/**
 * (REPLACED FOR DYNAMIC TEMPLATES - PHASE 4)
 * Gets coaching history and joins scores from the V2 tables.
 */
function webGetCoachingHistory(filter) {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    const role = userData.emailToRole[userEmail] || 'agent';

    // --- 1. Get Template Criteria Definitions ---
    const criteriaSheet = getOrCreateSheet(ss, "CoachingTemplateCriteria");
    const criteriaData = criteriaSheet.getRange(2, 1, criteriaSheet.getLastRow() - 1, 7).getValues();
    const criteriaMap = {}; // Key: ItemID, Value: { category, criteriaText, ... }
    for (let i = 0; i < criteriaData.length; i++) {
      const row = criteriaData[i];
      criteriaMap[row[1]] = { // row[1] is ItemID
        templateID: row[0],
        itemID: row[1],
        category: row[2],
        criteriaText: row[3],
        inputType: row[4],
        weight: row[5]
      };
    }

    // --- 2. Get All Scores from V2 Sheet ---
    const scoreSheet = getOrCreateSheet(ss, "CoachingScores_V2");
    const allScoresData = scoreSheet.getRange(2, 1, scoreSheet.getLastRow() - 1, 5).getValues();
    const scoresMap = {}; // Key: SessionID, Value: [ { score details } ]
    for (let i = 0; i < allScoresData.length; i++) {
      const row = allScoresData[i];
      const sessionID = row[0];
      const itemID = row[1];
      const scoreValue = row[2];
      const commentText = row[3];

      if (!scoresMap[sessionID]) {
        scoresMap[sessionID] = [];
      }

      const criteriaDef = criteriaMap[itemID] || {};
      scoresMap[sessionID].push({
        itemID: itemID,
        scoreValue: scoreValue,
        commentText: commentText,
        category: criteriaDef.category || 'Unknown Category',
        criteriaText: criteriaDef.criteriaText || 'Unknown Criteria'
      });
    }

    // --- 3. Get All Session Headers ---
    const sessionSheet = getOrCreateSheet(ss, SHEET_NAMES.coachingSessions);
    const headers = sessionSheet.getRange(1, 1, 1, sessionSheet.getLastColumn()).getValues()[0];
    const allData = sessionSheet.getRange(2, 1, sessionSheet.getLastRow() - 1, sessionSheet.getLastColumn()).getValues();

    const allSessions = allData.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      obj['scores'] = scoresMap[obj['SessionID']] || [];
      return obj;
    });

    // --- 4. Filter Sessions for the current user ---
    const results = [];
    let myTeamEmails = new Set();
    if (role === 'admin' || role === 'superadmin') {
      const myTeamList = webGetAllSubordinateEmails(userEmail);
      myTeamList.forEach(email => myTeamEmails.add(email.toLowerCase()));
    }

    for (let i = allSessions.length - 1; i >= 0; i--) {
      const session = allSessions[i];
      if (!session || !session.AgentEmail) continue;

      const agentEmail = session.AgentEmail.toLowerCase();
      let canView = false;

      if (role === 'agent' && agentEmail === userEmail) canView = true;
      else if (role === 'admin' && myTeamEmails.has(agentEmail)) canView = true;
      else if (role === 'superadmin') canView = true;

      if (canView) {
        results.push({
          sessionID: session.SessionID,
          agentName: session.AgentName,
          coachName: session.CoachName,
          sessionDate: convertDateToString(new Date(session.SessionDate)),
          weekNumber: session.WeekNumber,
          overallScore: session.OverallScore,
          followUpComment: session.FollowUpComment,
          followUpDate: convertDateToString(new Date(session.FollowUpDate)),
          followUpStatus: session.FollowUpStatus,
          templateID: session.TemplateID,
          scores: session.scores
        });
      }
    }
    return results;

  } catch (err) {
    Logger.log("webGetCoachingHistory Error: " + err.message);
    return { error: err.message };
  }
}
// ==========================================================
// === NEW/REPLACED COACHING FUNCTIONS (END) ===
// ==========================================================


// === NEW: Web App API for Manager Hierarchy ===
function webGetManagerHierarchy() {
  try {
    const managerEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    const userData = getUserDataFromDb(dbSheet);
    
    const managerRole = userData.emailToRole[managerEmail] || 'agent';
    if (managerRole === 'agent') {
      return { error: "Permission denied. Only managers can view the hierarchy." };
    }
    
    // --- Step 1: Build the direct reporting map (Supervisor -> [Subordinates]) ---
    const reportsMap = {};
    const userEmailMap = {}; // Map email -> {name, role}

    userData.userList.forEach(user => {
      userEmailMap[user.email] = { name: user.name, role: user.role };
      const supervisorEmail = user.supervisor;
      
      if (supervisorEmail) {
        if (!reportsMap[supervisorEmail]) {
          reportsMap[supervisorEmail] = [];
        }
        reportsMap[supervisorEmail].push(user.email);
      }
    });

    // --- Step 2: Recursive function to build the tree (Hierarchy) ---
    // MODIFIED: Added `visited` Set to track users in the current path.
    function buildHierarchy(currentEmail, depth = 0, visited = new Set()) {
      const user = userEmailMap[currentEmail];
      
      // If the email doesn't map to a user, it's likely a blank entry in the DB, so return null
      if (!user) return null; 
      
      // CRITICAL CHECK: Detect circular reference
      if (visited.has(currentEmail)) {
        Logger.log(`Circular reference detected at user: ${currentEmail}`);
        return {
          email: currentEmail,
          name: user.name,
          role: user.role,
          subordinates: [],
          circularError: true
        };
      }
      
      // Add current user to visited set for this path
      const newVisited = new Set(visited).add(currentEmail);


      const subordinates = reportsMap[currentEmail] || [];
      
      // Separate managers/admins from agents
      const adminSubordinates = subordinates
        .filter(email => userData.emailToRole[email] === 'admin' || userData.emailToRole[email] === 'superadmin')
        .map(email => buildHierarchy(email, depth + 1, newVisited))
        .filter(s => s !== null); // Build sub-teams for managers

      const agentSubordinates = subordinates
        .filter(email => userData.emailToRole[email] === 'agent')
        .map(email => ({
          email: email,
          name: userEmailMap[email].name,
          role: userEmailMap[email].role,
          subordinates: [] // Agents have no subordinates
        }));
        
      // Combine and sort: Managers first, then Agents, then alphabetically
      const combinedSubordinates = [...adminSubordinates, ...agentSubordinates];
      
      combinedSubordinates.sort((a, b) => {
          // Sort by role (manager/admin first)
          const aIsManager = a.role !== 'agent';
          const bIsManager = b.role !== 'agent';
          
          if (aIsManager && !bIsManager) return -1;
          if (!aIsManager && bIsManager) return 1;
          
          // Then sort by name
          return a.name.localeCompare(b.name);
      });


      return {
        email: currentEmail,
        name: user.name,
        role: user.role,
        subordinates: combinedSubordinates,
        depth: depth
      };
    }

    // Start building the hierarchy from the manager's email
    const hierarchy = buildHierarchy(managerEmail);
    
    // Check if the root node returned a circular error
    if (hierarchy && hierarchy.circularError) {
        throw new Error("Critical Error: Circular reporting line detected at the top level.");
    }

    return hierarchy;

  } catch (err) {
    Logger.log("webGetManagerHierarchy Error: " + err.message);
    throw new Error(err.message);
  }
}


// === NEW: Web App API to get all reports (flat list) ===
function webGetAllSubordinateEmails(managerEmail) {
    try {
        const ss = getSpreadsheet();
        const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
        const userData = getUserDataFromDb(dbSheet);

        const managerRole = userData.emailToRole[managerEmail] || 'agent';
        if (managerRole === 'agent') {
            throw new Error("Permission denied.");
        }

        // --- Build the direct reporting map ---
        const reportsMap = {};
        userData.userList.forEach(user => {
            const supervisorEmail = user.supervisor;
            if (supervisorEmail) {
                if (!reportsMap[supervisorEmail]) {
                    reportsMap[supervisorEmail] = [];
                }
                reportsMap[supervisorEmail].push(user.email);
            }
        });

        const allSubordinates = new Set();
        const queue = [managerEmail];

        // Use a set to track users we've already processed (including the manager him/herself)
        const processed = new Set();
        while (queue.length > 0) {
            const currentEmail = queue.shift();
            // Check for processing loop (shouldn't happen in BFS, but safe check)
            if (processed.has(currentEmail)) continue;
            processed.add(currentEmail);

            const directReports = reportsMap[currentEmail] || [];

            directReports.forEach(reportEmail => {
                if (!allSubordinates.has(reportEmail)) {
                    allSubordinates.add(reportEmail);
                    // If the report is a manager, add them to the queue to find their reports
                    if (userData.emailToRole[reportEmail] !== 'agent') {
                        // *** THIS WAS THE FIX ***
                        queue.push(reportEmail); 
                        // ************************
                    }
                }

            });
        }

        // Return all subordinates *plus* the manager
        allSubordinates.add(managerEmail);
        return Array.from(allSubordinates);

    } catch (err) {
        Logger.log("webGetAllSubordinateEmails Error: " + err.message);
        return [];
    }
}
// --- END OF WEB APP API SECTION ---


// Get user info for front-end display
function getUserInfo() { 
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const timeZone = Session.getScriptTimeZone(); 
    const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
    
    const userData = getUserDataFromDb(dbSheet); 
    const userName = userData.emailToName[userEmail] || ""; 
    const role = userData.emailToRole[userEmail] || 'agent'; 

    let allUsers = []; 
    let allAdmins = []; 
    
    if (role === 'admin' || role === 'superadmin') {
      allUsers = userData.userList; 
    }
    
    allAdmins = userData.userList.filter(u => u.role === 'admin' || u.role === 'superadmin');
    
    const myBalances = userData.emailToBalances[userEmail] || { annual: 0, sick: 0, casual: 0 };

    return {
      name: userName, 
      email: userEmail,
      role: role,
      allUsers: allUsers,
      allAdmins: allAdmins,
      myBalances: myBalances 
    };
  } catch (e) {
    throw new Error("Failed in getUserInfo: " + e.message); 
  }
}

// ================= PUNCH MAIN FUNCTION =================
function punch(action, targetUserName, puncherEmail, adminTimestamp) { 
  const ss = getSpreadsheet();
  const adherenceSheet = getOrCreateSheet(ss, SHEET_NAMES.adherence);
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  const otherCodesSheet = getOrCreateSheet(ss, SHEET_NAMES.otherCodes);
  const timeZone = Session.getScriptTimeZone(); 

  // === 1. GET ALL USER DATA ===
  const userData = getUserDataFromDb(dbSheet); 
  
  // === 2. IDENTIFY PUNCHER & TARGET ===
  const puncherRole = userData.emailToRole[puncherEmail] || 'agent'; 
  const puncherIsAdmin = (puncherRole === 'admin' || puncherRole === 'superadmin');
  
  const userName = targetUserName; 
  const userEmail = userData.nameToEmail[userName]; 

  if (!puncherIsAdmin && puncherEmail !== userEmail) { 
    throw new Error("Permission denied. You can only submit punches for yourself.");
  }
  const isAdmin = puncherIsAdmin; 
  
  // === 3. VALIDATE TARGET USER ===
  if (!userEmail) { 
     throw new Error(`User "${userName}" not found in Data Base.`); 
  }
  if (!userName && !puncherIsAdmin) { 
    throw new Error("Your email is not registered in the Data Base sheet. Contact your supervisor.");
  }
  
  const nowTimestamp = adminTimestamp ? new Date(adminTimestamp) : new Date();
  const shiftDate = getShiftDate(new Date(nowTimestamp), SHIFT_CUTOFF_HOUR);
  const formattedDate = Utilities.formatDate(shiftDate, timeZone, "MM/dd/yyyy");

  // === 4. HANDLE "OTHER CODES" ===
  const otherCodeActions = ["Meeting", "Personal", "Coaching"];
  for (const code of otherCodeActions) {
    if (action.startsWith(code)) {
      const resultMsg = logOtherCode(
        otherCodesSheet, userName, action, nowTimestamp, 
        isAdmin && (puncherEmail !== userEmail || adminTimestamp) ? puncherEmail : null 
      );
      logsSheet.appendRow([new Date(), userName, userEmail, action, nowTimestamp]); 
      return resultMsg;
    }
  }

  // === 5. PROCEED WITH ADHERENCE PUNCH ===
  const scheduleData = scheduleSheet.getDataRange().getValues();
  let shiftStartStr = "", shiftEndStr = "", leaveType = "";
  for (let i = 1; i < scheduleData.length; i++) {
    const [schName, schDate, schStart, schEnd, schLeave, schEmail] = scheduleData[i];
    const dateObj = new Date(schDate);
    if (isNaN(dateObj.getTime())) continue;
    const dateStr = Utilities.formatDate(dateObj, timeZone, "MM/dd/yyyy");
    
    if (schEmail && schEmail.toLowerCase() === userEmail && dateStr === formattedDate) { 
      
      if (schStart instanceof Date) {
        shiftStartStr = Utilities.formatDate(schStart, timeZone, "HH:mm:ss");
      } else {
        shiftStartStr = (schStart || "").toString(); 
      }
      
      if (schEnd instanceof Date) {
        shiftEndStr = Utilities.formatDate(schEnd, timeZone, "HH:mm:ss");
      } else {
        shiftEndStr = (schEnd || "").toString(); 
       }
          
      leaveType = (schLeave || "").toString().trim();
      break;
    }
  }

  if (!shiftStartStr && !leaveType) {
    throw new Error(`User ${userName} is not scheduled today (${formattedDate}), or their schedule has not been uploaded. Please contact your manager.`); 
  }

  const row = findOrCreateRow(adherenceSheet, userName, shiftDate, formattedDate); 
  const leaveTypeLower = leaveType.toLowerCase();
  
  if (leaveType && leaveTypeLower !== "present" && leaveTypeLower !== "") {
    adherenceSheet.getRange(row, 14).setValue(leaveType);
    if (leaveTypeLower === "absent") {
      adherenceSheet.getRange(row, 20).setValue("Yes");
    }
    return `${userName}: Leave type "${leaveType}" recorded. No further punches needed.`; 
  } else {
    adherenceSheet.getRange(row, 14).setValue("Present");
  }

  const columns = {
    "Login": 3, "First Break In": 4, "First Break Out": 5, "Lunch In": 6, 
    "Lunch Out": 7, "Last Break In": 8, "Last Break Out": 9, "Logout": 10
  };
  const col = columns[action];
  if (!col) throw new Error("Invalid action: " + action);

  const currentPunches = adherenceSheet.getRange(row, 3, 1, 8).getValues()[0];
  const punches = {
    login: currentPunches[0], firstBreakIn: currentPunches[1], firstBreakOut: currentPunches[2],
    lunchIn: currentPunches[3], lunchOut: currentPunches[4], lastBreakIn: currentPunches[5],
    lastBreakOut: currentPunches[6], logout: currentPunches[7]
  };

  // --- VALIDATION ---
  if (!isAdmin) {
    if (action !== "Login" && !punches.login) {
      throw new Error("You must 'Login' first.");
    }
    const sequentialErrors = {
      "First Break Out": { required: punches.firstBreakIn, msg: "You must punch 'First Break In' first." },
      "Lunch Out":       { required: punches.lunchIn,     msg: "You must punch 'Lunch In' first." },
      "Last Break Out":  { required: punches.lastBreakIn,   msg: "You must punch 'Last Break In' first." }
    };
    if (sequentialErrors[action] && !sequentialErrors[action].required) {
      throw new Error(sequentialErrors[action].msg);
    }
    const existingValue = adherenceSheet.getRange(row, col).getValue();
    if (existingValue) {
      throw new Error(`"${action}" already punched today.`);
    }
  }

  // --- ADMIN AUDIT ---
  if (isAdmin && (puncherEmail !== userEmail || adminTimestamp)) { 
    adherenceSheet.getRange(row, 15).setValue("Yes");
    adherenceSheet.getRange(row, 21).setValue(puncherEmail);
  }

  // === SAVE PUNCH ===
  adherenceSheet.getRange(row, col).setValue(nowTimestamp);
  logsSheet.appendRow([new Date(), userName, userEmail, action, nowTimestamp]); 
  
  const actionKey = Object.keys(columns).find(key => columns[key] === col);
  if (actionKey) {
    punches[actionKey.replace(/\s+/g, '').replace('1st', 'first').replace('Out', 'Out').replace('In', 'In').toLowerCase()] = nowTimestamp;
  }

  // === DATE-AWARE SHIFT METRICS ===
  const shiftStartDate = createDateTime(shiftDate, shiftStartStr);
  let shiftEndDate = createDateTime(shiftDate, shiftEndStr);
  
  if (!shiftStartDate) {
    throw new Error(`Could not parse Shift Start Time ("${shiftStartStr}"). Please check the schedule.`);
  }

  if (shiftEndDate <= shiftStartDate) {
    shiftEndDate.setDate(shiftEndDate.getDate() + 1);
  }
  
  if (action === "Login" || punches.login) {
    const loginTime = (action === "Login") ? nowTimestamp : punches.login;
    const diff = timeDiffInSeconds(shiftStartDate, loginTime);
    
    const timeFormat = "HH:mm";
    const scheduledTime = Utilities.formatDate(shiftStartDate, timeZone, timeFormat);
    const punchTime = Utilities.formatDate(loginTime, timeZone, timeFormat);

    if (action === "Login") {
      if (diff > (4 * 60 * 60)) {
        throw new Error(`Login is over 4 hours late (Shift: ${scheduledTime}, Punch: ${punchTime}). Please check your shift schedule or contact your manager.`);
      }
      if (diff < -(2 * 60 * 60)) {
        throw new Error(`Login is over 2 hours early (Shift: ${scheduledTime}, Punch: ${punchTime}). Please check your shift schedule or contact your manager.`);
      }
    }
    adherenceSheet.getRange(row, 11).setValue(diff > 0 ? diff : 0);
  }

  if (action === "Logout" || punches.logout) {
    if (!shiftEndDate) {
      throw new Error(`Could not parse Shift End Time ("${shiftEndStr}"). Please check the schedule.`);
    }
    const logoutTime = (action === "Logout") ? nowTimestamp : punches.logout;
    const diff = timeDiffInSeconds(shiftEndDate, logoutTime); 
    if (diff > 0) {
      adherenceSheet.getRange(row, 12).setValue(diff);
      adherenceSheet.getRange(row, 13).setValue(0);
    } else {
      adherenceSheet.getRange(row, 12).setValue(0);
      adherenceSheet.getRange(row, 13).setValue(Math.abs(diff));
    }
  }

  // === BREAK EXCEED CALCULATIONS ===
  let exceedMsg = "No";
  let duration = 0;
  let diff = 0;
  try {
    duration = timeDiffInSeconds(punches.firstBreakIn, punches.firstBreakOut);
    diff = duration - PLANNED_BREAK_SECONDS;
    if (diff > 0 && duration > 0) exceedMsg = diff; else exceedMsg = "No";
    adherenceSheet.getRange(row, 17).setValue(exceedMsg);

    duration = timeDiffInSeconds(punches.lunchIn, punches.lunchOut);
    diff = duration - PLANNED_LUNCH_SECONDS;
    if (diff > 0 && duration > 0) exceedMsg = diff; else exceedMsg = "No";
    adherenceSheet.getRange(row, 18).setValue(exceedMsg);

    duration = timeDiffInSeconds(punches.lastBreakIn, punches.lastBreakOut);
    diff = duration - PLANNED_BREAK_SECONDS;
    if (diff > 0 && duration > 0) exceedMsg = diff; else exceedMsg = "No";
    adherenceSheet.getRange(row, 19).setValue(exceedMsg);
    
  } catch (e) {
    logsSheet.appendRow([new Date(), userName, userEmail, "Break Exceed Error", e.message]); 
  }

  return `${userName}: ${action} recorded at ${Utilities.formatDate(nowTimestamp, timeZone, "HH:mm:ss")}`; 
}


// ================= SCHEDULE RANGE SUBMIT FUNCTION =================
function submitScheduleRange(puncherEmail, userEmail, userName, startDateStr, endDateStr, startTime, endTime, leaveType) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const puncherRole = userData.emailToRole[puncherEmail] || 'agent';
  const timeZone = Session.getScriptTimeZone();
  
  if (puncherRole !== 'admin' && puncherRole !== 'superadmin') {
    throw new Error("Permission denied. Only admins can submit schedules.");
  }
  
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  
  const userScheduleMap = {};
  for (let i = 1; i < scheduleData.length; i++) {
    const rowEmail = scheduleData[i][5];
    const rowDateRaw = scheduleData[i][1];
    if (rowEmail && rowDateRaw && rowEmail.toLowerCase() === userEmail) {
      const rowDate = new Date(rowDateRaw);
      const rowDateStr = Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy");
      userScheduleMap[rowDateStr] = i + 1; 
    }
  }
  
  const startDate = new Date(startDateStr);
  const endDate = new Date(endDateStr);
  
  let currentDate = new Date(startDate);
  let daysProcessed = 0;
  let daysUpdated = 0;
  let daysCreated = 0;
  
  const oneDayInMs = 24 * 60 * 60 * 1000;
  
  currentDate = new Date(currentDate.valueOf() + currentDate.getTimezoneOffset() * 60000);
  const finalDate = new Date(endDate.valueOf() + endDate.getTimezoneOffset() * 60000);
  
  while (currentDate <= finalDate) {
    const currentDateStr = Utilities.formatDate(currentDate, timeZone, "MM/dd/yyyy");
    
    const result = updateOrAddSingleSchedule(
      scheduleSheet, userScheduleMap, logsSheet,
      userEmail, userName, 
      currentDate, currentDateStr, 
      startTime, endTime, leaveType, puncherEmail
    );
    
    if (result === "UPDATED") daysUpdated++;
    if (result === "CREATED") daysCreated++;
    
    daysProcessed++;
    currentDate.setTime(currentDate.getTime() + oneDayInMs);
  }
  
  if (daysProcessed === 0) {
    throw new Error("No dates were processed. Check date range.");
  }
  
  return `Schedule submission complete for ${userName}. Days processed: ${daysProcessed} (Updated: ${daysUpdated}, Created: ${daysCreated}).`;
}

// (Helper for above)
function updateOrAddSingleSchedule(scheduleSheet, userScheduleMap, logsSheet, userEmail, userName, targetDate, targetDateStr, startTime, endTime, leaveType, puncherEmail) {
  
  const existingRow = userScheduleMap[targetDateStr];
  
  let startTimeObj = startTime ? new Date(`1899-12-30T${startTime}`) : "";
  let endTimeObj = endTime ? new Date(`1899-12-30T${endTime}`) : "";
  
  if (existingRow) {
    scheduleSheet.getRange(existingRow, 1, 1, 5).setValues([[
      userName, targetDate, startTimeObj, endTimeObj, leaveType
    ]]);
    logsSheet.appendRow([new Date(), userName, puncherEmail, "Schedule UPDATE", `Set to: ${leaveType}, ${startTime}-${endTime}`]);
    return "UPDATED";
  } else {
    scheduleSheet.appendRow([
      userName, targetDate, startTimeObj, endTimeObj, leaveType, userEmail
    ]);
    logsSheet.appendRow([new Date(), userName, puncherEmail, "Schedule CREATE", `Set to: ${leaveType}, ${startTime}-${endTime}`]);
    return "CREATED";
  }
}


// ================= HELPER FUNCTIONS =================

function getShiftDate(dateObj, cutoffHour) {
  if (dateObj.getHours() < cutoffHour) {
    dateObj.setDate(dateObj.getDate() - 1);
  }
  return dateObj;
}

function createDateTime(dateObj, timeStr) {
  if (!timeStr) return null;
  const parts = timeStr.split(':');
  if (parts.length < 2) return null;
  
  const [hours, minutes, seconds] = parts.map(Number);
  if (isNaN(hours) || isNaN(minutes)) return null; 

  const newDate = new Date(dateObj);
  newDate.setHours(hours, minutes, seconds || 0, 0);
  return newDate;
}

// UPDATED to read new SupervisorEmail column
function getUserDataFromDb(dbSheet) { 
  const dbData = dbSheet.getDataRange().getValues();
  const nameToEmail = {};
  const emailToName = {};
  const emailToRole = {}; 
  const emailToBalances = {}; 
  const emailToRow = {}; 
  const emailToSupervisor = {}; 
  const userList = []; 
  
  for (let i = 1; i < dbData.length; i++) {
    let [name, email, role, annual, sick, casual, supervisor] = dbData[i]; 
    
    if (name && email) {
      const cleanName = name.toString().trim();
      const cleanEmail = email.toString().trim().toLowerCase();
      const userRole = (role || 'agent').toString().trim().toLowerCase(); 
      const supervisorEmail = (supervisor || "").toString().trim().toLowerCase(); 
      
      nameToEmail[cleanName] = cleanEmail; 
      emailToName[cleanEmail] = cleanName;
      emailToRole[cleanEmail] = userRole; 
      emailToRow[cleanEmail] = i + 1; 
      emailToSupervisor[cleanEmail] = supervisorEmail; 
      
      emailToBalances[cleanEmail] = {
        annual: parseFloat(annual) || 0,
        sick: parseFloat(sick) || 0,
        casual: parseFloat(casual) || 0
      };
      
      userList.push({ 
        name: cleanName, 
        email: cleanEmail, 
        role: userRole,
        balances: emailToBalances[cleanEmail],
        supervisor: supervisorEmail 
      }); 
    }
  }
  userList.sort((a, b) => a.name.localeCompare(b.name)); 
  return { nameToEmail, emailToName, emailToRole, emailToBalances, emailToRow, emailToSupervisor, userList }; 
}

// (No Change)
function logOtherCode(sheet, userName, action, nowTimestamp, adminEmail) { 
  const [code, type] = action.split(" ");
  const data = sheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  
  const shiftDate = getShiftDate(new Date(nowTimestamp), SHIFT_CUTOFF_HOUR);
  const dateStr = Utilities.formatDate(shiftDate, timeZone, "MM/dd/yyyy");
  
  if (type === "In") {
    if (adminEmail) { 
       sheet.appendRow([nowTimestamp, userName, code, nowTimestamp, "", "", adminEmail]);
       return `${userName}: ${action} recorded at ${Utilities.formatDate(nowTimestamp, timeZone, "HH:mm:ss")}.`;
    }
    
    for (let i = data.length - 1; i > 0; i--) {
      const [rowDateRaw, rowName, rowCode, rowIn, rowOut] = data[i];
      if (!rowDateRaw || !rowName) continue;
      
      const rowShiftDate = getShiftDate(new Date(rowDateRaw), SHIFT_CUTOFF_HOUR);
      const rowDateStr = Utilities.formatDate(rowShiftDate, timeZone, "MM/dd/yyyy");
      
      if (rowName === userName && rowDateStr === dateStr && rowCode === code && rowIn && !rowOut) { 
        throw new Error(`You must punch "${code} Out" before punching "In" again.`);
      }
    }
    sheet.appendRow([nowTimestamp, userName, code, nowTimestamp, "", "", adminEmail || ""]); 
  } else if (type === "Out") {
    let matchingInPunch = null;
    let matchingInRow = -1;
    
    for (let i = data.length - 1; i > 0; i--) {
      const [rowDateRaw, rowName, rowCode, rowIn, rowOut] = data[i];
      if (!rowDateRaw || !rowName || !rowIn) continue;
      
      const rowShiftDate = getShiftDate(new Date(rowDateRaw), SHIFT_CUTOFF_HOUR);
      const rowDateStr = Utilities.formatDate(rowShiftDate, timeZone, "MM/dd/yyyy");
      
      if (rowName === userName && rowDateStr === dateStr && rowCode === code && rowIn && !rowOut) { 
        matchingInPunch = rowIn; // This is a Date object
        matchingInRow = i + 1;
        break;
      }
    }
    
    if (matchingInPunch) {
      const duration = timeDiffInSeconds(matchingInPunch, nowTimestamp);
      sheet.getRange(matchingInRow, 5).setValue(nowTimestamp);
      sheet.getRange(matchingInRow, 6).setValue(duration);
      if (adminEmail) {
        sheet.getRange(matchingInRow, 7).setValue(adminEmail);
      }
      return `${userName}: ${action} recorded. Duration: ${Math.round(duration/60)} mins.`;
    } else {
      if (adminEmail) { 
        sheet.appendRow([nowTimestamp, userName, code, "", nowTimestamp, 0, adminEmail]);
        return `${userName}: ${action} (Out) recorded without matching In.`;
      }
      throw new Error(`You must punch "${code} In" first.`);
    }
  }
  return `${userName}: ${action} recorded at ${Utilities.formatDate(nowTimestamp, timeZone, "HH:mm:ss")}.`; 
}

// (No Change)
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// (No Change)
function findOrCreateRow(sheet, userName, shiftDate, formattedDate) { 
  const data = sheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  let row = -1;
  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][0]);
    const rowUser = data[i][1]; 
    if (
      rowUser && 
      rowUser.toString().toLowerCase() === userName.toLowerCase() && 
      Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy") === formattedDate
    ) {
      row = i + 1;
      break;
    }
  }

  if (row === -1) {
    row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1).setValue(shiftDate);
    sheet.getRange(row, 2).setValue(userName); 
  }
  return row;
}

// *** UPDATED getOrCreateSheet ***
function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SHEET_NAMES.database) {
      // NEW: Added SupervisorEmail
      sheet.getRange("A1:G1").setValues([["User Name", "Email", "Role", "Annual Balance", "Sick Balance", "Casual Balance", "SupervisorEmail"]]);
    } else if (name === SHEET_NAMES.schedule) {
      sheet.getRange("A1:F1").setValues([["Name", "Date", "Shift Start Time", "Shift End Time", "Leave Type", "agent email"]]);
      sheet.getRange("C:D").setNumberFormat("hh:mm");
    } else if (name === SHEET_NAMES.adherence) {
      sheet.getRange("A1:U1").setValues([[ 
        "Date", "User Name", "Login", "First Break In", "First Break Out", "Lunch In", "Lunch Out", 
        "Last Break In", "Last Break Out", "Logout", "Tardy (Seconds)", "Overtime (Seconds)", "Early Leave (Seconds)",
        "Leave Type", "Admin Audit", "—", "1st Break Exceed", "Lunch Exceed", "Last Break Exceed", "Absent", "Admin Code"
      ]]);
    } else if (name === SHEET_NAMES.logs) {
      sheet.getRange("A1:E1").setValues([["Timestamp", "User Name", "Email", "Action", "Time"]]);
    } else if (name === SHEET_NAMES.otherCodes) { 
      sheet.getRange("A1:G1").setValues([["Date", "User Name", "Code", "Time In", "Time Out", "Duration (Seconds)", "Admin Audit (Email)"]]);
    } else if (name === SHEET_NAMES.leaveRequests) { 
      sheet.getRange("A1:L1").setValues([
        ["RequestID", "Status", "RequestedByEmail", "RequestedByName", 
         "LeaveType", "StartDate", "EndDate", "TotalDays", "Reason", 
         "ActionDate", "ActionBy", "SupervisorEmail"] 
      ]);
      sheet.getRange("F:G").setNumberFormat("mm/dd/yyyy");
          // ... inside getOrCreateSheet function
      sheet.getRange("J:J").setNumberFormat("mm/dd/yyyy");
    } else if (name === SHEET_NAMES.coaching_OLD) { // <-- Renamed this
      sheet.getRange("A1:R1").setValues([[
        "SessionID", "Status", "AgentEmail", "AgentName", "CoachEmail", "CoachName",
        "WeekNumber", "SessionDate", "AreaOfConcern", "RootCause", "CoachingTopic",
        "ActionsTaken", "AgentFeedback", "FollowUpPlan", "NextReviewDate",
        "QA_ID", "QA_Score", "LoggedByAdmin"
      ]]);
      sheet.getRange("H:H").setNumberFormat("mm/dd/yyyy");
      sheet.getRange("O:O").setNumberFormat("mm/dd/yyyy");
    } 
    // +++ ADDED NEW SHEET DEFINITIONS +++
    else if (name === SHEET_NAMES.coachingSessions) { 
      sheet.getRange("A1:M1").setValues([[ // *** CHANGED from L1 to M1 ***
        "SessionID", "AgentEmail", "AgentName", "CoachEmail", "CoachName",
        "SessionDate", "WeekNumber", "OverallScore", "FollowUpComment", "SubmissionTimestamp",
        "FollowUpDate", "FollowUpStatus",
        "TemplateID" // *** NEW: Track which template was used ***
      ]]);
      sheet.getRange("F:F").setNumberFormat("mm/dd/yyyy");
      sheet.getRange("J:J").setNumberFormat("mm/dd/yyyy hh:mm:ss");
      sheet.getRange("K:K").setNumberFormat("mm/dd/yyyy"); // *** ADDED FORMAT for K ***
    
    // *** DEPRECATED OLD SHEET, REPLACED WITH V2 ***
    } else if (name === SHEET_NAMES.coachingScores) { 
      sheet.getRange("A1:E1").setValues([[
        "SessionID", "Category", "Criteria", "Score", "Comment"
      ]]);
    
    // *** NEW SHEETS FOR DYNAMIC TEMPLATES ***
    } else if (name === "CoachingTemplates") { // NEW SHEET
      sheet.getRange("A1:E1").setValues([[
        "TemplateID", "TemplateName", "CreatedBy", "DateCreated", "IsArchived"
      ]]);
    } else if (name === "CoachingTemplateCriteria") { // NEW SHEET
      sheet.getRange("A1:G1").setValues([[
        "TemplateID", "ItemID", "Category", "CriteriaText", "InputType", "Weight", "ItemOrder"
        // InputType: 'score_0-1' (0, 0.5, 1), 'yes_no', 'text_comment', 'number'
      ]]);
    } else if (name === "CoachingScores_V2") { // NEW SHEET
       sheet.getRange("A1:E1").setValues([[
        "SessionID", "ItemID", "ScoreValue", "CommentText", "TemplateID"
      ]]);
    }
    // +++ END OF NEW SHEET DEFINITIONS +++
  }
  if (name === SHEET_NAMES.adherence) {
    sheet.getRange("C:J").setNumberFormat("hh:mm:ss");
  }
  if (name === SHEET_NAMES.otherCodes) {
    sheet.getRange("D:E").setNumberFormat("hh:mm:ss");
  }
  if (name === SHEET_NAMES.schedule) {
    sheet.getRange("C:D").setNumberFormat("hh:mm");
  }
  return sheet;
}

// (No Change)
function timeDiffInSeconds(start, end) {
  if (!start || !end || !(start instanceof Date) || !(end instanceof Date)) {
    return 0;
  }
  return Math.round((end.getTime() - start.getTime()) / 1000);
}


// ================= DAILY AUTO-LOG FUNCTION =================
function dailyLeaveSweeper() {
  const ss = getSpreadsheet();
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const adherenceSheet = getOrCreateSheet(ss, SHEET_NAMES.adherence);
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  const timeZone = Session.getScriptTimeZone();

  // 1. Define the 7-day lookback period
  const lookbackDays = 7;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const endDate = new Date(today); // Today
  endDate.setDate(endDate.getDate() - 1); // End date is yesterday
  
  const startDate = new Date(endDate);
  startDate.setDate(startDate.getDate() - (lookbackDays - 1)); // Start date is 7 days ago
  
  const startDateStr = Utilities.formatDate(startDate, timeZone, "MM/dd/yyyy");
  const endDateStr = Utilities.formatDate(endDate, timeZone, "MM/dd/yyyy");

  Logger.log(`Starting dailyLeaveSweeper for date range: ${startDateStr} to ${endDateStr}`);

  // 2. Get all Adherence rows for the past 7 days and create a lookup Set
  // The key will be a combo: "username:mm/dd/yyyy"
  const allAdherence = adherenceSheet.getDataRange().getValues();
  const adherenceLookup = new Set();
  for (let i = 1; i < allAdherence.length; i++) {
    try {
      const rowDate = new Date(allAdherence[i][0]);
      if (rowDate >= startDate && rowDate <= endDate) {
        const rowDateStr = Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy");
        const userName = allAdherence[i][1].toString().trim().toLowerCase();
        adherenceLookup.add(`${userName}:${rowDateStr}`);
      }
    } catch (e) {
      Logger.log(`Skipping adherence row ${i+1}: ${e.message}`);
    }
  }
  Logger.log(`Found ${adherenceLookup.size} existing adherence records in the date range.`);

  // 3. Get all Schedules and loop through them
  const allSchedules = scheduleSheet.getDataRange().getValues();
  let missedLogs = 0;

  for (let i = 1; i < allSchedules.length; i++) {
    try {
      const [schName, schDate, schStart, schEnd, schLeave, schEmail] = allSchedules[i];
      const leaveType = (schLeave || "").toString().trim();
      
      // Skip if "Present", no leave type, or no name/email
      if (leaveType === "" || leaveType.toLowerCase() === "present" || !schName || !schEmail) {
        continue;
      }

      const schDateObj = new Date(schDate);
      
      // Check if the schedule date is within our 7-day lookback period
      if (schDateObj >= startDate && schDateObj <= endDate) {
        const schDateStr = Utilities.formatDate(schDateObj, timeZone, "MM/dd/yyyy");
        const userName = schName.toString().trim();
        const userNameLower = userName.toLowerCase();
        
        const lookupKey = `${userNameLower}:${schDateStr}`;

        // 4. Check if this user is *already* in the Adherence sheet
        if (adherenceLookup.has(lookupKey)) {
          continue; // We found them, so skip
        }

        // 5. We found a missed user! Create their row and log their leave.
        Logger.log(`Found missed user: ${userName} for ${schDateStr}. Logging leave: ${leaveType}`);
        
        // We use findOrCreateRow to be safe, but it should just create it
        const row = findOrCreateRow(adherenceSheet, userName, schDateObj, schDateStr);
        
        adherenceSheet.getRange(row, 14).setValue(leaveType);
        
        if (leaveType.toLowerCase() === "absent") {
          adherenceSheet.getRange(row, 20).setValue("Yes");
        }
        
        logsSheet.appendRow([new Date(), userName, schEmail, "Auto-Log Leave", leaveType]);
        missedLogs++;
        
         // Add to lookup so we don't process them again if they have duplicate schedules
        adherenceLookup.add(lookupKey); 
      }
    } catch (e) {
      Logger.log(`Skipping schedule row ${i+1}: ${e.message}`);
    }
  }
  
  Logger.log(`dailyLeaveSweeper finished. Logged ${missedLogs} missed users.`);
}

// ================= LEAVE REQUEST FUNCTIONS =================

// (Helper - No Change)
function convertDateToString(dateObj) {
  if (dateObj instanceof Date && !isNaN(dateObj)) {
    return dateObj.toISOString(); // "2025-11-06T18:30:00.000Z"
  }
  return null; // Return null if it's not a valid date
}

// (No Change)
function getMyRequests(userEmail) {
  const ss = getSpreadsheet();
  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.leaveRequests);
  const allData = reqSheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const myRequests = [];
  
  for (let i = allData.length - 1; i > 0; i--) { 
    const row = allData[i];
    if (String(row[2] || "").trim().toLowerCase() === userEmail) {
      try { 
        const startDate = new Date(row[5]);
        const endDate = new Date(row[6]);
        const requestedDateNum = Number(row[0].split('_')[1]);

        if (isNaN(startDate.getTime()) || isNaN(endDate.getTime()) || isNaN(requestedDateNum)) {
          Logger.log(`Skipping Row ${i+1}. It contains invalid date data.`);
          continue; 
        }

        const supervisorEmail = row[11]; 
        myRequests.push({
          requestID: row[0],
          status: row[1],
          leaveType: row[4],
          startDate: convertDateToString(startDate),
          endDate: convertDateToString(endDate),
          totalDays: row[7],
          reason: row[8],
          requestedDate: convertDateToString(new Date(requestedDateNum)),
          supervisorName: userData.emailToName[supervisorEmail] || supervisorEmail 
        });
      } catch (e) {
        Logger.log(`CRITICAL ERROR processing row ${i+1} for getMyRequests. Error: ${e.message}`);
      }
    }
  }
  return myRequests;
}

// (No Change)
function getPendingRequests(adminEmail, filter) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  
  if (adminRole !== 'admin' && adminRole !== 'superadmin') {
    return []; 
  }
  
  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.leaveRequests);
  const allData = reqSheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  const pendingRequests = [];
  
  for (let i = allData.length - 1; i > 0; i--) {
    const row = allData[i];
    if (row[1] && row[1].toString().trim().toLowerCase() === 'pending') { 
      try { 
        const startDate = new Date(row[5]);
        const endDate = new Date(row[6]);
        const requestedDateNum = Number(row[0].split('_')[1]);

        if (isNaN(startDate.getTime()) || isNaN(endDate.getTime()) || isNaN(requestedDateNum)) {
          Logger.log(`Skipping Row ${i+1}. It contains invalid date data.`);
          continue; 
        }
        
        const supervisorEmail = row[11]; 
        if (adminRole !== 'superadmin' && filter === 'mine' && supervisorEmail !== adminEmail) {
          continue; 
        }
        
        const requesterEmail = row[2];
        const requesterBalance = userData.emailToBalances[requesterEmail];
        
        pendingRequests.push({
          requestID: row[0],
          status: row[1],
          requestedByName: row[3],
          leaveType: row[4],
          startDate: convertDateToString(startDate),
          endDate: convertDateToString(endDate),
          totalDays: row[7],
          reason: row[8],
          requestedDate: convertDateToString(new Date(requestedDateNum)),
          supervisorName: userData.emailToName[supervisorEmail] || supervisorEmail,
          requesterBalance: requesterBalance 
        });
      } catch (e) {
         Logger.log(`Failed to process row ${i+1} for getPendingRequests. Error: ${e.message}`);
      }
    }
  }
  return pendingRequests;
}

// UPDATED: Now automatically finds supervisor
function submitLeaveRequest(submitterEmail, request, targetUserEmail) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  
  const isSelfRequest = !targetUserEmail;
  const requestEmail = isSelfRequest ? submitterEmail : targetUserEmail;
  const requestName = userData.emailToName[requestEmail];
  
  if (!requestName) {
    throw new Error(`Could not find user ${requestEmail} in the Data Base.`);
  }
  
  // --- NEW: Auto-find supervisor ---
  const supervisorEmail = userData.emailToSupervisor[requestEmail];
  if (!supervisorEmail) {
     throw new Error(`Cannot submit request. User ${requestName} does not have a supervisor assigned in the Data Base.`);
  }
  // --- END NEW ---
  
  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.leaveRequests);
  
  const startDate = new Date(request.startDate + 'T00:00:00');
  let endDate;
  if (request.endDate) {
    endDate = new Date(request.endDate + 'T00:00:00');
  } else {
    endDate = startDate; 
  }

  const ONE_DAY_MS = 24 * 60 * 60 * 1000;
  const totalDays = Math.round((endDate.getTime() - startDate.getTime()) / ONE_DAY_MS) + 1;

  if (totalDays < 1) {
    throw new Error("Invalid date range.");
  }
  
  const balanceKey = request.leaveType.toLowerCase(); 
  const userBalances = userData.emailToBalances[requestEmail];
  
  if (!userBalances || userBalances[balanceKey] === undefined) {
    throw new Error(`Could not find balance information for user ${requestName}. Check the Data Base sheet.`);
  }

  if (userBalances[balanceKey] < totalDays) {
    throw new Error(`Insufficient balance for ${requestName}. User has ${userBalances[balanceKey]} ${request.leaveType} days, but are requesting ${totalDays}.`);
  }

  const requestID = `req_${new Date().getTime()}`;
  
  reqSheet.appendRow([
    requestID,
    "Pending",
    requestEmail,
    requestName,
    request.leaveType,
    startDate, 
    endDate,   
    totalDays,
    request.reason,
    "", // ActionDate
    "", // ActionBy
    supervisorEmail // Use the auto-found supervisor
  ]);
  
  SpreadsheetApp.flush(); 
  
  return `Leave request submitted successfully for ${requestName}.`;
}

// (No Change)
function approveDenyRequest(adminEmail, requestID, newStatus) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  const adminName = userData.emailToName[adminEmail] || adminEmail;
  
  if (adminRole !== 'admin' && adminRole !== 'superadmin') {
    throw new Error("Permission denied. Only admins can take this action.");
  }
  
  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.leaveRequests);
  const allData = reqSheet.getDataRange().getValues();
  
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] === requestID) {
      const row = allData[i];
      const status = row[1];
      
      if (status !== 'Pending') {
        throw new Error(`This request has already been ${status.toLowerCase()}.`);
      }
      
      const supervisorEmail = row[11];
      if (adminRole !== 'superadmin' && supervisorEmail !== adminEmail) {
        throw new Error("Permission denied. This request is assigned to a different supervisor.");
      }
      
      const reqEmail = row[2];
      const reqName = row[3];
      const reqLeaveType = row[4];
      const reqStartDate = new Date(row[5]);
      const reqEndDate = new Date(row[6]);
      const reqStartDateStr = Utilities.formatDate(reqStartDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      const reqEndDateStr = Utilities.formatDate(reqEndDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      const totalDays = row[7];
      let scheduleResult = "";
      
      if (newStatus === 'Approved') {
      
        const balanceKey = reqLeaveType.toLowerCase();
        const balanceCol = { annual: 4, sick: 5, casual: 6 }[balanceKey];
        if (!balanceCol) {
          throw new Error(`Unknown leave type: ${reqLeaveType}. Balance cannot be deducted.`);
        }
        
        const userRow = userData.emailToRow[reqEmail];
        if (!userRow) {
          throw new Error(`Could not find user ${reqName} in Data Base to deduct balance.`);
        }
        
        const balanceRange = dbSheet.getRange(userRow, balanceCol);
        const currentBalance = parseFloat(balanceRange.getValue()) || 0;
        
        if (currentBalance < totalDays) {
          throw new Error(`Cannot approve. User only has ${currentBalance} ${reqLeaveType} days, but request is for ${totalDays}.`);
        }
        
        balanceRange.setValue(currentBalance - totalDays);
      
        scheduleResult = submitScheduleRange(
          adminEmail, reqEmail, reqName,
          reqStartDateStr, reqEndDateStr,
          "", "", reqLeaveType
        );
        
        reqSheet.getRange(i + 1, 2).setValue(newStatus); 
        reqSheet.getRange(i + 1, 10).setValue(new Date()); 
        reqSheet.getRange(i + 1, 11).setValue(adminEmail); 
        
      } else {
        reqSheet.getRange(i + 1, 2).setValue(newStatus); 
        reqSheet.getRange(i + 1, 10).setValue(new Date()); 
        reqSheet.getRange(i + 1, 11).setValue(adminEmail); 
      }

      if (newStatus === 'Approved') {
        return `Request approved. ${scheduleResult}`;
      } else {
        return "Request has been denied.";
      }
    }
  }
  
  throw new Error("Could not find the request ID.");
}

// ================= NEW/MODIFIED FUNCTIONS =================

// (No Change)
function getAdherenceRange(adminEmail, userNames, startDateStr, endDateStr) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  
  let targetUserNames = [];
  
  // Security Check: If user is an agent, force userNames to be only them
  if (adminRole === 'agent') {
    const selfName = userData.emailToName[adminEmail];
    if (!selfName) throw new Error("Your user account was not found.");
    targetUserNames = [selfName];
  } else {
    targetUserNames = userNames; // Admin can view the list they provided
  }
  
  const targetUserSet = new Set(targetUserNames.map(name => name.toLowerCase()));
  const timeZone = Session.getScriptTimeZone();
  const startDate = new Date(startDateStr);
  const endDate = new Date(endDateStr);
  startDate.setHours(0, 0, 0, 0);
  endDate.setHours(23, 59, 59, 999);
  
  const results = [];

  // 1. Get Adherence Data
  const adherenceSheet = getOrCreateSheet(ss, SHEET_NAMES.adherence);
  const adherenceData = adherenceSheet.getDataRange().getValues();
  
  for (let i = 1; i < adherenceData.length; i++) {
    const row = adherenceData[i];
    const rowUser = (row[1] || "").toString().trim().toLowerCase();
    
    if (targetUserSet.has(rowUser)) {
      try {
        const rowDate = new Date(row[0]);
        if (rowDate >= startDate && rowDate <= endDate) {
          results.push({
            date: convertDateToString(row[0]),
            userName: row[1],
            login: convertDateToString(row[2]),
            firstBreakIn: convertDateToString(row[3]),
            firstBreakOut: convertDateToString(row[4]),
            lunchIn: convertDateToString(row[5]),
            lunchOut: convertDateToString(row[6]),
            lastBreakIn: convertDateToString(row[7]),
            lastBreakOut: convertDateToString(row[8]),
            logout: convertDateToString(row[9]),
            tardy: row[10],
            overtime: row[11],
            earlyLeave: row[12],
            leaveType: row[13],
            firstBreakExceed: row[16],
            lunchExceed: row[17],
            lastBreakExceed: row[18],
            otherCodes: [] 
          });
        }
      } catch (e) {
        Logger.log(`Skipping adherence row ${i+1}. Invalid date. Error: ${e.message}`);
      }
    }
  }
  
  // Sort by date, then by user name
  results.sort((a, b) => {
    if (a.date < b.date) return -1;
    if (a.date > b.date) return 1;
    if (a.userName < b.userName) return -1;
    if (a.userName > b.userName) return 1;
    return 0;
  });
  
  if (results.length === 0) {
    return { error: `No adherence records found for the selected criteria.` };
  }
  
  return results;
}


// (No Change)
function getMySchedule(userEmail) {
  const ss = getSpreadsheet();
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const nextSevenDays = new Date(today);
  nextSevenDays.setDate(today.getDate() + 7);
  
  // *** NEW: DEBUG LOGGING ***
  Logger.log(`getMySchedule for ${userEmail}: Today: ${today.toISOString()}, 7-Day-Cutoff: ${nextSevenDays.toISOString()}`);

  const mySchedule = [];
  
  for (let i = 1; i < scheduleData.length; i++) {
    const row = scheduleData[i];
    const schEmail = (row[5] || "").toString().trim().toLowerCase();
    
    if (schEmail === userEmail) {
      try {
        // *** NEW: DEBUG LOGGING ***
        Logger.log(`Checking row ${i+1}: Email: ${schEmail}, Date: ${row[1]}, IsMatch: ${schEmail === userEmail}`);
        const schDate = new Date(row[1]);
        
        // *** NEW: DEBUG LOGGING ***
        Logger.log(`DateCheck: ${schDate.toISOString()} >= ${today.toISOString()} && ${schDate.toISOString()} < ${nextSevenDays.toISOString()} = ${schDate >= today && schDate < nextSevenDays}`);

        if (schDate >= today && schDate < nextSevenDays) { 
          
          let startTime = row[2];
          let endTime = row[3];
          
          if (startTime instanceof Date) {
            startTime = Utilities.formatDate(startTime, timeZone, "HH:mm");
          }
          if (endTime instanceof Date) {
            endTime = Utilities.formatDate(endTime, timeZone, "HH:mm");
          }
          
          mySchedule.push({
            date: convertDateToString(schDate),
            leaveType: row[4] || 'Present',
            startTime: startTime,
            endTime: endTime
          });
        }
      } catch(e) {
        Logger.log(`Skipping schedule row ${i+1}. Invalid date. Error: ${e.message}`);
      }
    }
  }
  
  // Sort by date
  mySchedule.sort((a, b) => new Date(a.date) - new Date(b.date));
  
  return mySchedule;
}


// (No Change)
function adjustLeaveBalance(adminEmail, userEmail, leaveType, amount, reason) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  if (adminRole !== 'admin' && adminRole !== 'superadmin') {
    throw new Error("Permission denied. Only admins can adjust balances.");
  }
  
  const balanceKey = leaveType.toLowerCase();
  const balanceCol = { annual: 4, sick: 5, casual: 6 }[balanceKey];
  if (!balanceCol) {
    throw new Error(`Unknown leave type: ${leaveType}.`);
  }
  
  const userRow = userData.emailToRow[userEmail];
  const userName = userData.emailToName[userEmail];
  if (!userRow) {
    throw new Error(`Could not find user ${userName} in Data Base.`);
  }
  
  const balanceRange = dbSheet.getRange(userRow, balanceCol);
  const currentBalance = parseFloat(balanceRange.getValue()) || 0;
  const newBalance = currentBalance + amount;
  
  balanceRange.setValue(newBalance);
  
  // Log the adjustment
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  logsSheet.appendRow([
    new Date(), 
    userName, 
    adminEmail, 
    "Balance Adjustment", 
    `Admin: ${adminEmail} | User: ${userName} | Type: ${leaveType} | Amount: ${amount} | Reason: ${reason} | Old: ${currentBalance} | New: ${newBalance}`
  ]);
  
  return `Successfully adjusted ${userName}'s ${leaveType} balance from ${currentBalance} to ${newBalance}.`;
}

// (No Change)
function importScheduleCSV(adminEmail, csvData) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  
  if (adminRole !== 'admin' && adminRole !== 'superadmin') {
    throw new Error("Permission denied. Only admins can import schedules.");
  }
  
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  const timeZone = Session.getScriptTimeZone();
  
  // Build a map of existing schedules
  const userScheduleMap = {};
  for (let i = 1; i < scheduleData.length; i++) {
    const rowEmail = scheduleData[i][5];
    const rowDateRaw = scheduleData[i][1];
    if (rowEmail && rowDateRaw) {
      const email = rowEmail.toLowerCase();
      if (!userScheduleMap[email]) {
        userScheduleMap[email] = {};
      }
      const rowDate = new Date(rowDateRaw);
      const rowDateStr = Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy");
      userScheduleMap[email][rowDateStr] = i + 1; 
    }
  }
  
  let daysUpdated = 0;
  let daysCreated = 0;
  let errors = 0;
  let errorLog = [];

  for (const row of csvData) {
    try {
      const userName = row.Name;
      const userEmail = (row.Email || "").toLowerCase();
      const dateStr = row.Date; // Expects MM/dd/yyyy
      let startTime = row.StartTime || ""; // Expects HH:mm
      let endTime = row.EndTime || ""; // Expects HH:mm
      let leaveType = row.LeaveType || "Present";
      
      if (!userName || !userEmail || !dateStr) {
        throw new Error("Missing required field (Name, Email, or Date).");
      }
      
      // If leave type is not Present, clear times
      if (leaveType.toLowerCase() !== "present") {
        startTime = "";
        endTime = "";
      }

      const targetDate = new Date(dateStr);
      if (isNaN(targetDate.getTime())) {
        throw new Error(`Invalid date format: ${dateStr}. Use MM/dd/yyyy`);
      }
      
      // Get the map for the specific user
      const emailMap = userScheduleMap[userEmail] || {};
      
      const result = updateOrAddSingleSchedule(
        scheduleSheet, emailMap, logsSheet,
        userEmail, userName,
        targetDate, dateStr,
        startTime, endTime, leaveType, adminEmail
      );
      
      if (result === "UPDATED") daysUpdated++;
      if (result === "CREATED") daysCreated++;

    } catch (e) {
      errors++;
      errorLog.push(`Row ${row.Name}/${row.Date}: ${e.message}`);
    }
  }

  if (errors > 0) {
    return `Error: Import complete with ${errors} errors. (Created: ${daysCreated}, Updated: ${daysUpdated}). Errors: ${errorLog.join(' | ')}`;
  }
  
  return `Import successful. Records Created: ${daysCreated}, Records Updated: ${daysUpdated}.`;
}

// MODIFIED: Logic updated to handle 'ALL_USERS' and load team data if needed.
function getDashboardData(adminEmail, userEmails, date) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  if (adminRole !== 'admin' && adminRole !== 'superadmin') {
    throw new Error("Permission denied.");
  }
  
  const timeZone = Session.getScriptTimeZone();
  
  // Use the date from the date picker
  const targetDate = new Date(date);
  const targetDateStr = Utilities.formatDate(targetDate, timeZone, "MM/dd/yyyy");
  
  // --- UPDATED: Filter by selected users ---
  let targetUserEmails = userEmails;

  // Use the specific list provided by the custom dropdown
  const targetUserSet = new Set(targetUserEmails);
  // --- END UPDATED FILTER ---
  
  const statusMap = {
    "Logged In": 0,
    "On Break/Other": 0,
    "On Leave": 0,
    "Absent": 0,
    "Pending Login": 0,
    "Logged Out": 0
  };
  
  const totalAdherenceMetrics = {
    totalTardy: 0,
    totalEarlyLeave: 0,
    totalOvertime: 0,
    totalBreakExceed: 0,
    totalLunchExceed: 0
  };
  
  // Individual metrics
  const userMetricsMap = {}; 
  targetUserEmails.forEach(email => {
    const name = userData.emailToName[email] || email;
    userMetricsMap[name] = {
      name: name,
      tardy: 0,
      earlyLeave: 0,
      overtime: 0,
      breakExceed: 0,
      lunchExceed: 0
    };
  });
  // ---

  // 1. Get Today's Schedule
  const scheduleSheet = getOrCreateSheet(ss, SHEET_NAMES.schedule);
  const scheduleData = scheduleSheet.getDataRange().getValues();
  const usersScheduledToday = new Set();
  
  for (let i = 1; i < scheduleData.length; i++) {
    const row = scheduleData[i];
    const schEmail = (row[5] || "").toLowerCase();
    
    // FILTER: Only check users in our target set
    if (!targetUserSet.has(schEmail)) continue;
    
    const schDate = new Date(row[1]);
    const schDateStr = Utilities.formatDate(schDate, timeZone, "MM/dd/yyyy");
    
    if (schDateStr === targetDateStr) { // Use targetDateStr
      const leaveType = (row[4] || "Present").toLowerCase();
      
      if (leaveType === "present") {
        usersScheduledToday.add(schEmail);
      } else if (leaveType === "absent") {
        statusMap["Absent"]++;
      } else {
        statusMap["On Leave"]++;
      }
    }
  }
  
  // 2. Get Today's Adherence
  const adherenceSheet = getOrCreateSheet(ss, SHEET_NAMES.adherence);
  const adherenceData = adherenceSheet.getDataRange().getValues();
  
  // NEW: Get "Other Codes" for real-time status
  const otherCodesSheet = getOrCreateSheet(ss, SHEET_NAMES.otherCodes);
  const otherCodesData = otherCodesSheet.getDataRange().getValues();
  const userLastOtherCode = {}; // Map user -> { code: "Meeting", type: "In" }
  
  for (let i = otherCodesData.length - 1; i > 0; i--) { // Go backwards
    const row = otherCodesData[i];
    const rowDate = new Date(row[0]);
    const rowShiftDate = getShiftDate(rowDate, SHIFT_CUTOFF_HOUR);
    const rowDateStr = Utilities.formatDate(rowShiftDate, timeZone, "MM/dd/yyyy");
    
    if (rowDateStr === targetDateStr) {
      const userName = row[1];
      const userEmail = userData.nameToEmail[userName];
      
      if (userEmail && targetUserSet.has(userEmail.toLowerCase())) {
        if (!userLastOtherCode[userEmail]) { // Only get the *last* punch
          const [code, type] = (row[2] || "").split(" ");
          userLastOtherCode[userEmail] = {
            code: code,
            type: type
          };
        }
      }
    }
  }
  // ---
  
  for (let i = 1; i < adherenceData.length; i++) {
    const row = adherenceData[i];
    const rowDate = new Date(row[0]);
    const rowDateStr = Utilities.formatDate(rowDate, timeZone, "MM/dd/yyyy");
    
    if (rowDateStr === targetDateStr) { // Use targetDateStr
      const userName = row[1];
      const userEmail = userData.nameToEmail[userName];
      
      // FILTER: Only check users in our target set
      if (userEmail && targetUserSet.has(userEmail.toLowerCase())) {
        
        // If user is in the "Present" set
        if (usersScheduledToday.has(userEmail.toLowerCase())) {
          const login = row[2];
          const b1_in = row[3];
          const b1_out = row[4];
          const l_in = row[5];
          const l_out = row[6];
          const b2_in = row[7];
          const b2_out = row[8];
          const logout = row[9];
          
          // --- UPDATED: Real-time Status Logic ---
          let agentStatus = "Pending Login";
          if (login && !logout) {
            agentStatus = "Logged In"; // Default logged-in state
            
            // Check Other Codes first
            const lastOther = userLastOtherCode[userEmail.toLowerCase()];
            if (lastOther && lastOther.type === 'In') {
              agentStatus = "On Break/Other"; // Combined status
            } else {
              // If not in Other Code, check breaks
              if (b1_in && !b1_out) agentStatus = "On Break/Other";
               if (l_in && !l_out) agentStatus = "On Break/Other";
              if (b2_in && !b2_out) agentStatus = "On Break/Other";
            }
          } else if (login && logout) {
            agentStatus = "Logged Out";
          }
          // --- END UPDATED ---
          
          if (statusMap[agentStatus] !== undefined) {
            statusMap[agentStatus]++;
          }
          
          usersScheduledToday.delete(userEmail.toLowerCase());
        }
        
        // 3. Sum Adherence Metrics
        const tardy = parseFloat(row[10]) || 0;
        const earlyLeave = parseFloat(row[12]) || 0;
        const overtime = parseFloat(row[11]) || 0;
        const breakExceed = (parseFloat(row[16]) || 0) + (parseFloat(row[18]) || 0);
        const lunchExceed = parseFloat(row[17]) || 0;

        totalAdherenceMetrics.totalTardy += tardy;
        totalAdherenceMetrics.totalEarlyLeave += earlyLeave;
        totalAdherenceMetrics.totalOvertime += overtime;
        totalAdherenceMetrics.totalBreakExceed += breakExceed;
        totalAdherenceMetrics.totalLunchExceed += lunchExceed;
        
        // Add to individual user
        if (userMetricsMap[userName]) {
          userMetricsMap[userName].tardy += tardy;
          userMetricsMap[userName].earlyLeave += earlyLeave;
          userMetricsMap[userName].overtime += overtime;
          userMetricsMap[userName].breakExceed += breakExceed;
          userMetricsMap[userName].lunchExceed += lunchExceed;
        }
      }
    }
  }
  
  // Any users left in this set are scheduled but have no adherence row
  statusMap["Pending Login"] += usersScheduledToday.size;
  
  // 4. Get Pending Leave Requests
  const reqSheet = getOrCreateSheet(ss, SHEET_NAMES.leaveRequests);
  const reqData = reqSheet.getDataRange().getValues();
  const pendingRequests = [];
  
  for (let i = 1; i < reqData.length; i++) {
    const row = reqData[i];
    // FILTER: Check if this request is from a user in our target set
   const reqEmail = (row[2] || "").toLowerCase();
    
    if (row[1] && row[1].toString().trim().toLowerCase() === 'pending' && targetUserSet.has(reqEmail)) {
      try {
        pendingRequests.push({
          name: row[3], // Name
          type: row[4], // Type
          startDate: convertDateToString(new Date(row[5])), // Start Date
          days: row[7]  // Total Days
        });
      } catch (e) {
        Logger.log(`Failed to parse pending request row ${i+1}. Error: ${e.message}`);
      }
    }
  }
  
  // Format status counts for Google Charts
  const statusCounts = Object.keys(statusMap).map(key => [key, statusMap[key]]);
  
  // Format individual metrics
  const individualAdherenceMetrics = Object.values(userMetricsMap);
  
  return {
    statusCounts: statusCounts,
    totalAdherenceMetrics: totalAdherenceMetrics, // Renamed for clarity
   individualAdherenceMetrics: individualAdherenceMetrics, // NEW
    pendingRequests: pendingRequests
  };
}

// --- NEW: "My Team" Helper Functions ---
function saveMyTeam(adminEmail, userEmails) {
  try {
    // Uses Google Apps Script's built-in User Properties for saving user-specific settings.
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('myTeam', JSON.stringify(userEmails));
    return "Successfully saved 'My Team' preference.";
  } catch (e) {
    throw new Error("Failed to save team preferences: " + e.message);
  }
}

function getMyTeam(adminEmail) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    // Getting properties implicitly forces the Google auth dialog if needed.
    const properties = userProperties.getProperties(); 
    const myTeam = properties['myTeam'];
    return myTeam ? JSON.parse(myTeam) : [];
  } catch (e) {
    Logger.log("Failed to load team preferences: " + e.message);
    // Throwing an error here would break the dashboard's initial load. 
    // We return an empty array instead, and let the front-end handle the fallback.
   return [];
  }
}

// --- NEW: Reporting Line Function ---
function updateReportingLine(adminEmail, userEmail, newSupervisorEmail) {
  const ss = getSpreadsheet();
  const dbSheet = getOrCreateSheet(ss, SHEET_NAMES.database);
  const userData = getUserDataFromDb(dbSheet);
  
  const adminRole = userData.emailToRole[adminEmail] || 'agent';
  if (adminRole !== 'admin' && adminRole !== 'superadmin') {
    throw new Error("Permission denied. Only admins can change reporting lines.");
  }
  
  const userName = userData.emailToName[userEmail];
  const newSupervisorName = userData.emailToName[newSupervisorEmail];
  if (!userName) throw new Error(`Could not find user: ${userEmail}`);
  if (!newSupervisorName) throw new Error(`Could not find new supervisor: ${newSupervisorEmail}`);

  const userRow = userData.emailToRow[userEmail];
  const currentUserSupervisor = userData.emailToSupervisor[userEmail];

  // Check for auto-approval
  let canAutoApprove = false;
  if (adminRole === 'superadmin') {
    canAutoApprove = true;
  } else if (adminRole === 'admin') {
    // Check if both the user's current supervisor AND the new supervisor report to this admin
    const currentSupervisorManager = userData.emailToSupervisor[currentUserSupervisor];
    const newSupervisorManager = userData.emailToSupervisor[newSupervisorEmail];
    
    if (currentSupervisorManager === adminEmail && newSupervisorManager === adminEmail) {
      canAutoApprove = true;
    }
  }

  if (!canAutoApprove) {
    // This is where we will build Phase 2 (requesting the change)
    // For now, we will just show a permission error.
    throw new Error("Permission Denied: You do not have authority to approve this change. (This will become a request in Phase 2).");
  }

  // --- Auto-Approval Logic ---
  // Update the SupervisorEmail column (Column G = 7)
  dbSheet.getRange(userRow, 7).setValue(newSupervisorEmail);
  
  // Log the change
  const logsSheet = getOrCreateSheet(ss, SHEET_NAMES.logs);
  logsSheet.appendRow([
    new Date(), 
    userName, 
    adminEmail, 
    "Reporting Line Change", 
    `User: ${userName} moved to Supervisor: ${newSupervisorName} by ${adminEmail}`
  ]);
  
  return `${userName} has been successfully reassigned to ${newSupervisorName}.`;
}// ==========================================================
// === NEW COACHING TEMPLATE FUNCTIONS (PHASE 3) ===
// ==========================================================

// ==========================================================
// === NEW COACHING TEMPLATE FUNCTIONS (PHASE 3) ===
// ==========================================================

/**
 * (MODIFIED - PHASE 4 Migration)
 * Gets all non-archived coaching templates.
 * Automatically creates the default template if it's missing.
 */
function webGetCoachingTemplates() {
  try {
    // --- NEW: Run the migration helper ---
    _createDefaultQualityTemplate();
    // ------------------------------------

    const ss = getSpreadsheet();
    const templateSheet = getOrCreateSheet(ss, "CoachingTemplates");
    const criteriaSheet = getOrCreateSheet(ss, "CoachingTemplateCriteria");

    // 1. Get all criteria and map them by TemplateID
    const criteriaData = criteriaSheet.getRange(2, 1, criteriaSheet.getLastRow() - 1, 7).getValues();
    const criteriaMap = {};
    for (let i = 0; i < criteriaData.length; i++) {
      const row = criteriaData[i];
      const templateID = row[0];
      if (!criteriaMap[templateID]) {
        criteriaMap[templateID] = [];
      }
      criteriaMap[templateID].push({
        itemID: row[1],
        category: row[2],
        criteriaText: row[3],
        inputType: row[4],
        weight: row[5],
        itemOrder: row[6]
      });
    }

    // 2. Get all templates and attach their criteria
    const templateData = templateSheet.getRange(2, 1, templateSheet.getLastRow() - 1, 5).getValues();
    const templates = [];
    for (let i = 0; i < templateData.length; i++) {
      const row = templateData[i];
      const templateID = row[0];
      const isArchived = row[4];

      if (isArchived !== true) {
        templates.push({
          templateID: templateID,
          templateName: row[1],
          createdBy: row[2],
          dateCreated: row[3],
          criteria: criteriaMap[templateID] || []
        });
      }
    }

    return templates;

  } catch (err) {
    Logger.log("webGetCoachingTemplates Error: " + err.message);
    return { error: err.message };
  }
}

/**
 * Saves a coaching template (both new and existing).
 */
function webSaveCoachingTemplate(templateObject) {
  try {
    const adminEmail = Session.getActiveUser().getEmail().toLowerCase();
    const ss = getSpreadsheet();
    const templateSheet = getOrCreateSheet(ss, "CoachingTemplates");
    const criteriaSheet = getOrCreateSheet(ss, "CoachingTemplateCriteria");

    let templateID = templateObject.templateID;

    if (templateID) {
      // It's an existing template
      const templateData = templateSheet.getRange(2, 1, templateSheet.getLastRow() - 1, 1).getValues();
      let rowToUpdate = -1;
      for (let i = 0; i < templateData.length; i++) {
        if (templateData[i][0] === templateID) {
          rowToUpdate = i + 2;
          break;
        }
      }
      if (rowToUpdate !== -1) {
        templateSheet.getRange(rowToUpdate, 2).setValue(templateObject.templateName); // Update name
      } else {
        throw new Error("Could not find TemplateID to update.");
      }

      // Delete all old criteria
      const criteriaData = criteriaSheet.getDataRange().getValues();
      const rowsToDelete = [];
      for (let i = criteriaData.length - 1; i >= 1; i--) { // Go backwards
        if (criteriaData[i][0] === templateID) {
          rowsToDelete.push(i + 1);
        }
      }
      for (const row of rowsToDelete) {
        criteriaSheet.deleteRow(row);
      }

    } else {
      // It's a new template
      templateID = templateObject.templateID || `TPL-${new Date().getTime()}`; // Use default-id or generate new
      templateSheet.appendRow([
        templateID,
        templateObject.templateName,
        adminEmail,
        new Date(),
        false // IsArchived
      ]);
    }

    // 2. Save all criteria
    const newCriteriaRows = [];
    if (templateObject.criteria && templateObject.criteria.length > 0) {
      templateObject.criteria.forEach((item, index) => {
        const itemID = `ITEM-${new Date().getTime()}-${index}`;
        newCriteriaRows.push([
          templateID,
          itemID,
          item.category,
          item.criteriaText,
          item.inputType,
          item.weight,
          index // Use the array index for ordering
        ]);
      });

      if (newCriteriaRows.length > 0) {
        criteriaSheet.getRange(criteriaSheet.getLastRow() + 1, 1, newCriteriaRows.length, 7).setValues(newCriteriaRows);
      }
    }

    return { success: true, newID: templateID, name: templateObject.templateName };

  } catch (err) {
    Logger.log("webSaveCoachingTemplate Error: " + err.message);
    return { error: err.message };
  }
}


/**
 * Soft-deletes a template by setting IsArchived = true.
 */
function webDeleteCoachingTemplate(templateID) {
  try {
    const ss = getSpreadsheet();
    const templateSheet = getOrCreateSheet(ss, "CoachingTemplates");

    const templateData = templateSheet.getRange(2, 1, templateSheet.getLastRow() - 1, 1).getValues();
    let rowToUpdate = -1;
    for (let i = 0; i < templateData.length; i++) {
      if (templateData[i][0] === templateID) {
        rowToUpdate = i + 2;
        break;
      }
    }

    if (rowToUpdate !== -1) {
      templateSheet.getRange(rowToUpdate, 5).setValue(true); // Set IsArchived to true
      return { success: true, id: templateID };
    } else {
      throw new Error("Could not find TemplateID to delete.");
    }

  } catch (err) {
    Logger.log("webDeleteCoachingTemplate Error: " + err.message);
    return { error: err.message };
  }
}

/**
 * (NEW - PHASE 4 Migration)
 * This is a helper function that automatically creates the
 * original "Quality Score" template in the new dynamic system.
 */
function _createDefaultQualityTemplate() {
  try {
    const qualityCategories = [
      { 
        category: "Greeting & Opening",
        criteria: [
          "Agent greeted the customer professionally and introduced themselves appropriately",
          "Agent confirmed the customer’s name & purpose of the call/chat"
        ]
      },
      {
        category: "Communication Skills & Understanding Needs",
        criteria: [
          "Agent conversed actively without interrupting",
          "Agent asked relevant questions to understand customer needs",
          "Agent acknowledged customer concerns appropriately",
          "Language was clear, understandable, and free of jargon",
          "Agent applied correct hold etiquettes",
          "Tone was confident, professional and engaging"
        ]
      },
      {
        category: "Product Knowledge & providing solution",
        criteria: [
          "Agent demonstrated strong knowledge of Lenovo products/services",
          "Agent offered the right solution based on customer's needs",
          "Agent was able to handle objections confidently & Highlighted Lenovo's competitive advantage"
        ]
      },
      {
        category: "Tools usage and Chat/ Call Logging",
        criteria: [
          "Agent applied correct disposition.",
          "Agent logged the chat with all relevant details in Dynamics 365 B2C"
        ]
      },
      {
        category: "Sales Closing & Call to Action",
        criteria: [
          "Agent clearly stated pricing, offers, and benefits.",
          "Agent confirmed next steps (e.g., sending a quote, scheduling a follow-up)"
        ]
      },
      {
        category: "Process Compliance",
        criteria: [
          "Agent followed Lenovo's sales process & compliance guidelines. OR Agent transfered the chat to the approriate que when applicable"
        ]
      },
      {
        category: "Wrap-Up & Closing",
        criteria: [
          "Agent confirmed if the customer’s query was fully addressed",
          "Agent ended the chat approprietly.",
          "Follow-up commitment created (if applicable)"
        ]
      }
    ];

    const templateObject = {
      templateID: "default-quality-score", // Use a fixed, special ID
      templateName: "Quality Score (Default)",
      criteria: []
    };

    qualityCategories.forEach(cat => {
      cat.criteria.forEach(crit => {
        templateObject.criteria.push({
          category: cat.category,
          criteriaText: crit,
          inputType: 'score_0-1',
          weight: 1
        });
      });
    });

    const ss = getSpreadsheet();
    const templateSheet = getOrCreateSheet(ss, "CoachingTemplates");

    const templateData = templateSheet.getRange(2, 1, templateSheet.getLastRow() - 1, 2).getValues();
    let exists = false;
    for (let i = 0; i < templateData.length; i++) {
      if (templateData[i][0] === templateObject.templateID || templateData[i][1] === templateObject.templateName) {
        exists = true;
        break;
      }
    }

    if (!exists) {
      Logger.log("Default template not found. Saving it now.");
      webSaveCoachingTemplate(templateObject); // This will call getOrCreateSheet for all new sheets
    } else {
      Logger.log("Default template already exists. Skipping creation.");
    }
  } catch (e) {
    Logger.log("Error creating default template: " + e.message);
  }
}
function testConnection() {
  Logger.log("Test connection was successful.");
  return "It works! The backend file is okay.";
}
