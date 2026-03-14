// ============================================
// HSE MANAGEMENT SYSTEM — Google Apps Script
// Developer: Mahesh | HSE Officer
// Company: Hayleys Fentons Limited
// Version: 4.0.0 (Briefing Photo + Toolbox Fix + Auto-suggest)
// ============================================
// CHANGELOG v4.0.0:
// ✅ Briefing/Toolbox Talk photo upload support
// ✅ Server-side personnel/company/project endpoints
// ✅ Updated Briefing headers with Photo URL
// ✅ Future-proof API architecture
// ============================================

// 📁 Drive Folder IDs
var INCIDENT_FOLDER_ID = "1EGJgsFc8TSEbAfQlaJ1zNRjPyoahGyWK"; 
var TRAINING_FOLDER_ID = "1Y8Pzufd4Yhl4X6nSyTP3Fu0MxQ9LxqPZ";
var BRIEFING_FOLDER_ID = "1nwRLiR9rw8VF56iBgaPxefJ5x93DFpR7"; // 🆕 Briefing/Toolbox photos

// 📧 Email Settings
var EMAIL_ADDRESS = "dayanmahesh123@gmail.com";

// ============================================
// 📥 doPost — Handle all form submissions
// ============================================
function doPost(e) {
  try {
    var parsedData = JSON.parse(e.postData.contents);
    var sheetName = parsedData.sheetName;
    var values = parsedData.values;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      var headers = getHeadersForSheet(sheetName);
      if (headers.length > 0) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length)
          .setBackground('#0d1b2a')
          .setFontColor('#ffffff')
          .setFontWeight('bold')
          .setFontSize(10);
        sheet.setFrozenRows(1);
      }
    }

    // 📸 Photo save helper
    function savePhotoToDrive(photoData, fileNamePrefix, targetFolderId) {
      if (photoData && photoData.toString().indexOf('data:image') !== -1) {
        try {
          var folder = DriveApp.getFolderById(targetFolderId);
          var contentType = photoData.split(';')[0].split(':')[1];
          var base64String = photoData.split(',')[1];
          var blob = Utilities.newBlob(
            Utilities.base64Decode(base64String), 
            contentType, 
            fileNamePrefix + "_" + new Date().getTime() + ".jpg"
          );
          var file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          return file.getUrl();
        } catch (photoError) {
          Logger.log("Photo save error: " + photoError.toString());
          return "Photo upload failed";
        }
      }
      return "";
    }

    // 🚨 INCIDENTS
    if (sheetName === 'Incidents') {
      var photoUrl = savePhotoToDrive(values[14], values[0], INCIDENT_FOLDER_ID);
      values[14] = photoUrl;
      sheet.appendRow(values);
      if (values[4] === "High" || values[4] === "Critical") {
        sendIncidentEmail(values, photoUrl);
      }
    }
    // 📚 TRAINING
    else if (sheetName === 'Training') {
      var photoUrl = savePhotoToDrive(values[10], values[0], TRAINING_FOLDER_ID);
      values[10] = photoUrl;
      sheet.appendRow(values);
    }
    // 📢 BRIEFING — 🆕 UPDATED with Photo
    else if (sheetName === 'Briefing') {
      var photoUrl = savePhotoToDrive(values[14], values[0], BRIEFING_FOLDER_ID);
      values[14] = photoUrl;
      sheet.appendRow(values);
      Logger.log("Briefing: " + values[3] + " — " + values[9] + " attendees, Photo: " + (photoUrl ? "Yes" : "No"));
    }
    // 📋 EMPLOYEE ATTENDANCE
    else if (sheetName === 'EmployeeAttendance') {
      sheet.appendRow(values);
    }
    // 👷 SUB-CONTRACTOR ATTENDANCE
    else if (sheetName === 'SubContractorAttendance') {
      sheet.appendRow(values);
      if (values[10] === 'No' || values[11] === 'No') {
        sendSCAlertEmail(values, values[10], values[11]);
      }
    }
    // 🏥 HEALTH INFO
    else if (sheetName === 'HealthInfo') {
      sheet.appendRow(values);
      if (values[23] === 'No' || values[20] === 'Expired') {
        sendHealthAlertEmail(values);
      }
      if (values[15] && (values[15].indexOf('Epilepsy') !== -1 || values[15].indexOf('Vertigo') !== -1)) {
        sendHealthRestrictionEmail(values);
      }
    }
    // 🟢 ALL OTHER SHEETS
    else {
      sheet.appendRow(values);
    }
    
    return ContentService.createTextOutput(
      JSON.stringify({"status":"success","message":"Data saved to "+sheetName,"timestamp":new Date().toISOString()})
    ).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log("doPost Error: " + error.toString());
    return ContentService.createTextOutput(
      JSON.stringify({"status":"error","message":error.toString()})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// 📤 doGet — Dashboard + 🆕 Data Endpoints
// ============================================
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
    
    // 🆕 Server-side autocomplete endpoints
    if (action === 'getPersonnel') return getPersonnelData();
    if (action === 'getCompanies') return getCompanyData();
    if (action === 'getProjects') return getProjectData();
    
    // Default: Dashboard data
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var response = {};
    
    // 1️⃣ Last Incident Date
    var incSheet = ss.getSheetByName("Incidents");
    if (incSheet && incSheet.getLastRow() > 1) {
      response.lastDate = incSheet.getRange(incSheet.getLastRow(), 2).getValue();
    } else {
      response.lastDate = new Date().toISOString();
    }
    
    // 2️⃣ Employee Attendance
    var eaSheet = ss.getSheetByName("EmployeeAttendance");
    if (eaSheet && eaSheet.getLastRow() > 1) {
      var todayStr = Utilities.formatDate(new Date(), "Asia/Colombo", "yyyy-MM-dd");
      var eaData = eaSheet.getDataRange().getValues();
      var empPresent=0, empAbsent=0, empTotalHours=0;
      for (var i=1; i<eaData.length; i++) {
        var rowDate = eaData[i][1] instanceof Date ? 
          Utilities.formatDate(eaData[i][1], "Asia/Colombo", "yyyy-MM-dd") : eaData[i][1].toString();
        if (rowDate === todayStr) {
          if (eaData[i][12]==="Present"||eaData[i][12]==="Late") empPresent++;
          else if (eaData[i][12]==="Absent") empAbsent++;
          empTotalHours += parseFloat(eaData[i][10]) || 0;
        }
      }
      response.employeePresent = empPresent;
      response.employeeAbsent = empAbsent;
      response.employeeTotalHours = empTotalHours;
    }
    
    // 3️⃣ SC Attendance
    var scSheet = ss.getSheetByName("SubContractorAttendance");
    if (scSheet && scSheet.getLastRow() > 1) {
      var todayStr2 = Utilities.formatDate(new Date(), "Asia/Colombo", "yyyy-MM-dd");
      var scData = scSheet.getDataRange().getValues();
      var scPresent=0, scNoInduction=0, scNoPPE=0;
      for (var j=1; j<scData.length; j++) {
        var scRowDate = scData[j][1] instanceof Date ?
          Utilities.formatDate(scData[j][1], "Asia/Colombo", "yyyy-MM-dd") : scData[j][1].toString();
        if (scRowDate === todayStr2) {
          if (scData[j][13]==="Present") scPresent++;
          if (scData[j][10]==="No") scNoInduction++;
          if (scData[j][11]==="No") scNoPPE++;
        }
      }
      response.scPresent = scPresent;
      response.scNoInduction = scNoInduction;
      response.scNoPPE = scNoPPE;
    }
    
    // 4️⃣ Health Info
    var hiSheet = ss.getSheetByName("HealthInfo");
    if (hiSheet && hiSheet.getLastRow() > 1) {
      var hiData = hiSheet.getDataRange().getValues();
      var expiredCerts=0, unfitCount=0;
      for (var k=1; k<hiData.length; k++) {
        if (hiData[k][20]==="Expired") expiredCerts++;
        if (hiData[k][23]==="No") unfitCount++;
      }
      response.expiredCerts = expiredCerts;
      response.unfitWorkers = unfitCount;
      response.totalHealthRecords = hiData.length - 1;
    }
    
    // 5️⃣ Briefing Count
    var brSheet = ss.getSheetByName("Briefing");
    if (brSheet && brSheet.getLastRow() > 1) {
      var todayStr3 = Utilities.formatDate(new Date(), "Asia/Colombo", "yyyy-MM-dd");
      var brData = brSheet.getDataRange().getValues();
      var brCount=0, brTotalAttendees=0;
      for (var m=1; m<brData.length; m++) {
        var brRowDate = brData[m][1] instanceof Date ?
          Utilities.formatDate(brData[m][1], "Asia/Colombo", "yyyy-MM-dd") : brData[m][1].toString();
        if (brRowDate === todayStr3) {
          brCount++;
          brTotalAttendees += parseInt(brData[m][9]) || 0;
        }
      }
      response.todayBriefings = brCount;
      response.todayBriefingAttendees = brTotalAttendees;
    }
    
    return ContentService.createTextOutput(
      JSON.stringify(response)
    ).setMimeType(ContentService.MimeType.JSON);
    
  } catch(ex) {
    Logger.log("doGet Error: " + ex.toString());
    return ContentService.createTextOutput(
      JSON.stringify({"error":ex.toString()})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// 🆕 SERVER-SIDE DATA FUNCTIONS (Future-proof)
// ============================================
function getPersonnelData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var personnel = [];
  var seen = {};
  
  var eaSheet = ss.getSheetByName("EmployeeAttendance");
  if (eaSheet && eaSheet.getLastRow() > 1) {
    var eaData = eaSheet.getDataRange().getValues();
    for (var i=1; i<eaData.length; i++) {
      var name = (eaData[i][2]||'').toString().trim();
      if (name && !seen[name.toLowerCase()]) {
        seen[name.toLowerCase()] = true;
        personnel.push({name:name,empId:eaData[i][3]||'',nic:eaData[i][4]||'',
          designation:eaData[i][5]||'',department:eaData[i][6]||'',
          project:eaData[i][7]||'',type:'employee'});
      }
    }
  }
  
  var scSheet = ss.getSheetByName("SubContractorAttendance");
  if (scSheet && scSheet.getLastRow() > 1) {
    var scData = scSheet.getDataRange().getValues();
    for (var j=1; j<scData.length; j++) {
      var scName = (scData[j][3]||'').toString().trim();
      if (scName && !seen[scName.toLowerCase()]) {
        seen[scName.toLowerCase()] = true;
        personnel.push({name:scName,nic:scData[j][4]||'',trade:scData[j][5]||'',
          company:scData[j][2]||'',project:scData[j][6]||'',type:'subcontractor'});
      }
    }
  }
  
  return ContentService.createTextOutput(
    JSON.stringify({status:'success',data:personnel,count:personnel.length})
  ).setMimeType(ContentService.MimeType.JSON);
}

function getCompanyData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var companies = {};
  
  var scSheet = ss.getSheetByName("SubContractorAttendance");
  if (scSheet && scSheet.getLastRow() > 1) {
    var scData = scSheet.getDataRange().getValues();
    for (var i=1; i<scData.length; i++) {
      var co = (scData[i][2]||'').toString().trim();
      if (co) companies[co] = true;
    }
  }
  
  var cidaSheet = ss.getSheetByName("CIDA");
  if (cidaSheet && cidaSheet.getLastRow() > 1) {
    var cidaData = cidaSheet.getDataRange().getValues();
    for (var j=1; j<cidaData.length; j++) {
      var cn = (cidaData[j][2]||'').toString().trim();
      if (cn) companies[cn] = true;
    }
  }
  
  return ContentService.createTextOutput(
    JSON.stringify({status:'success',data:Object.keys(companies)})
  ).setMimeType(ContentService.MimeType.JSON);
}

function getProjectData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var projects = {};
  
  ['EmployeeAttendance','SubContractorAttendance'].forEach(function(sn) {
    var sheet = ss.getSheetByName(sn);
    if (sheet && sheet.getLastRow() > 1) {
      var data = sheet.getDataRange().getValues();
      var col = sn==='EmployeeAttendance' ? 7 : 6;
      for (var i=1; i<data.length; i++) {
        var p = (data[i][col]||'').toString().trim();
        if (p) projects[p] = true;
      }
    }
  });
  
  return ContentService.createTextOutput(
    JSON.stringify({status:'success',data:Object.keys(projects)})
  ).setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// 📧 EMAIL FUNCTIONS
// ============================================
function sendIncidentEmail(values, photoUrl) {
  try {
    var severity = values[4];
    var emailHtml = 
      "<div style='font-family:Arial;max-width:600px;margin:0 auto;'>" +
      "<div style='background:#e63946;color:white;padding:20px;border-radius:8px 8px 0 0;'>" +
      "<h2 style='margin:0;'>🚨 " + severity + " Incident Alert</h2>" +
      "<p style='margin:5px 0 0;'>HSE Management System — Hayleys Fentons Ltd</p></div>" +
      "<div style='background:white;padding:20px;border:1px solid #ddd;'>" +
      "<table style='width:100%;border-collapse:collapse;'>" +
      "<tr><td style='padding:8px;font-weight:bold;border-bottom:1px solid #eee;'>Incident ID:</td><td style='padding:8px;border-bottom:1px solid #eee;'>" + values[0] + "</td></tr>" +
      "<tr><td style='padding:8px;font-weight:bold;border-bottom:1px solid #eee;'>Date & Time:</td><td style='padding:8px;border-bottom:1px solid #eee;'>" + values[1] + " " + values[2] + "</td></tr>" +
      "<tr><td style='padding:8px;font-weight:bold;border-bottom:1px solid #eee;'>Type:</td><td style='padding:8px;border-bottom:1px solid #eee;'>" + values[3] + "</td></tr>" +
      "<tr><td style='padding:8px;font-weight:bold;border-bottom:1px solid #eee;'>Severity:</td><td style='padding:8px;border-bottom:1px solid #eee;color:" + (severity==='Critical'?'#e63946':'#f77f00') + ";font-weight:bold;'>" + severity + "</td></tr>" +
      "<tr><td style='padding:8px;font-weight:bold;border-bottom:1px solid #eee;'>Location:</td><td style='padding:8px;border-bottom:1px solid #eee;'>" + values[5] + "</td></tr>" +
      "<tr><td style='padding:8px;font-weight:bold;border-bottom:1px solid #eee;'>Reported By:</td><td style='padding:8px;border-bottom:1px solid #eee;'>" + values[6] + "</td></tr>" +
      "<tr><td style='padding:8px;font-weight:bold;border-bottom:1px solid #eee;'>Injured Person:</td><td style='padding:8px;border-bottom:1px solid #eee;'>" + (values[7]||'N/A') + "</td></tr>" +
      "<tr><td style='padding:8px;font-weight:bold;border-bottom:1px solid #eee;'>Body Part:</td><td style='padding:8px;border-bottom:1px solid #eee;'>" + (values[8]||'N/A') + "</td></tr>" +
      "<tr><td style='padding:8px;font-weight:bold;' colspan='2'>Description:</td></tr>" +
      "<tr><td style='padding:8px;background:#f8f9fa;' colspan='2'>" + values[9] + "</td></tr></table>";
    if (photoUrl) emailHtml += "<p style='margin-top:15px;'><strong>📸 Photo:</strong> <a href='" + photoUrl + "'>View</a></p>";
    emailHtml += "</div>" +
      "<div style='background:#f8f9fa;padding:12px 20px;border:1px solid #ddd;border-top:none;border-radius:0 0 8px 8px;font-size:12px;color:#666;'>" +
      "<p>Factory Ordinance Sec. 61 | ISO 45001 Clause 10.2</p>" +
      "<p>Generated: " + new Date().toLocaleString() + "</p></div></div>";
    MailApp.sendEmail({to:EMAIL_ADDRESS,subject:"🚨 "+severity+" Incident — "+values[0],htmlBody:emailHtml});
  } catch(e) { Logger.log("Email error: "+e); }
}

function sendSCAlertEmail(values, induction, ppe) {
  try {
    var items = [];
    if (induction==='No') items.push("❌ Safety Induction NOT completed");
    if (ppe==='No') items.push("❌ PPE NOT compliant");
    var emailHtml = 
      "<div style='font-family:Arial;max-width:600px;margin:0 auto;'>" +
      "<div style='background:#f77f00;color:white;padding:20px;border-radius:8px 8px 0 0;'>" +
      "<h2 style='margin:0;'>⚠️ Sub-Contractor Safety Alert</h2></div>" +
      "<div style='background:white;padding:20px;border:1px solid #ddd;'>" +
      "<p><strong>Date:</strong> "+values[1]+"</p>" +
      "<p><strong>Company:</strong> "+values[2]+"</p>" +
      "<p><strong>Worker:</strong> "+values[3]+"</p>" +
      "<p><strong>NIC:</strong> "+values[4]+"</p>" +
      "<div style='background:#fff3cd;padding:15px;border-radius:8px;margin-top:10px;'>" +
      "<h3 style='color:#856404;'>🚫 Violations:</h3><ul>" +
      items.map(function(i){return "<li style='color:#721c24;font-weight:bold;'>"+i+"</li>";}).join('') +
      "</ul></div>" +
      "<p style='color:#e63946;font-weight:bold;margin-top:15px;'>⚠️ ACTION REQUIRED</p>" +
      "</div></div>";
    MailApp.sendEmail({to:EMAIL_ADDRESS,subject:"⚠️ SC Alert — "+values[3],htmlBody:emailHtml});
  } catch(e) { Logger.log("SC email error: "+e); }
}

function sendHealthAlertEmail(values) {
  try {
    var emailHtml = 
      "<div style='font-family:Arial;max-width:600px;margin:0 auto;'>" +
      "<div style='background:#e63946;color:white;padding:20px;border-radius:8px 8px 0 0;'>" +
      "<h2 style='margin:0;'>🏥 Health Alert</h2></div>" +
      "<div style='background:white;padding:20px;border:1px solid #ddd;'>" +
      "<p><strong>Name:</strong> "+values[2]+"</p>" +
      "<p><strong>Type:</strong> "+values[1]+"</p>" +
      "<p><strong>NIC:</strong> "+values[3]+"</p>" +
      "<div style='background:#f8d7da;padding:15px;border-radius:8px;margin-top:10px;'>" +
      "<h3 style='color:#721c24;'>🚫 Alert:</h3><ul>";
    if (values[23]==='No') emailHtml += "<li style='color:#721c24;'>❌ NOT FIT TO WORK</li>";
    if (values[20]==='Expired') emailHtml += "<li style='color:#721c24;'>❌ Certificate EXPIRED</li>";
    emailHtml += "</ul></div>" +
      "<p style='color:#e63946;font-weight:bold;margin-top:15px;'>⚠️ Do not allow to work without clearance</p>" +
      "</div></div>";
    MailApp.sendEmail({to:EMAIL_ADDRESS,subject:"🏥 HEALTH ALERT — "+values[2],htmlBody:emailHtml});
  } catch(e) { Logger.log("Health email error: "+e); }
}

function sendHealthRestrictionEmail(values) {
  try {
    var emailHtml = 
      "<div style='font-family:Arial;max-width:600px;margin:0 auto;'>" +
      "<div style='background:#f77f00;color:white;padding:20px;border-radius:8px 8px 0 0;'>" +
      "<h2 style='margin:0;'>⚠️ Work Restriction Required</h2></div>" +
      "<div style='background:white;padding:20px;border:1px solid #ddd;'>" +
      "<p><strong>Worker:</strong> "+values[2]+"</p>" +
      "<p><strong>Conditions:</strong> "+values[15]+"</p>" +
      "<div style='background:#fff3cd;padding:15px;border-radius:8px;'>" +
      "<h4 style='color:#856404;'>🚫 Restrictions:</h4>" +
      "<ul><li><strong>NO Working at Height</strong></li>" +
      "<li><strong>NO Crane/Heavy Equipment</strong></li>" +
      "<li><strong>NO Confined Space</strong></li>" +
      "<li><strong>Ground-level duties ONLY</strong></li></ul>" +
      "</div></div></div>";
    MailApp.sendEmail({to:EMAIL_ADDRESS,subject:"⚠️ RESTRICTION — "+values[2]+" — NO Heights",htmlBody:emailHtml});
  } catch(e) { Logger.log("Restriction email error: "+e); }
}

// ============================================
// 📋 HEADERS — 🆕 Briefing updated with Photo URL
// ============================================
function getHeadersForSheet(sheetName) {
  var headersMap = {
    'Incidents': ['ID','Date','Time','Type','Severity','Location','Reported By','Injured Person','Body Part','Description','Root Cause','Corrective Actions','Status','Timestamp','Photo URL'],
    'HIRA': ['ID','Activity','Location','Hazard','Likelihood','Severity','Risk Score','Controls','Res. Likelihood','Res. Severity','Residual Score','Assessor','Date','Timestamp'],
    'Inspections': ['ID','Date','Type','Location','Inspector','Findings','Non-Conformities','Actions','Priority','Due Date','Status','Timestamp'],
    'PowerTools': ['ID','Name','Voltage','Guard','Cable','ELCB','PPE','Last Inspection','Next Due','Inspector','Remarks','Status','Timestamp'],
    'FireEquipment': ['ID','Type','Capacity','Location','Install Date','Expiry Date','Last Inspection','Next Due','Pressure','Pin & Seal','Hose','Status','Inspector','Remarks','Timestamp'],
    'FireHydrants': ['ID','Type','Location','Date','Pressure Test','Hose Condition','Valve Condition','Accessibility','Obstruction','Nozzle','Signage','Status','Inspector','Remarks','Timestamp'],
    'FireAlarms': ['ID','Type','Zone/Location','Date','Functional Test','Condition','Panel Status','Battery Status','Battery Replacement Date','Status','Inspector','Remarks','Timestamp'],
    'FireDrills': ['ID','Date','Time','Type','Location','Occupants','Participants','Participation %','Evacuation Time','Assembly Point','Wardens','Alarm Type','Coordinator','Observations','Actions','Timestamp'],
    'EmergencyContacts': ['ID','Name','Role','Phone','Alt. Phone','Category','Priority','Timestamp'],
    'Training': ['ID','Date','Title','Type','Trainer','Duration','Attendees','Topics','Status','Timestamp','Photo URL'],
    'PPERecords': ['ID','Date','Employee Name','Employee ID','PPE Item','Action','Condition','Quantity','Remarks','Timestamp'],
    'PermitToWork': ['ID','Date','Valid Until','Type','Location','Requester','Contractor','Description','Controls','Authorizer','Status','Timestamp'],
    'CIDA': ['ID','Reg. No.','Contractor','Grade','Expiry','Project','Tech Officer','Safety Officer','Safety Name','CAR Insurance','WCI','TPL','Compliance','Non-Compliance Details','Timestamp'],
    'LegalRegister': ['ID','Legislation','Section','Category','Description','Status','Review Date','Evidence','Timestamp'],
    'EmployeeAttendance': ['ID','Date','Name','Employee ID','NIC','Designation','Department','Project/Site','Check-In','Check-Out','Work Hours','Overtime','Status','Remarks','Timestamp'],
    'SubContractorAttendance': ['ID','Date','Company','Worker Name','NIC','Trade','Project/Site','Check-In','Check-Out','Work Hours','Safety Induction','PPE Compliance','Supervisor','Status','Remarks','Timestamp'],
    'HealthInfo': ['ID','Person Type','Name','ID/NIC','Company','Designation','DOB','Age','Gender','Blood Group','Phone','Emergency Name','Emergency Relation','Emergency Phone','Emergency Address','Medical Conditions','Allergies','Medications','Previous Injuries','Last Medical Date','Fitness Cert Status','Cert Expiry','Doctor Name','Fit to Work','Restrictions','Restriction Details','Record Date','Remarks','Timestamp'],
    // 🆕 UPDATED: Added Photo URL column
    'Briefing': ['ID','Date','Time','Topic','Type','Conductor','Location','Duration (min)','Attendees','Attendee Count','Key Points','Actions','Status','Timestamp','Photo URL'],
    'ConnectionTest': ['Test Data','Timestamp','System Version']
  };
  return headersMap[sheetName] || [];
}

// ============================================
// 🔧 SETUP ALL SHEETS
// ============================================
function setupAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = [
    'Incidents','HIRA','Inspections','PowerTools',
    'FireEquipment','FireHydrants','FireAlarms','FireDrills',
    'EmergencyContacts','Training','PPERecords',
    'PermitToWork','CIDA','LegalRegister',
    'EmployeeAttendance','SubContractorAttendance','HealthInfo',
    'Briefing'
  ];
  
  var created=0, existed=0;
  sheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      var headers = getHeadersForSheet(name);
      if (headers.length > 0) {
        sheet.getRange(1,1,1,headers.length).setValues([headers]);
        sheet.getRange(1,1,1,headers.length)
          .setBackground('#0d1b2a').setFontColor('#ffffff')
          .setFontWeight('bold').setFontSize(10);
        sheet.setFrozenRows(1);
      }
      created++;
    } else { existed++; }
  });
  
  try {
    SpreadsheetApp.getUi().alert('✅ Setup Complete!\nCreated: '+created+'\nExisted: '+existed+'\nTotal: '+sheetNames.length);
  } catch(e) {
    Logger.log('Setup: Created '+created+', Existed '+existed);
  }
}
