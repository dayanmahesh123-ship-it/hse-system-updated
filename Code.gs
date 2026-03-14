// ============================================
// HSE MANAGEMENT SYSTEM — Google Apps Script
// Developer: Mahesh | HSE Officer
// Company: Hayleys Fentons Limited
// Version: 4.0.0
// ============================================

// 📁 Drive Folder IDs
var INCIDENT_FOLDER_ID = "1EGJgsFc8TSEbAfQlaJ1zNRjPyoahGyWK";
var TRAINING_FOLDER_ID = "1Y8Pzufd4Yhl4X6nSyTP3Fu0MxQ9LxqPZ";
var BRIEFING_FOLDER_ID = "1nwRLiR9rw8VF56iBgaPxefJ5x93DFpR7";

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

    if (sheetName === 'Incidents') {
      var photoUrl = savePhotoToDrive(values[14], values[0], INCIDENT_FOLDER_ID);
      values[14] = photoUrl;
      sheet.appendRow(values);
      if (values[4] === "High" || values[4] === "Critical") {
        sendIncidentEmail(values, photoUrl);
      }
    }
    else if (sheetName === 'Training') {
      var photoUrl = savePhotoToDrive(values[10], values[0], TRAINING_FOLDER_ID);
      values[10] = photoUrl;
      sheet.appendRow(values);
    }
    else if (sheetName === 'Briefing') {
      var photoUrl = savePhotoToDrive(values[14], values[0], BRIEFING_FOLDER_ID);
      values[14] = photoUrl;
      sheet.appendRow(values);
      Logger.log("Briefing: " + values[3] + " — " + values[9] + " attendees, Photo: " + (photoUrl ? "Yes" : "No"));
    }
    else if (sheetName === 'EmployeeAttendance') {
      sheet.appendRow(values);
    }
    else if (sheetName === 'SubContractorAttendance') {
      sheet.appendRow(values);
      if (values[10] === 'No' || values[11] === 'No') {
        sendSCAlertEmail(values, values[10], values[11]);
      }
    }
    else if (sheetName === 'HealthInfo') {
      sheet.appendRow(values);
      if (values[23] === 'No' || values[20] === 'Expired') {
        sendHealthAlertEmail(values);
      }
      if (values[15] && (values[15].indexOf('Epilepsy') !== -1 || values[15].indexOf('Vertigo') !== -1)) {
        sendHealthRestrictionEmail(values);
      }
    }
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
// 📤 doGet — Dashboard + Data Endpoints
// ============================================
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';

    if (action === 'getPersonnel') return getPersonnelData();
    if (action === 'getCompanies') return getCompanyData();
    if (action === 'getProjects') return getProjectData();

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var response = {};

    var incSheet = ss.getSheetByName("Incidents");
    if (incSheet && incSheet.getLastRow() > 1) {
      response.lastDate = incSheet.getRange(incSheet.getLastRow(), 2).getValue();
    } else {
      response.lastDate = new Date().toISOString();
    }

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
// 🆕 SERVER-SIDE DATA FUNCTIONS
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
// 📋 HEADERS
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

// ============================================
// 🎨 AUTO FORMAT ALL SHEETS
// ============================================
function formatAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();

    if (lastCol === 0) return;

    var headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setBackground('#0d1b2a')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setFontSize(10)
               .setFontFamily('Arial')
               .setHorizontalAlignment('center')
               .setVerticalAlignment('middle')
               .setWrap(true);
    sheet.setRowHeight(1, 40);
    sheet.setFrozenRows(1);

    if (lastRow > 1) {
      var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
      dataRange.setFontSize(10)
               .setFontFamily('Arial')
               .setVerticalAlignment('middle')
               .setWrap(true);

      for (var i = 2; i <= lastRow; i++) {
        var rowRange = sheet.getRange(i, 1, 1, lastCol);
        if (i % 2 === 0) {
          rowRange.setBackground('#f8f9fa');
        } else {
          rowRange.setBackground('#ffffff');
        }
      }
    }

    if (lastRow > 0) {
      var allRange = sheet.getRange(1, 1, lastRow, lastCol);
      allRange.setBorder(true, true, true, true, true, true, '#dee2e6', SpreadsheetApp.BorderStyle.SOLID);
    }

    var columnWidths = {
      'Incidents': [80,100,80,140,90,120,120,120,120,250,200,200,100,150,200],
      'Training': [80,100,180,140,120,80,80,250,90,150,200],
      'Briefing': [80,100,80,180,120,120,120,80,200,80,250,200,90,150,200],
      'EmployeeAttendance': [80,100,150,100,120,120,100,120,80,80,80,60,80,200,150],
      'SubContractorAttendance': [80,100,140,140,120,100,120,80,80,80,100,100,120,80,200,150],
      'HealthInfo': [80,100,150,120,140,120,100,50,70,70,100,140,100,100,200,180,150,150,200,100,100,100,120,80,120,200,100,200,150],
      'HIRA': [80,150,120,250,50,50,80,250,50,50,80,100,100,150],
      'Inspections': [80,100,180,120,120,250,200,200,80,100,80,150],
      'PowerTools': [80,140,70,80,80,60,60,100,100,100,200,70,150],
      'FireEquipment': [80,120,80,140,100,100,100,100,100,80,80,100,100,200,150],
      'FireHydrants': [80,100,120,100,80,80,100,100,80,80,80,100,100,200,150],
      'FireAlarms': [80,120,120,100,80,100,100,80,100,100,100,200,150],
      'FireDrills': [80,100,80,120,120,80,80,80,100,100,100,80,100,250,200,150],
      'EmergencyContacts': [80,150,150,120,120,100,80,150],
      'PPERecords': [80,100,150,100,160,80,80,60,200,150],
      'PermitToWork': [80,100,100,120,120,120,140,250,250,120,80,150],
      'CIDA': [80,120,180,60,100,120,120,120,140,80,80,80,100,200,150],
      'LegalRegister': [80,150,100,100,250,100,100,250,150]
    };

    if (columnWidths[name]) {
      columnWidths[name].forEach(function(w, idx) {
        if (idx < lastCol) sheet.setColumnWidth(idx + 1, w);
      });
    } else {
      for (var c = 1; c <= lastCol; c++) {
        sheet.setColumnWidth(c, 120);
      }
      if (lastCol >= 1) sheet.setColumnWidth(1, 80);
    }

    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 1)
           .setFontWeight('bold')
           .setHorizontalAlignment('center');
    }

    var dateColumns = {
      'Incidents': [2], 'Training': [2], 'Briefing': [2],
      'EmployeeAttendance': [2], 'SubContractorAttendance': [2],
      'HealthInfo': [7, 20, 22, 27], 'HIRA': [13],
      'Inspections': [2, 10], 'PowerTools': [8, 9],
      'FireEquipment': [5, 6, 7, 8], 'FireHydrants': [4],
      'FireAlarms': [4, 9], 'FireDrills': [2],
      'PPERecords': [2], 'PermitToWork': [2, 3],
      'CIDA': [5], 'LegalRegister': [7]
    };

    if (dateColumns[name] && lastRow > 1) {
      dateColumns[name].forEach(function(col) {
        if (col <= lastCol) {
          sheet.getRange(2, col, lastRow - 1, 1)
               .setHorizontalAlignment('center')
               .setNumberFormat('yyyy-mm-dd');
        }
      });
    }

    Logger.log('✅ Formatted: ' + name);
  });

  var tabColors = {
    'Incidents': '#e63946', 'HIRA': '#f77f00', 'Inspections': '#457b9d',
    'PowerTools': '#6c757d', 'FireEquipment': '#e63946', 'FireHydrants': '#e76f51',
    'FireAlarms': '#f4a261', 'FireDrills': '#e63946', 'EmergencyContacts': '#e63946',
    'Training': '#2a9d8f', 'Briefing': '#457b9d', 'PPERecords': '#2a9d8f',
    'PermitToWork': '#f77f00', 'CIDA': '#2a9d8f', 'LegalRegister': '#457b9d',
    'EmployeeAttendance': '#1b263b', 'SubContractorAttendance': '#f77f00',
    'HealthInfo': '#e63946'
  };

  Object.keys(tabColors).forEach(function(sn) {
    var s = ss.getSheetByName(sn);
    if (s) s.setTabColor(tabColors[sn]);
  });

  SpreadsheetApp.getUi().alert(
    '🎨 Formatting Complete!\n\n' +
    '✅ Headers styled\n✅ Zebra stripes added\n✅ Borders added\n' +
    '✅ Column widths set\n✅ Date formats applied\n✅ Tab colors set\n\n' +
    'Now run: addConditionalFormatting()'
  );
}

// ============================================
// 🎨 CONDITIONAL FORMATTING
// ============================================
function addConditionalFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var incSheet = ss.getSheetByName('Incidents');
  if (incSheet) {
    clearConditionalFormatting(incSheet);
    var lastRow = Math.max(incSheet.getLastRow(), 100);
    addStatusRule(incSheet, 'E2:E' + lastRow, 'Critical', '#f8d7da', '#721c24');
    addStatusRule(incSheet, 'E2:E' + lastRow, 'High', '#fff3cd', '#856404');
    addStatusRule(incSheet, 'E2:E' + lastRow, 'Medium', '#fff3cd', '#856404');
    addStatusRule(incSheet, 'E2:E' + lastRow, 'Low', '#d4edda', '#155724');
    addStatusRule(incSheet, 'M2:M' + lastRow, 'Open', '#f8d7da', '#721c24');
    addStatusRule(incSheet, 'M2:M' + lastRow, 'Closed', '#d4edda', '#155724');
    addStatusRule(incSheet, 'M2:M' + lastRow, 'Under Investigation', '#fff3cd', '#856404');
  }

  var ptSheet = ss.getSheetByName('PowerTools');
  if (ptSheet) {
    clearConditionalFormatting(ptSheet);
    var lr = Math.max(ptSheet.getLastRow(), 100);
    addStatusRule(ptSheet, 'L2:L' + lr, 'SAFE', '#d4edda', '#155724');
    addStatusRule(ptSheet, 'L2:L' + lr, 'UNSAFE', '#f8d7da', '#721c24');
  }

  var feSheet = ss.getSheetByName('FireEquipment');
  if (feSheet) {
    clearConditionalFormatting(feSheet);
    var lr2 = Math.max(feSheet.getLastRow(), 100);
    addStatusRule(feSheet, 'L2:L' + lr2, 'Serviceable', '#d4edda', '#155724');
    addStatusRule(feSheet, 'L2:L' + lr2, 'Expired', '#f8d7da', '#721c24');
    addStatusRule(feSheet, 'L2:L' + lr2, 'Recharge Required', '#fff3cd', '#856404');
  }

  var eaSheet = ss.getSheetByName('EmployeeAttendance');
  if (eaSheet) {
    clearConditionalFormatting(eaSheet);
    var lr3 = Math.max(eaSheet.getLastRow(), 100);
    addStatusRule(eaSheet, 'M2:M' + lr3, 'Present', '#d4edda', '#155724');
    addStatusRule(eaSheet, 'M2:M' + lr3, 'Absent', '#f8d7da', '#721c24');
    addStatusRule(eaSheet, 'M2:M' + lr3, 'Late', '#fff3cd', '#856404');
    addStatusRule(eaSheet, 'M2:M' + lr3, 'Leave', '#cce5ff', '#004085');
    addStatusRule(eaSheet, 'M2:M' + lr3, 'Sick Leave', '#fff3cd', '#856404');
  }

  var scSheet = ss.getSheetByName('SubContractorAttendance');
  if (scSheet) {
    clearConditionalFormatting(scSheet);
    var lr4 = Math.max(scSheet.getLastRow(), 100);
    addStatusRule(scSheet, 'K2:K' + lr4, 'Yes', '#d4edda', '#155724');
    addStatusRule(scSheet, 'K2:K' + lr4, 'No', '#f8d7da', '#721c24');
    addStatusRule(scSheet, 'L2:L' + lr4, 'Yes', '#d4edda', '#155724');
    addStatusRule(scSheet, 'L2:L' + lr4, 'No', '#f8d7da', '#721c24');
    addStatusRule(scSheet, 'L2:L' + lr4, 'Partial', '#fff3cd', '#856404');
    addStatusRule(scSheet, 'N2:N' + lr4, 'Present', '#d4edda', '#155724');
    addStatusRule(scSheet, 'N2:N' + lr4, 'Absent', '#f8d7da', '#721c24');
  }

  var hiSheet = ss.getSheetByName('HealthInfo');
  if (hiSheet) {
    clearConditionalFormatting(hiSheet);
    var lr5 = Math.max(hiSheet.getLastRow(), 100);
    addStatusRule(hiSheet, 'U2:U' + lr5, 'Valid', '#d4edda', '#155724');
    addStatusRule(hiSheet, 'U2:U' + lr5, 'Expired', '#f8d7da', '#721c24');
    addStatusRule(hiSheet, 'U2:U' + lr5, 'Pending', '#fff3cd', '#856404');
    addStatusRule(hiSheet, 'X2:X' + lr5, 'Yes', '#d4edda', '#155724');
    addStatusRule(hiSheet, 'X2:X' + lr5, 'No', '#f8d7da', '#721c24');
    addStatusRule(hiSheet, 'X2:X' + lr5, 'Conditional', '#fff3cd', '#856404');
  }

  var brSheet = ss.getSheetByName('Briefing');
  if (brSheet) {
    clearConditionalFormatting(brSheet);
    var lr6 = Math.max(brSheet.getLastRow(), 100);
    addStatusRule(brSheet, 'M2:M' + lr6, 'Completed', '#d4edda', '#155724');
    addStatusRule(brSheet, 'M2:M' + lr6, 'Scheduled', '#cce5ff', '#004085');
    addStatusRule(brSheet, 'M2:M' + lr6, 'Cancelled', '#f8d7da', '#721c24');
  }

  var trSheet = ss.getSheetByName('Training');
  if (trSheet) {
    clearConditionalFormatting(trSheet);
    var lr7 = Math.max(trSheet.getLastRow(), 100);
    addStatusRule(trSheet, 'I2:I' + lr7, 'Completed', '#d4edda', '#155724');
    addStatusRule(trSheet, 'I2:I' + lr7, 'Scheduled', '#cce5ff', '#004085');
    addStatusRule(trSheet, 'I2:I' + lr7, 'Cancelled', '#f8d7da', '#721c24');
  }

  var insSheet = ss.getSheetByName('Inspections');
  if (insSheet) {
    clearConditionalFormatting(insSheet);
    var lr8 = Math.max(insSheet.getLastRow(), 100);
    addStatusRule(insSheet, 'I2:I' + lr8, 'Critical', '#f8d7da', '#721c24');
    addStatusRule(insSheet, 'I2:I' + lr8, 'High', '#fff3cd', '#856404');
    addStatusRule(insSheet, 'K2:K' + lr8, 'Open', '#f8d7da', '#721c24');
    addStatusRule(insSheet, 'K2:K' + lr8, 'Closed', '#d4edda', '#155724');
  }

  var ptwSheet = ss.getSheetByName('PermitToWork');
  if (ptwSheet) {
    clearConditionalFormatting(ptwSheet);
    var lr9 = Math.max(ptwSheet.getLastRow(), 100);
    addStatusRule(ptwSheet, 'K2:K' + lr9, 'Active', '#fff3cd', '#856404');
    addStatusRule(ptwSheet, 'K2:K' + lr9, 'Completed', '#d4edda', '#155724');
    addStatusRule(ptwSheet, 'K2:K' + lr9, 'Cancelled', '#f8d7da', '#721c24');
    addStatusRule(ptwSheet, 'K2:K' + lr9, 'Suspended', '#f8d7da', '#721c24');
  }

  var cidaSheet = ss.getSheetByName('CIDA');
  if (cidaSheet) {
    clearConditionalFormatting(cidaSheet);
    var lr10 = Math.max(cidaSheet.getLastRow(), 100);
    addStatusRule(cidaSheet, 'M2:M' + lr10, 'Compliant', '#d4edda', '#155724');
    addStatusRule(cidaSheet, 'M2:M' + lr10, 'Non-Compliant', '#f8d7da', '#721c24');
    addStatusRule(cidaSheet, 'M2:M' + lr10, 'Partially Compliant', '#fff3cd', '#856404');
  }

  var legSheet = ss.getSheetByName('LegalRegister');
  if (legSheet) {
    clearConditionalFormatting(legSheet);
    var lr11 = Math.max(legSheet.getLastRow(), 100);
    addStatusRule(legSheet, 'F2:F' + lr11, 'Compliant', '#d4edda', '#155724');
    addStatusRule(legSheet, 'F2:F' + lr11, 'Non-Compliant', '#f8d7da', '#721c24');
    addStatusRule(legSheet, 'F2:F' + lr11, 'Partially Compliant', '#fff3cd', '#856404');
  }

  SpreadsheetApp.getUi().alert(
    '🎨 Conditional Formatting Complete!\n\n' +
    '✅ All status columns colored\n\n' +
    'Now run: addDataValidation()'
  );
}

function addStatusRule(sheet, range, value, bgColor, fontColor) {
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(value)
    .setBackground(bgColor)
    .setFontColor(fontColor)
    .setBold(true)
    .setRanges([sheet.getRange(range)])
    .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

function clearConditionalFormatting(sheet) {
  sheet.setConditionalFormatRules([]);
}

// ============================================
// 📋 DATA VALIDATION
// ============================================
function addDataValidation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var incSheet = ss.getSheetByName('Incidents');
  if (incSheet) {
    var lr = Math.max(incSheet.getLastRow(), 100);
    setDropdown(incSheet, 'E2:E' + lr, ['Low','Medium','High','Critical']);
    setDropdown(incSheet, 'M2:M' + lr, ['Open','Under Investigation','Corrective Action Pending','Closed']);
  }

  var ptSheet = ss.getSheetByName('PowerTools');
  if (ptSheet) {
    var lr2 = Math.max(ptSheet.getLastRow(), 100);
    setDropdown(ptSheet, 'D2:D' + lr2, ['Good','Fair','Damaged','Missing']);
    setDropdown(ptSheet, 'E2:E' + lr2, ['Good','Fair','Damaged','Exposed Wires']);
    setDropdown(ptSheet, 'F2:F' + lr2, ['Yes','No']);
    setDropdown(ptSheet, 'L2:L' + lr2, ['SAFE','UNSAFE']);
  }

  var feSheet = ss.getSheetByName('FireEquipment');
  if (feSheet) {
    var lr3 = Math.max(feSheet.getLastRow(), 100);
    setDropdown(feSheet, 'L2:L' + lr3, ['Serviceable','Recharge Required','Expired','Out of Service']);
  }

  var eaSheet = ss.getSheetByName('EmployeeAttendance');
  if (eaSheet) {
    var lr4 = Math.max(eaSheet.getLastRow(), 100);
    setDropdown(eaSheet, 'M2:M' + lr4, ['Present','Absent','Half Day','Late','Leave','Sick Leave','Holiday']);
  }

  var scSheet = ss.getSheetByName('SubContractorAttendance');
  if (scSheet) {
    var lr5 = Math.max(scSheet.getLastRow(), 100);
    setDropdown(scSheet, 'K2:K' + lr5, ['Yes','No','Pending']);
    setDropdown(scSheet, 'L2:L' + lr5, ['Yes','No','Partial']);
    setDropdown(scSheet, 'N2:N' + lr5, ['Present','Absent','Half Day']);
  }

  var hiSheet = ss.getSheetByName('HealthInfo');
  if (hiSheet) {
    var lr6 = Math.max(hiSheet.getLastRow(), 100);
    setDropdown(hiSheet, 'B2:B' + lr6, ['Employee','Sub-Contractor','Visitor']);
    setDropdown(hiSheet, 'I2:I' + lr6, ['Male','Female','Other']);
    setDropdown(hiSheet, 'J2:J' + lr6, ['A+','A-','B+','B-','AB+','AB-','O+','O-','Unknown']);
    setDropdown(hiSheet, 'U2:U' + lr6, ['Valid','Expired','Pending','Not Required']);
    setDropdown(hiSheet, 'X2:X' + lr6, ['Yes','No','Conditional']);
  }

  var brSheet = ss.getSheetByName('Briefing');
  if (brSheet) {
    var lr7 = Math.max(brSheet.getLastRow(), 100);
    setDropdown(brSheet, 'E2:E' + lr7, ['Toolbox Talk','Safety Briefing','Morning Brief','Pre-Task Brief','Weekly Safety Meeting','Emergency Brief','Method Statement Brief','Hazard Alert']);
    setDropdown(brSheet, 'M2:M' + lr7, ['Completed','Scheduled','Cancelled']);
  }

  var trSheet = ss.getSheetByName('Training');
  if (trSheet) {
    var lr8 = Math.max(trSheet.getLastRow(), 100);
    setDropdown(trSheet, 'I2:I' + lr8, ['Completed','Scheduled','Cancelled']);
  }

  var insSheet = ss.getSheetByName('Inspections');
  if (insSheet) {
    var lr9 = Math.max(insSheet.getLastRow(), 100);
    setDropdown(insSheet, 'I2:I' + lr9, ['Low','Medium','High','Critical']);
    setDropdown(insSheet, 'K2:K' + lr9, ['Open','In Progress','Closed']);
  }

  var ptwSheet = ss.getSheetByName('PermitToWork');
  if (ptwSheet) {
    var lr10 = Math.max(ptwSheet.getLastRow(), 100);
    setDropdown(ptwSheet, 'D2:D' + lr10, ['Hot Work','Working at Height','Confined Space','Excavation','Electrical','Lifting','Demolition','Cold Work','Radiography']);
    setDropdown(ptwSheet, 'K2:K' + lr10, ['Active','Completed','Cancelled','Suspended']);
  }

  var cidaSheet = ss.getSheetByName('CIDA');
  if (cidaSheet) {
    var lr11 = Math.max(cidaSheet.getLastRow(), 100);
    setDropdown(cidaSheet, 'M2:M' + lr11, ['Compliant','Non-Compliant','Partially Compliant','Under Review']);
  }

  var legSheet = ss.getSheetByName('LegalRegister');
  if (legSheet) {
    var lr12 = Math.max(legSheet.getLastRow(), 100);
    setDropdown(legSheet, 'F2:F' + lr12, ['Compliant','Non-Compliant','Partially Compliant','Under Review']);
  }

  SpreadsheetApp.getUi().alert(
    '📋 Data Validation Complete!\n\n' +
    '✅ All dropdown validations added\n\n' +
    'Now run: createDashboardSheet()'
  );
}

function setDropdown(sheet, range, values) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(range).setDataValidation(rule);
}

// ============================================
// 📊 DASHBOARD SHEET
// ============================================
function createDashboardSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName('📊 Dashboard');
  if (dash) ss.deleteSheet(dash);

  dash = ss.insertSheet('📊 Dashboard', 0);

  dash.getRange('A1:F1').merge()
    .setValue('🛡️ HSE MANAGEMENT SYSTEM — DASHBOARD')
    .setBackground('#0d1b2a').setFontColor('#ffffff')
    .setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center');
  dash.setRowHeight(1, 50);

  dash.getRange('A2:F2').merge()
    .setValue('Hayleys Fentons Limited — ' + new Date().toLocaleDateString())
    .setBackground('#1b263b').setFontColor('#ffffff')
    .setFontSize(11).setHorizontalAlignment('center');
  dash.setRowHeight(2, 30);

  var sections = [
    [4, '📊 KEY METRICS', '#457b9d'],
    [12, '🚨 SAFETY ALERTS', '#e63946'],
    [20, '📢 BRIEFING & TRAINING', '#2a9d8f'],
    [26, '👥 ATTENDANCE', '#1b263b'],
    [32, '📋 COMPLIANCE', '#f77f00']
  ];

  sections.forEach(function(s) {
    dash.getRange(s[0], 1, 1, 6).merge()
      .setValue(s[1]).setBackground(s[2])
      .setFontColor('#ffffff').setFontWeight('bold')
      .setFontSize(12).setHorizontalAlignment('center');
    dash.setRowHeight(s[0], 35);
  });

  var metrics = [
    [5, 'A', 'Total Incidents', "=COUNTA(Incidents!A:A)-1"],
    [5, 'C', 'Open Incidents', '=COUNTIF(Incidents!M:M,"Open")'],
    [5, 'E', 'Near Misses', '=COUNTIF(Incidents!D:D,"Near Miss")'],
    [6, 'A', 'Fire Equipment', "=COUNTA(FireEquipment!A:A)-1"],
    [6, 'C', 'Expired Equipment', '=COUNTIF(FireEquipment!L:L,"Expired")'],
    [6, 'E', 'Unsafe Tools', '=COUNTIF(PowerTools!L:L,"UNSAFE")'],
    [7, 'A', 'Total HIRA', "=COUNTA(HIRA!A:A)-1"],
    [7, 'C', 'Inspections', "=COUNTA(Inspections!A:A)-1"],
    [7, 'E', 'Emergency Contacts', "=COUNTA(EmergencyContacts!A:A)-1"],
    [13, 'A', 'Critical Incidents', '=COUNTIF(Incidents!E:E,"Critical")'],
    [13, 'C', 'High Incidents', '=COUNTIF(Incidents!E:E,"High")'],
    [13, 'E', 'SC No Induction', '=COUNTIF(SubContractorAttendance!K:K,"No")'],
    [14, 'A', 'SC No PPE', '=COUNTIF(SubContractorAttendance!L:L,"No")'],
    [14, 'C', 'Unfit Workers', '=COUNTIF(HealthInfo!X:X,"No")'],
    [14, 'E', 'Expired Certs', '=COUNTIF(HealthInfo!U:U,"Expired")'],
    [21, 'A', 'Total Briefings', "=COUNTA(Briefing!A:A)-1"],
    [21, 'C', 'Total Training', "=COUNTA(Training!A:A)-1"],
    [21, 'E', 'Avg Attendees', '=IF(COUNTA(Briefing!A:A)>1,ROUND(AVERAGE(Briefing!J2:J),0),0)'],
    [27, 'A', 'Employee Records', "=COUNTA(EmployeeAttendance!A:A)-1"],
    [27, 'C', 'SC Records', "=COUNTA(SubContractorAttendance!A:A)-1"],
    [27, 'E', 'Health Records', "=COUNTA(HealthInfo!A:A)-1"],
    [33, 'A', 'CIDA Contractors', "=COUNTA(CIDA!A:A)-1"],
    [33, 'C', 'Legal Requirements', "=COUNTA(LegalRegister!A:A)-1"],
    [33, 'E', 'Compliant Legal', '=COUNTIF(LegalRegister!F:F,"Compliant")']
  ];

  metrics.forEach(function(m) {
    var row = m[0], col = m[1], label = m[2], formula = m[3];
    var colNum = col === 'A' ? 1 : col === 'C' ? 3 : 5;

    dash.getRange(row, colNum).setValue(label)
      .setFontWeight('bold').setFontSize(10)
      .setBackground('#f8f9fa');
    dash.getRange(row, colNum + 1).setFormula(formula)
      .setFontWeight('bold').setFontSize(14)
      .setHorizontalAlignment('center')
      .setBackground('#ffffff');
  });

  [180, 80, 180, 80, 180, 80].forEach(function(w, i) {
    dash.setColumnWidth(i + 1, w);
  });

  dash.getRange(1, 1, 35, 6).setBorder(true, true, true, true, true, true,
    '#dee2e6', SpreadsheetApp.BorderStyle.SOLID);

  dash.setTabColor('#0d1b2a');

  SpreadsheetApp.getUi().alert(
    '📊 Dashboard Sheet Created!\n\n' +
    '✅ All metrics with live formulas\n' +
    'All values auto-update!'
  );
}

// ============================================
// 🏆 RUN ALL AT ONCE
// ============================================
function makeSheetsProfessional() {
  formatAllSheets();
  Utilities.sleep(2000);
  addConditionalFormatting();
  Utilities.sleep(2000);
  addDataValidation();
  Utilities.sleep(2000);
  createDashboardSheet();

  SpreadsheetApp.getUi().alert(
    '🏆 ALL DONE! Sheets are PROFESSIONAL!\n\n' +
    '✅ Headers & formatting\n' +
    '✅ Conditional colors\n' +
    '✅ Data validation\n' +
    '✅ Dashboard created\n\n' +
    'v4.0.0 | Mahesh | Hayleys Fentons'
  );
}
