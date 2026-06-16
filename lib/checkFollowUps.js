/**
 * Automatically checks lead follow-up reminders in "LeadsTable" named range
 * (or falls back to Sheet1 data) and sends email notifications for leads requiring attention.
 */
function checkFollowUps() {
  // Ensure the daily trigger is set up
  setupTrigger();
  
  // Get active spreadsheet and the sheet tab "Sheet1"
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet1");
  if (!sheet) {
    Logger.log("Sheet 'Sheet1' not found.");
    return;
  }
  
  // Get named range "LeadsTable", fall back to entire sheet if not found
  var namedRange = ss.getRangeByName("LeadsTable");
  var values, startRow;
  
  if (!namedRange) {
    Logger.log("Named range 'LeadsTable' not found. Falling back to reading Sheet1 data range directly.");
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("No lead data found on Sheet1 (headers on row 1, data starts row 2).");
      return;
    }
    // Get columns A to F (1 to 6) from row 2 to the end
    values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    startRow = 2;
  } else {
    values = namedRange.getValues();
    startRow = namedRange.getRow();
  }
  
  var flaggedLeads = [];
  
  for (var i = 0; i < values.length; i++) {
    var currentRowNum = startRow + i;
    // Skip header row if the range starts at row 1
    if (currentRowNum < 2) {
      continue;
    }
    
    var rowData = values[i];
    var leadName = rowData[0];
    var contactInfo = rowData[1];
    var lastContactedVal = rowData[4];
    var status = rowData[5];
    
    // Skip if lead name is empty (blank rows at the bottom)
    if (!leadName || leadName.toString().trim() === "") {
      continue;
    }
    
    // If status is "Dead" or "Closed" -> skip, do nothing
    if (status === "Dead" || status === "Closed") {
      // Clear highlight just in case it was previously yellow
      sheet.getRange(currentRowNum, 1, 1, 6).setBackground(null);
      continue;
    }
    
    var isFlagged = false;
    var daysSinceContact = null;
    
    // Check if column E (Last Contacted) is empty
    if (!lastContactedVal || lastContactedVal.toString().trim() === "") {
      isFlagged = true;
    } else {
      var lastContactedDate = (lastContactedVal instanceof Date) ? lastContactedVal : new Date(lastContactedVal);
      if (isNaN(lastContactedDate.getTime())) {
        // Handle invalid date format as empty/never contacted
        isFlagged = true;
      } else {
        var today = new Date();
        var diffTime = today.getTime() - lastContactedDate.getTime();
        var diffDays = diffTime / (1000 * 60 * 60 * 24);
        if (diffDays > 3) {
          isFlagged = true;
          daysSinceContact = Math.floor(diffDays);
        }
      }
    }
    
    if (isFlagged) {
      var daysText = (daysSinceContact !== null) ? daysSinceContact + " days" : "Never contacted";
      flaggedLeads.push("- " + leadName + " (" + contactInfo + "): " + daysText);
      // Highlight the entire row (columns A-F) yellow
      sheet.getRange(currentRowNum, 1, 1, 6).setBackground("yellow");
    } else {
      // Clear highlight if not flagged
      sheet.getRange(currentRowNum, 1, 1, 6).setBackground(null);
    }
  }
  
  // Send email if there are flagged leads
  if (flaggedLeads.length > 0) {
    var emailBody = "The following leads require follow-up attention:\n\n" + 
                    flaggedLeads.join("\n") + 
                    "\n\nThis is an automated reminder from your Google Sheets Lead tracker.";
    MailApp.sendEmail("abdul0280ahad@gmail.com", "Follow-up Reminder: Leads Needing Attention", emailBody);
  }
}

/**
 * Sets up a daily time-based trigger to run checkFollowUps at 9 AM
 * if it doesn't already exist.
 */
function setupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var triggerExists = false;
  
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "checkFollowUps") {
      triggerExists = true;
      break;
    }
  }
  
  if (!triggerExists) {
    ScriptApp.newTrigger("checkFollowUps")
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();
    Logger.log("Daily trigger created for 9 AM.");
  } else {
    Logger.log("Trigger already exists.");
  }
}
