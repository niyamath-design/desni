function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Responses");
    
    if (!sheet) {
      sheet = ss.insertSheet("Responses");
      var headers = ["Timestamp", "Name", "Email", "Phone", "Service", "Message"];
      sheet.appendRow(headers);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
    } else if (sheet.getLastRow() === 0) {
      var headers = ["Timestamp", "Name", "Email", "Phone", "Service", "Message"];
      sheet.appendRow(headers);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
    }
    
    var params = (e && e.parameter) ? e.parameter : {};
    var action = params.action || "";
    var callback = params.callback || "";

    // Action 1: Return stored responses for Admin Dashboard
    if (action === "getResponses") {
      var lastRow = sheet.getLastRow();
      var responses = [];
      if (lastRow > 1) {
        var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
        for (var i = data.length - 1; i >= 0; i--) {
          responses.push({
            id: "sheet-lead-" + (i + 1),
            timestamp: data[i][0] ? data[i][0].toString() : "",
            name: data[i][1] || "",
            email: data[i][2] || "",
            phone: data[i][3] || "",
            service: data[i][4] || "",
            message: data[i][5] || "",
            status: "New"
          });
        }
      }

      var resultJson = JSON.stringify({ success: true, data: responses });
      if (callback) {
        return ContentService
          .createTextOutput(callback + "(" + resultJson + ")")
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService
        .createTextOutput(resultJson)
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Action 2: Append new client lead submission
    var name = params.name || "";
    var email = params.email || "";
    var phone = params.phone || "";
    var service = params.service || "";
    var message = params.message || "";
    
    if (name || email || service) {
      var timestamp = new Date();
      sheet.appendRow([timestamp, name, email, phone, service, message]);
    }
    
    var output = JSON.stringify({ success: true });
    if (callback) {
      return ContentService
        .createTextOutput(callback + "(" + output + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService
      .createTextOutput(output)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    var errOutput = JSON.stringify({ success: false, error: error.toString() });
    return ContentService
      .createTextOutput(errOutput)
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
