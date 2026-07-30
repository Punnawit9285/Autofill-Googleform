function doGet(e) {
  var studentEmail = Session.getActiveUser().getEmail();

  // 1. Read requested form key from URL (e.g., ?form=event)
  var requestedFormKey = (e && e.parameter && e.parameter.form) ? e.parameter.form.toLowerCase().trim() : "";

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 2. Read configuration from "Forms" tab
  var configSheet = ss.getSheetByName("Forms");
  if (!configSheet) {
    return HtmlService.createHtmlOutput("<h3>Error: 'Forms' tab not found in spreadsheet.</h3>");
  }

  var configData = configSheet.getDataRange().getValues();
  var templateUrl = "";
  var blankFormUrl = "";

  // Search for matching form key (skip header row 0)
  for (var k = 1; k < configData.length; k++) {
    var key = configData[k][0].toString().toLowerCase().trim();
    
    // Match requested key (or default to the first form if no key specified)
    if (key === requestedFormKey || (requestedFormKey === "" && k === 1)) {
      templateUrl = configData[k][1];  // Column B
      blankFormUrl = configData[k][2]; // Column C
      break;
    }
  }

  if (!templateUrl) {
    return HtmlService.createHtmlOutput("<h3>Error: Form key '" + requestedFormKey + "' not found in Forms tab.</h3>");
  }

  // 3. Search for student in "Sheet1" tab if an email is detected
  if (studentEmail) {
    var studentSheet = ss.getSheetByName("Sheet1");
    var studentData = studentSheet.getDataRange().getValues();

    for (var i = 1; i < studentData.length; i++) {
      var rowEmail = studentData[i][0]; // Column A: email
      
      if (rowEmail && rowEmail.toString().toLowerCase().trim() === studentEmail.toLowerCase().trim()) {
        var title    = studentData[i][1]; // Column B: title
        var name     = studentData[i][2]; // Column C: name
        var lastname = studentData[i][3]; // Column D: lastname
        var nickname = studentData[i][4]; // Column E: nickname
        var tel      = studentData[i][5]; // Column F: tel. number
        
        // Personalize the template URL dynamically
        var personalUrl = templateUrl
          .replace("DummyTitle", encodeURIComponent(title))
          .replace("DummyName", encodeURIComponent(name))
          .replace("DummyLastName", encodeURIComponent(lastname))
          .replace("DummyNickname", encodeURIComponent(nickname))
          .replace("0812345678", encodeURIComponent(tel));
        
        return HtmlService.createHtmlOutput("<script>window.location.href='" + personalUrl + "';</script><p style='font-family:sans-serif; text-align:center; margin-top:50px;'>Loading your personalized form...</p>");
      }
    }
  }

  // FALLBACK: If email isn't detected, non-docchula email, or isn't in Sheet1 -> redirect to blank form
  return HtmlService.createHtmlOutput("<script>window.location.href='" + blankFormUrl + "';</script><p style='font-family:sans-serif; text-align:center; margin-top:50px;'>Loading form...</p>");
}
