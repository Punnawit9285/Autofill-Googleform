/**
 * Universal Dynamic Google Form Autofill & Redirect
 * Configured specifically for the "SYNCDATA" tab:
 * - Col D (index 3): title
 * - Col E (index 4): first_name
 * - Col F (index 5): last_name
 * - Full Name: Col E + " " + Col F
 * - Col G (index 6): nickname
 * - Col K (index 10): docchula_email
 * - Col L (index 11): phone_number (formatted to 0XX-XXX-XXXX)
 * - Col M (index 12): line_id
 * - Col N (index 13): line_display (Line Display Name)
 * - Student ID: Left completely BLANK (user manually fills in)
 */

function doGet(e) {
  try {
    var studentEmail = Session.getActiveUser().getEmail();

    // 1. Read requested form key from URL (e.g., ?form=event)
    var requestedFormKey = (e && e.parameter && e.parameter.form) 
      ? e.parameter.form.toString().toLowerCase().trim() 
      : "";

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 2. Read configuration from "Forms" tab
    var configSheet = ss.getSheetByName("Forms") 
                   || ss.getSheetByName("forms")
                   || ss.getSheetByName("Config");
    if (!configSheet) {
      return HtmlService.createHtmlOutput("<h3>Error: 'Forms' tab not found in spreadsheet.</h3>");
    }

    var configData = configSheet.getDataRange().getValues();
    var templateUrl = "";
    var blankFormUrl = "";

    // Search for matching form key (skip header row 0)
    for (var k = 1; k < configData.length; k++) {
      var key = (configData[k][0] || "").toString().toLowerCase().trim();
      
      // Match requested key (or default to the first configured form if no key specified)
      if (key === requestedFormKey || (requestedFormKey === "" && k === 1)) {
        templateUrl = (configData[k][1] || "").toString().trim();  // Column B
        blankFormUrl = (configData[k][2] || "").toString().trim(); // Column C
        break;
      }
    }

    if (!templateUrl) {
      return HtmlService.createHtmlOutput("<h3>Error: Form key '" + requestedFormKey + "' not found in Forms tab.</h3>");
    }

    // 3. Search for student in "SYNCDATA" tab (with fallback to "Sheet1")
    var studentSheet = ss.getSheetByName("SYNCDATA") 
                    || ss.getSheetByName("SyncData") 
                    || ss.getSheetByName("syncdata")
                    || ss.getSheetByName("Sheet1");

    if (!studentSheet) {
      return HtmlService.createHtmlOutput("<h3>Error: 'SYNCDATA' tab not found in spreadsheet.</h3>");
    }

    var studentData = studentSheet.getDataRange().getValues();

    if (studentEmail && studentData.length > 1) {
      var colMap = getSyncDataColumnMapping(studentData[0]);

      for (var i = 1; i < studentData.length; i++) {
        var rowEmail = (studentData[i][colMap.email] || "").toString().toLowerCase().trim();
        
        // Match email address exactly or match username prefix before '@'
        if (isEmailMatch(studentEmail, rowEmail)) {
          var studentObj = extractSyncData(studentData[i], colMap);
          var personalUrl = personalizeFormUrl(templateUrl, studentObj);
          
          return HtmlService.createHtmlOutput(
            "<script>window.location.href='" + personalUrl + "';</script>" +
            "<div style='font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif; text-align:center; margin-top:80px;'>" +
            "  <p style='font-size:18px; color:#333;'>กำลังนำท่านไปยังแบบฟอร์มเฉพาะบุคคล...</p>" +
            "  <p style='font-size:14px; color:#777;'>Loading your personalized form, please wait...</p>" +
            "  <p style='font-size:13px; color:#aaa;'>หากหน้าเว็บไม่เปลี่ยนอัตโนมัติ <a href='" + personalUrl + "'>คลิกที่นี่</a></p>" +
            "</div>"
          );
        }
      }
    }

    // FALLBACK: If email not detected or student not in SYNCDATA
    var fallback = blankFormUrl || templateUrl;
    return HtmlService.createHtmlOutput(
      "<script>window.location.href='" + fallback + "';</script>" +
      "<div style='font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif; text-align:center; margin-top:80px;'>" +
      "  <p style='font-size:18px; color:#333;'>กำลังนำท่านไปยังแบบฟอร์ม...</p>" +
      "  <p style='font-size:14px; color:#777;'>Redirecting to form...</p>" +
      "</div>"
    );

  } catch (err) {
    return HtmlService.createHtmlOutput("<h3>An error occurred: " + err.message + "</h3>");
  }
}

/**
 * Checks email match (exact or matching username prefix before '@')
 */
function isEmailMatch(userEmail, sheetEmail) {
  if (!userEmail || !sheetEmail) return false;
  var u = userEmail.toString().toLowerCase().trim();
  var s = sheetEmail.toString().toLowerCase().trim();

  if (u === s) return true;

  var uPrefix = u.split('@')[0];
  var sPrefix = s.split('@')[0];
  if (uPrefix && sPrefix && uPrefix === sPrefix) return true;

  return false;
}

/**
 * Maps SYNCDATA column indices:
 * - Col D: title (3)
 * - Col E: first_name (4)
 * - Col F: last_name (5)
 * - Col G: nickname (6)
 * - Col K: docchula_email (10)
 * - Col L: phone_number (11)
 * - Col M: line_id (12)
 * - Col N: line_display (13)
 */
function getSyncDataColumnMapping(headerRow) {
  var map = {
    title: 3,        // Col D
    firstName: 4,    // Col E
    lastName: 5,     // Col F
    nickname: 6,     // Col G
    email: 10,       // Col K
    tel: 11,         // Col L
    lineId: 12,      // Col M
    lineDisplay: 13  // Col N (line_display)
  };

  if (!headerRow || !headerRow.length) return map;

  for (var c = 0; c < headerRow.length; c++) {
    var h = (headerRow[c] || "").toString().trim().toLowerCase();
    if (!h) continue;

    if (/email/i.test(h) || h === "docchula_email") {
      map.email = c;
    } else if (/^(first_name|firstname|first name|ชื่อ)$/i.test(h)) {
      map.firstName = c;
    } else if (/^(last_name|lastname|last name|นามสกุล)$/i.test(h)) {
      map.lastName = c;
    } else if (/^(full_name|fullname|full name|ชื่อ-สกุล|ชื่อ - สกุล)$/i.test(h)) {
      map.fullName = c;
    } else if (/^(nickname|nick name|ชื่อเล่น)$/i.test(h)) {
      map.nickname = c;
    } else if (/^(phone_number|phone|telephone|tel|เบอร์โทรศัพท์|เบอร์โทร)$/i.test(h)) {
      map.tel = c;
    } else if (/^(line_id|lineid|line id)$/i.test(h)) {
      map.lineId = c;
    } else if (/^(line_display|linedisplay|line display|line_name|linename|line name)$/i.test(h)) {
      map.lineDisplay = c;
    } else if (/^(title|คำนำหน้า|คำนำหน้าชื่อ)$/i.test(h)) {
      map.title = c;
    } else if (/^(year|class|ชั้นปี)$/i.test(h)) {
      map.year = c;
    }
  }

  return map;
}

/**
 * Extracts student fields according to SYNCDATA specifications
 */
function extractSyncData(row, colMap) {
  var title = (colMap.title !== undefined && row[colMap.title] !== undefined) 
    ? row[colMap.title].toString().trim() : "";
  var firstName = (colMap.firstName !== undefined && row[colMap.firstName] !== undefined) 
    ? row[colMap.firstName].toString().trim() : "";
  var lastName = (colMap.lastName !== undefined && row[colMap.lastName] !== undefined) 
    ? row[colMap.lastName].toString().trim() : "";
  var nickname = (colMap.nickname !== undefined && row[colMap.nickname] !== undefined) 
    ? row[colMap.nickname].toString().trim() : "";
  var email = (colMap.email !== undefined && row[colMap.email] !== undefined) 
    ? row[colMap.email].toString().trim() : "";
  var rawTel = (colMap.tel !== undefined && row[colMap.tel] !== undefined) 
    ? row[colMap.tel].toString().trim() : "";
  var lineId = (colMap.lineId !== undefined && row[colMap.lineId] !== undefined) 
    ? row[colMap.lineId].toString().trim() : "";
  var lineDisplay = (colMap.lineDisplay !== undefined && row[colMap.lineDisplay] !== undefined) 
    ? row[colMap.lineDisplay].toString().trim() : "";
  var year = (colMap.year !== undefined && row[colMap.year] !== undefined) 
    ? row[colMap.year].toString().trim() : "";

  // Strip prefix title if someone typed title into first_name
  var cleanFirstName = firstName.replace(/^(นาย|นางสาว|นาง|น\.ส\.|ด\.ช\.|ด\.ญ\.|นายแพทย์|นพ\.|พญ\.)\s*/i, "");

  // Full Name: Column E + " " + Column F
  var fullName = (cleanFirstName + " " + lastName).trim();
  if (!fullName && colMap.fullName !== undefined && row[colMap.fullName]) {
    fullName = row[colMap.fullName].toString().replace(/^(นาย|นางสาว|นาง|น\.ส\.|ด\.ช\.|ด\.ญ\.|นายแพทย์|นพ\.|พญ\.)\s*/i, "").trim();
  }

  // Format telephone to 0XX-XXX-XXXX
  var digits = rawTel.replace(/\D/g, "");
  if (digits.length === 9 && digits.charAt(0) !== "0") {
    digits = "0" + digits; // restore leading 0
  }
  var telFormatted = digits;
  if (digits.length === 10) {
    telFormatted = digits.replace(/(\d{3})(\d{3})(\d{4})/, "$1-$2-$3");
  }

  return {
    email: email,
    title: title,
    firstName: cleanFirstName,
    lastName: lastName,
    fullName: fullName,
    nickname: nickname,
    tel: telFormatted,
    lineId: lineId,
    lineDisplay: lineDisplay,
    year: year
  };
}

/**
 * Replaces placeholders in template URL:
 * - Populates Full Name, Nickname, Line ID, Line Display, Tel, Title
 * - Strips Student ID parameter completely so student fills it in manually
 * - Deduplicates URL parameters
 */
function personalizeFormUrl(templateUrl, s) {
  if (!templateUrl) return "";

  var url = templateUrl;

  function replacePlaceholder(targetUrl, placeholder, value) {
    if (!value && value !== 0) return targetUrl;
    var encoded = encodeURIComponent(value);

    var escaped = placeholder.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&')
                             .replace(/\\\+/g, '(\\+|%20|\\s)')
                             .replace(/%20/g, '(\\+|%20|\\s)')
                             .replace(/\\s/g, '(\\+|%20|\\s)');

    var re = new RegExp(escaped, 'gi');
    return targetUrl.replace(re, encoded);
  }

  // 1. Full Name (Column E + Column F)
  url = replacePlaceholder(url, "DummyFullName", s.fullName);
  url = replacePlaceholder(url, "DummyFullname", s.fullName);
  url = replacePlaceholder(url, "DummyName DummyLastName", s.fullName);
  url = replacePlaceholder(url, "DummyName+DummyLastName", s.fullName);
  url = replacePlaceholder(url, "DummyName%20DummyLastName", s.fullName);

  var hasDummyLastName = /DummyLastName/i.test(url);
  if (!hasDummyLastName) {
    url = replacePlaceholder(url, "DummyName", s.fullName);
  } else {
    url = replacePlaceholder(url, "DummyName", s.firstName);
    url = replacePlaceholder(url, "DummyLastName", s.lastName);
  }

  // 2. Title (Column D)
  url = replacePlaceholder(url, "DummyTitle", s.title);
  url = replacePlaceholder(url, "DummyPrefix", s.title);

  // 3. Nickname (Column G)
  url = replacePlaceholder(url, "DummyNickname", s.nickname);
  url = replacePlaceholder(url, "DummyNickName", s.nickname);
  url = replacePlaceholder(url, "DummyNick", s.nickname);

  // 4. Line Display / Line Name (Column N) - replaced before Line ID to avoid substring collisions
  url = replacePlaceholder(url, "DummyLineDisplay", s.lineDisplay);
  url = replacePlaceholder(url, "DummyLineName", s.lineDisplay);
  url = replacePlaceholder(url, "DummyDisplay", s.lineDisplay);

  // 5. Line ID (Column M)
  url = replacePlaceholder(url, "DummyLineID", s.lineId);
  url = replacePlaceholder(url, "DummyLineId", s.lineId);
  url = replacePlaceholder(url, "DummyLine", s.lineId);

  // 6. Telephone (Column L, formatted to 0XX-XXX-XXXX)
  url = replacePlaceholder(url, "0XX-XXX-XXXX", s.tel);
  url = replacePlaceholder(url, "0xx-xxx-xxxx", s.tel);
  url = replacePlaceholder(url, "081-234-5678", s.tel);
  url = replacePlaceholder(url, "0812345678", s.tel);
  url = replacePlaceholder(url, "DummyTel", s.tel);
  url = replacePlaceholder(url, "DummyPhone", s.tel);

  // 7. Academic Year in 2569 (if available)
  if (s.year) {
    url = replacePlaceholder(url, "DummyAcademicYear", s.year);
    url = replacePlaceholder(url, "DummyYear", s.year);
    url = replacePlaceholder(url, "DummyClass", s.year);
  }

  // 8. Email (Column K)
  url = replacePlaceholder(url, "DummyEmail", s.email);

  // 9. Process URL Query Parameters:
  // - STRIP Student ID parameter completely so student fills it in manually
  // - DEDUPLICATE all remaining query parameters
  if (url.indexOf('?') !== -1) {
    var urlParts = url.split('?');
    var base = urlParts[0];
    var query = urlParts.slice(1).join('?');
    var params = query.split('&');
    var paramMap = {};
    var paramOrder = [];

    // All possible placeholders for student ID
    var studentIdPlaceholders = [
      "6xxxxxxx30", "dummystudentid", "dummystudent", "dummyid", "dummystudent_id"
    ];

    for (var p = 0; p < params.length; p++) {
      if (!params[p]) continue;
      var kv = params[p].split('=');
      var key = kv[0];
      var val = kv.slice(1).join('=');
      var decodedVal = decodeURIComponent(val || "").toLowerCase().trim();

      // If this parameter is for Student ID, omit it so the question remains completely blank
      var isStudentIdParam = false;
      for (var j = 0; j < studentIdPlaceholders.length; j++) {
        if (decodedVal === studentIdPlaceholders[j]) {
          isStudentIdParam = true;
          break;
        }
      }
      if (isStudentIdParam) continue;

      if (!paramMap.hasOwnProperty(key)) {
        paramOrder.push(key);
      }
      paramMap[key] = val;
    }

    url = base + '?' + paramOrder.map(function(k) {
      return k + '=' + paramMap[k];
    }).join('&');
  }

  return url;
}
