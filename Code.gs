/**
 * Universal Dynamic Google Form Autofill & Redirect
 * ------------------------------------------------
 * Reads a roster tab (default "SYNCDATA") and a form registry tab (default
 * "Forms"), then redirects the signed-in visitor to a Google Form that has
 * been pre-filled with their own data.
 *
 * Usage:  .../exec?form=<FORM_KEY>
 *         .../exec?form=<FORM_KEY>&debug=1     <- diagnostics instead of redirect
 *
 * The "Forms" tab
 * ---------------
 *   A: key           short name used in ?form=...
 *   B: template_url  the pre-filled Google Form link
 *   C: blank_url     (optional) plain form link, used when the visitor is
 *                    not on the roster
 *   D+: (optional) one column per form question, where the HEADER is the
 *       field name (fullname, nickname, tel, line_id, ...) and the CELL is
 *       that question's entry id ("entry.1234567890" or just "1234567890").
 *       A single column headed "mapping" also works, holding
 *       "fullname=entry.111; nickname=entry.222".
 *
 * Two ways to say which answer goes where — use either, or mix them:
 *
 *   1. PLACEHOLDERS (no extra setup). Pre-fill the form by typing marker
 *      text into each question, e.g. DummyName, DummyNickname,
 *      0XX-XXX-XXXX. Every form may use its own subset; unresolved markers
 *      are dropped so the question is left blank rather than showing
 *      "DummyNickname" to the student.
 *
 *   2. ENTRY-ID MAPPING (columns D+ above). Explicit, and the only option
 *      that works when a form's questions cannot hold marker text
 *      (dropdowns, multiple choice, linear scale).
 *
 * Run validateSetup() from the Apps Script editor to check every configured
 * form at once.
 */

/** Tab names accepted for each role, in priority order (case-insensitive). */
var FORM_SHEET_NAMES = ['forms', 'form', 'config', 'formconfig', 'form_config'];
var DATA_SHEET_NAMES = ['syncdata', 'sync_data', 'sheet1', 'data', 'students', 'roster'];

/**
 * Column positions used when a roster header cannot be recognised.
 * D=3 title, E=4 first_name, F=5 last_name, G=6 nickname,
 * K=10 docchula_email, L=11 phone_number, M=12 line_id, N=13 line_display.
 */
var DEFAULT_DATA_COLUMNS = {
  title: 3,
  firstName: 4,
  lastName: 5,
  nickname: 6,
  email: 10,
  tel: 11,
  lineId: 12,
  lineDisplay: 13
};

/* ------------------------------------------------------------------ */
/* Entry point                                                         */
/* ------------------------------------------------------------------ */

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var debug = isTruthyParam(params.debug);
  var diag = { steps: [] };

  try {
    var requestedKey = normalizeKey(params.form);
    diag.requestedKey = requestedKey;

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    /* --- 1. Form registry ------------------------------------------ */
    var configSheet = findSheet(ss, FORM_SHEET_NAMES);
    if (!configSheet) {
      return renderError(
        "Form registry tab not found",
        "Add a tab named <b>Forms</b> with columns: key | template_url | blank_url.",
        listSheetNames(ss)
      );
    }
    diag.configSheet = configSheet.getName();

    var registry = readFormRegistry(configSheet);
    diag.formKeys = registry.forms.map(function (f) { return f.displayKey; });

    if (!registry.forms.length) {
      return renderError(
        "No forms configured",
        "The <b>" + escapeHtml(configSheet.getName()) + "</b> tab has no data rows. " +
        "Put the form key in column A and the pre-filled form link in column B.",
        []
      );
    }

    var form = pickForm(registry, requestedKey);
    if (!form) {
      return renderError(
        "Unknown form key: " + escapeHtml(requestedKey || "(none)"),
        "Add <code>?form=&lt;key&gt;</code> to the link, using one of the keys " +
        "configured in the <b>" + escapeHtml(configSheet.getName()) + "</b> tab.",
        diag.formKeys
      );
    }
    diag.form = form;

    if (!form.templateUrl) {
      return renderError(
        "Form '" + escapeHtml(form.displayKey) + "' has no template URL",
        "Row " + form.row + " of the <b>" + escapeHtml(configSheet.getName()) +
        "</b> tab is missing the pre-filled form link in column B.",
        diag.formKeys
      );
    }

    /* --- 2. Roster -------------------------------------------------- */
    var dataSheet = findSheet(ss, DATA_SHEET_NAMES);
    if (!dataSheet) {
      return renderError(
        "Roster tab not found",
        "Add a tab named <b>SYNCDATA</b> holding the student data.",
        listSheetNames(ss)
      );
    }
    diag.dataSheet = dataSheet.getName();

    var roster = readRoster(dataSheet);
    diag.emailColumns = roster.emailColumns.map(function (c) { return columnLabel(c) + " (" + (roster.headers[c] || "?") + ")"; });

    /* --- 3. Identify the visitor ------------------------------------ */
    var visitorEmail = getVisitorEmail();
    diag.visitorEmail = visitorEmail;

    var match = visitorEmail ? findStudentRow(roster, visitorEmail) : null;
    diag.matchedRow = match ? match.rowNumber : null;

    /* --- 4. Build the destination ----------------------------------- */
    if (match) {
      var student = extractStudent(match.row, roster);
      var built = personalizeFormUrl(form.templateUrl, student, form.mapping, roster, match.row);
      diag.student = student;
      diag.trace = built.trace;
      diag.finalUrl = built.url;

      if (debug) return renderDebug(diag, registry);
      return renderRedirect(built.url, true);
    }

    /* --- 5. Not on the roster (or not signed in) --------------------- */
    diag.finalUrl = form.blankUrl || form.templateUrl;
    diag.fallbackReason = visitorEmail
      ? "Signed in as " + visitorEmail + ", but no roster row matched that address."
      : "The web app could not read the visitor's email address.";

    if (debug) return renderDebug(diag, registry);
    return renderRedirect(diag.finalUrl, false);

  } catch (err) {
    if (debug) {
      diag.error = (err && err.stack) ? err.stack : String(err);
      return renderDebug(diag, null);
    }
    return renderError("An error occurred", escapeHtml(err && err.message ? err.message : String(err)), []);
  }
}

/* ------------------------------------------------------------------ */
/* Spreadsheet access                                                  */
/* ------------------------------------------------------------------ */

/** Finds a sheet by name, ignoring case and surrounding whitespace. */
function findSheet(ss, candidates) {
  var sheets = ss.getSheets();
  for (var i = 0; i < candidates.length; i++) {
    for (var j = 0; j < sheets.length; j++) {
      if (normalizeKey(sheets[j].getName()) === candidates[i]) return sheets[j];
    }
  }
  return null;
}

function listSheetNames(ss) {
  return ss.getSheets().map(function (s) { return s.getName(); });
}

/**
 * Reads the Forms tab into { forms: [...], hasHeader: bool, headers: [...] }.
 * A header row is assumed unless column B of the first row already looks
 * like a URL, so the tab works with or without one.
 */
function readFormRegistry(sheet) {
  var values = sheet.getDataRange().getValues();
  var forms = [];
  var headers = [];
  var hasHeader = false;

  if (!values.length) return { forms: forms, headers: headers, hasHeader: hasHeader };

  hasHeader = !looksLikeUrl(cellText(values[0][1]));
  if (hasHeader) {
    headers = values[0].map(function (h) { return cellText(h); });
  }

  for (var r = hasHeader ? 1 : 0; r < values.length; r++) {
    var row = values[r];
    var key = normalizeKey(row[0]);
    var templateUrl = cleanUrl(row[1]);
    var blankUrl = cleanUrl(row[2]);

    // Skip rows that carry neither a key nor a link.
    if (!key && !templateUrl) continue;

    var mappingIssues = [];
    forms.push({
      key: key,
      displayKey: cellText(row[0]) || key,
      row: r + 1,
      templateUrl: templateUrl,
      blankUrl: blankUrl,
      mapping: readRowMapping(headers, row, mappingIssues),
      mappingIssues: mappingIssues
    });
  }

  return { forms: forms, headers: headers, hasHeader: hasHeader };
}

/**
 * Builds { 'entry.123': 'fieldname' } from the optional mapping columns of
 * one registry row. Supports both a column-per-field layout and a single
 * "mapping" cell.
 */
function readRowMapping(headers, row, issues) {
  var mapping = {};

  for (var c = 3; c < row.length; c++) {
    var cell = cellText(row[c]);
    if (!cell) continue;

    var header = normalizeField(headers[c] || "");
    if (!header) {
      if (issues) issues.push("column " + columnLabel(c) + " has a value but no header, so it is ignored");
      continue;
    }

    if (header === 'mapping' || header === 'fields' || header === 'entries' || header === 'entry') {
      parseMappingText(cell, mapping, issues);
    } else if (looksLikeEntryId(cell)) {
      mapping[normalizeEntryId(cell)] = header;
    } else if (issues) {
      issues.push("column " + columnLabel(c) + " ('" + (headers[c] || "") + "') holds '" + cell +
                  "', which is not an entry id; expected something like entry.1234567890");
    }
  }

  return mapping;
}

/** Parses "fullname=entry.111; nickname=entry.222" in either direction. */
function parseMappingText(text, mapping, issues) {
  var pairs = String(text).split(/[;\n,]+/);
  for (var i = 0; i < pairs.length; i++) {
    var pair = pairs[i].trim();
    if (!pair) continue;

    var sep = pair.search(/[=:]/);
    if (sep === -1) {
      if (issues) issues.push("mapping entry '" + pair + "' is missing an '='");
      continue;
    }

    var left = pair.slice(0, sep).trim();
    var right = pair.slice(sep + 1).trim();
    if (!left || !right) continue;

    if (looksLikeEntryId(left)) {
      mapping[normalizeEntryId(left)] = normalizeField(right);
    } else if (looksLikeEntryId(right)) {
      mapping[normalizeEntryId(right)] = normalizeField(left);
    } else if (issues) {
      issues.push("mapping entry '" + pair + "' names no entry id " +
                  "(expected something like fullname=entry.1234567890)");
    }
  }
  return mapping;
}

/** Chooses the requested form, or the first configured one when ?form= is absent. */
function pickForm(registry, requestedKey) {
  var forms = registry.forms;

  if (requestedKey) {
    for (var i = 0; i < forms.length; i++) {
      if (forms[i].key === requestedKey) return forms[i];
    }
    return null;
  }

  // No key given: fall back to the first row that actually has a link.
  for (var j = 0; j < forms.length; j++) {
    if (forms[j].templateUrl) return forms[j];
  }
  return null;
}

/* ------------------------------------------------------------------ */
/* Roster                                                              */
/* ------------------------------------------------------------------ */

/**
 * Reads the roster tab and works out where each field lives.
 * Returns { headers, rows, columns, headerIndex, emailColumns }.
 */
function readRoster(sheet) {
  var values = sheet.getDataRange().getValues();
  var headers = values.length ? values[0].map(function (h) { return cellText(h); }) : [];
  var rows = values.slice(1);

  var columns = {};
  for (var k in DEFAULT_DATA_COLUMNS) {
    if (DEFAULT_DATA_COLUMNS.hasOwnProperty(k)) columns[k] = DEFAULT_DATA_COLUMNS[k];
  }

  var headerIndex = {};
  var emailColumns = [];
  var emailScore = -1;
  var fromHeader = {};

  for (var c = 0; c < headers.length; c++) {
    var raw = headers[c];
    if (!raw) continue;

    var h = normalizeField(raw);
    if (!headerIndex.hasOwnProperty(h)) headerIndex[h] = c;

    // Every address-like column is remembered, so a student can be matched
    // on any of them rather than on one guessed column.
    var score = emailHeaderScore(h);
    if (score > 0) {
      emailColumns.push(c);
      // The best-scoring column is the one used for the student's own email.
      // A bare "email" therefore beats a later "parent_email" or "email_status".
      if (score > emailScore) {
        emailScore = score;
        columns.email = c;
        fromHeader.email = true;
      }
    }

    // First matching header wins, so a column appended later by a sync job
    // cannot silently take over a field that was already resolved.
    var field = matchRosterField(h);
    if (field && !fromHeader[field]) {
      columns[field] = c;
      fromHeader[field] = true;
    }
  }

  // Always consider the documented default column too.
  if (emailColumns.indexOf(DEFAULT_DATA_COLUMNS.email) === -1) {
    emailColumns.push(DEFAULT_DATA_COLUMNS.email);
  }
  if (!fromHeader.email) columns.email = DEFAULT_DATA_COLUMNS.email;

  return {
    headers: headers,
    rows: rows,
    columns: columns,
    headerIndex: headerIndex,
    emailColumns: emailColumns
  };
}

/**
 * Ranks how likely a header is to hold the student's institutional address.
 * An exact "docchula_email" beats a bare "email", which beats anything that
 * merely contains the word. 0 means "not an email column".
 */
function emailHeaderScore(h) {
  if (h === 'docchulaemail' || h === 'docchula') return 100;
  if (h === 'email' || h === 'emailaddress' || h === 'mail') return 90;
  if (h === 'useremail' || h === 'studentemail' || h === 'schoolemail') return 80;
  if (h === 'timestamp' || h === 'emailverified' || h === 'emailstatus') return 0;
  if (/mail/.test(h)) return 10;
  return 0;
}

/** Maps a normalised roster header to one of the known field names. */
function matchRosterField(h) {
  var table = {
    firstname: 'firstName', firstnames: 'firstName', 'ชื่อ': 'firstName', givenname: 'firstName',
    lastname: 'lastName', surname: 'lastName', 'นามสกุล': 'lastName', familyname: 'lastName',
    fullname: 'fullName', name: 'fullName', 'ชื่อสกุล': 'fullName',
    nickname: 'nickname', nick: 'nickname', 'ชื่อเล่น': 'nickname',
    phonenumber: 'tel', phone: 'tel', telephone: 'tel', tel: 'tel', mobile: 'tel',
    'เบอร์โทรศัพท์': 'tel', 'เบอร์โทร': 'tel',
    lineid: 'lineId', line: 'lineId',
    linedisplay: 'lineDisplay', linedisplayname: 'lineDisplay', linename: 'lineDisplay',
    displayname: 'lineDisplay',
    title: 'title', prefix: 'title', 'คำนำหน้า': 'title', 'คำนำหน้าชื่อ': 'title',
    year: 'year', 'class': 'year', 'ชั้นปี': 'year',
    studentid: 'studentId', 'รหัสนิสิต': 'studentId', 'รหัสนักศึกษา': 'studentId'
  };
  return table.hasOwnProperty(h) ? table[h] : null;
}

/** Finds the roster row belonging to the visitor, checking every email column. */
function findStudentRow(roster, visitorEmail) {
  for (var i = 0; i < roster.rows.length; i++) {
    for (var j = 0; j < roster.emailColumns.length; j++) {
      var c = roster.emailColumns[j];
      if (isEmailMatch(visitorEmail, roster.rows[i][c])) {
        return { row: roster.rows[i], rowNumber: i + 2, matchedColumn: c };
      }
    }
  }
  return null;
}

/** True when the two addresses identify the same person. */
function isEmailMatch(visitorEmail, sheetEmail) {
  var u = normalizeEmail(visitorEmail);
  var s = normalizeEmail(sheetEmail);
  if (!u || !s) return false;
  if (u === s) return true;

  var uLocal = u.split('@')[0];
  var sLocal = s.split('@')[0];
  return !!uLocal && uLocal === sLocal;
}

/** Lowercases, trims, strips invisible characters and any "+tag" suffix. */
function normalizeEmail(value) {
  if (value === null || value === undefined) return "";
  var s = String(value).replace(/[\u200B-\u200D\uFEFF]/g, '').replace(/\u00A0/g, ' ').trim().toLowerCase();
  if (!s) return "";
  return s.replace(/\+[^@]*(?=@|$)/, '');
}

/** Reads the visitor's address; returns "" when the deployment hides it. */
function getVisitorEmail() {
  try {
    var email = Session.getActiveUser().getEmail();
    return email ? email.trim() : "";
  } catch (err) {
    return "";
  }
}

/** Pulls the student's fields out of their roster row. */
function extractStudent(row, roster) {
  var cols = roster.columns;

  function get(field) {
    var c = cols[field];
    if (c === undefined || c === null) return "";
    return cellText(row[c]);
  }

  var titlePattern = /^(นาย|นางสาว|นาง|น\.ส\.|ด\.ช\.|ด\.ญ\.|นายแพทย์|แพทย์หญิง|นพ\.|พญ\.|mr\.?|mrs\.?|ms\.?|miss)\s*/i;

  var firstName = get('firstName').replace(titlePattern, '').trim();
  var lastName = get('lastName');

  var fullName = (firstName + " " + lastName).trim();
  if (!fullName) {
    fullName = get('fullName').replace(titlePattern, '').trim();
  }

  return {
    email: get('email'),
    title: get('title'),
    firstName: firstName,
    lastName: lastName,
    fullName: fullName,
    nickname: get('nickname'),
    tel: formatThaiPhone(get('tel')),
    lineId: get('lineId'),
    lineDisplay: get('lineDisplay'),
    year: get('year'),
    studentId: get('studentId')
  };
}

/** Formats a Thai mobile number as 0XX-XXX-XXXX; leaves anything else alone. */
function formatThaiPhone(raw) {
  if (!raw) return "";
  var digits = String(raw).replace(/\D/g, "");

  // Sheets drops the leading zero when the cell is stored as a number.
  if (digits.length === 9 && digits.charAt(0) !== "0") digits = "0" + digits;
  // +66 country code.
  if (digits.length === 11 && digits.indexOf("66") === 0) digits = "0" + digits.slice(2);

  if (digits.length === 10) {
    return digits.slice(0, 3) + "-" + digits.slice(3, 6) + "-" + digits.slice(6);
  }
  return digits || String(raw).trim();
}

/* ------------------------------------------------------------------ */
/* URL personalisation                                                 */
/* ------------------------------------------------------------------ */

/**
 * Marker text -> field name. Longest markers are applied first so that
 * DummyLineDisplay is never eaten by DummyLine, and DummyNickname is never
 * eaten by DummyNick.
 */
var PLACEHOLDERS = [
  ['DummyAcademicYear', 'year'],
  ['DummyLineDisplay', 'lineDisplay'],
  ['DummyStudentID', 'studentId'],
  ['DummyStudentId', 'studentId'],
  ['DummyStudent_Id', 'studentId'],
  ['DummyFullName', 'fullName'],
  ['DummyLastName', 'lastName'],
  ['DummyFirstName', 'firstName'],
  ['DummyLineName', 'lineDisplay'],
  ['DummyNickname', 'nickname'],
  ['DummyStudent', 'studentId'],
  ['DummyDisplay', 'lineDisplay'],
  ['DummyLineID', 'lineId'],
  ['0XX-XXX-XXXX', 'tel'],
  ['081-234-5678', 'tel'],
  ['DummyPrefix', 'title'],
  ['DummyEmail', 'email'],
  ['DummyPhone', 'tel'],
  ['DummyTitle', 'title'],
  ['DummyClass', 'year'],
  ['6XXXXXXX30', 'studentId'],
  ['0812345678', 'tel'],
  ['DummyNick', 'nickname'],
  ['DummyLine', 'lineId'],
  ['DummyYear', 'year'],
  ['DummyName', 'fullName'],   // becomes firstName when DummyLastName is present
  ['DummyTel', 'tel'],
  ['DummyId', 'studentId']
].sort(function (a, b) { return b[0].length - a[0].length; });

/** Fields that are deliberately left blank for the student to type in. */
var BLANK_FIELDS = { studentId: true };

/**
 * Rewrites the template URL for one student.
 *
 * Each query parameter is handled on its own, so one question can never
 * clobber another, duplicate keys (checkbox answers) survive, and a
 * parameter whose marker could not be resolved is dropped so the question
 * shows up empty instead of showing "DummyNickname".
 *
 * Returns { url: string, trace: [{ key, before, after, action }] }.
 */
function personalizeFormUrl(templateUrl, student, mapping, roster, row) {
  var trace = [];
  if (!templateUrl) return { url: "", trace: trace };

  var hashAt = templateUrl.indexOf('#');
  var hash = hashAt === -1 ? "" : templateUrl.slice(hashAt);
  var withoutHash = hashAt === -1 ? templateUrl : templateUrl.slice(0, hashAt);

  var queryAt = withoutHash.indexOf('?');
  if (queryAt === -1) return { url: templateUrl, trace: trace };

  var base = withoutHash.slice(0, queryAt);
  var query = withoutHash.slice(queryAt + 1);

  // "DummyName" means the first name only when the form asks for the surname
  // in a separate question; otherwise it stands for the whole name.
  var splitName = /DummyLastName/i.test(query);

  var parts = query.split('&');
  var kept = [];

  for (var i = 0; i < parts.length; i++) {
    var part = parts[i];
    if (!part) continue;

    var eq = part.indexOf('=');
    var key = eq === -1 ? part : part.slice(0, eq);
    var rawValue = eq === -1 ? null : part.slice(eq + 1);

    if (rawValue === null) {           // bare flag, e.g. "pageHistory"
      kept.push(part);
      continue;
    }

    var before = decodeParam(rawValue);
    var resolved = resolveParamValue(key, before, student, mapping, roster, row, splitName);

    if (resolved.drop) {
      trace.push({ key: key, before: before, after: "", action: resolved.reason });
      continue;
    }

    kept.push(key + '=' + encodeURIComponent(resolved.value));
    trace.push({ key: key, before: before, after: resolved.value, action: resolved.reason });
  }

  var url = kept.length ? base + '?' + kept.join('&') : base;
  return { url: url + hash, trace: trace };
}

/**
 * Works out what one query parameter should carry.
 * Returns { value, drop, reason }.
 */
function resolveParamValue(key, value, student, mapping, roster, row, splitName) {
  // 1. An explicit entry-id mapping always wins.
  var entryKey = normalizeEntryId(key);
  if (mapping && mapping.hasOwnProperty(entryKey)) {
    var field = mapping[entryKey];
    if (isBlankDirective(field)) {
      return { drop: true, reason: "left blank (mapping)" };
    }
    var mapped = lookupFieldValue(field, student, roster, row);
    if (!mapped) return { drop: true, reason: "no data for '" + field + "'" };
    return { value: mapped, drop: false, reason: "mapped to " + field };
  }

  // 2. Otherwise look for marker text inside the value.
  var replaced = value;
  var sawMarker = false;
  var sawBlankField = false;

  for (var i = 0; i < PLACEHOLDERS.length; i++) {
    var marker = PLACEHOLDERS[i][0];
    var fieldName = PLACEHOLDERS[i][1];

    var re = new RegExp(escapeRegExp(marker), 'gi');
    if (!re.test(replaced)) continue;
    re.lastIndex = 0;

    sawMarker = true;

    if (BLANK_FIELDS[fieldName]) {
      sawBlankField = true;
      break;
    }

    if (fieldName === 'fullName' && splitName && /^dummyname$/i.test(marker)) {
      fieldName = 'firstName';
    }

    replaced = replaced.replace(re, lookupFieldValue(fieldName, student, roster, row));
  }

  if (sawBlankField) {
    return { drop: true, reason: "left blank for the student" };
  }

  if (!sawMarker) {
    // Marker text this script does not know about must not reach the student.
    if (looksLikeUnresolvedMarker(value)) {
      return { drop: true, reason: "unrecognised placeholder" };
    }
    // A literal value the form author wants on every submission.
    return { value: value, drop: false, reason: "kept as-is" };
  }

  replaced = replaced.replace(/\s+/g, ' ').trim();

  if (!replaced) {
    return { drop: true, reason: "no data on the roster" };
  }
  if (looksLikeUnresolvedMarker(replaced)) {
    return { drop: true, reason: "unrecognised placeholder" };
  }

  return { value: replaced, drop: false, reason: "filled in" };
}

/**
 * Resolves a mapping target: a known field name, "col:C", a roster header
 * name, or a literal fallback of "".
 */
function lookupFieldValue(field, student, roster, row) {
  if (!field) return "";

  if (student.hasOwnProperty(field)) {
    return BLANK_FIELDS[field] ? "" : student[field];
  }

  var normalized = normalizeField(field);

  var known = matchRosterField(normalized);
  if (known && student.hasOwnProperty(known)) {
    return BLANK_FIELDS[known] ? "" : student[known];
  }

  var colMatch = /^col:?([a-z]+)$/.exec(normalized);
  if (colMatch) {
    var index = columnLetterToIndex(colMatch[1]);
    return (row && index < row.length) ? cellText(row[index]) : "";
  }

  if (roster && roster.headerIndex.hasOwnProperty(normalized)) {
    return cellText(row[roster.headerIndex[normalized]]);
  }

  return "";
}

function isBlankDirective(field) {
  var f = normalizeField(field);
  return f === 'skip' || f === 'blank' || f === 'none' || f === 'empty' || f === '' || f === '-';
}

/** Catches leftovers such as "DummyWhatever" or "0xx-xxx-xxxx". */
function looksLikeUnresolvedMarker(value) {
  if (/dummy/i.test(value)) return true;
  if (/^[0-9]?x{2,}[-\s]?x{3}[-\s]?x{3,4}$/i.test(value)) return true;
  if (/^x{3,}$/i.test(value)) return true;
  return false;
}

/* ------------------------------------------------------------------ */
/* Rendering                                                           */
/* ------------------------------------------------------------------ */

var PAGE_FONT = "-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif";

/** Sends the visitor on to the form. */
function renderRedirect(url, personalized) {
  var headline = personalized
    ? "กำลังนำท่านไปยังแบบฟอร์มเฉพาะบุคคล..."
    : "กำลังนำท่านไปยังแบบฟอร์ม...";
  var sub = personalized
    ? "Loading your personalized form, please wait..."
    : "Redirecting to form...";

  var html =
    "<script>" +
    "  var target = '" + escapeJs(url) + "';" +
    "  try { window.top.location.href = target; }" +
    "  catch (err) { window.location.href = target; }" +
    "</script>" +
    "<div style=\"font-family:" + PAGE_FONT + "; text-align:center; margin-top:80px;\">" +
    "  <p style='font-size:18px; color:#333;'>" + headline + "</p>" +
    "  <p style='font-size:14px; color:#777;'>" + sub + "</p>" +
    "  <p style='font-size:13px; color:#aaa;'>หากหน้าเว็บไม่เปลี่ยนอัตโนมัติ " +
    "<a target='_top' href=\"" + escapeHtml(url) + "\">คลิกที่นี่</a></p>" +
    "</div>";

  return htmlPage(html);
}

/** Shows a readable problem report instead of silently sending people to a blank form. */
function renderError(title, detail, keys) {
  var html =
    "<div style=\"font-family:" + PAGE_FONT + "; max-width:640px; margin:60px auto; padding:0 20px; color:#333;\">" +
    "  <h3 style='color:#c5221f; margin-bottom:8px;'>" + title + "</h3>" +
    "  <p style='color:#555; line-height:1.6;'>" + detail + "</p>";

  if (keys && keys.length) {
    html += "<p style='color:#555;'>Available: " +
            keys.map(function (k) { return "<code>" + escapeHtml(k) + "</code>"; }).join(", ") +
            "</p>";
  }

  html += "<p style='color:#888; font-size:13px;'>Add <code>&amp;debug=1</code> to the link for a full diagnostic report.</p>" +
          "</div>";

  return htmlPage(html);
}

/** The ?debug=1 report: everything needed to see why a form is not filling in. */
function renderDebug(diag, registry) {
  var html =
    "<div style=\"font-family:" + PAGE_FONT + "; max-width:900px; margin:32px auto; padding:0 20px; color:#222;\">" +
    "<h2 style='margin-bottom:4px;'>Autofill diagnostics</h2>" +
    "<p style='color:#777; font-size:13px; margin-top:0;'>No redirect was performed because <code>debug=1</code> is set.</p>";

  html += debugSection("Request", [
    ["Requested form key", escapeHtml(diag.requestedKey || "(none \u2014 first configured form is used)")],
    ["Form registry tab", escapeHtml(diag.configSheet || "not found")],
    ["Roster tab", escapeHtml(diag.dataSheet || "not found")],
    ["Configured form keys", escapeHtml((diag.formKeys || []).join(", ")) || "(none)"]
  ]);

  if (diag.form) {
    var mappingKeys = [];
    for (var m in diag.form.mapping) {
      if (diag.form.mapping.hasOwnProperty(m)) {
        mappingKeys.push(escapeHtml(m + " \u2192 " + diag.form.mapping[m]));
      }
    }
    var issues = (diag.form.mappingIssues || []).map(escapeHtml);

    html += debugSection("Selected form", [
      ["Key", escapeHtml(diag.form.displayKey)],
      ["Registry row", escapeHtml(diag.form.row)],
      ["Template URL", escapeHtml(diag.form.templateUrl)],
      ["Blank URL", escapeHtml(diag.form.blankUrl) || "(none)"],
      ["Entry-id mapping", mappingKeys.length
        ? mappingKeys.join("<br>")
        : "(none \u2014 using placeholder text)"],
      ["Mapping problems", issues.length ? problem(issues.join("<br>")) : "(none)"]
    ]);
  }

  html += debugSection("Visitor", [
    ["Detected email", escapeHtml(diag.visitorEmail) || problem("empty")],
    ["Roster email columns searched", escapeHtml((diag.emailColumns || []).join(", ")) || "(none)"],
    ["Matched roster row", diag.matchedRow ? escapeHtml(diag.matchedRow) : problem("no match")]
  ]);

  if (!diag.visitorEmail) {
    html +=
      "<div style='background:#fef7e0; border-left:4px solid #f9ab00; padding:12px 16px; margin:16px 0; line-height:1.6;'>" +
      "<b>The visitor's email address is not visible to the script.</b><br>" +
      "Re-deploy the web app with <b>Execute as: User accessing the web app</b>, and make sure the " +
      "student opens the link in a browser where they are signed in to the right Google account. " +
      "In-app browsers (LINE, Messenger) are often signed out \u2014 append " +
      "<code>&amp;openExternalBrowser=1</code> to the link." +
      "</div>";
  }

  if (diag.fallbackReason) {
    html += "<p style='color:#c5221f;'>" + escapeHtml(diag.fallbackReason) + "</p>";
  }

  if (diag.student) {
    var fields = [];
    for (var f in diag.student) {
      if (diag.student.hasOwnProperty(f)) {
        fields.push([escapeHtml(f), escapeHtml(diag.student[f]) || "(empty)"]);
      }
    }
    html += debugSection("Student data read from the roster", fields);
  }

  if (diag.trace && diag.trace.length) {
    html += "<h3 style='margin-top:28px;'>Question by question</h3>" +
            "<table style='border-collapse:collapse; width:100%; font-size:13px;'>" +
            "<tr style='background:#f1f3f4; text-align:left;'>" +
            "<th style='padding:8px; border:1px solid #ddd;'>Parameter</th>" +
            "<th style='padding:8px; border:1px solid #ddd;'>Template value</th>" +
            "<th style='padding:8px; border:1px solid #ddd;'>Sent to the form</th>" +
            "<th style='padding:8px; border:1px solid #ddd;'>What happened</th></tr>";

    for (var t = 0; t < diag.trace.length; t++) {
      var e = diag.trace[t];
      var isProblem = /unrecognised|no data/.test(e.action);
      html += "<tr" + (isProblem ? " style='background:#fce8e6;'" : "") + ">" +
              "<td style='padding:8px; border:1px solid #ddd;'><code>" + escapeHtml(e.key) + "</code></td>" +
              "<td style='padding:8px; border:1px solid #ddd;'>" + escapeHtml(e.before) + "</td>" +
              "<td style='padding:8px; border:1px solid #ddd;'>" + escapeHtml(e.after) + "</td>" +
              "<td style='padding:8px; border:1px solid #ddd;'>" + escapeHtml(e.action) + "</td></tr>";
    }
    html += "</table>";
  }

  if (diag.finalUrl) {
    html += "<h3 style='margin-top:28px;'>Destination</h3>" +
            "<p style='word-break:break-all; font-size:13px;'>" +
            "<a target='_top' href=\"" + escapeHtml(diag.finalUrl) + "\">" + escapeHtml(diag.finalUrl) + "</a></p>";
  }

  if (diag.error) {
    html += "<h3 style='margin-top:28px; color:#c5221f;'>Error</h3><pre style='white-space:pre-wrap; font-size:12px;'>" +
            escapeHtml(diag.error) + "</pre>";
  }

  html += "</div>";
  return htmlPage(html);
}

/** Renders one label/value table. Values must already be HTML-safe. */
function debugSection(title, rows) {
  var html = "<h3 style='margin-top:28px;'>" + title + "</h3>" +
             "<table style='border-collapse:collapse; width:100%; font-size:13px;'>";
  for (var i = 0; i < rows.length; i++) {
    html += "<tr>" +
            "<td style='padding:8px; border:1px solid #ddd; width:230px; background:#f8f9fa;'>" + rows[i][0] + "</td>" +
            "<td style='padding:8px; border:1px solid #ddd; word-break:break-all;'>" + rows[i][1] + "</td>" +
            "</tr>";
  }
  return html + "</table>";
}

function problem(safeHtml) {
  return "<span style='color:#c5221f;'>" + safeHtml + "</span>";
}

function htmlPage(bodyHtml) {
  return HtmlService.createHtmlOutput(bodyHtml)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ------------------------------------------------------------------ */
/* Setup check (run from the Apps Script editor)                       */
/* ------------------------------------------------------------------ */

/**
 * Checks every configured form and logs anything that would stop it from
 * filling in. Run this after adding a form.
 */
function validateSetup() {
  var report = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var configSheet = findSheet(ss, FORM_SHEET_NAMES);
  if (!configSheet) {
    report.push("ERROR  No form registry tab. Expected one named 'Forms'. Tabs present: " + listSheetNames(ss).join(", "));
    Logger.log(report.join("\n"));
    return report;
  }

  var dataSheet = findSheet(ss, DATA_SHEET_NAMES);
  if (!dataSheet) {
    report.push("ERROR  No roster tab. Expected one named 'SYNCDATA'. Tabs present: " + listSheetNames(ss).join(", "));
    Logger.log(report.join("\n"));
    return report;
  }

  var roster = readRoster(dataSheet);
  report.push("Roster tab: " + dataSheet.getName() + " (" + roster.rows.length + " rows)");
  report.push("Email columns searched: " + roster.emailColumns.map(function (c) {
    return columnLabel(c) + " '" + (roster.headers[c] || "") + "'";
  }).join(", "));

  for (var field in roster.columns) {
    if (roster.columns.hasOwnProperty(field)) {
      report.push("  " + field + " -> column " + columnLabel(roster.columns[field]) +
                  " '" + (roster.headers[roster.columns[field]] || "") + "'");
    }
  }

  var registry = readFormRegistry(configSheet);
  report.push("");
  report.push("Form registry tab: " + configSheet.getName() + " (" + registry.forms.length + " forms)");

  var seen = {};
  for (var i = 0; i < registry.forms.length; i++) {
    var form = registry.forms[i];
    report.push("");
    report.push("[row " + form.row + "] key='" + form.key + "'");

    if (!form.key) {
      report.push("  ERROR  Column A is empty, so ?form=... can never select this row.");
    } else if (seen[form.key]) {
      report.push("  ERROR  Duplicate key; the row at line " + seen[form.key] + " wins.");
    } else {
      seen[form.key] = form.row;
    }

    for (var m = 0; m < form.mappingIssues.length; m++) {
      report.push("  WARN   " + form.mappingIssues[m]);
    }

    if (!form.templateUrl) {
      report.push("  ERROR  Column B has no pre-filled form link.");
      continue;
    }
    if (!looksLikeUrl(form.templateUrl)) {
      report.push("  ERROR  Column B is not a URL: " + form.templateUrl);
      continue;
    }
    if (form.templateUrl.indexOf('?') === -1) {
      report.push("  ERROR  Column B has no query string. Use the link from " +
                  "'Get pre-filled link', not the plain form link.");
      continue;
    }
    if (!form.blankUrl) {
      report.push("  WARN   Column C is empty; visitors who are not on the roster " +
                  "will be sent to the pre-filled link instead of a blank form.");
    }

    // Which questions will actually receive data?
    var probe = personalizeFormUrl(
      form.templateUrl,
      { email: 'a@b.c', title: 'x', firstName: 'x', lastName: 'x', fullName: 'x x',
        nickname: 'x', tel: '012-345-6789', lineId: 'x', lineDisplay: 'x',
        year: 'x', studentId: '' },
      form.mapping,
      roster,
      []
    );

    var filled = 0;
    var unresolved = [];
    for (var t = 0; t < probe.trace.length; t++) {
      if (probe.trace[t].action === 'filled in' || probe.trace[t].action.indexOf('mapped to') === 0) filled++;
      if (probe.trace[t].action === 'unrecognised placeholder') unresolved.push(probe.trace[t].before);
    }

    report.push("  " + filled + " question(s) will be filled in.");
    if (!filled) {
      report.push("  ERROR  Nothing in this URL is recognised. Either pre-fill the form " +
                  "using the marker words (DummyName, DummyNickname, 0XX-XXX-XXXX, ...) " +
                  "or add entry-id mapping columns to the registry row.");
    }
    if (unresolved.length) {
      report.push("  WARN   Unrecognised markers left blank: " + unresolved.join(", "));
    }
  }

  Logger.log(report.join("\n"));
  return report;
}

/* ------------------------------------------------------------------ */
/* Small helpers                                                       */
/* ------------------------------------------------------------------ */

function cellText(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(value).replace(/[\u200B-\u200D\uFEFF]/g, '').replace(/\u00A0/g, ' ').trim();
}

function cleanUrl(value) {
  return cellText(value).replace(/\s+/g, '');
}

function looksLikeUrl(value) {
  return /^https?:\/\//i.test(String(value || "").trim());
}

/** Lowercases and strips spaces/punctuation so key lookups are forgiving. */
function normalizeKey(value) {
  return cellText(value).toLowerCase().replace(/\s+/g, '');
}

/** Same as normalizeKey but also drops separators, for header matching. */
function normalizeField(value) {
  return cellText(value).toLowerCase().replace(/[\s_\-.()\[\]]/g, '');
}

function looksLikeEntryId(value) {
  var v = cellText(value);
  return /^entry[._]?\d+$/i.test(v) || /^\d{4,}$/.test(v);
}

function normalizeEntryId(value) {
  var digits = cellText(value).replace(/\D/g, '');
  return digits ? 'entry.' + digits : normalizeField(value);
}

function columnLetterToIndex(letters) {
  var s = String(letters).toUpperCase().replace(/[^A-Z]/g, '');
  var index = 0;
  for (var i = 0; i < s.length; i++) {
    index = index * 26 + (s.charCodeAt(i) - 64);
  }
  return index - 1;
}

function columnLabel(index) {
  var n = Number(index) + 1;
  var label = "";
  while (n > 0) {
    var rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

function decodeParam(raw) {
  try {
    return decodeURIComponent(String(raw).replace(/\+/g, ' '));
  } catch (err) {
    return String(raw);
  }
}

function escapeRegExp(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\\-]/g, '\\$&');
}

function isTruthyParam(value) {
  if (value === undefined || value === null) return false;
  var v = String(value).toLowerCase().trim();
  return v === '1' || v === 'true' || v === 'yes' || v === 'on' || v === '';
}

function escapeHtml(text) {
  return String(text === null || text === undefined ? "" : text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** Escapes a value for use inside a single-quoted JavaScript string literal. */
function escapeJs(text) {
  return String(text === null || text === undefined ? "" : text)
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/"/g, '\\"')
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026')
    .replace(/\r/g, '')
    .replace(/\n/g, '');
}
