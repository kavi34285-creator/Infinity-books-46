/**
 * Infinity Books — Google Sheets backend (Apps Script Web App)
 *
 * ONE deployed web app + ONE spreadsheet. The app sends every form as a
 * POST with a "type" field, and this script files it in its own tab:
 *
 *   type "signup" -> "Personal Details"  (Name, Email, Password, Phone)
 *   type "admin"  -> "Admin Accounts"    (Create the admin account form)
 *   type "member" -> "Members"           (Add member form)
 *   type "book"   -> "Books"             (Add book form)
 *
 * Tabs and header rows are created automatically the first time each form
 * is submitted (or all at once by running setupSheets(), see step 5).
 *
 * ---------------------------------------------------------------------
 * SETUP
 * ---------------------------------------------------------------------
 * 1. Open the Google Sheet that should hold the data.
 *    Extensions -> Apps Script.
 * 2. Delete the starter code and paste this whole file in as Code.gs.
 * 3. SPREADSHEET_ID below is already filled in. If you opened Apps Script
 *    from inside the sheet (step 1) you can also set it to "" and the
 *    script will use that sheet.
 * 4. SHARED_SECRET must match SHEETS_SYNC_TOKEN in the app (it already
 *    does: "infinity-books-2026"). Change both together if you change it.
 * 5. Pick "setupSheets" in the function dropdown and press Run. Approve
 *    the permissions when asked. The four tabs appear in your sheet.
 * 6. Deploy -> New deployment -> gear icon -> Web app
 *      Execute as:      Me
 *      Who has access:  Anyone
 *    Deploy, then copy the Web app URL (ends in /exec).
 * 7. Paste that URL in the app: Staff -> Google Sheets sync -> Save.
 *    (Or set DEFAULT_SHEETS_WEBAPP_URL in the HTML file.)
 *
 * AFTER EDITING THIS SCRIPT you must publish a new version or the old code
 * keeps running: Deploy -> Manage deployments -> pencil icon -> Version:
 * "New version" -> Deploy. The /exec URL stays the same.
 *
 * ---------------------------------------------------------------------
 * PASSWORDS
 * ---------------------------------------------------------------------
 * PASSWORD_STORAGE is set to "plain": the Password column in both
 * "Personal Details" and "Admin Accounts" shows the real password exactly
 * as it was typed. Only share the sheet with people you fully trust.
 * To go back to scrambled (salted SHA-256) passwords, set it to "hash".
 * Rows saved earlier as "sha256:..." cannot be turned back into the real
 * password - those users/admins must submit the form again.
 */

var SPREADSHEET_ID = "1ARlalweb09ib50zqfKm9Kl8NUtcQFCpQ2JxQLF9sStY";
var SHARED_SECRET = "infinity-books-2026"; // must match SHEETS_SYNC_TOKEN in the app
var PASSWORD_STORAGE = "plain";            // "plain" = real password shown in sheet, "hash" = scrambled
var PASSWORD_SALT = "infinity-books-2026-salt";

// Tab definitions. `legacyName` lets an older tab be renamed in place so
// existing rows are kept (the personal-details tab used to be "Signups").
var TABS = {
  signup: {
    name: "Personal Details",
    legacyName: "Signups",
    headers: ["Timestamp", "Name", "Email", "Password", "Phone"]
  },
  admin: {
    name: "Admin Accounts",
    headers: ["Timestamp", "Name", "Email (sign-in id)", "Password", "Phone", "Role", "Created"]
  },
  member: {
    name: "Members",
    headers: ["Timestamp", "Member ID", "Full name", "Email", "Phone", "Joined"]
  },
  book: {
    name: "Books",
    headers: ["Timestamp", "Book ID", "Title", "Author", "ISBN", "Category", "Total copies", "Issued to"]
  }
};

/* ===================================================================
 * Entry points
 * =================================================================== */

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (parseErr) {
    return err_("Invalid JSON body");
  }
  if (!data || data.token !== SHARED_SECRET) {
    return err_("Invalid token");
  }

  var type = String(data.type || "member");
  var handler = {
    signup: handleSignup_,
    admin: handleAdmin_,
    member: handleMember_,
    book: handleBook_
  }[type];
  if (!handler) return err_("Unknown form type: " + type);

  // One write at a time, so two people submitting together can't
  // produce duplicate rows or overwrite each other.
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (lockErr) {
    return err_("Server busy, try again");
  }
  try {
    return handler(openSpreadsheet_(), data);
  } catch (ex) {
    return err_("Server error: " + ex.message);
  } finally {
    lock.releaseLock();
  }
}

// Open the /exec URL in a browser to check the deployment is live.
function doGet() {
  return json_({ status: "ok", message: "Infinity Books sync endpoint is live.", tabs: tabNames_() });
}

// Run this once from the editor to create all four tabs with headers.
function setupSheets() {
  var ss = openSpreadsheet_();
  Object.keys(TABS).forEach(function (key) { getTab_(ss, TABS[key]); });
  Logger.log("Ready: " + tabNames_().join(", "));
}

/* ===================================================================
 * Form handlers
 * =================================================================== */

// Personal details form: Name -> Email -> Password -> Phone number
function handleSignup_(ss, data) {
  var name = clip_(data.name, 200);
  var email = clip_(data.email, 254);
  var password = String(data.password || "");
  var phone = clip_(data.phone, 40);

  if (!name) return err_("Name is required");
  if (!isValidEmail_(email)) return err_("Invalid email");
  if (!isValidPassword_(password)) return err_("Password must be at least 6 characters and include a letter and a number");
  if (!isValidPhone_(phone)) return err_("Invalid phone number");

  var sh = getTab_(ss, TABS.signup);
  ensurePasswordColumnText_(sh, 4);
  if (findRowByEmail_(sh, 3, email) > 0) return err_("An account with this email already exists");

  sh.appendRow([
    new Date(),
    cell_(name),
    cell_(email),
    passwordCell_(password, email),
    cell_(phone, true)
  ]);
  return ok_();
}

// Create the admin account form. The same email updates its row instead
// of adding a duplicate (an admin can be re-created after "Reset all logins").
function handleAdmin_(ss, data) {
  var name = clip_(data.name, 200);
  var email = clip_(data.email, 254);
  var password = String(data.password || "");
  var phone = clip_(data.phone, 40);

  if (!name) return err_("Name is required");
  if (!isValidEmail_(email)) return err_("Invalid email");
  if (!isValidPassword_(password)) return err_("Password must be at least 6 characters and include a letter and a number");
  if (!isValidPhone_(phone)) return err_("Invalid phone number");

  var sh = getTab_(ss, TABS.admin);
  ensurePasswordColumnText_(sh, 4);
  var row = [
    new Date(),
    cell_(name),
    cell_(email),
    passwordCell_(password, email),
    cell_(phone, true),
    "Admin",
    cell_(clip_(data.createdAt, 40))
  ];
  var existing = findRowByEmail_(sh, 3, email);
  if (existing > 0) {
    sh.getRange(existing, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
  return ok_();
}

// Add member form
function handleMember_(ss, data) {
  var name = clip_(data.name, 200);
  var email = clip_(data.email, 254);
  if (!name) return err_("Name is required");
  if (email && !isValidEmail_(email)) return err_("Invalid email");

  getTab_(ss, TABS.member).appendRow([
    new Date(),
    cell_(data.memberId),
    cell_(name),
    cell_(email),
    cell_(clip_(data.phone, 40), true),
    cell_(clip_(data.joinedAt, 40))
  ]);
  return ok_();
}

// Add book form
function handleBook_(ss, data) {
  var title = clip_(data.title, 300);
  if (!title) return err_("Title is required");

  getTab_(ss, TABS.book).appendRow([
    new Date(),
    cell_(data.bookId),
    cell_(title),
    cell_(clip_(data.author, 300)),
    cell_(clip_(data.isbn, 40), true),
    cell_(clip_(data.category, 100)),
    cell_(data.totalCopies),
    cell_(clip_(data.issuedTo, 1000))
  ]);
  return ok_();
}

/* ===================================================================
 * Sheet helpers
 * =================================================================== */

function openSpreadsheet_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error("Set SPREADSHEET_ID at the top of the script.");
  return active;
}

function tabNames_() {
  return Object.keys(TABS).map(function (k) { return TABS[k].name; });
}

// Returns the tab, creating it (with a styled header row) if needed.
function getTab_(ss, def) {
  var sh = ss.getSheetByName(def.name);
  if (!sh && def.legacyName) {
    var old = ss.getSheetByName(def.legacyName);
    if (old) { old.setName(def.name); sh = old; }
  }
  if (!sh) sh = ss.insertSheet(def.name);

  if (sh.getLastRow() === 0) {
    sh.appendRow(def.headers);
    var header = sh.getRange(1, 1, 1, def.headers.length);
    header.setFontWeight("bold");
    header.setBackground("#e8f0fe");
    sh.setFrozenRows(1);
    sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  }
  return sh;
}

// Row number (1-based) of the first data row whose `col` matches `email`
// (case-insensitive), or 0 if there is none. Row 1 is the header.
function findRowByEmail_(sh, col, email) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var values = sh.getRange(2, col, last - 1, 1).getValues();
  var want = String(email).trim().toLowerCase();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === want) return i + 2;
  }
  return 0;
}

/* ===================================================================
 * Value helpers
 * =================================================================== */

function clip_(v, max) {
  var s = (v === undefined || v === null) ? "" : String(v).trim();
  return s.length > max ? s.substring(0, max) : s;
}

// Makes a value safe to write into a cell.
//  - Text starting with = + - @ would be run as a formula when the sheet
//    opens (formula injection), so a leading apostrophe keeps it as text.
//  - forceText also protects phone numbers and ISBNs from being turned
//    into numbers (which drops leading zeros and shows 9.78E+12).
function cell_(v, forceText) {
  var s = (v === undefined || v === null) ? "" : String(v).trim();
  if (s === "") return "";
  if (/^[=+\-@\t\r]/.test(s) || (forceText && /^[0-9.,+\-() ]+$/.test(s))) s = "'" + s;
  return s;
}

// Password cell value. In "plain" mode the password is written untouched
// (no trimming, no apostrophe) into a column formatted as plain text, so
// what you see in the sheet is exactly what the person typed.
function passwordCell_(password, email) {
  return passwordForSheet_(password, email);
}

// Makes the whole Password column plain text so values like 1e10, 007abc
// or =abc1 are never turned into numbers/formulas. Also fixes older tabs.
function ensurePasswordColumnText_(sh, col) {
  sh.getRange(1, col, Math.max(sh.getMaxRows(), 1), 1).setNumberFormat("@");
}

function passwordForSheet_(password, email) {
  if (PASSWORD_STORAGE === "plain") return password;
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    PASSWORD_SALT + "|" + String(email).toLowerCase() + "|" + password,
    Utilities.Charset.UTF_8
  );
  return "sha256:" + bytes.map(function (b) {
    return ("0" + (b & 0xff).toString(16)).slice(-2);
  }).join("");
}

function isValidEmail_(v) {
  return typeof v === "string" && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v);
}
function isValidPassword_(v) {
  return typeof v === "string" && v.length >= 6 && /[A-Za-z]/.test(v) && /[0-9]/.test(v);
}
function isValidPhone_(v) {
  return typeof v === "string" && v.replace(/\D/g, "").length >= 7;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function ok_() {
  return json_({ status: "ok" });
}
function err_(message) {
  return json_({ status: "error", message: message });
}
