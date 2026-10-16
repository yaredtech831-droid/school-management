/**
 * ============================================================================
 *  የተማሪ ውጤት እና ክፍያ ስርዓት - ነጻ Backend (Google Apps Script)
 * ============================================================================
 */

function getAdminPassword_() {
  return PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD") || "";
}

function getAdminEmail_() {
  return PropertiesService.getScriptProperties().getProperty("ADMIN_EMAIL") || "";
}

const SHEET_NAME = "Results";
const HEADER = ["Year", "Semester", "Grade", "Stream", "Section", "Code", "Name", "ScoresJSON", "Commented", "SubjectsJSON", "BreakdownJSON", "AssessmentConfigJSON", "Specialization"];

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const savedId = props.getProperty("spreadsheetId");
  if (savedId) {
    try { ss = SpreadsheetApp.openById(savedId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create("StudentResultsDB");
    props.setProperty("spreadsheetId", ss.getId());
  }
  return ss;
}

function getSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
  } else {
    // Add the optional specialization field without shifting any existing Results columns.
    const lastCol = Math.max(sh.getLastColumn(), HEADER.length - 1);
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
    if (headers.indexOf('Specialization') === -1) sh.getRange(1, lastCol + 1).setValue('Specialization');
  }
  return sh;
}

const CARDS_SHEET_NAME = "Cards";
const CARDS_HEADER = ["Year", "Grade", "Stream", "Section", "Code", "Name", "Status", "DriveFileId", "UpdatedAt"];

function getCardsSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(CARDS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CARDS_SHEET_NAME);
    sh.appendRow(CARDS_HEADER);
  }
  return sh;
}

// Accounts ሠንጠረዥ: Host የተባለ አዲስ ረድፍ ጨምረናል (የተለያዩ ሆስቶች እንዳይጋጩ)
const ACCOUNTS_SHEET_NAME = "Accounts";
const ACCOUNTS_HEADER = ["Email", "Host", "PasswordHash", "BoundCode", "BoundName", "BoundYear", "BoundGrade", "BoundSection", "CreatedAt", "DeviceId"];

function getAccountsSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(ACCOUNTS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(ACCOUNTS_SHEET_NAME);
    sh.appendRow(ACCOUNTS_HEADER);
  } else {
    // Keep existing account data and add the device-binding column at the end.
    const lastCol = Math.max(1, sh.getLastColumn());
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim());
    if (headers.indexOf("DeviceId") === -1) {
      sh.getRange(1, lastCol + 1).setValue("DeviceId");
    }
  }
  return sh;
}

function deviceIdColumn_() {
  const sh = getAccountsSheet_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v || "").trim());
  return headers.indexOf("DeviceId") + 1;
}

function checkParentDeviceRequest_(body) {
  const email = normalizeEmail_(body.email);
  const host = normalizeHost_(body.host || body.origin || "default");
  if (host !== "parent") return null;
  const action = String(body.action || "").trim();
  // Signup and login have their own checks; login is the only place that may bind a device.
  if (["requestSignupCode", "parentSignup", "signup", "parentLogin", "login"].indexOf(action) !== -1) return null;
  if (!email) return { ok: true, status: "no_account" };
  const found = findAccountRow_(email, host);
  if (!found) return { ok: true, status: "no_account" };
  if (String(found.row[2] || "") !== hashPassword_(String(body.password || ""))) return { ok: true, status: "wrong_password" };
  const deviceId = String(body.deviceId || "").trim();
  if (!deviceId) return { ok: true, status: "device_required" };
  const col = deviceIdColumn_();
  const savedDeviceId = String(found.row[col - 1] || "").trim();
  if (!savedDeviceId) return { ok: true, status: "device_not_registered" };
  if (savedDeviceId !== deviceId) return { ok: true, status: "device_mismatch" };
  return null;
}

const PAYMENTS_SHEET_NAME = "Payments";
const PAYMENTS_HEADER = ["PaymentID", "ParentEmail", "Host", "StudentCode", "StudentName", "AcademicYear", "Grade", "Section", "BankID", "Amount", "PhotoURL", "Status", "AdminNote", "SubmittedAt"];

function getPaymentsSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(PAYMENTS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PAYMENTS_SHEET_NAME);
    sh.appendRow(PAYMENTS_HEADER);
  }
  return sh;
}

function getCardsFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty("cardsFolderId");
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (e) {}
  }
  const folder = DriveApp.createFolder("StudentCards");
  props.setProperty("cardsFolderId", folder.getId());
  return folder;
}

function hashPassword_(pw) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, "srsSalt_v1_" + String(pw || ""));
  return bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, "0")).join("");
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function normalizeName_(v) {
  return String(v || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function normalizeEmail_(v) {
  return String(v || "").trim().toLowerCase();
}

function normalizeHost_(v) {
  return String(v || "default").trim().toLowerCase();
}

function safeParse_(v, fallback) {
  if (v === null || v === undefined) return fallback;
  try {
    const parsed = JSON.parse(v);
    return parsed === undefined ? fallback : parsed;
  } catch (e) { return fallback; }
}

function generateOtp_() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

function otpCacheKey_(purpose, email, host) {
  return "otp_" + purpose + "_" + normalizeEmail_(email) + "_" + normalizeHost_(host);
}

function sendOtpEmail_(email, code, purpose) {
  const subject = purpose === "admin"
    ? "Admin Login Verification Code"
    : "Verify your email — Student Result Portal";
  const body =
    "Your verification code is: " + code + "\n\n" +
    "This code expires in 10 minutes. If you did not request this, you can safely ignore this email.";
  MailApp.sendEmail(email, subject, body);
}

function doGet(e) {
  try {
    const action = (e.parameter.action || "").trim();
    if (action === "meta") return jsonOut_(getMeta_());
    if (action === "lookup") return jsonOut_(lookupStudent_(e.parameter));
    if (action === "semCompare") return jsonOut_(semCompare_(e.parameter));
    if (action === "paymentSettings") return jsonOut_(getPaymentSettings_());
    if (action === "ping") return jsonOut_({ ok: true, time: new Date().toISOString() });
    if (action === "diag") return jsonOut_(diag_());
    return jsonOut_({ ok: false, error: "unknown_action" });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    let body = {};
    try { body = JSON.parse(e.postData.contents); } catch (parseErr) { body = e.parameter; }
    const action = (body.action || "").trim();

    const deviceGuard = checkParentDeviceRequest_(body);
    if (deviceGuard) return jsonOut_(deviceGuard);

    if (action === "publish") return jsonOut_(publishData_(body));
    if (action === "publishSection") return jsonOut_(publishSection_(body));
    if (action === "setComment") return jsonOut_(setCommentStatus_(body));
    if (action === "verifyPassword") return jsonOut_(verifyAdminPassword_(body));
    if (action === "requestAdminCode") return jsonOut_(requestAdminCode_(body));
    if (action === "verifyAdminCode") return jsonOut_(verifyAdminCode_(body));

    // Result Lookups
    if (action === "parentLookupResult" || action === "lookup") return jsonOut_(lookupStudent_(body));
    if (action === "parentVerifyPaymentStudent") return jsonOut_(verifyPaymentStudent_(body));
    if (action === "semCompare") return jsonOut_(semCompare_(body));

    // Accounts (Host-aware)
    if (action === "requestSignupCode") return jsonOut_(requestSignupCode_(body));
    if (action === "parentSignup" || action === "signup") return jsonOut_(handleParentSignup_(body));
    if (action === "parentLogin" || action === "login") return jsonOut_(handleParentLogin_(body));
    if (action === "deleteAccount") return jsonOut_(deleteAccount_(body));

    // Parent Portal Payments
    if (action === "parentSubmitPayment") return jsonOut_(handleParentSubmitPayment_(body));
    if (action === "parentCheckPaymentStatus") return jsonOut_(handleParentCheckPaymentStatus_(body));
    if (action === "parentListMyPayments") return jsonOut_(handleParentListMyPayments_(body));
    if (action === "parentGetNotices") return jsonOut_(handleParentGetNotices_(body));
    if (action === "parentMarkNoticesRead") return jsonOut_(handleParentMarkNoticesRead_(body));

    // Payment Admin
    if (action === "adminSavePaymentAmount") return jsonOut_(adminSavePaymentAmount_(body));
    if (action === "adminAddBankAccount") return jsonOut_(adminAddBankAccount_(body));
    if (action === "adminDeleteBankAccount") return jsonOut_(adminDeleteBankAccount_(body));
    if (action === "adminAddNotice") return jsonOut_(adminAddNotice_(body));
    if (action === "adminDeleteNotice") return jsonOut_(adminDeleteNotice_(body));
    if (action === "adminListPayments") return jsonOut_(adminListPayments_(body));
    if (action === "adminUpdatePaymentStatus") return jsonOut_(adminUpdatePaymentStatus_(body));
    if (action === "requestPaymentAdminCode") return jsonOut_(requestPaymentAdminCode_(body));
    if (action === "verifyPaymentAdminCode") return jsonOut_(verifyPaymentAdminCode_(body));
    if (action === "verifyPaymentAdminPassword") return jsonOut_(verifyPaymentAdminPassword_(body));

    // Cards
    if (action === "saveCardTemplate") return jsonOut_(saveCardTemplate_(body));
    if (action === "generateCards") return jsonOut_(generateCards_(body));
    if (action === "claimCard") return jsonOut_(claimCard_(body));

    if (action === "requestPromotionCode") return jsonOut_(requestPromotionCode_(body));
    if (action === "verifyPromotionCode") return jsonOut_(verifyPromotionCode_(body));
    if (action === "submitPromotionRegistration") return jsonOut_(submitPromotionRegistration_(body));
    if (action === "adminListPromotions") return jsonOut_(adminListPromotions_(body));
    if (action === "adminListPromotionGroups") return jsonOut_(adminListPromotionGroups_(body));

    return jsonOut_({ ok: false, error: "unknown_action" });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function diag_() {
  const sh = getSheet_();
  const lastRow = sh.getLastRow();
  return {
    ok: true,
    spreadsheetUrl: sh.getParent().getUrl(),
    sheetName: sh.getName(),
    studentRowCount: Math.max(0, lastRow - 1)
  };
}

function getMeta_() {
  const props = PropertiesService.getScriptProperties();
  let meta = {};
  try { meta = JSON.parse(props.getProperty("meta") || "{}"); } catch (e) { meta = {}; }
  return {
    ok: true,
    academicYearsList: meta.academicYearsList || [],
    currentAcademicYear: meta.currentAcademicYear || "",
    gradeSectionsData: meta.gradeSectionsData || {},
    sectionRegistry: meta.sectionRegistry || {}
  };
}

function requestAdminCode_(body) {
  const email = normalizeEmail_(body.email);
  const adminEmail = normalizeEmail_(getAdminEmail_());
  if (!adminEmail) return { ok: true, status: "not_configured" };
  if (!email || email !== adminEmail) return { ok: true, status: "invalid_email" };

  const code = generateOtp_();
  CacheService.getScriptCache().put(otpCacheKey_("admin", email, "admin"), code, 600);
  try {
    sendOtpEmail_(email, code, "admin");
  } catch (e) {
    return { ok: true, status: "send_failed", detail: String(e) };
  }
  return { ok: true, status: "sent" };
}

function verifyAdminCode_(body) {
  const email = normalizeEmail_(body.email);
  const code = String(body.code || "").trim();
  const adminEmail = normalizeEmail_(getAdminEmail_());
  if (!adminEmail || !email || email !== adminEmail) return { ok: true, status: "invalid_email" };

  const cache = CacheService.getScriptCache();
  const key = otpCacheKey_("admin", email, "admin");
  const stored = cache.get(key);
  if (!stored || stored !== code) return { ok: true, status: "invalid_code" };

  cache.remove(key);
  return { ok: true, status: "verified" };
}

const ADMIN_PW_MAX_ATTEMPTS = 2;
const ADMIN_PW_LOCKOUT_MS = 15 * 60 * 1000;

function verifyAdminPassword_(body) {
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();

  const lockUntilStr = props.getProperty("adminLockUntil");
  if (lockUntilStr) {
    const lockUntil = Number(lockUntilStr);
    if (now < lockUntil) {
      return { ok: true, valid: false, locked: true, lockUntil: lockUntil };
    }
    props.deleteProperty("adminLockUntil");
    props.deleteProperty("adminFailedAttempts");
  }

  const correct = String(body.password || "") === getAdminPassword_();
  if (correct) {
    props.deleteProperty("adminFailedAttempts");
    return { ok: true, valid: true };
  }

  const attempts = Number(props.getProperty("adminFailedAttempts") || "0") + 1;
  if (attempts >= ADMIN_PW_MAX_ATTEMPTS) {
    const lockUntil = now + ADMIN_PW_LOCKOUT_MS;
    props.setProperty("adminLockUntil", String(lockUntil));
    props.deleteProperty("adminFailedAttempts");
    const adminEmail = getAdminEmail_();
    if (adminEmail) {
      try {
        MailApp.sendEmail(
          adminEmail,
          "⚠️ Admin login locked — repeated wrong password",
          "Someone entered the wrong admin password " + ADMIN_PW_MAX_ATTEMPTS + " times in a row."
        );
      } catch (e) {}
    }
    return { ok: true, valid: false, locked: true, lockUntil: lockUntil };
  }

  props.setProperty("adminFailedAttempts", String(attempts));
  return { ok: true, valid: false, locked: false };
}

/* ==========================================================================
 * ACCOUNTS & AUTHENTICATION (በ Host የተነጠለ)
 * ========================================================================== */
function findAccountRow_(email, host) {
  const sh = getAccountsSheet_();
  const values = sh.getDataRange().getValues();
  const wantedEmail = normalizeEmail_(email);
  const wantedHost = normalizeHost_(host);

  for (let i = 1; i < values.length; i++) {
    const rowEmail = normalizeEmail_(values[i][0]);
    // አሮጌ ዴታዎች Host ሳይኖራቸው ከቀሩ ከኢሜይል ጋር ብቻ ይዛመዳል፣ ካለ ግን በኢሜይል እና በሆስት ይለያል
    const rowHost = normalizeHost_(values[i][1]);
    
    if (rowEmail === wantedEmail && (rowHost === wantedHost || rowHost === "default" || !values[i][1])) {
      return { rowIndex: i + 1, row: values[i] };
    }
  }
  return null;
}

function requestSignupCode_(body) {
  const email = normalizeEmail_(body.email);
  const host = normalizeHost_(body.host || body.origin || "default");

  if (!email || email.indexOf("@") === -1 || email.indexOf(".") === -1) {
    return { ok: true, status: "bad_email" };
  }
  
  const existingAcc = findAccountRow_(email, host);
  if (existingAcc) {
    return { ok: true, status: "already_registered" };
  }

  const code = generateOtp_();
  CacheService.getScriptCache().put(otpCacheKey_("signup", email, host), code, 600);
  try {
    sendOtpEmail_(email, code, "signup");
  } catch (e) {
    return { ok: true, status: "send_failed", detail: String(e) };
  }
  return { ok: true, status: "sent" };
}

function handleParentSignup_(body) {
  const email = normalizeEmail_(body.email);
  const host = normalizeHost_(body.host || body.origin || "default");
  const password = String(body.password || "");

  if (!email || email.indexOf("@") === -1) return { ok: true, status: "bad_email" };
  if (password.length < 4) return { ok: true, status: "weak_password" };
  
  const existingAcc = findAccountRow_(email, host);
  if (existingAcc) {
    const sh = getAccountsSheet_();
    sh.getRange(existingAcc.rowIndex, 3).setValue(hashPassword_(password)); // PasswordHash column
    return { 
      ok: true, 
      status: "created", 
      boundCode: String(existingAcc.row[3] || ""),
      boundName: String(existingAcc.row[4] || "")
    };
  }

  // ሠንጠረዥ፡ [Email, Host, PasswordHash, BoundCode, BoundName, BoundYear, BoundGrade, BoundSection, CreatedAt]
  getAccountsSheet_().appendRow([email, host, hashPassword_(password), "", "", "", "", "", new Date().toISOString()]);
  return { ok: true, status: "created" };
}

function handleParentLogin_(body) {
  const email = normalizeEmail_(body.email);
  const host = normalizeHost_(body.host || body.origin || "default");
  const password = String(body.password || "");
  if (!email || !password) return { ok: true, status: "bad_request" };

  const found = findAccountRow_(email, host);
  if (!found) return { ok: true, status: "no_account" };
  if (String(found.row[2]) !== hashPassword_(password)) return { ok: true, status: "wrong_password" };

  const deviceId = String(body.deviceId || "").trim();
  if (!deviceId) return { ok: true, status: "device_required" };
  const deviceCol = deviceIdColumn_();
  const savedDeviceId = String(found.row[deviceCol - 1] || "").trim();
  if (savedDeviceId && savedDeviceId !== deviceId) return { ok: true, status: "device_mismatch" };
  if (!savedDeviceId) getAccountsSheet_().getRange(found.rowIndex, deviceCol).setValue(deviceId);

  return {
    ok: true,
    status: "ok",
    boundCode: String(found.row[3] || ""),
    boundName: String(found.row[4] || "")
  };
}

function deleteAccount_(body) {
  const email = normalizeEmail_(body.email);
  const host = normalizeHost_(body.host || body.origin || "default");
  const password = String(body.password || "");
  const found = findAccountRow_(email, host);
  if (!found) return { ok: true, status: "no_account" };
  if (String(found.row[2]) !== hashPassword_(password)) return { ok: true, status: "wrong_password" };

  getAccountsSheet_().deleteRow(found.rowIndex);
  return { ok: true, status: "deleted" };
}

function checkAccountBinding_(email, hostParam, password, code, name) {
  const em = normalizeEmail_(email);
  const host = normalizeHost_(hostParam || "default");
  const found = findAccountRow_(em, host);
  if (!found) return { allowed: false, status: "no_account" };
  if (String(found.row[2]) !== hashPassword_(password)) return { allowed: false, status: "wrong_password" };

  const wantCode = String(code || "").trim();
  const wantName = normalizeName_(name || "");
  if (!wantCode || !wantName) return { allowed: false, status: "bad_request" };

  // One parent account may be used for more than one child. Do not reject a
  // second child merely because the account's legacy single-child fields are set.
  // However, the same student code + normalized name cannot be claimed by a
  // different account on the same host.
  const sh = getAccountsSheet_();
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const rowEmail = normalizeEmail_(values[i][0]);
    const rowHost = normalizeHost_(values[i][1]);
    if (rowEmail === em && rowHost === host) continue;
    const rowCode = String(values[i][3] || "").trim();
    const rowName = normalizeName_(values[i][4] || "");
    if (rowHost === host && rowCode === wantCode && rowName === wantName) {
      return { allowed: false, status: "student_already_claimed" };
    }
  }

  const existingBoundCode = String(found.row[3] || "").trim();
  return {
    allowed: true,
    rowIndex: found.rowIndex,
    // Only bind the legacy fields for the first child. Never overwrite the
    // first child's binding when the same parent adds another child.
    fresh: !existingBoundCode
  };
}

function bindAccountToStudent_(rowIndex, code, name, year, grade, section) {
  const sh = getAccountsSheet_();
  sh.getRange(rowIndex, 4).setValue(String(code || ""));
  sh.getRange(rowIndex, 5).setValue(String(name || ""));
  sh.getRange(rowIndex, 6).setValue(String(year || ""));
  sh.getRange(rowIndex, 7).setValue(String(grade || ""));
  sh.getRange(rowIndex, 8).setValue(String(section || ""));
}

/* ==========================================================================
 * RESULT LOOKUP
 * ========================================================================== */
function lookupStudent_(p) {
  const year = String(p.year || "");
  const semester = String(p.semester || "");
  const grade = String(p.grade || "");
  const stream = String(p.stream || "");
  const section = String(p.section || "").toUpperCase();
  const code = String(p.code || "").trim();
  const wantedName = normalizeName_(p.name || "");
  const host = p.host || p.origin || "default";

  if (!year || !semester || !grade || !section || !code || !wantedName) {
    return { ok: false, status: "bad_request" };
  }

  const guard = checkAccountBinding_(p.email, host, p.password, code, p.name);
  if (!guard.allowed) return { ok: true, status: guard.status };

  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const matches = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rCode = String(row[5]).trim();
    const rName = normalizeName_(row[6]);
    if (rCode !== code) continue;
    if (!rName.includes(wantedName)) continue;
    matches.push({
      year: String(row[0]), semester: String(row[1]), grade: String(row[2]),
      stream: String(row[3]), section: String(row[4]).toUpperCase(),
      code: rCode, name: row[6], scores: safeParse_(row[7], []),
      commented: row[8] === true || row[8] === "TRUE" || row[8] === "true",
      subjects: safeParse_(row[9], []),
      breakdown: safeParse_(row[10], []),
      assessmentConfig: safeParse_(row[11], null)
    });
  }

  if (matches.length === 0) return { ok: true, status: "not_found" };

  const exact = matches.find(m =>
    m.year === year && m.semester === semester && m.grade === grade && m.section === section &&
    (parseInt(grade, 10) < 11 || m.stream.toLowerCase() === stream.toLowerCase())
  );

  if (!exact) {
    const uniqueLocations = [];
    matches.forEach(m => {
      const label = `${m.year}|${m.semester}|${m.grade}|${m.stream}|${m.section}`;
      if (!uniqueLocations.some(u => u.label === label)) uniqueLocations.push({ label, loc: m });
    });
    return {
      ok: true,
      status: "location_mismatch",
      locations: uniqueLocations.map(u => ({
        year: u.loc.year, semester: u.loc.semester, grade: u.loc.grade,
        stream: u.loc.stream, section: u.loc.section
      }))
    };
  }

  if (exact.commented) return { ok: true, status: "blocked" };

  if (guard.fresh) bindAccountToStudent_(guard.rowIndex, exact.code, exact.name, exact.year, exact.grade, exact.section);

  return {
    ok: true,
    status: "found",
    student: {
      name: exact.name, code: exact.code, grade: exact.grade, stream: exact.stream,
      section: exact.section, semester: exact.semester, scores: exact.scores, subjects: exact.subjects,
      breakdown: exact.breakdown, assessmentConfig: exact.assessmentConfig
    }
  };
}

/* ==========================================================================
 * PAYMENT MANAGEMENT
 * ========================================================================== */
// Notices are stored separately in small Script Properties chunks. Apps Script limits
// each property value to about 9 KB, so storing long notices inside paymentSettings
// can make publishing fail even when the UI looks normal.
const NOTICE_CHUNK_SIZE_ = 7000;
function readAllNotices_() {
  const props = PropertiesService.getScriptProperties();
  const keys = props.getProperties();
  const count = Math.max(0, parseInt(keys.noticeChunkCount || '0', 10) || 0);
  if (count > 0) {
    let raw = '';
    for (let i = 0; i < count; i++) raw += props.getProperty('noticeChunk_' + i) || '';
    const parsed = safeParse_(raw, []);
    return Array.isArray(parsed) ? parsed : [];
  }
  // One-time migration of notices saved by the previous version.
  const legacy = safeParse_(props.getProperty('paymentSettings'), { notices: [] });
  const oldNotices = legacy && Array.isArray(legacy.notices) ? legacy.notices : [];
  if (oldNotices.length) {
    writeAllNotices_(oldNotices);
    delete legacy.notices;
    props.setProperty('paymentSettings', JSON.stringify(legacy));
  }
  return oldNotices;
}
function writeAllNotices_(notices) {
  const props = PropertiesService.getScriptProperties();
  const serialized = JSON.stringify(Array.isArray(notices) ? notices : []);
  const chunks = [];
  for (let i = 0; i < serialized.length; i += NOTICE_CHUNK_SIZE_) chunks.push(serialized.slice(i, i + NOTICE_CHUNK_SIZE_));
  const existing = Math.max(0, parseInt(props.getProperty('noticeChunkCount') || '0', 10) || 0);
  chunks.forEach((chunk, i) => props.setProperty('noticeChunk_' + i, chunk));
  for (let i = chunks.length; i < existing; i++) props.deleteProperty('noticeChunk_' + i);
  props.setProperty('noticeChunkCount', String(chunks.length));
  return Array.isArray(notices) ? notices : [];
}
function paymentSettingsWithNotices_() {
  const props = PropertiesService.getScriptProperties();
  const settings = safeParse_(props.getProperty('paymentSettings'), { requiredAmount: 1500, bankAccounts: [] });
  settings.notices = readAllNotices_();
  return settings;
}
function savePaymentSettingsWithoutNotices_(settings) {
  const copy = Object.assign({}, settings || {});
  delete copy.notices;
  PropertiesService.getScriptProperties().setProperty('paymentSettings', JSON.stringify(copy));
}

function getPaymentSettings_() {
  return { ok: true, settings: paymentSettingsWithNotices_() };
}

function verifyPaymentStudent_(body) {
  const host = normalizeHost_(body.host || body.origin || "default");
  const email = normalizeEmail_(body.email);
  const password = String(body.password || "");
  const account = findAccountRow_(email, host);
  if (!account) return { ok: true, status: "no_account" };
  if (String(account.row[2]) !== hashPassword_(password)) return { ok: true, status: "wrong_password" };
  const year = String(body.ecYear || body.year || "").trim();
  const grade = String(body.grade || "").trim();
  const stream = String(body.stream || "").trim().toLowerCase();
  const section = String(body.section || "").trim().toUpperCase();
  const code = String(body.code || "").trim();
  const wantedName = normalizeName_(body.name || "");
  if (!year || !grade || !section || !/^\d{6}$/.test(code) || !wantedName) return { ok: false, status: "bad_request" };
  const values = getSheet_().getDataRange().getValues();
  const found = values.slice(1).some(row => {
    const rowYear = String(row[0] || '').trim();
    const rowGrade = String(row[2] || '').trim();
    const rowStream = String(row[3] || '').trim().toLowerCase();
    const rowSection = String(row[4] || '').trim().toUpperCase();
    const rowCode = String(row[5] || '').trim();
    const rowName = normalizeName_(row[6] || '');
    return rowYear === year && rowGrade === grade && rowSection === section && rowCode === code && rowName === wantedName && (parseInt(grade, 10) < 11 || rowStream === stream);
  });
  return found ? { ok: true, status: "found" } : { ok: true, status: "not_found" };
}

function handleParentSubmitPayment_(body) {
  const host = normalizeHost_(body.host || body.origin || "default");
  const email = normalizeEmail_(body.email);
  const password = String(body.password || "");
  const account = findAccountRow_(email, host);
  if (!account) return { ok: true, status: "no_account" };
  if (String(account.row[2]) !== hashPassword_(password)) return { ok: true, status: "wrong_password" };
  const children = Array.isArray(body.children) ? body.children : [body];
  if (!children.length) return { ok: false, status: "bad_request" };
  const verified = [];
  for (const child of children) {
    const check = verifyPaymentStudent_(Object.assign({}, child, { email, password, host }));
    if (!check.ok || check.status !== "found") return { ok: true, status: check.status || "not_found" };
    verified.push(child);
  }
  const settings = paymentSettingsWithNotices_();
  const perChildAmount = Number(settings.requiredAmount ?? settings.amount ?? 0);
  if (!(perChildAmount > 0)) return { ok: true, status: "bad_amount" };
  const bankId = String(body.bankId || "");
  const bank = (settings.bankAccounts || []).find(b => String(b.id || b.bankId || b.name || '') === bankId);
  if (!bank && (settings.bankAccounts || []).length) return { ok: true, status: "bad_bank" };
  let photoUrl = "";
  if (body.photoBase64) {
    const folderName = "School_Payment_Receipts";
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const blob = Utilities.newBlob(Utilities.base64Decode(body.photoBase64), body.photoMime || "image/jpeg", `Receipt_${Date.now()}`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    photoUrl = file.getUrl();
  }
  const sheet = getPaymentsSheet_();
  const ids = [];
  verified.forEach(child => {
    const paymentId = "PAY-" + Math.floor(100000 + Math.random() * 900000);
    sheet.appendRow([paymentId, email, host, String(child.code), String(child.name), String(child.ecYear || child.year), String(child.grade), String(child.section), bankId, perChildAmount, photoUrl, "pending", "", new Date().toISOString()]);
    ids.push(paymentId);
  });
  return { ok: true, status: "submitted", paymentIds: ids, count: ids.length, totalAmount: perChildAmount * ids.length };
}

function handleParentCheckPaymentStatus_(body) {
  const host = body.host || body.origin || "default";
  const guard = checkAccountBinding_(body.email, host, body.password, body.code, body.name);
  if (!guard.allowed) return { ok: true, status: guard.status };

  const sheet = getPaymentsSheet_();
  const rows = sheet.getDataRange().getValues();

  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][3]) === String(body.code) && String(rows[i][5]) === String(body.ecYear) && normalizeHost_(rows[i][2]) === normalizeHost_(host)) {
      return {
        ok: true,
        status: "found",
        payment: {
          paymentId: rows[i][0],
          amount: rows[i][9],
          bankName: rows[i][8],
          status: rows[i][11],
          adminNote: rows[i][12]
        }
      };
    }
  }
  return { ok: true, status: "none" };
}

function handleParentListMyPayments_(body) {
  const host = body.host || body.origin || "default";
  const found = findAccountRow_(body.email, host);
  if (!found || String(found.row[2]) !== hashPassword_(body.password)) return { ok: true, status: "wrong_password" };

  const sheet = getPaymentsSheet_();
  const rows = sheet.getDataRange().getValues();
  let list = [];

  for (let i = rows.length - 1; i >= 1; i--) {
    if (normalizeEmail_(rows[i][1]) === normalizeEmail_(body.email) && normalizeHost_(rows[i][2]) === normalizeHost_(host)) {
      list.push({
        paymentId: rows[i][0], code: rows[i][3], name: rows[i][4], ecYear: rows[i][5],
        grade: rows[i][6], section: rows[i][7], bankName: rows[i][8], amount: rows[i][9],
        photoUrl: rows[i][10], status: rows[i][11], adminNote: rows[i][12], submittedAt: rows[i][13]
      });
    }
  }
  return { ok: true, payments: list };
}

function noticeReadStateKey_(email, host) {
  // Store read state separately for each parent account (email + host).
  return "noticeRead_" + hashPassword_(normalizeEmail_(email) + "|" + normalizeHost_(host));
}

function getAuthenticatedNoticeAccount_(body) {
  const email = normalizeEmail_(body.email);
  const host = normalizeHost_(body.host || body.origin || "default");
  const password = String(body.password || "");
  if (!email || !password) return { error: { ok: false, status: "bad_request" } };
  const found = findAccountRow_(email, host);
  if (!found || String(found.row[2] || "") !== hashPassword_(password)) {
    return { error: { ok: false, status: "wrong_password" } };
  }
  return { email: email, host: host };
}

function handleParentGetNotices_(body) {
  const auth = getAuthenticatedNoticeAccount_(body);
  if (auth.error) return auth.error;

  const props = PropertiesService.getScriptProperties();
  const allNotices = readAllNotices_();
  // Parent Portal sends the academic years of children they have selected/linked.
  // Strict academic-year filtering: unassigned legacy notices are not broadcast to every year.
  const requestedYears = Array.isArray(body.academicYears) ? body.academicYears.map(y => String(y || '').trim()).filter(Boolean) : [];
  const visibleNotices = requestedYears.length
    ? allNotices.filter(n => n && n.academicYear && requestedYears.includes(String(n.academicYear)))
    : [];
  const readIds = safeParse_(props.getProperty(noticeReadStateKey_(auth.email, auth.host)), []);
  const readSet = {};
  (Array.isArray(readIds) ? readIds : []).forEach(id => { readSet[String(id)] = true; });

  const notices = visibleNotices.map(n => Object.assign({}, n, {
    read: !!(n && n.id && readSet[String(n.id)])
  }));
  const unreadCount = notices.filter(n => !n.read).length;
  return { ok: true, notices: notices, unreadCount: unreadCount };
}

function handleParentMarkNoticesRead_(body) {
  const auth = getAuthenticatedNoticeAccount_(body);
  if (auth.error) return auth.error;

  const props = PropertiesService.getScriptProperties();
  const key = noticeReadStateKey_(auth.email, auth.host);
  const existing = safeParse_(props.getProperty(key), []);
  const merged = {};
  (Array.isArray(existing) ? existing : []).forEach(id => { merged[String(id)] = true; });

  const requestedIds = Array.isArray(body.noticeIds) ? body.noticeIds : [];
  const validIds = {};
  readAllNotices_().forEach(n => {
    if (n && n.id) validIds[String(n.id)] = true;
  });
  let markedCount = 0;
  requestedIds.forEach(id => {
    const cleanId = String(id || "").trim();
    if (cleanId && validIds[cleanId] && !merged[cleanId]) {
      merged[cleanId] = true;
      markedCount++;
    }
  });

  props.setProperty(key, JSON.stringify(Object.keys(merged)));
  return { ok: true, markedCount: markedCount };
}

/* ==========================================================================
 * PAYMENT ADMIN AUTH (Finance admin login)
 * Script Properties ውስጥ PAYMENT_ADMIN_EMAIL እና PAYMENT_ADMIN_PASSWORD ካልተቀመጡ
 * ዋናው ADMIN_EMAIL / ADMIN_PASSWORD ጥቅም ላይ ይውላል።
 * ========================================================================== */
function getPaymentAdminEmail_() {
  return PropertiesService.getScriptProperties().getProperty("PAYMENT_ADMIN_EMAIL") || getAdminEmail_();
}

function getPaymentAdminPassword_() {
  return PropertiesService.getScriptProperties().getProperty("PAYMENT_ADMIN_PASSWORD") || getAdminPassword_();
}

function isPaymentAdminPassword_(pw) {
  const p = String(pw || "");
  if (!p) return false;
  return p === getPaymentAdminPassword_() || p === getAdminPassword_();
}

function requestPaymentAdminCode_(body) {
  const email = normalizeEmail_(body.email);
  const adminEmail = normalizeEmail_(getPaymentAdminEmail_());
  if (!adminEmail) return { ok: true, status: "not_configured" };
  if (!email || email !== adminEmail) return { ok: true, status: "invalid_email" };

  const code = generateOtp_();
  CacheService.getScriptCache().put(otpCacheKey_("payadmin", email, "admin"), code, 600);
  try {
    sendOtpEmail_(email, code, "admin");
  } catch (e) {
    return { ok: true, status: "send_failed", detail: String(e) };
  }
  return { ok: true, status: "sent" };
}

function verifyPaymentAdminCode_(body) {
  const email = normalizeEmail_(body.email);
  const code = String(body.code || "").trim();
  const adminEmail = normalizeEmail_(getPaymentAdminEmail_());
  if (!adminEmail || !email || email !== adminEmail) return { ok: true, status: "invalid_email" };

  const cache = CacheService.getScriptCache();
  const key = otpCacheKey_("payadmin", email, "admin");
  const stored = cache.get(key);
  if (!stored || stored !== code) return { ok: true, status: "invalid_code" };

  cache.remove(key);
  return { ok: true, status: "verified" };
}

function verifyPaymentAdminPassword_(body) {
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();

  const lockUntilStr = props.getProperty("payAdminLockUntil");
  if (lockUntilStr) {
    const lockUntil = Number(lockUntilStr);
    if (now < lockUntil) {
      return { ok: true, valid: false, locked: true, lockUntil: lockUntil };
    }
    props.deleteProperty("payAdminLockUntil");
    props.deleteProperty("payAdminFailedAttempts");
  }

  if (isPaymentAdminPassword_(body.password)) {
    props.deleteProperty("payAdminFailedAttempts");
    return { ok: true, valid: true };
  }

  const attempts = Number(props.getProperty("payAdminFailedAttempts") || "0") + 1;
  if (attempts >= ADMIN_PW_MAX_ATTEMPTS) {
    const lockUntil = now + ADMIN_PW_LOCKOUT_MS;
    props.setProperty("payAdminLockUntil", String(lockUntil));
    props.deleteProperty("payAdminFailedAttempts");
    return { ok: true, valid: false, locked: true, lockUntil: lockUntil };
  }

  props.setProperty("payAdminFailedAttempts", String(attempts));
  return { ok: true, valid: false, locked: false };
}

/* ==========================================================================
 * PAYMENT ADMIN BACKEND
 * ========================================================================== */
function adminSavePaymentAmount_(body) {
  if (!isPaymentAdminPassword_(body.password)) return { ok: false, error: "wrong_password" };
  const props = PropertiesService.getScriptProperties();
  let settings = safeParse_(props.getProperty("paymentSettings"), { requiredAmount: 1500, bankAccounts: [] });
  settings.requiredAmount = parseFloat(body.amount) || 0;
  savePaymentSettingsWithoutNotices_(settings);
  settings.notices = readAllNotices_();
  return { ok: true, settings: settings };
}

function adminAddBankAccount_(body) {
  if (!isPaymentAdminPassword_(body.password)) return { ok: false, error: "wrong_password" };
  const props = PropertiesService.getScriptProperties();
  let settings = safeParse_(props.getProperty("paymentSettings"), { requiredAmount: 1500, bankAccounts: [] });
  const newBank = {
    id: "BANK-" + Date.now(),
    bankName: body.bankName,
    accountNumber: body.accountNumber,
    accountHolder: body.accountHolder
  };
  settings.bankAccounts = Array.isArray(settings.bankAccounts) ? settings.bankAccounts : [];
  settings.bankAccounts.push(newBank);
  savePaymentSettingsWithoutNotices_(settings);
  settings.notices = readAllNotices_();
  return { ok: true, settings: settings };
}

function adminDeleteBankAccount_(body) {
  if (!isPaymentAdminPassword_(body.password)) return { ok: false, error: "wrong_password" };
  const props = PropertiesService.getScriptProperties();
  let settings = safeParse_(props.getProperty("paymentSettings"), { requiredAmount: 1500, bankAccounts: [] });
  settings.bankAccounts = (settings.bankAccounts || []).filter(b => b.id !== body.id);
  savePaymentSettingsWithoutNotices_(settings);
  settings.notices = readAllNotices_();
  return { ok: true, settings: settings };
}

function adminAddNotice_(body) {
  const password = String(body.password || "");
  if (!isPaymentAdminPassword_(password) && password !== getAdminPassword_()) return { ok: false, error: "wrong_password" };
  const title = String(body.title || "").trim();
  const message = String(body.message || "").trim();
  const academicYear = String(body.academicYear || "").trim();
  if (!title || !message) return { ok: false, error: "missing_title_or_message" };
  if (!academicYear) return { ok: false, error: "missing_academic_year" };
  const newNotice = {
    id: "NOT-" + Date.now() + "-" + Math.random().toString(36).slice(2, 7),
    title: title, message: message, academicYear: academicYear,
    createdAt: new Date().toISOString()
  };
  try {
    const notices = readAllNotices_();
    notices.push(newNotice);
    writeAllNotices_(notices);
  } catch (e) {
    return { ok: false, error: "notice_save_failed: " + String(e) };
  }
  // Email errors must not undo a notice that was saved successfully.
  try { sendNoticeEmailsToRegisteredAccounts_(newNotice); }
  catch (e) { console.warn("School notice email failed: " + e); }
  return { ok: true, settings: paymentSettingsWithNotices_(), notice: newNotice };
}

function sendNoticeEmailsToRegisteredAccounts_(notice) {
  const sh = getAccountsSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const values = sh.getRange(2, 1, lastRow - 1, Math.min(sh.getLastColumn(), 9)).getValues();
  const recipients = [];
  const seen = {};

  values.forEach(row => {
    const email = normalizeEmail_(row[0]);
    const host = normalizeHost_(row[1]);
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return;
    // Send only to registered Parent Portal and Host Portal accounts.
    if (host !== "parent" && host !== "host") return;
    if (seen[email]) return;
    seen[email] = true;
    recipients.push(email);
  });

  if (!recipients.length) return;

  const adminEmail = normalizeEmail_(getAdminEmail_());
  const to = adminEmail || recipients[0];
  const bcc = adminEmail ? recipients.join(",") : recipients.slice(1).join(",");
  const subject = "📢 " + String(notice.title || "New School Notice");
  const body = String(notice.message || "") + "\n\n" +
    "SMART-SCHOOL School Notice\n" +
    "Notice time: " + new Date(notice.createdAt || new Date()).toLocaleString() + "\n\n" +
    "This message was sent to registered Parent Portal and Host Portal email accounts.";

  MailApp.sendEmail({
    to: to,
    bcc: bcc,
    subject: subject,
    body: body
  });
}

function adminDeleteNotice_(body) {
  const password = String(body.password || "");
  if (!isPaymentAdminPassword_(password) && password !== getAdminPassword_()) return { ok: false, error: "wrong_password" };
  const notices = readAllNotices_().filter(n => n.id !== body.id);
  try { writeAllNotices_(notices); }
  catch (e) { return { ok: false, error: "notice_delete_failed: " + String(e) }; }
  return { ok: true, settings: paymentSettingsWithNotices_() };
}

function adminListPayments_(body) {
  if (!isPaymentAdminPassword_(body.password)) return { ok: false, error: "wrong_password" };
  const sheet = getPaymentsSheet_();
  const rows = sheet.getDataRange().getValues();
  let list = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (body.grade && String(row[6]) !== String(body.grade)) continue;
    if (body.section && String(row[7]).toUpperCase() !== String(body.section).toUpperCase()) continue;
    if (body.status && String(row[11]).toLowerCase() !== String(body.status).toLowerCase()) continue;

    list.push({
      id: row[0], email: row[1], host: row[2], code: row[3], name: row[4], ecYear: row[5],
      grade: row[6], section: row[7], bankName: row[8], amount: row[9],
      photoUrl: row[10], status: row[11], adminNote: row[12], submittedAt: row[13]
    });
  }
  return { ok: true, payments: list };
}

function adminUpdatePaymentStatus_(body) {
  if (!isPaymentAdminPassword_(body.password)) return { ok: false, error: "wrong_password" };
  const sheet = getPaymentsSheet_();
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(body.id)) {
      sheet.getRange(i + 1, 12).setValue(body.status);
      sheet.getRange(i + 1, 13).setValue(body.note || "");
      return { ok: true };
    }
  }
  return { ok: false, error: "not_found" };
}

/* ==========================================================================
 * PUBLISH & CARDS
 * ========================================================================== */
function publishData_(body) {
  if (String(body.password || "") !== getAdminPassword_()) return { ok: false, error: "wrong_password" };

  const schoolDatabase = body.schoolDatabase || { sem1: {}, sem2: {} };
  const gradeSubjects = body.gradeSubjects || {};
  const rows = [];

  ["sem1", "sem2"].forEach(semKey => {
    const semStore = schoolDatabase[semKey] || {};
    Object.keys(semStore).forEach(dbKey => {
      const parts = dbKey.split("_");
      let year, grade, stream, section;
      if (parts.length >= 4) { year = parts[0]; grade = parts[1]; stream = parts[2]; section = parts[3]; }
      else { year = parts[0]; grade = parts[1]; stream = ""; section = parts[2]; }

      let subjects = [];
      if (parseInt(grade, 10) >= 11) subjects = (gradeSubjects[grade] && gradeSubjects[grade][stream]) || [];
      else subjects = gradeSubjects[grade] || [];

      (semStore[dbKey] || []).forEach(st => {
        rows.push([
          year, semKey, grade, stream, section,
          String(st.code || ""), String(st.name || ""),
          JSON.stringify(st.scores || []),
          !!(st.commented === true || st.status === "disabled" || st.hidden === true || st.disabled === true),
          JSON.stringify(subjects),
          JSON.stringify(st.breakdown || []),
          JSON.stringify(body.assessmentConfig || null),
          String(st.specialization || st.speciality || "")
        ]);
      });
    });
  });

  const sh = getSheet_();
  sh.clearContents();
  sh.appendRow(HEADER);
  if (rows.length) sh.getRange(2, 1, rows.length, HEADER.length).setValues(rows);

  PropertiesService.getScriptProperties().setProperty("meta", JSON.stringify({
    academicYearsList: body.academicYearsList || [],
    currentAcademicYear: body.currentAcademicYear || "",
    gradeSectionsData: body.gradeSectionsData || {},
    sectionRegistry: body.sectionRegistry || {}
  }));

  return { ok: true, rowsWritten: rows.length };
}

function publishSection_(body) {
  if (String(body.password || "") !== getAdminPassword_()) return { ok: false, error: "wrong_password" };

  const year = String(body.year || "");
  const semester = String(body.semester || "");
  const grade = String(body.grade || "");
  const stream = String(body.stream || "");
  const section = String(body.section || "").toUpperCase();
  const students = Array.isArray(body.students) ? body.students : [];
  const subjects = Array.isArray(body.subjects) ? body.subjects : [];

  if (!year || !semester || !grade || !section) return { ok: false, error: "bad_request" };

  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const keptRows = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const isThisSection =
      String(row[0]) === year && String(row[1]) === semester && String(row[2]) === grade &&
      String(row[4]).toUpperCase() === section &&
      (parseInt(grade, 10) < 11 || String(row[3]).toLowerCase() === stream.toLowerCase());
    if (!isThisSection) keptRows.push(row);
  }

  // 2ኛ semester ላይ 1ኛ semester የነበረውን 6-digit code እንደገና አንፈጥርም።
  // ስም በመመሳሰል የ1ኛ semester code ይጠቀማል።
  const sem1CodesByName = {};
  if (semester === "sem2" || semester === "2") {
    values.forEach(row => {
      if (String(row[0]) !== year || !(String(row[1]) === "sem1" || String(row[1]) === "1")) return;
      if (String(row[2]) !== grade) return;
      if (String(row[4]).toUpperCase() !== section) return;
      if (parseInt(grade, 10) >= 11 && String(row[3]).toLowerCase() !== stream.toLowerCase()) return;
      sem1CodesByName[normalizeName_(row[6])] = String(row[5] || "").trim();
    });
  }

  students.forEach(st => {
    const studentName = String(st.name || "").trim();
    const stableCode = (semester === "sem2" || semester === "2")
      ? (sem1CodesByName[normalizeName_(studentName)] || String(st.code || ""))
      : String(st.code || "");
    keptRows.push([
      year, semester, grade, stream, section,
      stableCode, studentName,
      JSON.stringify(st.scores || []),
      !!(st.commented === true || st.status === "disabled" || st.hidden === true || st.disabled === true),
      JSON.stringify(subjects),
      JSON.stringify(st.breakdown || []),
      JSON.stringify(body.assessmentConfig || null),
      String(st.specialization || st.speciality || "")
    ]);
  });

  sh.clearContents();
  sh.appendRow(HEADER);
  if (keptRows.length) sh.getRange(2, 1, keptRows.length, HEADER.length).setValues(keptRows);

  const props = PropertiesService.getScriptProperties();
  let meta = {};
  try { meta = JSON.parse(props.getProperty("meta") || "{}"); } catch (e) { meta = {}; }
  meta.academicYearsList = Array.isArray(meta.academicYearsList) ? meta.academicYearsList : [];
  if (!meta.academicYearsList.includes(year)) meta.academicYearsList.push(year);
  meta.currentAcademicYear = meta.currentAcademicYear || year;
  meta.gradeSectionsData = meta.gradeSectionsData || {};
  meta.sectionRegistry = meta.sectionRegistry || {};
  const regKey = `${year}_${semester}_${grade}_${parseInt(grade, 10) >= 11 ? stream : "none"}`;
  if (!Array.isArray(meta.sectionRegistry[regKey])) meta.sectionRegistry[regKey] = [];
  if (!meta.sectionRegistry[regKey].includes(section)) meta.sectionRegistry[regKey].push(section);
  props.setProperty("meta", JSON.stringify(meta));

  return { ok: true, rowsWritten: students.length };
}

function setCommentStatus_(body) {
  if (String(body.password || "") !== getAdminPassword_()) return { ok: false, error: "wrong_password" };
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const code = String(body.code || "").trim();
  const name = normalizeName_(body.name || "");
  let updated = 0;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][5]).trim() === code && normalizeName_(values[i][6]) === name) {
      sh.getRange(i + 1, 9).setValue(!!body.commented);
      updated++;
    }
  }
  return { ok: true, updated: updated };
}

function saveCardTemplate_(body) {
  if (String(body.password || "") !== getAdminPassword_()) return { ok: false, error: "wrong_password" };
  const base64 = String(body.templateBase64 || "");
  if (!base64) return { ok: false, error: "no_file" };

  const props = PropertiesService.getScriptProperties();
  const oldId = props.getProperty("cardTemplateDocId");
  if (oldId) {
    try { DriveApp.getFileById(oldId).setTrashed(true); } catch (e) {}
  }

  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64),
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    String(body.fileName || "CardTemplate.docx")
  );

  const file = Drive.Files.insert({ title: "CardTemplate", mimeType: MimeType.GOOGLE_DOCS }, blob);
  props.setProperty("cardTemplateDocId", file.id);
  return { ok: true, docId: file.id };
}

function normalizeLabel_(v) {
  return String(v || "").toLowerCase().replace(/[^a-z0-9\u1200-\u137F]/g, "");
}

function fillCardDoc_(doc, info) {
  const body = doc.getBody();
  const tables = body.getTables();

  const labelMap = {
    "studentsname": info.name, "studentname": info.name, "name": info.name,
    "የተማሪስም": info.name, "ስም": info.name,
    "academicyear": info.year, "የትምህርትዘመን": info.year,
    "gradeandsection": info.gradeSection, "gradesection": info.gradeSection,
    "ክፍልናሴክሽን": info.gradeSection,
    "promoteddetainedtograde": info.status, "promoteddetained": info.status,
    "code": info.code, "ኮድ": info.code,
    "id": info.code
  };

  const subjectIndex = {};
  (info.subjects || []).forEach((s, i) => { subjectIndex[normalizeLabel_(s)] = i; });

  tables.forEach(table => {
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      const nCells = row.getNumCells();
      if (nCells < 2) continue;

      const first = normalizeLabel_(row.getCell(0).getText());
      if (!first) continue;

      if (Object.prototype.hasOwnProperty.call(labelMap, first)) {
        const val = labelMap[first];
        if (val !== undefined && val !== null && String(row.getCell(1).getText()).trim() === "") {
          row.getCell(1).setText(String(val));
        }
        continue;
      }

      if (Object.prototype.hasOwnProperty.call(subjectIndex, first)) {
        const idx = subjectIndex[first];
        const s1 = info.sem1[idx];
        const s2 = info.sem2[idx];
        if (nCells > 1 && s1 !== undefined) row.getCell(1).setText(String(s1));
        if (nCells > 2 && s2 !== undefined) row.getCell(2).setText(String(s2));
        if (nCells > 3) {
          const vals = [s1, s2].filter(v => typeof v === "number");
          if (vals.length) row.getCell(3).setText((vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(1));
        }
        continue;
      }

      if (first === "total" || first === "ድምር") {
        if (nCells > 1) row.getCell(1).setText(String(info.total1));
        if (nCells > 2) row.getCell(2).setText(String(info.total2));
        if (nCells > 3) row.getCell(3).setText(String(((info.total1 + info.total2) / 2).toFixed(1)));
      } else if (first === "average" || first === "አማካይ") {
        if (nCells > 1) row.getCell(1).setText(String(info.avg1));
        if (nCells > 2) row.getCell(2).setText(String(info.avg2));
        if (nCells > 3) row.getCell(3).setText(String(((info.avg1 + info.avg2) / 2).toFixed(1)));
      }
    }
  });

  body.replaceText("\\{ተማሪ_ስም\\}", info.name);
  body.replaceText("\\{student_name\\}", info.name);
  body.replaceText("\\{code\\}", String(info.code));
  body.replaceText("\\{status\\}", info.status);

  doc.saveAndClose();
}

function avgOf_(scores) {
  const nums = (scores || []).map(n => parseFloat(n)).filter(n => Number.isFinite(n));
  if (!nums.length) return 0;
  return parseFloat((nums.reduce((a, b) => a + b, 0) / nums.length).toFixed(1));
}

function sumOf_(scores) {
  return (scores || []).map(n => parseFloat(n) || 0).reduce((a, b) => a + b, 0);
}

function generateCards_(body) {
  if (String(body.password || "") !== getAdminPassword_()) return { ok: false, error: "wrong_password" };

  const props = PropertiesService.getScriptProperties();
  const templateId = props.getProperty("cardTemplateDocId");
  if (!templateId) return { ok: false, error: "no_template" };

  const sh = getSheet_();
  const values = sh.getDataRange().getValues();

  const students = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const commented = row[8] === true || row[8] === "TRUE" || row[8] === "true";
    if (commented) continue;
    const name = String(row[6] || "").trim();
    const code = String(row[5] || "").trim();
    if (!name || !code) continue;

    const key = [row[0], row[2], row[3], row[4], code].join("|");
    if (!students[key]) {
      students[key] = {
        year: String(row[0]), grade: String(row[2]), stream: String(row[3]), section: String(row[4]).toUpperCase(),
        code: code, name: name, subjects: safeParse_(row[9], []), sem1: [], sem2: []
      };
    }
    const scores = safeParse_(row[7], []);
    if (String(row[1]) === "sem1") students[key].sem1 = scores;
    else students[key].sem2 = scores;
    const subj = safeParse_(row[9], []);
    if (subj.length) students[key].subjects = subj;
  }

  const folder = getCardsFolder_();
  const cardsSheet = getCardsSheet_();
  cardsSheet.clearContents();
  cardsSheet.appendRow(CARDS_HEADER);

  const cardRows = [];
  let completeCount = 0, notCompleteCount = 0;

  Object.keys(students).forEach(key => {
    const s = students[key];
    const avg1 = avgOf_(s.sem1);
    const avg2 = avgOf_(s.sem2);
    const overall = (s.sem2 && s.sem2.length) ? parseFloat(((avg1 + avg2) / 2).toFixed(1)) : avg1;
    const status = overall >= 50 ? "PROMOTED" : "DETAINED";
    if (overall >= 50) completeCount++; else notCompleteCount++;

    const copy = DriveApp.getFileById(templateId).makeCopy(`Card_${s.code}_${s.name}`, folder);
    const doc = DocumentApp.openById(copy.getId());
    fillCardDoc_(doc, {
      name: s.name,
      code: s.code,
      year: s.year,
      gradeSection: `${s.grade}${s.stream ? " " + s.stream : ""} ${s.section}`,
      status: status,
      subjects: s.subjects,
      sem1: s.sem1,
      sem2: s.sem2,
      total1: sumOf_(s.sem1),
      total2: sumOf_(s.sem2),
      avg1: avg1,
      avg2: avg2
    });

    cardRows.push([s.year, s.grade, s.stream, s.section, s.code, s.name, status, copy.getId(), new Date().toISOString()]);
  });

  if (cardRows.length) cardsSheet.getRange(2, 1, cardRows.length, CARDS_HEADER.length).setValues(cardRows);

  return { ok: true, generated: cardRows.length, completeCount: completeCount, notCompleteCount: notCompleteCount };
}

function claimCard_(body) {
  const code = String(body.code || "").trim();
  const name = String(body.name || "").trim();
  const year = String(body.year || "");
  const grade = String(body.grade || "");
  const section = String(body.section || "").toUpperCase();
  const host = body.host || body.origin || "default";

  if (!code || !name || !year || !grade || !section) return { ok: true, status: "bad_request" };

  const guard = checkAccountBinding_(body.email, host, body.password, code, name);
  if (!guard.allowed) return { ok: true, status: guard.status };

  const sh = getCardsSheet_();
  const values = sh.getDataRange().getValues();
  const wantedName = normalizeName_(name);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[4]).trim() !== code) continue;
    if (!normalizeName_(row[5]).includes(wantedName)) continue;
    if (String(row[0]) !== year || String(row[1]) !== grade || String(row[3]).toUpperCase() !== section) continue;

    const fileId = String(row[7] || "");
    if (!fileId) return { ok: true, status: "card_not_ready" };

    let pdfBlob;
    try {
      pdfBlob = DriveApp.getFileById(fileId).getAs("application/pdf");
    } catch (e) {
      return { ok: true, status: "card_not_ready" };
    }

    if (guard.fresh) bindAccountToStudent_(guard.rowIndex, code, row[5], year, grade, section);

    return {
      ok: true,
      status: "found",
      fileName: `Card_${code}_${row[5]}.pdf`,
      pdfBase64: Utilities.base64Encode(pdfBlob.getBytes())
    };
  }

  return { ok: true, status: "card_not_ready" };
}

function semCompare_(p) {
  const code = String(p.code || "").trim();
  const year = String(p.year || "");
  const grade = String(p.grade || "");
  const section = String(p.section || "").toUpperCase();
  const host = p.host || p.origin || "default";
  if (!code) return { ok: true, status: "bad_request" };

  const guard = checkAccountBinding_(p.email, host, p.password, code, p.name);
  if (!guard.allowed) return { ok: true, status: guard.status };

  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  let subjects = [], sem1 = [], sem2 = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[5]).trim() !== code) continue;
    if (year && String(row[0]) !== year) continue;
    if (grade && String(row[2]) !== grade) continue;
    if (section && String(row[4]).toUpperCase() !== section) continue;

    const subj = safeParse_(row[9], []);
    if (subj.length) subjects = subj;
    if (String(row[1]) === "sem1") sem1 = safeParse_(row[7], []);
    else sem2 = safeParse_(row[7], []);
  }

  if (!subjects.length) return { ok: true, status: "not_found" };

  return {
    ok: true,
    status: "found",
    subjects: subjects,
    sem1: subjects.map((_, i) => parseFloat(sem1[i]) || 0),
    sem2: subjects.map((_, i) => parseFloat(sem2[i]) || 0)
  };
}

/**
 * ============================================================================
 *  ራስ-ሰር የክፍል ማሸጋገሪያ ምዝገባ (Self-Service Grade Promotion Registration)
 *  ============================================================================
 *  ይህ ፋይል ነባሩን ኮድ (getSheet_, HEADER, lookupStudent_, ወዘተ) ብቻ ይጠቀማል፤
 *  ምንም ነባር function አይቀይርም/አይሰርዝም። ከ message__4_.txt መጨረሻ ላይ ብቻ ለጥፈው።
 *
 *  ስለሚያደርገው ነገር:
 *  1) ተማሪ/ወላጅ ኢሜይል ያረጋግጣል (OTP) — ማንም ያለ ኢሜይል በነጻ እንዳይሞክር
 *  2) የቀድሞ ክፍል ኮድ + ስም ያስገባል — ስርዓቱ ያንን ውጤት ፈልጎ ያገኘዋል
 *  3) Average ≥ 50% ከሆነ ብቻ (PROMOTED) ወደ ቀጣይ ክፍል በራስ-ሰር ይመዘግባል
 *  4) አዲስ 6-digit ኮድ ይፈጥራል፣ በ Results sheet ውስጥ ያስገባል፣ በ Promotions
 *     log ውስጥም ሙሉ ታሪኩን ይመዘግባል (admin በኋላ መልሶ ማየት/መሰረዝ ይችላል)
 *  5) የቀድሞ ኮድ "used" ተብሎ ይመዝገባል — ዳግም ጥቅም ላይ አይውልም
 *  6) ተደጋጋሚ የተሳሳተ ሙከራ (wrong code/name) ካለ ለተወሰነ ጊዜ ይቆለፋል
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// SETTINGS — እነዚህን ለትምህርት ቤቱ እውነታ አስተካክል
// ---------------------------------------------------------------------------

// 9→10, 10→11, 11→12 — ግሬድ 8ኛ በዚህ ስርዓት ውስጥ የለም (በአካል ብቻ ነው)
const NEXT_GRADE_MAP_ = { "9": "10", "10": "11", "11": "12" };

// ግሬድ 11 ላይ stream (Natural/Social) አዲስ ስለሚመረጥ፣ ከ10→11 ብቻ stream ግዴታ ይጠየቃል
// 10→11 (አዲስ ምርጫ) እና 11→12 (ያለውን ዘርፍ ማረጋገጥ) — ሁለቱም stream ያስፈልጋቸዋል
const STREAM_REQUIRED_NEXT_GRADES_ = ["11"];

const PROMOTIONS_SHEET_NAME_ = "Promotions";
const PROMOTIONS_HEADER_ = [
  "Timestamp", "Email", "Host",
  "PrevYear", "PrevGrade", "PrevStream", "PrevSection", "PrevCode", "Name",
  "Average", "Status",
  "NewYear", "NewGrade", "NewStream", "NewSection", "NewCode",
  "ParentPhone", "FaydaNumber", "Specialization", "Gender", "Age"
];

const PROMO_MAX_ATTEMPTS_ = 5;
const PROMO_LOCKOUT_MS_ = 20 * 60 * 1000; // 20 ደቂቃ
const PROMO_VERIFICATION_TTL_ = 30 * 60; // OTP ከተረጋገጠ በኋላ 30 ደቂቃ የምዝገባ ፍቃድ

function promoVerificationKey_(email, host) {
  return "promoVerified_" + normalizeEmail_(email) + "_" + normalizeHost_(host);
}

function generateVerificationToken_() {
  return Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
}

// ---------------------------------------------------------------------------
// doPost ውስጥ እነዚህን ሁለት መስመሮች ጨምር (በእጅ)
// ---------------------------------------------------------------------------
// ነባሩ doPost function ውስጥ፣ ከ "return jsonOut_({ ok: false, error: "unknown_action" });"
// በፊት፣ እነዚህን ሁለት መስመሮች ጨምር፦
//
//   if (action === "requestPromotionCode") return jsonOut_(requestPromotionCode_(body));
//   if (action === "submitPromotionRegistration") return jsonOut_(submitPromotionRegistration_(body));
//
// (ይህ ብቸኛው በነባሩ ፋይል ላይ የምታደርገው ለውጥ ነው — ሁለት አዲስ መስመር መጨመር ብቻ)

// ---------------------------------------------------------------------------
// Promotions sheet
// ---------------------------------------------------------------------------
function getPromotionsSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(PROMOTIONS_SHEET_NAME_);
  if (!sh) {
    sh = ss.insertSheet(PROMOTIONS_SHEET_NAME_);
    sh.appendRow(PROMOTIONS_HEADER_);
  } else {
    // Keep old Promotion rows intact, but add any newer columns that are
    // missing from an older Promotions sheet.
    const lastCol = Math.max(sh.getLastColumn(), 1);
    const existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());
    PROMOTIONS_HEADER_.forEach((header, index) => {
      if (existing[index] !== header) {
        sh.getRange(1, index + 1).setValue(header);
      }
    });
  }
  return sh;
}

function logPromotion_(row) {
  getPromotionsSheet_().appendRow(row);
}

// ---------------------------------------------------------------------------
// STEP 1 — Email OTP (ልክ እንደ ነባሩ signup ፍሰት)
// ---------------------------------------------------------------------------
function requestPromotionCode_(body) {
  const email = normalizeEmail_(body.email);
  const host = normalizeHost_(body.host || body.origin || "default");

  if (!email || email.indexOf("@") === -1 || email.indexOf(".") === -1) {
    return { ok: true, status: "bad_email" };
  }

  if (isPromoLocked_(email)) {
    return { ok: true, status: "locked" };
  }

  const code = generateOtp_();
  CacheService.getScriptCache().put(otpCacheKey_("promotion", email, host), code, 600);
  try {
    MailApp.sendEmail(
      email,
      "Grade Registration Verification Code",
      "Your verification code is: " + code + "\n\nThis code expires in 10 minutes."
    );
  } catch (e) {
    return { ok: true, status: "send_failed", detail: String(e) };
  }
  return { ok: true, status: "sent" };
}

// ---------------------------------------------------------------------------
// STEP 1-B — OTP ማረጋገጥ፣ ከዚያ 30 ደቂቃ የሚቆይ verification token
// ---------------------------------------------------------------------------
function verifyPromotionCode_(body) {
  const email = normalizeEmail_(body.email);
  const host = normalizeHost_(body.host || body.origin || "default");
  const otp = String(body.otp || "").trim();

  if (!email || !/^\d{6}$/.test(otp)) return { ok: true, status: "invalid_otp" };
  if (isPromoLocked_(email)) return { ok: true, status: "locked" };

  const cache = CacheService.getScriptCache();
  const otpKey = otpCacheKey_("promotion", email, host);
  const storedOtp = cache.get(otpKey);
  if (!storedOtp || storedOtp !== otp) {
    recordPromoFailure_(email);
    return { ok: true, status: storedOtp ? "invalid_otp" : "expired" };
  }

  cache.remove(otpKey);
  clearPromoFailures_(email);

  const token = generateVerificationToken_();
  cache.put(promoVerificationKey_(email, host), token, PROMO_VERIFICATION_TTL_);
  return { ok: true, status: "verified", verificationToken: token, expiresInSeconds: PROMO_VERIFICATION_TTL_ };
}

// ---------------------------------------------------------------------------
// Rate limiting — ልክ እንደ ADMIN_PW_MAX_ATTEMPTS ያለው ንድፍ
// ---------------------------------------------------------------------------
function promoAttemptKey_(email) {
  return "promoAttempts_" + normalizeEmail_(email);
}

function isPromoLocked_(email) {
  const cache = CacheService.getScriptCache();
  const raw = cache.get(promoAttemptKey_(email));
  if (!raw) return false;
  const data = JSON.parse(raw);
  return data.count >= PROMO_MAX_ATTEMPTS_ && (Date.now() - data.firstAt) < PROMO_LOCKOUT_MS_;
}

function recordPromoFailure_(email) {
  const cache = CacheService.getScriptCache();
  const key = promoAttemptKey_(email);
  const raw = cache.get(key);
  let data = raw ? JSON.parse(raw) : { count: 0, firstAt: Date.now() };
  data.count += 1;
  cache.put(key, JSON.stringify(data), Math.ceil(PROMO_LOCKOUT_MS_ / 1000));

  if (data.count === PROMO_MAX_ATTEMPTS_) {
    const adminEmail = getAdminEmail_();
    if (adminEmail) {
      try {
        MailApp.sendEmail(
          adminEmail,
          "⚠️ Promotion registration locked — repeated wrong attempts",
          "Email " + email + " entered wrong code/name " + PROMO_MAX_ATTEMPTS_ + " times in a row and is now locked for 20 minutes."
        );
      } catch (e) { /* ignore mail errors */ }
    }
  }
}

function clearPromoFailures_(email) {
  CacheService.getScriptCache().remove(promoAttemptKey_(email));
}

// ---------------------------------------------------------------------------
// STEP 2 — የቀድሞ ክፍል ውጤት ማግኘት (ያለ account binding፣ ግን rate-limited)
// ---------------------------------------------------------------------------
function findPreviousRecords_(year, grade, stream, section, code, name) {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const wantedName = normalizeName_(name);
  const wantedCode = String(code || "").trim();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[0]) !== String(year)) continue;
    if (String(row[2]) !== String(grade)) continue;
    if (String(row[4]).toUpperCase() !== String(section).toUpperCase()) continue;
    if (String(row[5]).trim() !== wantedCode) continue;
    if (!normalizeName_(row[6]).includes(wantedName)) continue;
    if (parseInt(grade, 10) >= 11 && stream && String(row[3]).toLowerCase() !== String(stream).toLowerCase()) continue;

    rows.push({
      rowIndex: i + 1,
      year: String(row[0]), grade: String(row[2]), stream: String(row[3]),
      section: String(row[4]), code: String(row[5]), name: row[6],
      semester: String(row[1]), scores: safeParse_(row[7], []),
      subjects: safeParse_(row[9], [])
    });
  }

  return rows;
}

function computeAverage_(scores) {
  const raw = Array.isArray(scores) ? scores : [];
  // Blank/null marks mean that this student's semester is not complete yet.
  // A real numeric 0 is still a valid mark and is therefore kept.
  if (!raw.length || raw.some(v => v === null || v === undefined || String(v).trim() === "" || !Number.isFinite(parseFloat(v)))) {
    return null;
  }
  const nums = raw.map(v => parseFloat(v));
  return nums.reduce((a, b) => a + b, 0) / nums.length;
}

function computeAnnualAverage_(records) {
  const sem1Record = records.find(r => r.semester === "sem1" || r.semester === "1");
  const sem2Record = records.find(r => r.semester === "sem2" || r.semester === "2");
  const sem1Average = sem1Record ? computeAverage_(sem1Record.scores) : null;
  const sem2Average = sem2Record ? computeAverage_(sem2Record.scores) : null;
  if (sem1Average === null || sem2Average === null) {
    return { complete: false, sem1Average, sem2Average, annualAverage: null };
  }
  return {
    complete: true,
    sem1Average,
    sem2Average,
    annualAverage: (sem1Average + sem2Average) / 2
  };
}

// ---------------------------------------------------------------------------
// አዲስ 6-digit ኮድ — ከነባር ኮዶች ጋር እንዳይጋጭ እየፈተነ ይፈጥራል
// ---------------------------------------------------------------------------
function generateUniqueSixDigitCode_() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const existing = new Set(values.slice(1).map(r => String(r[5]).trim()));

  let code;
  let tries = 0;
  do {
    code = String(Math.floor(100000 + Math.random() * 900000));
    tries++;
  } while (existing.has(code) && tries < 50);

  return code;
}

// A student may complete registration only once for the same previous-year
// student record.  This is deliberately based on the previous academic-year
// identity, not on email, because the same student must not create a second
// registration by changing email or submitting the form again.
function findExistingPromotionRegistration_(prevYear, prevGrade, prevStream, prevSection, prevCode, name) {
  const values = getPromotionsSheet_().getDataRange().getValues();
  const wantedName = normalizeName_(name);
  const wantedCode = String(prevCode || '').trim();
  const wantedYear = String(prevYear || '').trim();
  const wantedGrade = String(prevGrade || '').trim();
  const wantedSection = String(prevSection || '').trim().toUpperCase();
  const wantedStream = String(prevStream || '').trim().toLowerCase();

  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const status = String(r[10] || '').trim();
    const rowYear = String(r[3] || '').trim();
    const rowGrade = String(r[4] || '').trim();
    const rowStream = String(r[5] || '').trim().toLowerCase();
    const rowSection = String(r[6] || '').trim().toUpperCase();
    const rowCode = String(r[7] || '').trim();
    const rowName = normalizeName_(r[8]);

    // Only a completed registration counts. Failed lookup attempts do not.
    if (status !== 'promoted' && status !== 'detained') continue;
    if (rowYear !== wantedYear || rowGrade !== wantedGrade) continue;
    if (rowSection !== wantedSection || rowCode !== wantedCode) continue;
    if (wantedStream && rowStream && rowStream !== wantedStream) continue;
    if (rowName !== wantedName && !rowName.includes(wantedName) && !wantedName.includes(rowName)) continue;

    return {
      status, newYear: String(r[11] || '').trim(), newGrade: String(r[12] || '').trim(),
      newCode: String(r[15] || '').trim(), timestamp: r[0]
    };
  }
  return null;
}

// ---------------------------------------------------------------------------
// STEP 3 — የቀጣዩ የትምህርት ዘመን ማስላት
// ---------------------------------------------------------------------------
function getNextAcademicYear_(prevYear) {
  const y = String(prevYear || '').trim();
  const slash = y.match(/^(\d{4})\/(\d{2}|\d{4})$/);
  if (slash) {
    const first = parseInt(slash[1], 10) + 1;
    const second = String(parseInt(slash[2], 10) + 1);
    return slash[2].length === 2 ? first + '/' + String(second).slice(-2) : first + '/' + second;
  }
  if (/^\d{4}$/.test(y)) return String(parseInt(y, 10) + 1);
  return y;
}

// ---------------------------------------------------------------------------
// STEP 3 — ዋናው ተግባር: ማረጋገጥ → ማስላት → በራስ-ሰር መመዝገብ
// ---------------------------------------------------------------------------
function submitPromotionRegistration_(body) {
  const email = normalizeEmail_(body.email);
  const host = normalizeHost_(body.host || body.origin || "default");
  const verificationToken = String(body.verificationToken || "").trim();

  const prevYear = String(body.prevYear || "").trim();
  const prevGrade = String(body.prevGrade || "").trim();
  const prevStream = String(body.prevStream || "").trim();
  const prevSection = String(body.prevSection || "").trim();
  const prevCode = String(body.prevCode || "").trim();
  const name = String(body.name || "").trim();
  const chosenStream = String(body.chosenStream || "").trim();
  const chosenSpecialization = String(body.chosenSpecialization || "").trim();
  const parentPhoneRaw = String(body.parentPhone || "");
  const faydaNumberRaw = String(body.faydaNumber || "");
  const gender = String(body.gender || "").trim();
  const ageRaw = String(body.age || "").trim();
  const age = Number(ageRaw);
  if (!["Male", "Female", "Other"].includes(gender)) return { ok: false, status: "invalid_gender" };
  if (!/^\d{1,2}$/.test(ageRaw) || age < 3 || age > 100) return { ok: false, status: "invalid_age" };

  // Parent phone: exactly 10 digits, numbers only.
  if (!/^\d{10}$/.test(parentPhoneRaw)) {
    return { ok: false, status: "invalid_parent_phone" };
  }

  // Fayda: exactly 12 digits (FAN) or exactly 16 digits (FCN), numbers only.
  if (!/^\d{12}$/.test(faydaNumberRaw) && !/^\d{16}$/.test(faydaNumberRaw)) {
    return { ok: false, status: "invalid_fayda_number" };
  }

  const parentPhone = parentPhoneRaw;
  const faydaNumber = faydaNumberRaw;

  if (!email || !verificationToken || !prevYear || !prevGrade || !prevSection || !prevCode || !name) {
    return { ok: false, status: "bad_request" };
  }
  if (isPromoLocked_(email)) return { ok: false, status: "locked" };

  // OTP አንድ ጊዜ ከተረጋገጠ በኋላ 30 ደቂቃ የሚቆይ token ብቻ ይጠቀማል።
  const verifyKey = promoVerificationKey_(email, host);
  const savedToken = CacheService.getScriptCache().get(verifyKey);
  if (!savedToken || savedToken !== verificationToken) {
    return { ok: false, status: "invalid_verification" };
  }

  const nextGrade = NEXT_GRADE_MAP_[prevGrade];
  if (!nextGrade) return { ok: false, status: "unsupported_grade" };
  const records = findPreviousRecords_(prevYear, prevGrade, prevStream, prevSection, prevCode, name);
  if (!records.length) {
    recordPromoFailure_(email);
    logPromotion_([new Date(), email, host, prevYear, prevGrade, prevStream, prevSection, prevCode, name,
      "", "not_found", "", "", "", "", ""]);
    return { ok: false, status: "record_not_found" };
  }

  // One student = one registration.  Check this BEFORE recalculating the
  // current result so a later mark correction cannot create a second
  // registration for the same previous-year student.
  const existingRegistration = findExistingPromotionRegistration_(
    prevYear, prevGrade, prevStream, prevSection, prevCode, name
  );
  if (existingRegistration) {
    return {
      ok: false,
      status: 'already_registered',
      registrationStatus: existingRegistration.status,
      newYear: existingRegistration.newYear,
      newGrade: existingRegistration.newGrade,
      newCode: existingRegistration.newCode
    };
  }

  const annual = computeAnnualAverage_(records);
  const sem1Average = annual.sem1Average;
  const sem2Average = annual.sem2Average;

  // ሁለቱም semester ሳይሞሉ ማለፊያ አይሰጥም።
  if (!annual.complete) {
    return { ok: true, status: "incomplete", sem1Average, sem2Average };
  }

  const annualAverage = annual.annualAverage;
  const promoted = annualAverage >= 50;

  // A Grade 10 student chooses a Grade 11 stream and specialization only if they pass.
  if (promoted && STREAM_REQUIRED_NEXT_GRADES_.indexOf(nextGrade) !== -1 && !chosenStream) {
    return { ok: false, status: "stream_required" };
  }
  const allowedSpecializations = {
    "Natural": ["Computer Maintenance", "Health"],
    "Social": ["Journalism", "Accounting"]
  };
  if (promoted && nextGrade === "11" && (!allowedSpecializations[chosenStream] || allowedSpecializations[chosenStream].indexOf(chosenSpecialization) === -1)) {
    return { ok: false, status: "specialization_required" };
  }

  if (!promoted) {
    // ያልፈው ተማሪ ደጋሚ ሆኖ በአዲሱ academic year በዚያው grade ይመዘገባል፣ አዲስ 6-digit code ይወስዳል።
    const currentYear = getNextAcademicYear_(prevYear);
    const repeatStream = prevGrade === "11" ? records[0].stream : "";
    const newCode = generateUniqueSixDigitCode_();
    const sh = getSheet_();
    sh.appendRow([
      currentYear, "sem1", prevGrade, repeatStream, "", newCode, records[0].name,
      JSON.stringify([]), false, JSON.stringify([]), JSON.stringify([]), JSON.stringify(null), ""
    ]);

    clearPromoFailures_(email);
    CacheService.getScriptCache().remove(verifyKey);
    logPromotion_([new Date(), email, host, prevYear, prevGrade, prevStream, prevSection, prevCode, name,
      annualAverage, "detained", currentYear, prevGrade, repeatStream, "", newCode, parentPhone, faydaNumber, "", gender, age]);
    return { ok: true, status: "detained", sem1Average, sem2Average, annualAverage,
      newCode, newGrade: prevGrade, newSection: "", newStream: repeatStream, newYear: currentYear };
  }

  // ተማሪው ወደ አዲሱ የትምህርት ዘመን ሲገባ አዲስ 6-digit code ያገኛል።
  // ነገር ግን በአንድ academic year ውስጥ 1ኛ እና 2ኛ semester ሁለቱም ይህንኑ code ይጠቀማሉ።
  const currentYear = getNextAcademicYear_(prevYear);
  const finalStream = nextGrade === "11" ? chosenStream : (prevGrade === "11" ? records[0].stream : "");
  const newCode = generateUniqueSixDigitCode_();
  const sh = getSheet_();

  sh.appendRow([
    currentYear, "sem1", nextGrade, finalStream, "", newCode, records[0].name,
    JSON.stringify([]), false, JSON.stringify([]), JSON.stringify([]), JSON.stringify(null), chosenSpecialization
  ]);

  clearPromoFailures_(email);
  CacheService.getScriptCache().remove(verifyKey);
  logPromotion_([new Date(), email, host, prevYear, prevGrade, prevStream, prevSection, prevCode, name,
    annualAverage, "promoted", currentYear, nextGrade, finalStream, "", newCode, parentPhone, faydaNumber, chosenSpecialization, gender, age]);

  return { ok: true, status: "promoted", sem1Average, sem2Average, annualAverage,
    newCode, newGrade: nextGrade, newSection: "", newStream: finalStream, newYear: currentYear };
}

// ---------------------------------------------------------------------------
// STEP 4 — Admin page ለ "Pending/ተመዝግበው ያለቁ ተማሪዎች" tab የሚያስፈልግ ዝርዝር
// ---------------------------------------------------------------------------
function adminListPromotions_(body) {
  if (String(body.password || "") !== getAdminPassword_()) return { ok: false, error: "wrong_password" };
  const values = getPromotionsSheet_().getDataRange().getValues();
  if (!values.length) return { ok: true, rows: [] };
  // Read by header name rather than fixed column number. This keeps Age/Gender
  // correct even if an older Promotions sheet had columns inserted/reordered.
  const headers = values[0].map(v => String(v || "").trim().toLowerCase());
  const indexOf = (...names) => {
    for (const name of names) {
      const i = headers.indexOf(String(name).toLowerCase());
      if (i >= 0) return i;
    }
    return -1;
  };
  const get = (r, ...names) => {
    const i = indexOf(...names);
    return i >= 0 && r[i] !== undefined && r[i] !== null ? r[i] : "";
  };
  const rows = values.slice(1).filter(r => r.some(v => v !== "" && v !== null)).map(r => ({
    timestamp: get(r, "Timestamp"), email: get(r, "Email"), host: get(r, "Host"),
    prevYear: get(r, "PrevYear"), prevGrade: get(r, "PrevGrade"), prevStream: get(r, "PrevStream"),
    prevSection: get(r, "PrevSection"), prevCode: get(r, "PrevCode"), name: get(r, "Name"),
    average: get(r, "Average"), status: get(r, "Status"), newYear: get(r, "NewYear"),
    newGrade: get(r, "NewGrade"), newStream: get(r, "NewStream"), newSection: get(r, "NewSection"),
    newCode: get(r, "NewCode"), parentPhone: String(get(r, "ParentPhone") || "").trim(),
    faydaNumber: String(get(r, "FaydaNumber") || "").trim(),
    specialization: String(get(r, "Specialization") || "").trim(),
    gender: String(get(r, "Gender") || "").trim(), age: String(get(r, "Age") || "").trim()
  }));
  return { ok: true, rows: rows };
}

function adminListPromotionGroups_(body) {
  if (String(body.password || "") !== getAdminPassword_()) return { ok: false, error: "wrong_password" };

  // The Admin page is year-based.  A student's result/failure belongs to the
  // academic year in which it happened.  A later registration can remove that
  // student from the older year's "failed" list, while the new registration
  // appears in the new year's registration group.
  const requestedYear = String(body.academicYear || getMeta_().currentAcademicYear || "").trim();
  const groups = { "9to10": [], "10to11": [], "11to12": [], "failed": [] };
  const promoValues = getPromotionsSheet_().getDataRange().getValues();
  const promotions = [];
  const laterRegistrations = {};

  promoValues.slice(1).forEach(r => {
    const row = {
      timestamp: r[0],
      prevYear: String(r[3] || "").trim(),
      prevGrade: String(r[4] || "").trim(),
      prevStream: String(r[5] || "").trim(),
      prevSection: String(r[6] || "").trim(),
      prevCode: String(r[7] || "").trim(),
      name: r[8],
      average: r[9],
      status: String(r[10] || "").trim(),
      newYear: String(r[11] || "").trim(),
      newGrade: String(r[12] || "").trim(),
      newStream: String(r[13] || "").trim(),
      newSection: String(r[14] || "").trim(),
      newCode: String(r[15] || "").trim(),
      parentPhone: String(r[16] || "").trim(),
      faydaNumber: String(r[17] || "").trim(),
      specialization: String(r[18] || "").trim(), gender: String(r[19] || "").trim(), age: String(r[20] || "").trim()
    };
    promotions.push(row);

    // If a failed student registers again in a later academic year, the old
    // year's failed list should no longer show that student.
    if (row.status === "detained" && row.newCode && row.prevYear && row.newYear) {
      const laterKey = [row.prevYear, row.prevGrade, row.prevStream, row.prevSection, row.prevCode, normalizeName_(row.name)].join("|");
      if (row.newYear !== row.prevYear) laterRegistrations[laterKey] = true;
    }
  });

  const groupKeyForGrade = grade => grade === "9" ? "9to10" : grade === "10" ? "10to11" : grade === "11" ? "11to12" : "";

  // Registrations for the selected academic year.
  // - promoted: goes to the next-grade group
  // - detained: stays in the same-grade group as a repeater
  // One previous-year student must appear only once. Older test data may contain
  // duplicate promotion rows, so the Admin view keeps the first completed
  // registration and ignores later duplicate rows. New submissions are blocked
  // by findExistingPromotionRegistration_ above.
  const registrationSeen = new Set();
  promotions.forEach(row => {
    if (requestedYear && row.newYear !== requestedYear) return;
    if ((row.status !== "promoted" && row.status !== "detained") || !row.newCode) return;

    const identityKey = [
      row.prevYear, row.prevGrade, row.prevStream, row.prevSection,
      row.prevCode, normalizeName_(row.name), row.newYear
    ].join("|");
    if (registrationSeen.has(identityKey)) return;
    registrationSeen.add(identityKey);

    const key = groupKeyForGrade(row.prevGrade);
    if (!key) return;

    if (row.status === "promoted") {
      groups[key].push(row);
    } else if (row.status === "detained") {
      groups[key].push({
        ...row,
        isRepeater: true,
        // The repeater is registered in the same grade, not promoted to the next grade.
        newGrade: row.newGrade || row.prevGrade,
        newStream: row.newStream || row.prevStream
      });
    }
  });

  // Build the failed list ONLY for the selected academic year.
  // Each student's two semester averages are independent of other students.
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const byStudent = {};

  values.slice(1).forEach(r => {
    const year = String(r[0] || "").trim();
    const sem = String(r[1] || "").trim();
    const grade = String(r[2] || "").trim();
    if (requestedYear && year !== requestedYear) return;
    const code = String(r[5] || "").trim();
    const name = String(r[6] || "").trim();
    const section = String(r[4] || "").trim();
    if (!year || !code || !name || !["9", "10", "11"].includes(grade)) return;

    const stream = String(r[3] || "").trim();
    const key = [year, grade, stream, section, code, normalizeName_(name)].join("|");
    if (!byStudent[key]) {
      byStudent[key] = {
        year, grade, stream, section, code, name,
        sem1: null, sem2: null
      };
    }

    const rec = computeAverage_(safeParse_(r[7], []));
    if (sem === "sem1" || sem === "1") byStudent[key].sem1 = rec;
    if (sem === "sem2" || sem === "2") byStudent[key].sem2 = rec;
  });

  Object.values(byStudent).forEach(st => {
    // A student is evaluated as soon as BOTH of that student's semesters are complete.
    // Other students do not have to be complete first.
    if (st.sem1 === null || st.sem2 === null) return;

    const annualAverage = (st.sem1 + st.sem2) / 2;
    if (annualAverage >= 50) return;

    const repeatKey = [st.year, st.grade, st.stream, st.section, st.code, normalizeName_(st.name)].join("|");

    // If the student has already registered in a later year, remove them from
    // this older year's failed list. Their later-year registration will appear
    // in that later year's appropriate group as REGISTERED + ደጋሚ.
    if (laterRegistrations[repeatKey]) return;

    groups.failed.push({
      ...st,
      sem1Average: st.sem1,
      sem2Average: st.sem2,
      annualAverage,
      status: "failed",
      newYear: "",
      newGrade: st.grade,
      newStream: st.stream,
      newSection: "",
      newCode: ""
    });
  });

  return { ok: true, academicYear: requestedYear, groups };
}

// ADMIN action ደግሞ doPost ውስጥ ጨምር፦
//   if (action === "adminListPromotions") return jsonOut_(adminListPromotions_(body));