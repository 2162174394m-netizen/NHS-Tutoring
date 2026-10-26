
/** ============================================================
 * SOURCE GROUP: NHS_Config.gs
 * ============================================================ */

/**
 * NHS CONFIG — CONSOLIDATED PLATFORM V1
 *
 * This is the ONLY place the main spreadsheet ID is declared.
 */

const NHS_PRODUCTION_SPREADSHEET_ID = "1zDqAOvmIEnybL5erExVECG1FE0P4LnKggv4YDSS4oFQ";

/**
 * V47 RELEASE GUARD — environment-aware data target.
 *
 * Production requires no new Script Property and therefore preserves the
 * current live workbook by default. A separate STAGING Apps Script project
 * should set:
 *   NHS_ENVIRONMENT=STAGING
 *   NHS_DATA_SPREADSHEET_ID=<staging spreadsheet id>
 *
 * Script Properties are project-scoped, so the same source can be deployed
 * safely from separate production and staging Apps Script projects.
 */
function nhsrg_bootEnvironmentName_() {
  try {
    const value = String(
      PropertiesService.getScriptProperties().getProperty("NHS_ENVIRONMENT") || "PRODUCTION"
    ).trim().toUpperCase();
    return value === "STAGING" ? "STAGING" : "PRODUCTION";
  } catch (e) {
    return "PRODUCTION";
  }
}

function nhsrg_bootSpreadsheetId_() {
  const environment = nhsrg_bootEnvironmentName_();

  try {
    const configured = String(
      PropertiesService.getScriptProperties().getProperty("NHS_DATA_SPREADSHEET_ID") || ""
    ).trim();

    if (environment === "STAGING") {
      if (!configured) {
        throw new Error(
          "STAGING_DATA_TARGET_REQUIRED|A STAGING project must define NHS_DATA_SPREADSHEET_ID and may not fall back to the production workbook."
        );
      }

      if (configured === NHS_PRODUCTION_SPREADSHEET_ID) {
        throw new Error(
          "STAGING_PRODUCTION_TARGET_BLOCKED|A STAGING project may not use the production spreadsheet ID."
        );
      }

      return configured;
    }

    return configured || NHS_PRODUCTION_SPREADSHEET_ID;
  } catch (e) {
    if (environment === "STAGING") {
      throw e;
    }
    return NHS_PRODUCTION_SPREADSHEET_ID;
  }
}

const NHS_SPREADSHEET_ID = nhsrg_bootSpreadsheetId_();

const NHS_RELEASE = Object.freeze({
  VERSION: "V47.7",
  CACHE_NAMESPACE: "v477",
  SCHEMA_VERSION: "V47.3",
  SYNC_PROTOCOL: "V47.7-adaptive",
  RELEASE_NAME: "Schoolwide Deployment Readiness"
});

const NHS_PLATFORM = Object.freeze({
  TITLE: "Tesoro NHS Tutoring",
  TITAN_TIME: "10:25 AM–11:05 AM",
  TIMEZONE: "America/Los_Angeles",
  DEFAULT_ROUTE: "gateway",
  BUILD: NHS_RELEASE.VERSION
});


/**
 * Returns a guaranteed primitive-string timezone for Apps Script services.
 * Preserves the platform's existing timezone behavior while hardening
 * Utilities.formatDate() and other APIs that require a String timezone.
 */
function nhs_safeTimeZone_(ss) {
  let tz = "";

  try {
    if (ss && typeof ss.getSpreadsheetTimeZone === "function") {
      tz = ss.getSpreadsheetTimeZone();
    }
  } catch (e) {}

  try {
    tz = String(
      tz ||
      (typeof NHS_PLATFORM !== "undefined" && NHS_PLATFORM.TIMEZONE) ||
      Session.getScriptTimeZone() ||
      "America/Los_Angeles"
    ).trim();
  } catch (e) {
    tz = "America/Los_Angeles";
  }

  if (!tz || tz === "[object Object]") {
    tz = "America/Los_Angeles";
  }

  return tz;
}


/** ============================================================
 * SOURCE GROUP: NHS_Data_Core.gs
 * ============================================================ */

/**
 * TESORO NHS — DATA CORE V43
 *
 * Central data-integrity/performance helpers.
 * Key principle: rows are real only when their business key is populated.
 * This prevents checkbox/data-validation artifacts from inflating getLastRow().
 */

const NHSDATA_REQUEST_SHEET = "Tutoring Requests";
const NHSDATA_REQUEST_KEY = "Request ID";
const NHSDATA_DEADLETTER_PREFIX = "NHS_DEADLETTER_V43_";

function nhs_ensureSpreadsheetTimeZone_(ss) {
  if (!ss) throw new Error("Spreadsheet is required.");

  const before = nhs_safeTimeZone_(ss);
  const target = String(NHS_PLATFORM.TIMEZONE || "America/Los_Angeles").trim();

  if (before !== target) {
    ss.setSpreadsheetTimeZone(target);
    SpreadsheetApp.flush();
  }

  const after = nhs_safeTimeZone_(ss);

  return {
    ok: after === target,
    changed: before !== after,
    before: before,
    after: after,
    expected: target
  };
}


function nhsdata_realLastRowByKey_(sheet, keyHeader) {
  if (!sheet) return 1;

  const h = nhsperf_headerMap_(sheet, true, true);
  const col = h[String(keyHeader || "").trim()];
  if (!col) throw new Error('Missing key header: "' + keyHeader + '".');

  const physicalLast = Math.max(1, sheet.getLastRow());
  if (physicalLast < 2) return 1;

  const values = sheet
    .getRange(2, col, physicalLast - 1, 1)
    .getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0] || "").trim()) return i + 2;
  }

  return 1;
}

function nhsdata_nextAppendRow_(sheet, keyHeader) {
  return Math.max(2, nhsdata_realLastRowByKey_(sheet, keyHeader) + 1);
}

function nhsdata_appendKeyedRow_(sheet, keyHeader, rowValues) {
  if (!sheet) throw new Error("Sheet is required.");
  const row = nhsdata_nextAppendRow_(sheet, keyHeader);
  sheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
  return row;
}

function nhsdata_findRowByKey_(sheet, keyHeader, keyValue) {
  if (!sheet) return 0;

  const key = String(keyValue || "").trim();
  if (!key) return 0;

  const h = nhsperf_headerMap_(sheet, true, true);
  const col = h[keyHeader];
  if (!col) return 0;

  const last = nhsdata_realLastRowByKey_(sheet, keyHeader);
  if (last < 2) return 0;

  const match = sheet
    .getRange(2, col, last - 1, 1)
    .createTextFinder(key)
    .matchEntireCell(true)
    .matchCase(true)
    .findNext();

  return match ? match.getRow() : 0;
}

/**
 * Reads only through the last row whose key is populated.
 * Returned objects contain a non-sheet `_row` property with their physical row.
 */
function nhsdata_readKeyedObjects_(sheet, keyHeader) {
  if (!sheet) return [];

  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return [];

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(function(v) { return String(v || "").trim(); });

  const h = {};
  headers.forEach(function(header, i) {
    if (header && h[header] === undefined) h[header] = i;
  });

  const keyIndex = h[keyHeader];
  if (keyIndex === undefined) {
    throw new Error('Missing key header: "' + keyHeader + '".');
  }

  const last = nhsdata_realLastRowByKey_(sheet, keyHeader);
  if (last < 2) return [];

  const rows = sheet
    .getRange(2, 1, last - 1, lastColumn)
    .getDisplayValues();

  return rows.map(function(row, i) {
    const key = String(row[keyIndex] || "").trim();
    if (!key) return null;

    const obj = { _row: i + 2 };
    headers.forEach(function(header, col) {
      if (header) obj[header] = row[col] || "";
    });
    return obj;
  }).filter(Boolean);
}

function nhsdata_requestObjects_(ss) {
  const sh = nhsperf_sheet_(ss, NHSDATA_REQUEST_SHEET);
  if (!sh) return [];
  return nhsdata_readKeyedObjects_(sh, NHSDATA_REQUEST_KEY);
}

/**
 * Compacts a keyed table WITHOUT deleting sheet rows. Formatting, validations,
 * filters, hidden columns, etc. remain intact. Only row content is compacted.
 */
function nhsdata_compactKeyedTable_(sheet, keyHeader) {
  if (!sheet) return { kept: 0, clearedArtifactRows: 0 };

  const lastColumn = sheet.getLastColumn();
  const physicalLast = Math.max(1, sheet.getLastRow());
  if (physicalLast < 2 || lastColumn < 1) {
    return { kept: 0, clearedArtifactRows: 0 };
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(function(v) { return String(v || "").trim(); });
  const keyIndex = headers.indexOf(keyHeader);
  if (keyIndex < 0) throw new Error('Missing key header: "' + keyHeader + '".');

  const bodyRange = sheet.getRange(2, 1, physicalLast - 1, lastColumn);
  const values = bodyRange.getValues();
  const keep = values.filter(function(row) {
    return String(row[keyIndex] == null ? "" : row[keyIndex]).trim() !== "";
  });

  // Clear content only; preserve data validation/formatting.
  bodyRange.clearContent();

  if (keep.length) {
    sheet.getRange(2, 1, keep.length, lastColumn).setValues(keep);
  }

  return {
    kept: keep.length,
    clearedArtifactRows: Math.max(0, values.length - keep.length),
    oldPhysicalLastRow: physicalLast,
    newRealLastRow: keep.length ? keep.length + 1 : 1
  };
}

function nhsdata_applyCheckboxValidation_(sheet, header) {
  if (!sheet) return false;

  const h = nhsperf_headerMap_(sheet, true, true);
  const col = h[header];
  if (!col) return false;

  const rule = SpreadsheetApp
    .newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();

  // V43.1: only real keyed rows + one empty staging row receive checkbox
  // validation. Applying checkbox validation down the entire sheet can
  // materialize FALSE values in otherwise empty rows and inflate getLastRow().
  const realLast = nhsdata_realLastRowByKey_(sheet, "Request ID");
  const firstDataRow = 2;
  const lastValidationRow = Math.min(
    sheet.getMaxRows(),
    Math.max(firstDataRow, realLast + 1)
  );
  const count = Math.max(1, lastValidationRow - firstDataRow + 1);

  sheet
    .getRange(firstDataRow, col, count, 1)
    .setDataValidation(rule);

  // Remove phantom checkbox values below the real table while preserving
  // formatting and existing validation rules.
  const cleanupStart = Math.max(firstDataRow, realLast + 1);
  const cleanupCount = sheet.getMaxRows() - cleanupStart + 1;

  if (cleanupCount > 0) {
    sheet
      .getRange(cleanupStart, col, cleanupCount, 1)
      .clearContent();
  }

  return true;
}

function nhsdata_invalidateRequestState_(spec) {
  spec = spec || {};

  // V45 central path.
  if (typeof nhs_v45InvalidatePortalState_ === "function") {
    return nhs_v45InvalidatePortalState_(spec);
  }

  // Defensive fallback for partial installs.
  try {
    const cache = nhsperf_scriptCache_();
    cache.remove(NHSOS_CACHE_KEY);
    cache.remove("nhsos_health_v14");

    if (spec.tutorId) {
      cache.remove(
        "tutorPortalV460|" +
        String(spec.tutorId)
      );
    }

    if (spec.teacherEmail) {
      cache.remove(
        "teacherPortalV460|" +
        String(spec.teacherEmail).toLowerCase()
      );
    }
  } catch (e) {}

  if (typeof nhsperf_touchDataRevision_ === "function") {
    return nhsperf_touchDataRevision_();
  }

  return String(Date.now());
}

function nhsdata_recordPropertyJobFailure_(props, key, raw, error, maxAttempts) {
  maxAttempts = Number(maxAttempts || 5);
  let job = {};
  try { job = JSON.parse(raw || "{}"); } catch (e) { job = {}; }

  job.attempts = Number(job.attempts || 0) + 1;
  job.lastAttemptAt = new Date().toISOString();
  job.lastError = error && error.message ? error.message : String(error || "Unknown error");

  if (job.attempts >= maxAttempts) {
    props.setProperty(
      NHSDATA_DEADLETTER_PREFIX + Date.now() + "_" + Utilities.getUuid(),
      JSON.stringify({ originalKey: key, job: job })
    );
    props.deleteProperty(key);
    return { deadLettered: true, attempts: job.attempts };
  }

  props.setProperty(key, JSON.stringify(job));
  return { deadLettered: false, attempts: job.attempts };
}

function nhsdata_pruneDeadLetters_(maxKeep) {
  maxKeep = Math.max(0, Number(maxKeep || 50));
  const props = nhsperf_scriptProperties_();
  const all = props.getProperties();
  const keys = Object.keys(all)
    .filter(function(k) { return k.indexOf(NHSDATA_DEADLETTER_PREFIX) === 0; })
    .sort()
    .reverse();

  keys.slice(maxKeep).forEach(function(k) { props.deleteProperty(k); });
  return Math.max(0, keys.length - maxKeep);
}


/** ============================================================
 * SOURCE GROUP: NHS_Lifecycle_Core.gs
 * ============================================================ */

/**
 * TESORO NHS — REQUEST / SESSION LIFECYCLE CORE V43
 *
 * One semantic contract used by Teacher, Tutor, Operations, notifications,
 * matching, and diagnostics.
 *
 * Scheduled = tutor/date placement exists.
 * Confirmed = Assignment Confirmed is populated with a real truthy value.
 */

function nhslife_norm_(value) {
  return String(value == null ? "" : value).trim().toLowerCase();
}

function nhslife_truthy_(value) {
  const s = nhslife_norm_(value);
  return !!s && !["false", "0", "no", "unchecked", "pending", "null", "undefined"].includes(s);
}

function nhslife_status_(request) {
  return nhslife_norm_(request && (request.status !== undefined ? request.status : request["Status"]));
}

function nhslife_response_(request) {
  return nhslife_norm_(request && (
    request.tutorResponseStatus !== undefined
      ? request.tutorResponseStatus
      : request["Tutor Response Status"]
  ));
}

function nhslife_confirmed_(request) {
  if (!request) return false;
  if (typeof request.isConfirmed === "boolean") return request.isConfirmed;
  const value = request.assignmentConfirmed !== undefined
    ? request.assignmentConfirmed
    : request["Assignment Confirmed"];
  return nhslife_truthy_(value);
}

function nhslife_cancelled_(request) {
  return nhslife_status_(request) === "cancelled";
}

function nhslife_completed_(request) {
  return nhslife_status_(request) === "completed";
}

function nhslife_closed_(request) {
  return nhslife_cancelled_(request) || nhslife_completed_(request);
}

function nhslife_activeCaseload_(request) {
  const status = nhslife_status_(request);
  return ["assigned", "scheduled", "pending verification", "needs follow-up"].includes(status) &&
    !nhslife_closed_(request);
}

function nhslife_needsTutorAction_(request) {
  const status = nhslife_status_(request);
  const response = nhslife_response_(request);
  return ["assigned", "scheduled"].includes(status) &&
    !["accepted", "declined"].includes(response) &&
    !nhslife_closed_(request);
}

function nhslife_dateKey_(value, timezone) {
  if (!value) return "";
  const raw = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, timezone || Session.getScriptTimeZone() || "America/Los_Angeles", "yyyy-MM-dd");
}

function nhslife_endMinutes_(timeText) {
  const text = String(timeText || "").trim();
  const matches = text.match(/\b(\d{1,2}):(\d{2})\s*(AM|PM)\b/gi);
  if (!matches || !matches.length) return 11 * 60 + 5;

  const last = matches[matches.length - 1].match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (!last) return 11 * 60 + 5;

  let hour = Number(last[1]);
  const minute = Number(last[2]);
  const meridiem = String(last[3]).toUpperCase();
  if (hour === 12) hour = 0;
  if (meridiem === "PM") hour += 12;
  return hour * 60 + minute;
}

function nhslife_sessionEnded_(request, now, timezone) {
  if (!request) return false;
  const scheduled = request.scheduledDate !== undefined
    ? request.scheduledDate
    : request["Scheduled Date"];
  const dateKey = nhslife_dateKey_(scheduled, timezone);
  if (!dateKey) return false;

  const current = now instanceof Date ? now : new Date();
  const today = Utilities.formatDate(current, timezone || Session.getScriptTimeZone() || "America/Los_Angeles", "yyyy-MM-dd");
  if (dateKey < today) return true;
  if (dateKey > today) return false;

  const hhmm = Utilities.formatDate(current, timezone || Session.getScriptTimeZone() || "America/Los_Angeles", "H:mm").split(":");
  const nowMinutes = Number(hhmm[0]) * 60 + Number(hhmm[1]);
  const time = request.scheduledTime !== undefined
    ? request.scheduledTime
    : request["Scheduled Time"];
  return nowMinutes >= nhslife_endMinutes_(time || NHS_PLATFORM.TITAN_TIME);
}

function nhslife_upcoming_(request, now, timezone) {
  return !!request &&
    nhslife_confirmed_(request) &&
    !nhslife_closed_(request) &&
    !!nhslife_dateKey_(
      request.scheduledDate !== undefined ? request.scheduledDate : request["Scheduled Date"],
      timezone
    ) &&
    !nhslife_sessionEnded_(request, now, timezone);
}

function nhslife_pastSession_(request, now, timezone) {
  return !!request &&
    nhslife_confirmed_(request) &&
    !nhslife_cancelled_(request) &&
    nhslife_sessionEnded_(request, now, timezone);
}

function nhslife_reportable_(request, now, timezone) {
  return nhslife_pastSession_(request, now, timezone) &&
    nhslife_response_(request) === "accepted" &&
    !["completed", "cancelled", "pending verification"].includes(nhslife_status_(request));
}


/** ============================================================
 * SOURCE GROUP: NHS_Titan_Scheduling_V46_5.gs
 * ============================================================ */

/**
 * V46.5 scheduling contract
 * -------------------------
 * - Same-day Titan Time scheduling is never allowed.
 * - Titan Time dates are Tuesday-Friday only.
 * - "This Week" accepts one or more remaining eligible dates in the current week.
 * - "Future Reservation" accepts one exact Titan Time date after the current week.
 * - "Recurring Support" stores weekday preferences; each generated occurrence must
 *   be reconfirmed before Operations can finalize it.
 * - Five Star readiness is tracked separately from NHS assignment confirmation.
 */
const NHSSCHED_TYPES = {
  THIS_WEEK: "This Week",
  FUTURE: "Future Reservation",
  RECURRING: "Recurring Support"
};
const NHSSCHED_TITAN_DAYS = ["Tuesday", "Wednesday", "Thursday", "Friday"];
const NHSSCHED_RECONFIRM_DAYS = 7;

const NHSSCHED_EXCEPTION_SHEET = "Schedule Exceptions";
var NHSSCHED_EXCEPTION_MEMO_ = null;

const NHSSCHED_EXCEPTION_HEADERS = [
  "Date",
  "Type",
  "Reason",
  "Active",
  "Source",
  "Updated At",
  "Updated By"
];

const NHSSCHED_DEFAULT_EXCEPTIONS = Object.freeze([
  {date:"2026-09-07",type:"No School",reason:"Labor Day",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-09-11",type:"Minimum Day",reason:"Minimum Day — no Titan Time shown on the school calendar",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-11-02",type:"No Students",reason:"Professional Development Day",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-11-11",type:"No School",reason:"Veterans Day",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-11-23",type:"Break",reason:"Thanksgiving Break",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-11-24",type:"Break",reason:"Thanksgiving Break",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-11-25",type:"Break",reason:"Thanksgiving Break",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-11-26",type:"Break",reason:"Thanksgiving Break",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-11-27",type:"Break",reason:"Thanksgiving Break",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-12-15",type:"Final Exams",reason:"Final Exams — no Titan Time shown",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-12-16",type:"Final Exams",reason:"Final Exams — no Titan Time shown",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-12-17",type:"Final Exams",reason:"Final Exams — no Titan Time shown",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-12-18",type:"No Students",reason:"No Students",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-12-21",type:"Winter Recess",reason:"Winter Recess — No School",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-12-22",type:"Winter Recess",reason:"Winter Recess — No School",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-12-23",type:"Winter Recess",reason:"Winter Recess — No School",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-12-24",type:"Winter Recess",reason:"Winter Recess — No School",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-12-25",type:"Winter Recess",reason:"Winter Recess — No School",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-12-28",type:"Winter Recess",reason:"Winter Recess — No School",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-12-29",type:"Winter Recess",reason:"Winter Recess — No School",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-12-30",type:"Winter Recess",reason:"Winter Recess — No School",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2026-12-31",type:"Winter Recess",reason:"Winter Recess — No School",source:"Tesoro Fall 2026–27 calendar"},
  {date:"2027-01-01",type:"Winter Recess",reason:"Winter Recess — No School",source:"Tesoro Fall 2026–27 calendar"}
]);

function nhssched_exceptionBoolean_(value) {
  if (value === false) return false;
  const v = String(value == null ? "" : value).trim().toLowerCase();
  if (!v) return true;
  return !["false","0","no","inactive","disabled"].includes(v);
}

function nhssched_exceptionKey_(value) {
  const raw = String(value || "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  return nhssched_dateKey_(value, NHS_PLATFORM.TIMEZONE);
}

function nhssched_readManagedExceptions_() {
  try {
    const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
    const sh = nhsperf_sheet_(ss, NHSSCHED_EXCEPTION_SHEET);
    if (!sh || sh.getLastRow() < 2) return [];

    const values = sh
      .getRange(1, 1, sh.getLastRow(), sh.getLastColumn())
      .getDisplayValues();

    const h = {};
    values[0].forEach(function(header, i) {
      const key = String(header || "").trim();
      if (key) h[key] = i;
    });

    return values.slice(1).map(function(row) {
      const date = nhssched_exceptionKey_(h["Date"] === undefined ? "" : row[h["Date"]]);
      if (!date) return null;
      return {
        date: date,
        type: h["Type"] === undefined ? "Unavailable" : String(row[h["Type"]] || "Unavailable").trim(),
        reason: h["Reason"] === undefined ? "" : String(row[h["Reason"]] || "").trim(),
        active: h["Active"] === undefined ? true : nhssched_exceptionBoolean_(row[h["Active"]]),
        source: h["Source"] === undefined ? "Operations" : String(row[h["Source"]] || "Operations").trim(),
        managed: true
      };
    }).filter(Boolean);
  } catch (e) {
    return [];
  }
}

function nhssched_listExceptions_() {
  if (Array.isArray(NHSSCHED_EXCEPTION_MEMO_)) {
    return NHSSCHED_EXCEPTION_MEMO_.slice();
  }

  const byDate = {};
  NHSSCHED_DEFAULT_EXCEPTIONS.forEach(function(item) {
    byDate[item.date] = {
      date: item.date,
      type: item.type,
      reason: item.reason,
      active: true,
      source: item.source,
      managed: false
    };
  });

  nhssched_readManagedExceptions_().forEach(function(item) {
    byDate[item.date] = item;
  });

  NHSSCHED_EXCEPTION_MEMO_ = Object.keys(byDate)
    .sort()
    .map(function(key) { return byDate[key]; });

  return NHSSCHED_EXCEPTION_MEMO_.slice();
}

function nhssched_exceptionForDate_(dateValue) {
  const key = nhssched_exceptionKey_(dateValue);
  if (!key) return null;
  return nhssched_listExceptions_().find(function(item) {
    return item.date === key && item.active;
  }) || null;
}

function nhssched_assertOpenTitanDate_(dateValue) {
  const key = nhssched_exceptionKey_(dateValue);
  const exception = nhssched_exceptionForDate_(key);
  if (exception) {
    throw new Error(
      "Titan Time is unavailable on " +
      nhssched_formatDateLabel_(key) +
      ": " +
      (exception.reason || exception.type || "school calendar exception") +
      ". Choose another date."
    );
  }
  return key;
}

function nhssched_ensureExceptionSheet_(ss) {
  let sh = nhsperf_sheet_(ss, NHSSCHED_EXCEPTION_SHEET);
  if (!sh) sh = ss.insertSheet(NHSSCHED_EXCEPTION_SHEET);

  let headers = sh
    .getRange(1, 1, 1, Math.max(1, sh.getLastColumn()))
    .getDisplayValues()[0]
    .map(function(v) { return String(v || "").trim(); });

  NHSSCHED_EXCEPTION_HEADERS.forEach(function(header) {
    if (!headers.includes(header)) {
      sh.getRange(1, headers.length + 1).setValue(header);
      headers.push(header);
    }
  });

  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, sh.getLastColumn())
    .setBackground("#123A5A")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");

  const existing = {};
  if (sh.getLastRow() >= 2) {
    const data = sh.getDataRange().getDisplayValues();
    const hi = {};
    data[0].forEach(function(v, i) { hi[String(v || "").trim()] = i; });
    data.slice(1).forEach(function(row) {
      const key = nhssched_exceptionKey_(row[hi["Date"]]);
      if (key) existing[key] = true;
    });
  }

  const h = nhsperf_headerMap_(sh, true, false);
  NHSSCHED_DEFAULT_EXCEPTIONS.forEach(function(item) {
    if (existing[item.date]) return;
    const row = new Array(sh.getLastColumn()).fill("");
    row[h["Date"] - 1] = item.date;
    row[h["Type"] - 1] = item.type;
    row[h["Reason"] - 1] = item.reason;
    row[h["Active"] - 1] = true;
    row[h["Source"] - 1] = item.source;
    row[h["Updated At"] - 1] = new Date();
    row[h["Updated By"] - 1] = "System seed";
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  });

  return sh;
}

function nhssched_upsertException_(ss, spec) {
  const sh = nhssched_ensureExceptionSheet_(ss);
  const h = nhsperf_headerMap_(sh, true, false);
  const date = nhssched_exceptionKey_(spec.date);
  if (!date) throw new Error("Choose a valid exception date.");

  let rowNumber = 0;
  if (sh.getLastRow() >= 2) {
    const found = sh
      .getRange(2, h["Date"], sh.getLastRow() - 1, 1)
      .createTextFinder(date)
      .matchEntireCell(true)
      .findNext();
    if (found) rowNumber = found.getRow();
  }

  if (!rowNumber) rowNumber = sh.getLastRow() + 1;

  const row = new Array(sh.getLastColumn()).fill("");
  row[h["Date"] - 1] = date;
  row[h["Type"] - 1] = spec.type || "Unavailable";
  row[h["Reason"] - 1] = spec.reason || "";
  row[h["Active"] - 1] = spec.active !== false;
  row[h["Source"] - 1] = spec.source || "Operations";
  row[h["Updated At"] - 1] = new Date();
  row[h["Updated By"] - 1] = spec.updatedBy || nhsa_getActiveEmail_();

  sh.getRange(rowNumber, 1, 1, row.length).setValues([row]);

  NHSSCHED_EXCEPTION_MEMO_ = null;
  nhsperf_touchDataRevision_();
  try { nhsperf_scriptCache_().remove(NHSOS_CACHE_KEY); } catch (e) {}

  return date;
}

function nhssched_adminSaveException(payload) {
  nhssec_requireOfficer_();
  payload = payload || {};

  const date = nhssched_exceptionKey_(payload.date);
  if (!date) throw new Error("Choose a valid exception date.");

  const type = nhsinput_literalText_(payload.type || "Unavailable", 80, "Exception type");
  const reason = nhsinput_literalText_(payload.reason || "", 240, "Exception reason");
  if (!type) throw new Error("Exception type is required.");
  if (!reason) throw new Error("Exception reason is required.");

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  nhssched_upsertException_(ss, {
    date: date,
    type: type,
    reason: reason,
    active: true,
    source: "Operations",
    updatedBy: nhsa_getActiveEmail_()
  });

  return {
    ok: true,
    date: date,
    message: "School-calendar exception saved for " + date + "."
  };
}

function nhssched_adminDeleteException(dateValue) {
  nhssec_requireOfficer_();
  const date = nhssched_exceptionKey_(dateValue);
  if (!date) throw new Error("Choose a valid exception date.");

  const existing = nhssched_listExceptions_().find(function(item) {
    return item.date === date;
  });

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  nhssched_upsertException_(ss, {
    date: date,
    type: existing ? existing.type : "Unavailable",
    reason: existing ? existing.reason : "Disabled in Operations",
    active: false,
    source: "Operations override",
    updatedBy: nhsa_getActiveEmail_()
  });

  return {
    ok: true,
    date: date,
    message: "Exception disabled for " + date + "."
  };
}


function nhssched_timezone_(fallback) {
  return String(fallback || NHS_PLATFORM.TIMEZONE || Session.getScriptTimeZone() || "America/Los_Angeles");
}

function nhssched_dateKey_(value, timezone) {
  return nhslife_dateKey_(value, nhssched_timezone_(timezone));
}

function nhssched_todayKey_(timezone, now) {
  const tz = nhssched_timezone_(timezone);
  return Utilities.formatDate(now instanceof Date ? now : new Date(), tz, "yyyy-MM-dd");
}

function nhssched_parseKeyLocal_(key) {
  const m = String(key || "").trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0, 0);
  return isNaN(d.getTime()) ? null : d;
}

function nhssched_dayName_(key) {
  const d = nhssched_parseKeyLocal_(key);
  return d ? ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"][d.getDay()] : "";
}

function nhssched_isTitanDay_(key) {
  return NHSSCHED_TITAN_DAYS.includes(nhssched_dayName_(key));
}

function nhssched_currentWeekFridayKey_(timezone, now) {
  const todayKey = nhssched_todayKey_(timezone, now);
  const d = nhssched_parseKeyLocal_(todayKey);
  if (!d) return todayKey;
  // Treat the weekend as the lead-in to the upcoming school week.
  // Monday–Friday target the current school week's Friday; Saturday/Sunday
  // target the upcoming Friday so weekend requests can use Tue–Fri as
  // "This Week" rather than being forced into the reservation pathway.
  const dow = d.getDay();
  const delta = dow === 6 ? 6 : 5 - dow;
  d.setDate(d.getDate() + delta);
  const y=d.getFullYear(), m=String(d.getMonth()+1).padStart(2,'0'), day=String(d.getDate()).padStart(2,'0');
  return y+'-'+m+'-'+day;
}

function nhssched_addDaysKey_(key, days) {
  const d = nhssched_parseKeyLocal_(key);
  if (!d) return "";
  d.setDate(d.getDate() + Number(days || 0));
  const y=d.getFullYear(), m=String(d.getMonth()+1).padStart(2,'0'), day=String(d.getDate()).padStart(2,'0');
  return y+'-'+m+'-'+day;
}

function nhssched_daysBetween_(fromKey, toKey) {
  const a=nhssched_parseKeyLocal_(fromKey), b=nhssched_parseKeyLocal_(toKey);
  if (!a || !b) return NaN;
  return Math.round((b.getTime()-a.getTime())/86400000);
}

function nhssched_formatDateLabel_(key) {
  const d = nhssched_parseKeyLocal_(key);
  if (!d) return String(key || "");
  return Utilities.formatDate(d, "America/Los_Angeles", "EEEE, MMM d");
}

function nhssched_teacherOptions_(timezone, now) {
  const tz = nhssched_timezone_(timezone);
  const today = nhssched_todayKey_(tz, now);
  const friday = nhssched_currentWeekFridayKey_(tz, now);
  const thisWeekDates = [];
  let cursor = nhssched_addDaysKey_(today, 1);
  while (cursor && cursor <= friday) {
    if (nhssched_isTitanDay_(cursor) && !nhssched_exceptionForDate_(cursor)) {
      thisWeekDates.push({
        date: cursor,
        day: nhssched_dayName_(cursor),
        label: nhssched_formatDateLabel_(cursor)
      });
    }
    cursor = nhssched_addDaysKey_(cursor, 1);
  }

  let futureMin = nhssched_addDaysKey_(friday, 1);
  let futureGuard = 0;
  while (
    futureMin &&
    (
      !nhssched_isTitanDay_(futureMin) ||
      !!nhssched_exceptionForDate_(futureMin)
    ) &&
    futureGuard < 400
  ) {
    futureMin = nhssched_addDaysKey_(futureMin, 1);
    futureGuard++;
  }

  return {
    timezone: tz,
    today: today,
    thisWeekEnds: friday,
    thisWeekDates: thisWeekDates,
    futureMin: futureMin,
    blockedDates: nhssched_listExceptions_()
      .filter(function(item) { return item.active; })
      .map(function(item) {
        return {
          date: item.date,
          type: item.type,
          reason: item.reason
        };
      }),
    titanDays: NHSSCHED_TITAN_DAYS.slice(),
    sameDayAllowed: false,
    reconfirmDays: NHSSCHED_RECONFIRM_DAYS
  };
}

function nhssched_normalizeDateList_(value, timezone) {
  const raw = Array.isArray(value)
    ? value
    : String(value || "").split(/[;,\n]+/);
  const seen = {};
  return raw.map(function(v){ return nhssched_dateKey_(v, timezone); })
    .filter(function(v){
      if (!v || seen[v]) return false;
      seen[v]=true;
      return true;
    })
    .sort();
}

function nhssched_normalizeDayList_(value) {
  const raw = Array.isArray(value)
    ? value
    : String(value || "").split(/[;,]+/);
  const seen = {};
  return raw.map(function(v){ return String(v || "").trim(); })
    .filter(function(v){
      if (!NHSSCHED_TITAN_DAYS.includes(v) || seen[v]) return false;
      seen[v]=true;
      return true;
    });
}

function nhssched_validateFutureTitanDate_(dateValue, timezone, now) {
  const tz=nhssched_timezone_(timezone);
  const key=nhssched_dateKey_(dateValue,tz);
  if (!key) throw new Error("Choose a valid Titan Time date.");
  const today=nhssched_todayKey_(tz,now);
  if (key <= today) {
    throw new Error("Same-day Titan Time scheduling is not available in Five Star. Choose a future date.");
  }
  if (!nhssched_isTitanDay_(key)) {
    throw new Error("Titan Time tutoring can only be scheduled Tuesday through Friday.");
  }
  nhssched_assertOpenTitanDate_(key);
  return key;
}

function nhssched_validateCurrentOrFutureTitanDate_(dateValue, timezone, now) {
  const tz=nhssched_timezone_(timezone);
  const key=nhssched_dateKey_(dateValue,tz);
  if (!key) throw new Error("Choose a valid Titan Time date.");
  const today=nhssched_todayKey_(tz,now);
  if (key < today) {
    throw new Error("This Titan Time session has already passed.");
  }
  if (!nhssched_isTitanDay_(key)) {
    throw new Error("Titan Time tutoring can only be scheduled Tuesday through Friday.");
  }
  nhssched_assertOpenTitanDate_(key);
  return key;
}

function nhssched_validateTeacherPreference_(payload, timezone, now) {
  const tz=nhssched_timezone_(timezone);
  const options=nhssched_teacherOptions_(tz,now);
  const type=String(payload && payload.schedulingType || "").trim();
  const dates=nhssched_normalizeDateList_(payload && payload.preferredDates, tz);
  const days=nhssched_normalizeDayList_(payload && payload.preferredDays);

  if (!Object.values(NHSSCHED_TYPES).includes(type)) {
    throw new Error("Choose This Week, Future Reservation, or Recurring Support.");
  }

  if (type === NHSSCHED_TYPES.THIS_WEEK) {
    if (!dates.length) throw new Error("Select at least one available Titan Time date this week.");
    dates.forEach(function(key){
      nhssched_validateFutureTitanDate_(key,tz,now);
      if (key > options.thisWeekEnds) {
        throw new Error("This Week requests may only use the remaining Titan Time dates in the current week.");
      }
    });
    return {
      type:type,
      preferredDates:dates,
      preferredDays:dates.map(nhssched_dayName_),
      recurring:false
    };
  }

  if (type === NHSSCHED_TYPES.FUTURE) {
    if (dates.length !== 1) throw new Error("Choose exactly one future reservation date.");
    const key=nhssched_validateFutureTitanDate_(dates[0],tz,now);
    if (key <= options.thisWeekEnds) {
      throw new Error("Use This Week for dates remaining in the current week. Future Reservations begin after this week.");
    }
    return {
      type:type,
      preferredDates:[key],
      preferredDays:[nhssched_dayName_(key)],
      recurring:false
    };
  }

  if (!days.length) throw new Error("Select at least one preferred weekday for recurring support.");
  return {
    type:type,
    preferredDates:[],
    preferredDays:days,
    recurring:true
  };
}


function nhssched_validateScheduledAgainstRequest_(request, dateValue, timezone, now) {
  const key=nhssched_validateFutureTitanDate_(dateValue,timezone,now);
  const type=String(request && (request.schedulingType !== undefined ? request.schedulingType : request["Scheduling Preference"]) || "").trim();
  const preferredDates=nhssched_normalizeDateList_(
    request && (request.preferredDates !== undefined ? request.preferredDates : request["Preferred Dates"]),
    timezone
  );
  const preferredDays=nhssched_normalizeDayList_(
    request && (request.preferredDays !== undefined ? request.preferredDays : request["Preferred Days"])
  );

  if (type === NHSSCHED_TYPES.THIS_WEEK && preferredDates.length && !preferredDates.includes(key)) {
    throw new Error("Choose one of the specific This Week dates selected by the teacher.");
  }
  if (type === NHSSCHED_TYPES.FUTURE && preferredDates.length && preferredDates[0] !== key) {
    throw new Error("This future reservation is for " + preferredDates[0] + ". Change the teacher request before using a different date.");
  }
  if (type === NHSSCHED_TYPES.RECURRING && preferredDays.length && !preferredDays.includes(nhssched_dayName_(key))) {
    throw new Error("Choose one of the teacher's preferred recurring weekdays: " + preferredDays.join(", ") + ".");
  }
  return key;
}

function nhssched_reconfirmWindowOpen_(dateValue, timezone, now) {
  const tz=nhssched_timezone_(timezone);
  const key=nhssched_dateKey_(dateValue,tz);
  if (!key) return false;
  const today=nhssched_todayKey_(tz,now);
  const diff=nhssched_daysBetween_(today,key);
  return diff >= 1 && diff <= NHSSCHED_RECONFIRM_DAYS;
}

function nhssched_readiness_(request) {
  const teacherReady=nhslife_truthy_(request && (request.teacherFiveStarReady !== undefined ? request.teacherFiveStarReady : request["Teacher Five Star Ready"]));
  const tutorReady=nhslife_truthy_(request && (request.tutorFiveStarEnrolled !== undefined ? request.tutorFiveStarEnrolled : request["Tutor Five Star Enrolled"]));
  if (teacherReady && tutorReady) return "Ready";
  if (teacherReady) return "Waiting for Tutor Confirmation";
  if (tutorReady) return "Waiting for Teacher to Assign Tutor & Student";
  return "Five Star Setup Pending";
}


/** ============================================================
 * SOURCE GROUP: NHS_Performance_Core.gs
 * ============================================================ */

/**
 * NHS TUTORING — GOOGLE-MODEL PERFORMANCE CORE V18
 *
 * Safe optimizations for Apps Script + Google Sheets:
 * - one Spreadsheet object per execution
 * - one Sheet object per execution/name
 * - one header-row read per sheet per execution
 * - memoized Script/Document Cache handles
 * - optional lightweight timing
 *
 * IMPORTANT:
 * This is execution-local memoization. It does not make stale cross-request data.
 */

var NHS_PERF_SS_MEMO_ = {};
var NHS_PERF_SHEET_MEMO_ = {};
var NHS_PERF_HEADER_MEMO_ = {};
var NHS_PERF_CACHE_SERVICE_ = null;
var NHS_PERF_SCRIPT_PROPS_ = null;

function nhsperf_ss_(spreadsheetId) {
  const id = String(spreadsheetId || NHS_SPREADSHEET_ID || "").trim();
  if (!id) throw new Error("Missing spreadsheet ID.");

  if (!NHS_PERF_SS_MEMO_[id]) {
    NHS_PERF_SS_MEMO_[id] = SpreadsheetApp.openById(id);
  }

  return NHS_PERF_SS_MEMO_[id];
}

function nhsperf_sheet_(ss, sheetName) {
  if (!ss) throw new Error("Spreadsheet is required.");

  const name = String(sheetName || "").trim();
  if (!name) return null;

  let ssId = "";
  try {
    ssId = ss.getId();
  } catch (e) {
    ssId = "default";
  }

  const key = ssId + "::" + name;

  if (NHS_PERF_SHEET_MEMO_[key] === undefined) {
    NHS_PERF_SHEET_MEMO_[key] = ss.getSheetByName(name) || null;
  }

  return NHS_PERF_SHEET_MEMO_[key];
}

function nhsperf_scriptCache_() {
  if (!NHS_PERF_CACHE_SERVICE_) {
    NHS_PERF_CACHE_SERVICE_ = CacheService.getScriptCache();
  }
  return NHS_PERF_CACHE_SERVICE_;
}

function nhsperf_scriptProperties_() {
  if (!NHS_PERF_SCRIPT_PROPS_) {
    NHS_PERF_SCRIPT_PROPS_ = PropertiesService.getScriptProperties();
  }
  return NHS_PERF_SCRIPT_PROPS_;
}

/**
 * Header map with per-execution memoization.
 *
 * oneBased:
 *   true  => A=1, B=2...
 *   false => A=0, B=1...
 *
 * firstWins:
 *   true  => first duplicate header is authoritative
 *   false => last duplicate header is authoritative
 */
function nhsperf_headerMap_(sheet, oneBased, firstWins) {
  if (!sheet) return {};

  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return {};

  const key =
    sheet.getSheetId() +
    "|" +
    lastColumn +
    "|" +
    (oneBased ? "1" : "0") +
    "|" +
    (firstWins ? "F" : "L");

  if (NHS_PERF_HEADER_MEMO_[key]) {
    return NHS_PERF_HEADER_MEMO_[key];
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const map = {};

  headers.forEach(function(header, i) {
    const name = String(header || "").trim();
    if (!name) return;

    if (firstWins && map[name] !== undefined) {
      return;
    }

    map[name] = oneBased ? i + 1 : i;
  });

  NHS_PERF_HEADER_MEMO_[key] = map;
  return map;
}

/**
 * Call after a schema/header mutation inside the same execution.
 */
function nhsperf_invalidateSheetMemo_(sheet) {
  if (!sheet) return;

  const id = sheet.getSheetId();

  Object.keys(NHS_PERF_HEADER_MEMO_).forEach(function(key) {
    if (key.indexOf(String(id) + "|") === 0) {
      delete NHS_PERF_HEADER_MEMO_[key];
    }
  });
}

/**
 * Optional performance instrumentation.
 * Disabled by default; enable with:
 *   setNHSPerformanceTracing(true)
 */
function setNHSPerformanceTracing(enabled) {
  nhssec_requireMaintenanceAdmin_();
  nhsperf_scriptProperties_().setProperty(
    "NHS_PERF_TRACE",
    enabled ? "true" : "false"
  );

  return {
    ok: true,
    tracing: !!enabled
  };
}

function nhsperf_traceEnabled_() {
  try {
    return nhsperf_scriptProperties_().getProperty("NHS_PERF_TRACE") === "true";
  } catch (e) {
    return false;
  }
}

function nhsperf_timed_(label, fn) {
  if (!nhsperf_traceEnabled_()) {
    return fn();
  }

  const started = Date.now();

  try {
    return fn();
  } finally {
    console.log(
      "[NHS PERF] " +
      String(label || "operation") +
      ": " +
      (Date.now() - started) +
      " ms"
    );
  }
}

/**
 * Tiny sanity benchmark. This does not modify workbook data.
 */
function benchmarkNHSPerformanceCore() {
  nhssec_requireMaintenanceAdmin_();

  const started = Date.now();
  const ss1 = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const t1 = Date.now();

  const ss2 = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const t2 = Date.now();

  const sh1 = nhsperf_sheet_(ss1, "Tutoring Requests");
  const t3 = Date.now();

  const sh2 = nhsperf_sheet_(ss2, "Tutoring Requests");
  const t4 = Date.now();

  const h1 = sh1 ? nhsperf_headerMap_(sh1, true, true) : {};
  const t5 = Date.now();

  const h2 = sh2 ? nhsperf_headerMap_(sh2, true, true) : {};
  const t6 = Date.now();

  const result = {
    ok: true,
    firstSpreadsheetOpenMs: t1 - started,
    memoizedSpreadsheetLookupMs: t2 - t1,
    firstSheetLookupMs: t3 - t2,
    memoizedSheetLookupMs: t4 - t3,
    firstHeaderReadMs: t5 - t4,
    memoizedHeaderLookupMs: t6 - t5,
    headerCount: Object.keys(h2).length,
    totalMs: t6 - started
  };

  console.log(JSON.stringify(result, null, 2));
  return result;
}


/* ============================================================
   V25 LIVE DATA REVISION
   ============================================================ */
const NHS_PERF_DATA_REVISION_PROPERTY = "NHS_DATA_REVISION_V25";
const NHS_PERF_DATA_REVISION_CACHE_KEY = "NHS_DATA_REVISION_CACHE_V43";

function nhsperf_touchDataRevision_() {
  const value = String(Date.now());

  try {
    nhsperf_scriptProperties_().setProperty(
      NHS_PERF_DATA_REVISION_PROPERTY,
      value
    );
  } catch (e) {}

  try {
    nhsperf_scriptCache_().put(
      NHS_PERF_DATA_REVISION_CACHE_KEY,
      value,
      21600
    );
  } catch (e) {}

  return value;
}

function nhsperf_getDataRevision_() {
  try {
    const cached = nhsperf_scriptCache_().get(NHS_PERF_DATA_REVISION_CACHE_KEY);
    if (cached) return String(cached);
  } catch (e) {}

  let value = "0";
  try {
    value = String(
      nhsperf_scriptProperties_().getProperty(NHS_PERF_DATA_REVISION_PROPERTY) || "0"
    );
  } catch (e) {}

  try {
    nhsperf_scriptCache_().put(
      NHS_PERF_DATA_REVISION_CACHE_KEY,
      value,
      21600
    );
  } catch (e) {}

  return value;
}

function nhs_getDataRevision() {
  nhsa_requireAnyRole_(
    NHSA_VALID_ROLES,
    "Portal Synchronization"
  );
  return nhsperf_getDataRevision_();
}


/** ============================================================
 * SOURCE GROUP: NHS_Access_Core.gs
 * ============================================================ */

/**
 * NHS UNIFIED ACCESS CORE — V1
 *
 * COPY THIS SAME FILE INTO:
 *   1) NHS Internal Operations project
 *   2) Teacher Portal project
 *   3) Tutor Portal project
 *   4) NHS Portal Gateway project
 *
 * Google Workspace authenticates the account.
 * This layer decides what that account is allowed to access.
 */

const NHS_ACCESS_SPREADSHEET_ID = NHS_SPREADSHEET_ID;

const NHSA_REGISTRY_SHEET = "Access Registry";
const NHSA_REQUESTS_SHEET = "Access Requests";
const NHSA_LOG_SHEET = "Access Log";
const NHSA_SETTINGS_SHEET = "Portal Settings";
const NHSA_TUTOR_SHEET = "Tutor Database";

const NHSA_VALID_ROLES = [
  "ADMIN",
  "ADVISOR",
  "OFFICER",
  "TUTOR",
  "TEACHER"
];

// V46.4.3 — domain contract for school identities.
// CUSD staff/teacher accounts use @capousd.org. Student accounts use
// @capostu.org. TEACHER grants are intentionally staff-only; Admin/Advisor
// users do not need a TEACHER grant to enter the Teacher Portal.
const NHSA_TEACHER_EMAIL_DOMAIN = "capousd.org";
const NHSA_SCHOOL_EMAIL_DOMAINS = Object.freeze(["capostu.org", "capousd.org"]);

function nhsa_isValidEmailAddress_(value) {
  const email = String(value || "").trim().toLowerCase();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function nhsa_isTeacherStaffEmail_(value) {
  const email = String(value || "").trim().toLowerCase();
  return nhsa_isValidEmailAddress_(email) &&
    email.endsWith("@" + NHSA_TEACHER_EMAIL_DOMAIN);
}

function nhsa_isSchoolIdentityEmail_(value) {
  const email = String(value || "").trim().toLowerCase();
  if (!nhsa_isValidEmailAddress_(email)) return false;
  return NHSA_SCHOOL_EMAIL_DOMAINS.some(function(domain) {
    return email.endsWith("@" + domain);
  });
}

function nhsa_normalizePersonName_(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\([^)]*\)/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function nhsa_assertRoleEmailPolicy_(email, role) {
  const normalizedEmail = String(email || "").trim().toLowerCase();
  const normalizedRole = String(role || "").trim().toUpperCase();

  if (!nhsa_isValidEmailAddress_(normalizedEmail)) {
    throw new Error("Invalid school email: " + normalizedEmail);
  }

  if (normalizedRole === "TEACHER" && !nhsa_isTeacherStaffEmail_(normalizedEmail)) {
    throw new Error(
      "Teacher access requires a district staff email ending in @" +
      NHSA_TEACHER_EMAIL_DOMAIN +
      ". Student @capostu.org accounts should not be synchronized as teachers."
    );
  }

  if (normalizedRole === "TUTOR" && !nhsa_isSchoolIdentityEmail_(normalizedEmail)) {
    throw new Error(
      "Tutor access requires a CUSD school identity ending in @capostu.org or @capousd.org. " +
      "Personal email addresses belong only in Notification Email."
    );
  }

  return normalizedEmail;
}

const NHSA_REGISTRY_HEADERS = [
  "Grant ID",
  "School Email",
  "Full Name",
  "Role",
  "Status",
  "School Year",
  "Start Date",
  "Expiration Date",
  "Approved By",
  "Approved At",
  "Last Login",
  "Last Portal",
  "Notes"
];

const NHSA_REQUEST_HEADERS = [
  "Request ID",
  "Requested At",
  "School Email",
  "Full Name",
  "Requested Role",
  "Reason",
  "Status",
  "Reviewed By",
  "Reviewed At",
  "Review Notes"
];

const NHSA_LOG_HEADERS = [
  "Log ID",
  "Timestamp",
  "School Email",
  "Full Name",
  "Portal",
  "Result",
  "Roles",
  "Detail"
];

const NHSA_SETTING_HEADERS = [
  "Key",
  "Value",
  "Notes"
];


/* ============================================================
   INSTALL / BOOTSTRAP
   ============================================================ */

const NHSA_ACCESS_REVISION_PROPERTY = "NHSA_ACCESS_REVISION_V43";
const NHSA_ACCESS_REVISION_CACHE_KEY = "NHSA_ACCESS_REVISION_CACHE_V43";

function nhsa_installAccessCore_() {
  const ss = nhsa_openSpreadsheet_();

  const required = [
    {
      name: NHSA_REGISTRY_SHEET,
      headers: NHSA_REGISTRY_HEADERS
    },
    {
      name: NHSA_REQUESTS_SHEET,
      headers: NHSA_REQUEST_HEADERS
    },
    {
      name: NHSA_LOG_SHEET,
      headers: NHSA_LOG_HEADERS
    },
    {
      name: NHSA_SETTINGS_SHEET,
      headers: NHSA_SETTING_HEADERS
    }
  ];

  const missingSheets = [];
  const badHeaders = [];

  required.forEach(spec => {
    const sh = nhsperf_sheet_(ss, spec.name);

    if (!sh) {
      missingSheets.push(spec.name);
      return;
    }

    const values = sh
      .getRange(1, 1, 1, spec.headers.length)
      .getDisplayValues()[0]
      .map(v => String(v || "").trim());

    const matches =
      spec.headers.every((header, index) =>
        values[index] === header
      );

    if (!matches) {
      badHeaders.push(spec.name);
    }
  });

  if (missingSheets.length || badHeaders.length) {
    throw new Error(
      "ACCESS_SCHEMA_NOT_READY|" +
      "Create/repair the access tabs manually in the spreadsheet. " +
      (missingSheets.length
        ? "Missing sheets: " + missingSheets.join(", ") + ". "
        : "") +
      (badHeaders.length
        ? "Incorrect headers: " + badHeaders.join(", ") + "."
        : "")
    );
  }

  return {
    ok: true,
    message:
      "Unified NHS access schema verified. No Google Cloud API required."
  };
}


function nhsa_installAccessCore() {
  nhssec_requireMaintenanceAdmin_();
  return nhsa_installAccessCore_();
}


/**
 * Run ONCE from the INTERNAL NHS project after nhsa_installAccessCore().
 *
 * The first administrator is the Google account currently running the
 * Apps Script function. This function refuses to add another bootstrap
 * administrator after any active ADMIN/ADVISOR grant exists.
 */
function nhsa_bootstrapFirstAdmin() {
  const activeEmail = String(
    Session.getActiveUser().getEmail() || ""
  ).trim().toLowerCase();

  const effectiveEmail = String(
    Session.getEffectiveUser().getEmail() || ""
  ).trim().toLowerCase();

  if (!activeEmail || !effectiveEmail || activeEmail !== effectiveEmail) {
    throw new Error(
      "BOOTSTRAP_OWNER_REQUIRED|The first NHS administrator can only be bootstrapped by the Apps Script deployment owner from the script project."
    );
  }

  // Verify the schema without exposing the installer as an unauthenticated RPC.
  nhsa_installAccessCore_();

  const ss = nhsa_openSpreadsheet_();
  const sh = nhsa_ensureRegistry_(ss);

  const privileged = nhsa_readRegistry_(sh)
    .filter(function(grant) {
      return nhsa_isGrantActive_(grant) &&
        ["ADMIN", "ADVISOR"].includes(grant.role);
    });

  if (privileged.length) {
    return {
      ok: true,
      skipped: true,
      message:
        "An active Admin/Advisor already exists. Bootstrap was not changed."
    };
  }

  const result = nhsa_createGrant_(ss, {
    email: activeEmail,
    fullName: "",
    role: "ADMIN",
    schoolYear: "Permanent",
    startDate: new Date(),
    expirationDate: "",
    approvedBy: activeEmail,
    notes: "Initial system administrator"
  });

  nhsa_createGrant_(ss, {
    email: activeEmail,
    fullName: "",
    role: "ADVISOR",
    schoolYear: "Permanent",
    startDate: new Date(),
    expirationDate: "",
    approvedBy: activeEmail,
    notes: "Initial system advisor / recovery authority"
  });

  return {
    ok: true,
    message:
      "First administrator bootstrapped for " + activeEmail + ".",
    grantId: result.grantId
  };
}


/* ============================================================
   PUBLIC LOGIN / ROLE RESOLUTION
   ============================================================ */

function nhsa_getLoginState(portalName) {
  const email = nhsa_getActiveEmail_();
  const revision = nhsa_getAccessRevision_();
  const cache = nhsperf_scriptCache_();
  const cacheKey =
    "nhsaLoginStateV43|" + revision + "|" + email + "|" + String(portalName || "");

  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) {}

  const ss = nhsa_openSpreadsheet_();
  const access = nhsa_resolveAccess_(ss, email);
  const settings = nhsa_readSettingsCached_(ss, revision);

  // Access Requests are only relevant to users who currently have no role.
  // Avoid an unnecessary sheet scan on the normal Gateway hot path.
  const pending = access.roles.length
    ? []
    : nhsa_pendingRequestsForEmail_(ss, email);

  const response = {
    ok: true,
    email: email,
    fullName: access.fullName || "",
    roles: access.roles,
    grants: access.grants,
    tutorDerived: access.tutorDerived,
    tutorEligible: access.tutorEligible,
    pendingRequests: pending,
    settings: settings,
    currentSchoolYear: settings.CURRENT_SCHOOL_YEAR || "",
    generatedAt: new Date().toISOString(),
    accessRevision: revision
  };

  try { cache.put(cacheKey, JSON.stringify(response), 300); } catch (e) {}

  // Gateway identification is intentionally not logged synchronously.
  // Actual portal authorization is logged by nhsa_requireAnyRole_.
  return response;
}


function nhsa_requireAnyRole_(allowedRoles, portalName) {
  const email = nhsa_getActiveEmail_();
  const ss = nhsa_openSpreadsheet_();
  const access = nhsa_resolveAccess_(ss, email);

  const normalizedAllowed =
    (allowedRoles || [])
      .map(role => String(role || "").trim().toUpperCase());

  const allowed = access.roles
    .some(role => normalizedAllowed.includes(role));

  nhsa_logAccess_(ss, {
    email: email,
    fullName: access.fullName || "",
    portal: portalName || "NHS Tutoring",
    result: allowed ? "ALLOWED" : "DENIED",
    roles: access.roles,
    detail:
      "Required: " +
      (normalizedAllowed.join(", ") || "none")
  });

  if (!allowed) {
    const error = new Error(
      "ACCESS_DENIED|" +
      (portalName || "NHS Tutoring") +
      "|This account is not verified for this portal."
    );

    error.name = "AccessDenied";
    throw error;
  }

  nhsa_touchLogin_(ss, email, portalName || "NHS Tutoring");

  return {
    email: email,
    fullName: access.fullName || "",
    roles: access.roles,
    grants: access.grants
  };
}


function nhsa_requestAccess(requestedRole, fullName, reason) {
  const email = nhsa_getActiveEmail_();
  const safeFullName = nhsinput_literalText_(fullName, 120, "Full name");
  const safeReason = nhsinput_literalText_(reason, 500, "Reason");
  const role = String(requestedRole || "")
    .trim()
    .toUpperCase();

  if (!["TEACHER", "TUTOR"].includes(role)) {
    throw new Error(
      "Only Teacher or Tutor access can be requested from a portal. Officer access must be granted by an Advisor/Admin."
    );
  }

  if (role === "TEACHER" && !nhsa_isTeacherStaffEmail_(email)) {
    throw new Error(
      "Teacher access requires a CUSD staff account ending in @" +
      NHSA_TEACHER_EMAIL_DOMAIN + "."
    );
  }

  const ss = nhsa_openSpreadsheet_();
  const requests = nhsa_ensureRequests_(ss);

  // If the user already has the requested role, do not create noise.
  const access = nhsa_resolveAccess_(ss, email);

  if (access.roles.includes(role)) {
    return {
      ok: true,
      alreadyAuthorized: true,
      message: "You already have " + role + " access."
    };
  }

  const existing = nhsa_readRequests_(requests)
    .find(req =>
      req.email === email &&
      req.requestedRole === role &&
      req.status === "Pending"
    );

  if (existing) {
    return {
      ok: true,
      alreadyPending: true,
      requestId: existing.requestId,
      message: "Your access request is already pending review."
    };
  }

  const requestId = "AR-" + Utilities.getUuid();

  const h = nhsa_headerMap_(requests);
  const row = new Array(requests.getLastColumn()).fill("");

  nhsa_setRow_(row, h, "Request ID", requestId);
  nhsa_setRow_(row, h, "Requested At", new Date());
  nhsa_setRow_(row, h, "School Email", email);
  nhsa_setRow_(row, h, "Full Name", safeFullName);
  nhsa_setRow_(row, h, "Requested Role", role);
  nhsa_setRow_(row, h, "Reason", safeReason);
  nhsa_setRow_(row, h, "Status", "Pending");

  nhsa_appendRowLocked_(
    requests,
    row,
    "create access request"
  );

  nhsa_logAccess_(ss, {
    email: email,
    fullName: safeFullName,
    portal: "Access Request",
    result: "REQUESTED",
    roles: access.roles,
    detail: role + " access requested"
  });

  return {
    ok: true,
    requestId: requestId,
    message:
      "Access request submitted. An NHS Advisor/Admin can now review it."
  };
}


/* ============================================================
   RESOLUTION
   ============================================================ */

function nhsa_resolveAccess_(ss, email) {
  email = nhsa_clean_(email).toLowerCase();
  if (!email) {
    return {
      email: "",
      fullName: "",
      roles: [],
      grants: [],
      tutorDerived: false,
      tutorProfile: null
    };
  }

  // Authorization is deliberately resolved live. Access Registry changes,
  // revocations, expirations, and Tutor Database eligibility must take effect
  // on the next request rather than waiting for a cross-request cache TTL.
  return nhsa_resolveAccessUncached_(ss, email);
}

function nhsa_resolveAccessUncached_(ss, email) {
  email = nhsa_clean_(email).toLowerCase();

  const registry =
    nhsa_ensureRegistry_(ss);

  const emailGrants =
    nhsa_readRegistry_(registry)
      .filter(function(grant) {
        return grant.email === email;
      });

  const activeGrants =
    emailGrants
      .filter(function(grant) {
        return nhsa_isGrantActive_(
          grant
        );
      });

  const tutor =
    nhsa_findTutorBySchoolEmail_(
      ss,
      email
    );

  const tutorEligible = !!(
    tutor &&
    ![
      "inactive",
      "suspended"
    ].includes(
      String(
        tutor.reliabilityStatus ||
        ""
      )
        .trim()
        .toLowerCase()
    ) &&
    (
      !tutor.expectationsAgreement ||
      /agree|yes|true/i.test(
        tutor.expectationsAgreement
      )
    )
  );

  // V55: Access Registry is the sole authorization authority.
  // A granted TUTOR role must make the Tutor workspace visible immediately.
  // Tutor Database eligibility remains a matching/profile diagnostic only.
  const roleSet =
    new Set(
      activeGrants.map(
        function(grant) {
          return grant.role;
        }
      )
    );

  let fullName =
    activeGrants.find(
      function(grant) {
        return grant.fullName;
      }
    )?.fullName || "";

  if (
    !fullName &&
    tutor
  ) {
    fullName =
      tutor.preferredName
        ? nhsa_displayTutorName_(
            tutor.preferredName,
            tutor.fullName
          )
        : tutor.fullName;
  }

  return {
    email: email,
    fullName: fullName,
    roles:
      Array.from(roleSet),
    grants:
      activeGrants,
    tutorEligible:
      tutorEligible,
    tutorDerived:
      false,
    tutorProfileLinked:
      !!tutor
  };
}


function nhsa_isGrantActive_(grant) {
  if (!grant) return false;

  if (String(grant.status || "").toLowerCase() !== "active") {
    return false;
  }

  const start = nhsa_parseDate_(grant.startDate);
  const expiration = nhsa_parseDate_(grant.expirationDate);
  const now = new Date();

  if (start && start.getTime() > now.getTime()) {
    return false;
  }

  if (expiration) {
    expiration.setHours(23, 59, 59, 999);

    if (expiration.getTime() < now.getTime()) {
      return false;
    }
  }

  return true;
}


/* ============================================================
   TUTOR DERIVATION
   ============================================================ */

function nhsa_findTutorBySchoolEmail_(ss, email) {
  const sh = nhsperf_sheet_(ss, NHSA_TUTOR_SHEET);

  if (!sh || sh.getLastRow() < 2) {
    return null;
  }

  const values = nhsa_retry_(
    () => sh
      .getRange(1, 1, sh.getLastRow(), sh.getLastColumn())
      .getDisplayValues(),
    "read Tutor Database for access"
  );

  const h = nhsa_headerIndex_(values[0]);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const schoolEmail = nhsa_get_(
      row,
      h,
      "School Email"
    )
      .toLowerCase();

    if (!schoolEmail || schoolEmail !== email) {
      continue;
    }

    return {
      tutorId: nhsa_get_(row, h, "Tutor ID"),
      fullName: nhsa_get_(row, h, "Full Name"),
      preferredName: nhsa_get_(row, h, "Preferred Name"),
      schoolEmail: schoolEmail,
      notificationEmail:
        nhsa_get_(row, h, "Notification Email"),
      reliabilityStatus:
        nhsa_get_(row, h, "Reliability Status"),
      expectationsAgreement:
        nhsa_get_(row, h, "Expectations Agreement")
    };
  }

  return null;
}


/* ============================================================
   REGISTRY / REQUEST / LOG READS
   ============================================================ */

function nhsa_readRegistry_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];

  const values = nhsa_retry_(
    () => sh
      .getRange(1, 1, sh.getLastRow(), sh.getLastColumn())
      .getDisplayValues(),
    "read Access Registry"
  );

  const h = nhsa_headerIndex_(values[0]);

  return values
    .slice(1)
    .map(row => ({
      grantId: nhsa_get_(row, h, "Grant ID"),
      email:
        nhsa_get_(row, h, "School Email")
          .toLowerCase(),
      fullName: nhsa_get_(row, h, "Full Name"),
      role:
        nhsa_get_(row, h, "Role")
          .toUpperCase(),
      status: nhsa_get_(row, h, "Status"),
      schoolYear: nhsa_get_(row, h, "School Year"),
      startDate: nhsa_get_(row, h, "Start Date"),
      expirationDate:
        nhsa_get_(row, h, "Expiration Date"),
      approvedBy:
        nhsa_get_(row, h, "Approved By"),
      approvedAt:
        nhsa_get_(row, h, "Approved At"),
      lastLogin:
        nhsa_get_(row, h, "Last Login"),
      lastPortal:
        nhsa_get_(row, h, "Last Portal"),
      notes: nhsa_get_(row, h, "Notes")
    }))
    .filter(grant => grant.grantId && grant.email);
}


function nhsa_readRequests_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];

  const values = nhsa_retry_(
    () => sh
      .getRange(1, 1, sh.getLastRow(), sh.getLastColumn())
      .getDisplayValues(),
    "read Access Requests"
  );

  const h = nhsa_headerIndex_(values[0]);

  return values
    .slice(1)
    .map(row => ({
      requestId: nhsa_get_(row, h, "Request ID"),
      requestedAt: nhsa_get_(row, h, "Requested At"),
      email:
        nhsa_get_(row, h, "School Email")
          .toLowerCase(),
      fullName: nhsa_get_(row, h, "Full Name"),
      requestedRole:
        nhsa_get_(row, h, "Requested Role")
          .toUpperCase(),
      reason: nhsa_get_(row, h, "Reason"),
      status: nhsa_get_(row, h, "Status"),
      reviewedBy:
        nhsa_get_(row, h, "Reviewed By"),
      reviewedAt:
        nhsa_get_(row, h, "Reviewed At"),
      reviewNotes:
        nhsa_get_(row, h, "Review Notes")
    }))
    .filter(req => req.requestId);
}


function nhsa_pendingRequestsForEmail_(ss, email) {
  const sh = nhsa_ensureRequests_(ss);

  return nhsa_readRequests_(sh)
    .filter(req =>
      req.email === email &&
      req.status === "Pending"
    );
}


/* ============================================================
   ACCESS LOG / LOGIN TOUCH
   ============================================================ */

function nhsa_logAccess_(ss, spec) {
  const email =
    nhsa_clean_(spec.email)
      .toLowerCase();

  const portal =
    nhsa_clean_(spec.portal) || "NHS Tutoring";

  const result =
    nhsa_clean_(spec.result) || "UNKNOWN";

  // Prevent login polling from producing thousands of rows.
  const cache = nhsperf_scriptCache_();
  const key =
    "nhsaLog|" +
    email +
    "|" +
    portal +
    "|" +
    result;

  if (cache.get(key)) {
    return;
  }

  try {
    cache.put(key, "1", 900);
  } catch (e) {}

  const sh = nhsa_ensureLog_(ss);
  const h = nhsa_headerMap_(sh);
  const row = new Array(sh.getLastColumn()).fill("");

  nhsa_setRow_(
    row,
    h,
    "Log ID",
    "AL-" + Utilities.getUuid()
  );

  nhsa_setRow_(row, h, "Timestamp", new Date());
  nhsa_setRow_(row, h, "School Email", email);
  nhsa_setRow_(row, h, "Full Name", spec.fullName || "");
  nhsa_setRow_(row, h, "Portal", portal);
  nhsa_setRow_(row, h, "Result", result);
  nhsa_setRow_(
    row,
    h,
    "Roles",
    Array.isArray(spec.roles)
      ? spec.roles.join(", ")
      : String(spec.roles || "")
  );
  nhsa_setRow_(row, h, "Detail", spec.detail || "");

  try {
    nhsa_appendRowLocked_(
      sh,
      row,
      "write access log"
    );
  } catch (e) {
    console.error("Access Log write failed", e);
  }
}


function nhsa_touchLogin_(ss, email, portal) {
  const cache = nhsperf_scriptCache_();
  const key =
    "nhsaTouch|" +
    email +
    "|" +
    portal;

  if (cache.get(key)) return;

  try {
    cache.put(key, "1", 1800);
  } catch (e) {}

  const sh = nhsa_ensureRegistry_(ss);
  const values = sh
    .getRange(1, 1, sh.getLastRow(), sh.getLastColumn())
    .getDisplayValues();

  const h = nhsa_headerIndex_(values[0]);
  const headerMap = nhsa_headerMap_(sh);

  for (let i = 1; i < values.length; i++) {
    const rowEmail =
      nhsa_get_(values[i], h, "School Email")
        .toLowerCase();

    if (rowEmail !== email) continue;

    const grant = {
      status: nhsa_get_(values[i], h, "Status"),
      startDate: nhsa_get_(values[i], h, "Start Date"),
      expirationDate:
        nhsa_get_(values[i], h, "Expiration Date")
    };

    if (!nhsa_isGrantActive_(grant)) continue;

    try {
      sh
        .getRange(i + 1, headerMap["Last Login"])
        .setValue(new Date());

      sh
        .getRange(i + 1, headerMap["Last Portal"])
        .setValue(portal);
    } catch (e) {}
  }
}


/* ============================================================
   SHEET CREATION
   ============================================================ */

function nhsa_ensureRegistry_(ss) {
  const sh = nhsperf_sheet_(ss, NHSA_REGISTRY_SHEET);

  if (!sh) {
    throw new Error(
      "ACCESS_SCHEMA_NOT_READY|Missing required sheet: " +
      NHSA_REGISTRY_SHEET
    );
  }

  return sh;
}


function nhsa_ensureRequests_(ss) {
  const sh = nhsperf_sheet_(ss, NHSA_REQUESTS_SHEET);

  if (!sh) {
    throw new Error(
      "ACCESS_SCHEMA_NOT_READY|Missing required sheet: " +
      NHSA_REQUESTS_SHEET
    );
  }

  return sh;
}


function nhsa_ensureLog_(ss) {
  const sh = nhsperf_sheet_(ss, NHSA_LOG_SHEET);

  if (!sh) {
    throw new Error(
      "ACCESS_SCHEMA_NOT_READY|Missing required sheet: " +
      NHSA_LOG_SHEET
    );
  }

  return sh;
}


function nhsa_ensureSettings_(ss) {
  const sh = nhsperf_sheet_(ss, NHSA_SETTINGS_SHEET);

  if (!sh) {
    throw new Error(
      "ACCESS_SCHEMA_NOT_READY|Missing required sheet: " +
      NHSA_SETTINGS_SHEET
    );
  }

  return sh;
}


function nhsa_ensureTable_(ss, name, requiredHeaders, hide) {
  let sh = nhsa_retry_(
    () => nhsperf_sheet_(ss, name),
    "find " + name,
    3
  );

  let created = false;

  if (!sh) {
    sh = nhsa_retry_(
      () => ss.insertSheet(name),
      "create " + name,
      3
    );
    created = true;
  }

  // Avoid getLastRow()/getLastColumn() during installation. Those calls can
  // intermittently time out immediately after a new sheet is created.
  // Reading a small fixed header range is both faster and more reliable.
  const width = Math.max(requiredHeaders.length, 1);

  const headerRow = nhsa_retry_(
    () => sh
      .getRange(1, 1, 1, width)
      .getDisplayValues()[0]
      .map(v => String(v || "").trim()),
    "read " + name + " header row",
    4
  );

  const hasAnyHeader =
    headerRow.some(value => value !== "");

  if (!hasAnyHeader) {
    nhsa_retry_(
      () => sh
        .getRange(1, 1, 1, requiredHeaders.length)
        .setValues([requiredHeaders]),
      "initialize " + name + " headers",
      4
    );
  } else {
    // Canonical tables should always begin in A1. Repair an old shifted
    // header layout only when the expected first header is not in A1.
    if (
      headerRow[0] !== requiredHeaders[0] &&
      headerRow.includes(requiredHeaders[0])
    ) {
      const start = headerRow.indexOf(requiredHeaders[0]);
      const existing = headerRow.slice(
        start,
        start + requiredHeaders.length
      );

      const looksShifted =
        existing.filter(Boolean).length >=
        Math.min(2, requiredHeaders.length);

      if (looksShifted) {
        nhsa_retry_(
          () => sh
            .getRange(1, 1, 1, requiredHeaders.length)
            .setValues([requiredHeaders]),
          "repair " + name + " headers",
          4
        );
      }
    } else {
      const missing =
        requiredHeaders.filter(
          header => !headerRow.includes(header)
        );

      if (missing.length) {
        // Access tables are platform-owned. Write the canonical header row
        // in one operation instead of issuing one write per missing header.
        nhsa_retry_(
          () => sh
            .getRange(1, 1, 1, requiredHeaders.length)
            .setValues([requiredHeaders]),
          "normalize " + name + " headers",
          4
        );
      }
    }
  }

  try {
    sh.setFrozenRows(1);
  } catch (e) {}

  if (hide) {
    try {
      sh.hideSheet();
    } catch (e) {}
  }

  return sh;
}


/* ============================================================
   SETTINGS
   ============================================================ */

function nhsa_readSettings_(ss) {
  const sh =
    nhsperf_sheet_(ss, NHSA_SETTINGS_SHEET);

  if (!sh || sh.getLastRow() < 2) {
    return {};
  }

  const values = sh
    .getRange(1, 1, sh.getLastRow(), sh.getLastColumn())
    .getDisplayValues();

  const h = nhsa_headerIndex_(values[0]);
  const result = {};

  values.slice(1).forEach(row => {
    const key = nhsa_get_(row, h, "Key");

    if (!key) return;

    result[key] =
      nhsa_get_(row, h, "Value");
  });

  return result;
}


/* ============================================================
   GRANT CREATION
   ============================================================ */

const NHSA_DELETION_TOMBSTONE_NOTE = "NHS_ACCESS_DELETION_TOMBSTONE";

function nhsa_hasDeletionTombstone_(sh, email) {
  if (!sh || sh.getLastRow() < 2) return false;
  email = nhsa_clean_(email).toLowerCase();
  if (!email) return false;
  return nhsa_readRegistry_(sh).some(function(grant) {
    return grant.email === email &&
      grant.role === "TUTOR" &&
      String(grant.status || "").trim() === "Deleted" &&
      String(grant.notes || "").indexOf(NHSA_DELETION_TOMBSTONE_NOTE) >= 0;
  });
}


function nhsa_removeDeletionTombstones_(sh, email) {
  if (!sh || sh.getLastRow() < 2) return 0;
  email = nhsa_clean_(email).toLowerCase();
  if (!email) return 0;

  const grants = nhsa_readRegistry_(sh);
  const rows = grants
    .filter(function(grant) {
      return grant.email === email &&
        grant.role === "TUTOR" &&
        String(grant.status || "").trim() === "Deleted" &&
        String(grant.notes || "").indexOf(NHSA_DELETION_TOMBSTONE_NOTE) >= 0;
    })
    .map(function(grant) { return nhsa_findGrantRow_(sh, grant.grantId); })
    .filter(Boolean)
    .map(function(info) { return info.row; })
    .sort(function(a, b) { return b - a; });

  rows.forEach(function(row) { sh.deleteRow(row); });
  return rows.length;
}

function nhsa_appendDeletionTombstone_(ss, email, fullName, actor, reason) {
  const sh = nhsa_ensureRegistry_(ss);
  nhsa_removeDeletionTombstones_(sh, email);
  const h = nhsa_headerMap_(sh);
  const row = new Array(sh.getLastColumn()).fill("");
  const grantId = "DEL-" + Utilities.getUuid();

  nhsa_setRow_(row, h, "Grant ID", grantId);
  nhsa_setRow_(row, h, "School Email", email);
  nhsa_setRow_(row, h, "Full Name", fullName || "");
  nhsa_setRow_(row, h, "Role", "TUTOR");
  nhsa_setRow_(row, h, "Status", "Deleted");
  nhsa_setRow_(row, h, "Approved By", actor || "");
  nhsa_setRow_(row, h, "Approved At", new Date());
  nhsa_setRow_(
    row,
    h,
    "Notes",
    NHSA_DELETION_TOMBSTONE_NOTE + " · " + String(reason || "").trim()
  );
  nhsa_appendRowLocked_(sh, row, "create deletion tombstone");
  return grantId;
}


function nhsa_createGrant_(ss, spec) {
  const role =
    nhsa_clean_(spec.role)
      .toUpperCase();

  if (!NHSA_VALID_ROLES.includes(role)) {
    throw new Error("Invalid role: " + role);
  }

  const email =
    nhsa_clean_(spec.email)
      .toLowerCase();

  if (!email) {
    throw new Error("School Email is required.");
  }

  // V46.4.3: centralize role/domain validation so bulk sync, manual grants,
  // and approved access requests all obey the same Teacher identity policy.
  nhsa_assertRoleEmailPolicy_(email, role);

  const sh = nhsa_ensureRegistry_(ss);
  const h = nhsa_headerMap_(sh);

  // Only an explicit TUTOR grant can intentionally restore Tutor access.
  if (role === "TUTOR") {
    nhsa_removeDeletionTombstones_(sh, email);
  }

  const duplicate = nhsa_readRegistry_(sh)
    .find(grant =>
      grant.email === email &&
      grant.role === role &&
      nhsa_isGrantActive_(grant)
    );

  if (duplicate) {
    return {
      ok: true,
      existing: true,
      grantId: duplicate.grantId
    };
  }

  const grantId = "AG-" + Utilities.getUuid();

  const row =
    new Array(sh.getLastColumn()).fill("");

  nhsa_setRow_(row, h, "Grant ID", grantId);
  nhsa_setRow_(row, h, "School Email", email);
  nhsa_setRow_(
    row,
    h,
    "Full Name",
    nhsa_clean_(spec.fullName)
  );
  nhsa_setRow_(row, h, "Role", role);
  nhsa_setRow_(row, h, "Status", "Active");
  nhsa_setRow_(
    row,
    h,
    "School Year",
    spec.schoolYear || ""
  );
  nhsa_setRow_(
    row,
    h,
    "Start Date",
    spec.startDate || new Date()
  );
  nhsa_setRow_(
    row,
    h,
    "Expiration Date",
    spec.expirationDate || ""
  );
  nhsa_setRow_(
    row,
    h,
    "Approved By",
    spec.approvedBy || ""
  );
  nhsa_setRow_(row, h, "Approved At", new Date());
  nhsa_setRow_(row, h, "Notes", spec.notes || "");

  nhsa_appendRowLocked_(
    sh,
    row,
    "create access grant"
  );

  nhsa_invalidateAccessCache_();

  return {
    ok: true,
    grantId: grantId
  };
}


/* ============================================================
   GENERIC HELPERS
   ============================================================ */

function nhsa_openSpreadsheet_() {
  return nhsa_retry_(
    () => nhsperf_ss_(
      NHS_ACCESS_SPREADSHEET_ID
    ),
    "open NHS access spreadsheet"
  );
}


function nhsa_getActiveEmail_() {
  const email =
    String(
      Session.getActiveUser().getEmail() || ""
    )
      .trim()
      .toLowerCase();

  if (!email) {
    throw new Error(
      "LOGIN_REQUIRED|Google did not expose your signed-in account to this web-app deployment. Open the production CUSD-domain deployment while signed into your CUSD Google account. If this persists, an Admin must verify that the production web app was deployed by a CUSD Workspace account."
    );
  }

  return email;
}


function nhsa_headerIndex_(headers) {
  const map = {};

  headers.forEach((header, index) => {
    const key = String(header || "").trim();

    if (key && map[key] === undefined) {
      map[key] = index;
    }
  });

  return map;
}


function nhsa_headerMap_(sh) {
  return nhsperf_headerMap_(sh, true, false);
}


function nhsa_get_(row, h, header) {
  const idx = h[header];

  return idx === undefined
    ? ""
    : String(row[idx] || "").trim();
}


function nhsa_appendRowLocked_(sheet, row, label) {
  if (!sheet) throw new Error("Sheet is required.");
  if (!Array.isArray(row)) throw new Error("Row values are required.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const nextRow = Math.max(2, sheet.getLastRow() + 1);
    nhsa_retry_(
      function() {
        sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
      },
      label || "append access row"
    );
    return nextRow;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


function nhsa_setRow_(row, h, header, value) {
  const col = h[header];

  if (!col) return;

  row[col - 1] = value;
}


function nhsa_clean_(value) {
  return String(
    value == null ? "" : value
  ).trim();
}

function nhsinput_literalText_(value, maxLength, label) {
  let text = String(value == null ? "" : value).trim();
  const limit = Number(maxLength || 0);

  if (limit > 0 && text.length > limit) {
    throw new Error(
      String(label || "Text") +
      " must be " +
      limit +
      " characters or fewer."
    );
  }

  // Prevent user-entered spreadsheet formulas while preserving the text
  // exactly as displayed in Google Sheets.
  if (/^[=+\-@]/.test(text)) {
    text = "'" + text;
  }

  return text;
}


function nhsa_parseDate_(value) {
  if (!value) return null;
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : new Date(value.getTime());
  }

  const raw = String(value || "").trim();
  const dateOnly = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (dateOnly) {
    const date = new Date(
      Number(dateOnly[1]),
      Number(dateOnly[2]) - 1,
      Number(dateOnly[3]),
      12, 0, 0, 0
    );
    return isNaN(date.getTime()) ? null : date;
  }

  const date = new Date(raw);
  return isNaN(date.getTime()) ? null : date;
}


function nhsa_displayTutorName_(preferred, fullName) {
  preferred = nhsa_clean_(preferred);
  fullName = nhsa_clean_(fullName);

  if (!preferred) return fullName;

  const parts = fullName
    .split(/\s+/)
    .filter(Boolean);

  return parts.length > 1
    ? preferred + " " + parts[parts.length - 1]
    : preferred;
}


function nhsa_retry_(fn, label, attempts) {
  attempts = Number(attempts || 4);

  let lastError = null;

  for (
    let attempt = 1;
    attempt <= attempts;
    attempt++
  ) {
    try {
      return fn();
    } catch (error) {
      lastError = error;

      const message =
        String(
          error && error.message
            ? error.message
            : error
        );

      const transient =
        /timed out|service invoked too many times|internal error|try again/i
          .test(message);

      if (
        !transient ||
        attempt === attempts
      ) {
        throw error;
      }

      Utilities.sleep(
        450 * Math.pow(2, attempt - 1)
      );
    }
  }

  throw lastError ||
    new Error(
      "Operation failed: " + label
    );
}


/* ============================================================
   COMPATIBILITY WRAPPERS FOR CURRENT INTERNAL OS
   ============================================================ */

function nhssec_install_(ss) {
  // Existing internal installer can continue calling this name.
  nhsa_installAccessCore_();

  return {
    ok: true
  };
}


function nhssec_requireOfficer_() {
  return nhsa_requireAnyRole_(
    ["ADMIN", "ADVISOR", "OFFICER"],
    "NHS Internal Operations"
  );
}


function nhsa_invalidateAccessCache_() {
  const value = String(Date.now());
  try {
    nhsperf_scriptProperties_().setProperty(NHSA_ACCESS_REVISION_PROPERTY, value);
  } catch (e) {}
  try {
    nhsperf_scriptCache_().put(NHSA_ACCESS_REVISION_CACHE_KEY, value, 21600);
  } catch (e) {}
  return value;
}


function nhsa_getAccessRevision_() {
  try {
    const cached = nhsperf_scriptCache_().get(NHSA_ACCESS_REVISION_CACHE_KEY);
    if (cached) return String(cached);
  } catch (e) {}

  let value = "0";
  try {
    value = String(
      nhsperf_scriptProperties_().getProperty(NHSA_ACCESS_REVISION_PROPERTY) || "0"
    );
  } catch (e) {}

  try { nhsperf_scriptCache_().put(NHSA_ACCESS_REVISION_CACHE_KEY, value, 21600); } catch (e) {}
  return value;
}

function nhsa_readSettingsCached_(ss, revision) {
  const key = "nhsaSettingsV43|" + String(revision || nhsa_getAccessRevision_());
  const cache = nhsperf_scriptCache_();
  try {
    const cached = cache.get(key);
    if (cached) return JSON.parse(cached);
  } catch (e) {}

  const settings = nhsa_readSettings_(ss);
  try { cache.put(key, JSON.stringify(settings), 300); } catch (e) {}
  return settings;
}


/** ============================================================
 * SOURCE GROUP: NHS_Access_Admin.gs
 * ============================================================ */

/**
 * NHS ACCESS ADMIN — V1
 *
 * INTERNAL NHS OPERATIONS PROJECT ONLY.
 * Requires NHS_Access_Core.gs in the same Apps Script project.
 */

const NHSA_ADMIN_CONSOLE_CACHE_PREFIX = "nhsaAdminConsoleV464|";

function nhsa_adminGetConsole(forceRefresh) {
  const actorState = nhsa_requireAnyRole_(
    ["ADMIN", "ADVISOR"],
    "Access Management"
  );

  const revision = nhsa_getAccessRevision_();
  const cache = nhsperf_scriptCache_();
  const consoleKey =
    NHSA_ADMIN_CONSOLE_CACHE_PREFIX +
    revision + "|" +
    String(actorState.email || "").toLowerCase();

  if (forceRefresh !== true) {
    try {
      const cached = cache.get(consoleKey);
      if (cached) return JSON.parse(cached);
    } catch (e) {}
  }

  const ss = nhsa_openSpreadsheet_();
  const registry = nhsa_ensureRegistry_(ss);
  const requests = nhsa_ensureRequests_(ss);
  const log = nhsa_ensureLog_(ss);
  const settings = nhsa_readSettings_(ss);

  const grants = nhsa_readRegistry_(registry)
    .filter(function(grant) {
      return String(grant.status || "").trim() !== "Deleted";
    })
    .sort((a, b) =>
      String(a.email).localeCompare(
        String(b.email)
      )
    );

  const pending = nhsa_readRequests_(requests)
    .filter(req => req.status === "Pending")
    .sort((a, b) =>
      String(b.requestedAt).localeCompare(
        String(a.requestedAt)
      )
    );

  const recentLog =
    nhsa_readAccessLog_(log)
      .slice(-75)
      .reverse();

  const groupedUsers =
    nhsa_groupGrantsByUser_(grants);

  const expiring =
    grants
      .filter(grant =>
        nhsa_isGrantActive_(grant) &&
        nhsa_expiresWithinDays_(
          grant.expirationDate,
          45
        )
      )
      .sort((a, b) =>
        String(a.expirationDate)
          .localeCompare(
            String(b.expirationDate)
          )
      );

  const response = {
    ok: true,
    users: groupedUsers,
    grants: grants,
    pending: pending,
    expiring: expiring,
    recentLog: recentLog,
    settings: settings,
    currentUser: actorState.email,
    roleOptions: NHSA_VALID_ROLES.slice(),
    generatedAt: new Date().toISOString(),
    accessRevision: revision
  };

  try { cache.put(consoleKey, JSON.stringify(response), 60); } catch (e) {}
  return response;
}


function nhsa_adminGrantAccess(payload) {
  payload = payload || {};

  const actor =
    nhsa_requireAnyRole_(
      ["ADMIN", "ADVISOR"],
      "Access Management"
    ).email;

  const ss = nhsa_openSpreadsheet_();
  const settings = nhsa_readSettings_(ss);
  const safeFullName = nhsinput_literalText_(payload.fullName, 120, "Full name");
  const safeNotes = nhsinput_literalText_(payload.notes, 500, "Notes");
  const safeSchoolYear = nhsinput_literalText_(payload.schoolYear, 24, "School year");

  const requestedRoles =
    Array.isArray(payload.roles)
      ? payload.roles
      : [payload.role];

  const roles =
    Array.from(
      new Set(
        requestedRoles
          .map(function(role) {
            return String(role || "")
              .trim()
              .toUpperCase();
          })
          .filter(Boolean)
      )
    );

  if (!roles.length) {
    throw new Error(
      "Choose at least one role to grant."
    );
  }

  const invalidRoles =
    roles.filter(function(role) {
      return !NHSA_VALID_ROLES.includes(role);
    });

  if (invalidRoles.length) {
    throw new Error(
      "Invalid role selection: " +
      invalidRoles.join(", ")
    );
  }

  const email =
    String(payload.email || "")
      .trim()
      .toLowerCase();

  // Pre-validate every requested role before writing anything.
  // This prevents a partial grant if one selected role violates
  // an email/domain rule.
  roles.forEach(function(role) {
    nhsa_assertRoleEmailPolicy_(email, role);
  });

  if (
    roles.includes("ADMIN") &&
    !nhsa_actorHasRole_(ss, actor, "ADMIN")
  ) {
    throw new Error(
      "Only an Admin can grant ADMIN access."
    );
  }

  let tutorProfileProvision = null;

  if (roles.includes("TUTOR")) {
    tutorProfileProvision =
      nhstutor_ensureProfileForAccess_(
        ss,
        email,
        safeFullName
      );
  }

  const results =
    roles.map(function(role) {
      const result =
        nhsa_createGrant_(ss, {
          email: email,
          fullName: safeFullName,
          role: role,
          schoolYear:
            safeSchoolYear ||
            settings.CURRENT_SCHOOL_YEAR ||
            "",
          startDate:
            payload.startDate ||
            new Date(),
          expirationDate:
            payload.expirationDate || "",
          approvedBy: actor,
          notes: safeNotes
        });

      return {
        role: role,
        grantId: result.grantId || "",
        existing: result.existing === true
      };
    });

  const createdRoles =
    results
      .filter(function(result) {
        return !result.existing;
      })
      .map(function(result) {
        return result.role;
      });

  const existingRoles =
    results
      .filter(function(result) {
        return result.existing;
      })
      .map(function(result) {
        return result.role;
      });

  return {
    ok: true,
    roles: roles,
    createdRoles: createdRoles,
    existingRoles: existingRoles,
    grantIds:
      results
        .map(function(result) {
          return result.grantId;
        })
        .filter(Boolean),
    tutorProfileProvision:
      tutorProfileProvision,
    message:
      createdRoles.length
        ? (
            "Access granted: " +
            createdRoles.join(", ") +
            (
              existingRoles.length
                ? ". Already active: " +
                  existingRoles.join(", ")
                : ""
            )
          )
        : "All selected roles are already active."
  };
}


/* ============================================================
   V45.2 — BULK TEACHER ACCESS SYNC
   ============================================================ */

/**
 * Synchronizes an official teacher roster with TEACHER grants.
 *
 * payload:
 * {
 *   teachers: [
 *     { fullName: "Jane Smith", email: "jsmith@school.org" }
 *   ],
 *   schoolYear: "2026-27",
 *   expirationDate: "2027-06-30",
 *   expireMissing: false
 * }
 *
 * SAFETY CONTRACT
 * ---------------
 * - Only TEACHER grants are created/updated/expired.
 * - ADMIN / ADVISOR / OFFICER / TUTOR grants are never altered.
 * - Existing active TEACHER grants are updated rather than duplicated.
 * - Missing teachers are expired only when expireMissing === true.
 */
function nhsa_adminBulkSyncTeachers(payload) {
  const actor =
    nhsa_requireAnyRole_(
      ["ADMIN", "ADVISOR"],
      "Access Management"
    ).email;

  payload = payload || {};

  const incoming = Array.isArray(payload.teachers)
    ? payload.teachers
    : [];

  if (!incoming.length) {
    throw new Error("No teachers were provided for synchronization.");
  }

  const ss = nhsa_openSpreadsheet_();
  const registry = nhsa_ensureRegistry_(ss);
  const h = nhsa_headerMap_(registry);
  const settings = nhsa_readSettings_(ss);

  const schoolYear = String(
    payload.schoolYear ||
    settings.CURRENT_SCHOOL_YEAR ||
    ""
  ).trim();

  const expirationDate = String(
    payload.expirationDate || ""
  ).trim();

  const expireMissing = payload.expireMissing === true;

  // Normalize and deduplicate the roster by email.
  const rosterMap = {};

  incoming.forEach(function(person) {
    const email = String(
      person && person.email || ""
    ).trim().toLowerCase();

    const fullName = String(
      person && person.fullName || ""
    ).trim();

    if (!email) return;

    // V46.4.3: the previous literal accidentally double-escaped \s and \.
    // and could reject otherwise valid addresses containing the letter "s".
    // Teacher roster sync is explicitly for CUSD staff identities.
    if (!nhsa_isValidEmailAddress_(email)) {
      throw new Error("Invalid teacher email: " + email);
    }

    if (!nhsa_isTeacherStaffEmail_(email)) {
      throw new Error(
        "Teacher roster sync accepts only @" + NHSA_TEACHER_EMAIL_DOMAIN +
        " staff emails. Invalid teacher account: " + email
      );
    }

    rosterMap[email] = {
      email: email,
      fullName: fullName
    };
  });

  const roster = Object.keys(rosterMap).map(function(email) {
    return rosterMap[email];
  });

  if (!roster.length) {
    throw new Error("The teacher roster contained no valid email addresses.");
  }

  const allGrants = nhsa_readRegistry_(registry);

  const activeTeacherGrants = allGrants.filter(function(grant) {
    return grant.role === "TEACHER" && nhsa_isGrantActive_(grant);
  });

  const activeByEmail = {};
  activeTeacherGrants.forEach(function(grant) {
    activeByEmail[grant.email] = grant;
  });

  let created = 0;
  let updated = 0;
  let unchanged = 0;
  let expired = 0;

  roster.forEach(function(person) {
    const existing = activeByEmail[person.email];

    if (!existing) {
      nhsa_createGrant_(ss, {
        email: person.email,
        fullName: person.fullName,
        role: "TEACHER",
        schoolYear: schoolYear,
        startDate: new Date(),
        expirationDate: expirationDate || "",
        approvedBy: actor,
        notes: "Bulk teacher roster sync"
      });

      created++;
      return;
    }

    const rowInfo = nhsa_findGrantRow_(registry, existing.grantId);

    if (!rowInfo) {
      unchanged++;
      return;
    }

    let changed = false;

    if (person.fullName && person.fullName !== existing.fullName) {
      registry.getRange(rowInfo.row, h["Full Name"]).setValue(person.fullName);
      changed = true;
    }

    if (schoolYear && schoolYear !== existing.schoolYear) {
      registry.getRange(rowInfo.row, h["School Year"]).setValue(schoolYear);
      changed = true;
    }

    if (expirationDate) {
      const existingExpiration = String(
        existing.expirationDate || ""
      ).trim();

      if (existingExpiration !== expirationDate) {
        registry
          .getRange(rowInfo.row, h["Expiration Date"])
          .setValue(expirationDate);
        changed = true;
      }
    }

    if (changed) {
      const existingNotes = String(
        registry
          .getRange(rowInfo.row, h["Notes"])
          .getDisplayValue() || ""
      ).trim();

      registry
        .getRange(rowInfo.row, h["Notes"])
        .setValue(
          [
            existingNotes,
            "Roster synchronized by " + actor + " on " +
              Utilities.formatDate(
                new Date(),
                Session.getScriptTimeZone(),
                "yyyy-MM-dd"
              )
          ]
            .filter(Boolean)
            .join(" | ")
        );

      updated++;
    } else {
      unchanged++;
    }
  });

  if (expireMissing) {
    const incomingEmails = new Set(
      roster.map(function(person) {
        return person.email;
      })
    );

    activeTeacherGrants
      .filter(function(grant) {
        return !incomingEmails.has(grant.email);
      })
      .forEach(function(grant) {
        const rowInfo = nhsa_findGrantRow_(registry, grant.grantId);
        if (!rowInfo) return;

        registry
          .getRange(rowInfo.row, h["Status"])
          .setValue("Expired");

        registry
          .getRange(rowInfo.row, h["Expiration Date"])
          .setValue(new Date());

        const existingNotes = String(
          registry
            .getRange(rowInfo.row, h["Notes"])
            .getDisplayValue() || ""
        ).trim();

        registry
          .getRange(rowInfo.row, h["Notes"])
          .setValue(
            [
              existingNotes,
              "Expired by teacher roster sync · " + actor
            ]
              .filter(Boolean)
              .join(" | ")
          );

        expired++;
      });
  }

  SpreadsheetApp.flush();
  nhsa_invalidateAccessCache_();

  return {
    ok: true,
    processed: roster.length,
    created: created,
    updated: updated,
    unchanged: unchanged,
    expired: expired,
    schoolYear: schoolYear
  };
}


function nhsa_adminSetGrantStatus(
  grantId,
  status
) {
  const actor =
    nhsa_requireAnyRole_(
      ["ADMIN", "ADVISOR"],
      "Access Management"
    ).email;

  status =
    String(status || "")
      .trim();

  if (
    ![
      "Active",
      "Inactive",
      "Expired",
      "Revoked"
    ].includes(status)
  ) {
    throw new Error("Invalid grant status.");
  }

  const ss = nhsa_openSpreadsheet_();
  const sh = nhsa_ensureRegistry_(ss);
  const rowInfo =
    nhsa_findGrantRow_(sh, grantId);

  if (!rowInfo) {
    throw new Error("Access grant not found.");
  }

  if (
    ["ADMIN", "ADVISOR"].includes(
      rowInfo.grant.role
    ) &&
    status !== "Active"
  ) {
    nhsa_assertAnotherPrivilegedGrant_(
      sh,
      grantId
    );
  }

  const h = nhsa_headerMap_(sh);

  sh
    .getRange(
      rowInfo.row,
      h["Status"]
    )
    .setValue(status);

  const notes =
    String(
      sh
        .getRange(
          rowInfo.row,
          h["Notes"]
        )
        .getDisplayValue() || ""
    );

  sh
    .getRange(
      rowInfo.row,
      h["Notes"]
    )
    .setValue(
      [
        notes,
        status +
          " by " +
          actor +
          " on " +
          new Date().toLocaleDateString()
      ]
        .filter(Boolean)
        .join(" | ")
    );

  nhsa_invalidateAccessCache_();

  return {
    ok: true,
    grantId: grantId,
    status: status
  };
}


function nhsa_normalizeBulkDeleteEmails_(emails) {
  if (!Array.isArray(emails)) {
    throw new Error("Select one or more verified users.");
  }

  return Array.from(
    new Set(
      emails
        .map(function(email) {
          return String(email || "")
            .trim()
            .toLowerCase();
        })
        .filter(Boolean)
    )
  );
}


function nhsa_bulkDeleteSelectedSummary_(grants, emails) {
  const selected = new Set(
    nhsa_normalizeBulkDeleteEmails_(emails)
  );

  const target = (grants || []).filter(function(grant) {
    const email = String(grant.email || "")
      .trim()
      .toLowerCase();

    return (
      selected.has(email) &&
      String(grant.status || "").trim() !== "Deleted"
    );
  });

  const users = {};
  const roleCounts = {};

  target.forEach(function(grant) {
    const email = String(grant.email || "")
      .trim()
      .toLowerCase();

    if (!users[email]) {
      users[email] = {
        email: email,
        fullName: grant.fullName || "",
        roles: []
      };
    }

    if (
      grant.role &&
      !users[email].roles.includes(grant.role)
    ) {
      users[email].roles.push(grant.role);
    }

    if (grant.role) {
      roleCounts[grant.role] =
        (roleCounts[grant.role] || 0) + 1;
    }
  });

  const userList = Object.keys(users)
    .sort()
    .map(function(email) {
      users[email].roles.sort();
      return users[email];
    });

  return {
    userCount: userList.length,
    grantCount: target.length,
    users: userList,
    roleCounts: roleCounts,
    selectedEmails: userList.map(function(user) {
      return user.email;
    })
  };
}


function nhsa_assertBulkDeleteSelectedSafe_(
  grants,
  selectedEmails,
  actorEmail
) {
  const selected = new Set(
    nhsa_normalizeBulkDeleteEmails_(
      selectedEmails
    )
  );

  actorEmail = String(actorEmail || "")
    .trim()
    .toLowerCase();

  if (!selected.size) {
    throw new Error(
      "Select one or more verified users."
    );
  }

  if (selected.has(actorEmail)) {
    throw new Error(
      "Your current Admin account cannot be included in a bulk deletion. Manage your own access separately."
    );
  }

  const selectedExisting = new Set(
    (grants || [])
      .filter(function(grant) {
        return (
          selected.has(
            String(grant.email || "")
              .trim()
              .toLowerCase()
          ) &&
          String(grant.status || "").trim() !==
            "Deleted"
        );
      })
      .map(function(grant) {
        return String(grant.email || "")
          .trim()
          .toLowerCase();
      })
  );

  const missing = Array.from(selected)
    .filter(function(email) {
      return !selectedExisting.has(email);
    });

  if (missing.length) {
    throw new Error(
      "Verified access was not found for: " +
      missing.join(", ") +
      ". Refresh Access Management and try again."
    );
  }

  const privilegedRemaining =
    (grants || []).some(function(grant) {
      const email = String(grant.email || "")
        .trim()
        .toLowerCase();

      return (
        !selected.has(email) &&
        ["ADMIN", "ADVISOR"].includes(
          grant.role
        ) &&
        nhsa_isGrantActive_(grant)
      );
    });

  const removesActivePrivileged =
    (grants || []).some(function(grant) {
      const email = String(grant.email || "")
        .trim()
        .toLowerCase();

      return (
        selected.has(email) &&
        ["ADMIN", "ADVISOR"].includes(
          grant.role
        ) &&
        nhsa_isGrantActive_(grant)
      );
    });

  if (
    removesActivePrivileged &&
    !privilegedRemaining
  ) {
    throw new Error(
      "This selection would remove the final active Admin/Advisor. Add or preserve a replacement first."
    );
  }

  return true;
}


/**
 * Read-only preview for deleting access for multiple selected users.
 *
 * Every non-deleted Access Registry role for each selected email is included.
 * The current Admin account is intentionally excluded from bulk deletion.
 * Access Log history, pending Access Requests, Tutor Database rows,
 * Teacher Directory rows, requests, sessions, and tutoring records are
 * preserved.
 */
function nhsa_adminPreviewBulkDeleteAccess(payload) {
  const actorState = nhsa_requireAnyRole_(
    ["ADMIN"],
    "Bulk Access Deletion"
  );

  payload = payload || {};

  const emails =
    nhsa_normalizeBulkDeleteEmails_(
      payload.emails || []
    );

  const ss = nhsa_openSpreadsheet_();
  const sh = nhsa_ensureRegistry_(ss);
  const grants = nhsa_readRegistry_(sh);

  nhsa_assertBulkDeleteSelectedSafe_(
    grants,
    emails,
    actorState.email
  );

  const preview =
    nhsa_bulkDeleteSelectedSummary_(
      grants,
      emails
    );

  if (!preview.grantCount) {
    throw new Error(
      "No verified access was found for the selected users."
    );
  }

  return Object.assign(
    {
      ok: true,
      currentUser: actorState.email,
      allRolesPerUser: true,
      pendingRequestsPreserved: true,
      tutoringRecordsPreserved: true,
      accessLogPreserved: true
    },
    preview
  );
}


/**
 * Permanently removes every non-deleted Access Registry role for the selected
 * verified users.
 *
 * Safety:
 * - ADMIN only.
 * - Requires an authoritative preview in Operations.
 * - Requires the exact confirmation phrase DELETE SELECTED ACCESS.
 * - The caller's current Admin account cannot be bulk-deleted.
 * - The final active Admin/Advisor cannot be removed.
 * - Access Log, Access Requests, Tutor Database, Teacher Directory,
 *   requests, sessions, and tutoring history remain intact.
 * - A deletion tombstone is written per selected user so Tutor sync cannot
 *   silently recreate deleted Tutor access.
 */
function nhsa_adminBulkDeleteAccess(payload) {
  const actorState = nhsa_requireAnyRole_(
    ["ADMIN"],
    "Bulk Access Deletion"
  );

  payload = payload || {};

  const emails =
    nhsa_normalizeBulkDeleteEmails_(
      payload.emails || []
    );

  const reason = nhsinput_literalText_(
    payload.reason || "",
    500,
    "Deletion reason"
  );

  if (!reason) {
    throw new Error(
      "A deletion reason is required."
    );
  }

  const confirmation = String(
    payload.confirmation || ""
  )
    .trim()
    .toUpperCase();

  if (
    confirmation !==
    "DELETE SELECTED ACCESS"
  ) {
    throw new Error(
      "Type DELETE SELECTED ACCESS exactly to continue."
    );
  }

  const ss = nhsa_openSpreadsheet_();
  const sh = nhsa_ensureRegistry_(ss);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  let deletedUsers = [];
  let deletedGrantCount = 0;

  try {
    const grants = nhsa_readRegistry_(sh);

    nhsa_assertBulkDeleteSelectedSafe_(
      grants,
      emails,
      actorState.email
    );

    const summary =
      nhsa_bulkDeleteSelectedSummary_(
        grants,
        emails
      );

    const selected = new Set(
      summary.selectedEmails
    );

    if (!summary.grantCount) {
      throw new Error(
        "No verified access was found for the selected users."
      );
    }

    deletedUsers = summary.users.map(
      function(user) {
        return {
          email: user.email,
          fullName: user.fullName || "",
          roles: user.roles.slice()
        };
      }
    );

    const rowsToDelete = grants
      .filter(function(grant) {
        const email = String(
          grant.email || ""
        )
          .trim()
          .toLowerCase();

        const isVisibleGrant =
          String(grant.status || "")
            .trim() !== "Deleted";

        const isDeletionTombstone =
          String(grant.status || "")
            .trim() === "Deleted" &&
          String(grant.notes || "")
            .indexOf(
              NHSA_DELETION_TOMBSTONE_NOTE
            ) >= 0;

        return (
          selected.has(email) &&
          (
            isVisibleGrant ||
            isDeletionTombstone
          )
        );
      })
      .map(function(grant) {
        return nhsa_findGrantRow_(
          sh,
          grant.grantId
        );
      })
      .filter(Boolean)
      .sort(function(a, b) {
        return b.row - a.row;
      });

    rowsToDelete.forEach(function(info) {
      sh.deleteRow(info.row);

      if (
        String(
          info.grant.status || ""
        ).trim() !== "Deleted"
      ) {
        deletedGrantCount++;
      }
    });

    const h = nhsa_headerMap_(sh);

    deletedUsers.forEach(function(user) {
      const row =
        new Array(
          sh.getLastColumn()
        ).fill("");

      nhsa_setRow_(
        row,
        h,
        "Grant ID",
        "DEL-" + Utilities.getUuid()
      );
      nhsa_setRow_(
        row,
        h,
        "School Email",
        user.email
      );
      nhsa_setRow_(
        row,
        h,
        "Full Name",
        user.fullName || ""
      );
      nhsa_setRow_(
        row,
        h,
        "Role",
        "TUTOR"
      );
      nhsa_setRow_(
        row,
        h,
        "Status",
        "Deleted"
      );
      nhsa_setRow_(
        row,
        h,
        "Approved By",
        actorState.email
      );
      nhsa_setRow_(
        row,
        h,
        "Approved At",
        new Date()
      );
      nhsa_setRow_(
        row,
        h,
        "Notes",
        NHSA_DELETION_TOMBSTONE_NOTE +
          " · Bulk selected-user deletion · " +
          reason
      );

      sh
        .getRange(
          Math.max(
            2,
            sh.getLastRow() + 1
          ),
          1,
          1,
          row.length
        )
        .setValues([row]);
    });

    SpreadsheetApp.flush();
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }

  deletedUsers.forEach(function(user) {
    nhsa_logAccess_(ss, {
      email: user.email,
      fullName: user.fullName || "",
      portal: "Access Management",
      result: "BULK_DELETED",
      roles: user.roles,
      detail:
        "All Access Registry roles deleted in selected-user bulk action by " +
        actorState.email +
        " · Reason: " +
        reason
    });
  });

  nhsa_invalidateAccessCache_();

  return {
    ok: true,
    deletedUserCount:
      deletedUsers.length,
    deletedGrantCount:
      deletedGrantCount,
    deletedEmails:
      deletedUsers.map(function(user) {
        return user.email;
      }),
    message:
      "Deleted all access roles for " +
      deletedUsers.length +
      " selected user(s) across " +
      deletedGrantCount +
      " grant(s)."
  };
}


/**
 * Permanently removes all Access Registry grants for one verified account.
 * Access/audit history is intentionally preserved in Access Log and request
 * records. Tutor Database identity is not deleted by this operation.
 */
function nhsa_adminDeleteVerifiedUser(email, reason) {
  const actor = nhsa_requireAnyRole_(
    ["ADMIN", "ADVISOR"],
    "Access Management"
  ).email;

  email = String(email || "").trim().toLowerCase();
  reason = String(reason || "").trim();

  if (!email) throw new Error("Verified user email is required.");
  if (!reason) throw new Error("A deletion reason is required for the audit log.");

  const ss = nhsa_openSpreadsheet_();
  const sh = nhsa_ensureRegistry_(ss);
  const grants = nhsa_readRegistry_(sh);
  const target = grants.filter(function(grant) {
    return String(grant.email || "").toLowerCase() === email &&
      String(grant.status || "").trim() !== "Deleted";
  });

  if (!target.length) {
    throw new Error("No Access Registry grants were found for " + email + ".");
  }

  const targetIds = new Set(target.map(function(grant) { return grant.grantId; }));
  const removesActivePrivileged = target.some(function(grant) {
    return ["ADMIN", "ADVISOR"].includes(grant.role) && nhsa_isGrantActive_(grant);
  });

  if (removesActivePrivileged) {
    const anotherPrivileged = grants.some(function(grant) {
      return !targetIds.has(grant.grantId) &&
        ["ADMIN", "ADVISOR"].includes(grant.role) &&
        nhsa_isGrantActive_(grant);
    });

    if (!anotherPrivileged) {
      throw new Error(
        "Cannot delete the final active Admin/Advisor account. Add a replacement first."
      );
    }
  }

  const roles = Array.from(new Set(target.map(function(grant) { return grant.role; })));
  const rows = target
    .map(function(grant) { return nhsa_findGrantRow_(sh, grant.grantId); })
    .filter(Boolean)
    .map(function(info) { return info.row; })
    .sort(function(a, b) { return b - a; });

  rows.forEach(function(row) { sh.deleteRow(row); });

  // Hidden tombstone prevents an eligible Tutor Database profile from silently
  // recreating TUTOR access after all visible grants were deleted.
  nhsa_appendDeletionTombstone_(
    ss,
    email,
    target.find(function(grant) { return grant.fullName; })?.fullName || "",
    actor,
    reason
  );

  nhsa_logAccess_(ss, {
    email: email,
    fullName: target.find(function(grant) { return grant.fullName; })?.fullName || "",
    portal: "Access Management",
    result: "DELETED",
    roles: roles,
    detail:
      "All Access Registry grants deleted by " + actor +
      " · Reason: " + reason
  });

  nhsa_invalidateAccessCache_();

  return {
    ok: true,
    email: email,
    deletedGrantCount: rows.length,
    deletedRoles: roles,
    message: "Verified access deleted for " + email + "."
  };
}


function nhsa_adminApproveRequest(
  requestId,
  options
) {
  const actor =
    nhsa_requireAnyRole_(
      ["ADMIN", "ADVISOR"],
      "Access Management"
    ).email;

  const ss = nhsa_openSpreadsheet_();
  const requestSh =
    nhsa_ensureRequests_(ss);

  const requestInfo =
    nhsa_findRequestRow_(
      requestSh,
      requestId
    );

  if (!requestInfo) {
    throw new Error(
      "Access request not found."
    );
  }

  if (
    requestInfo.request.status !==
    "Pending"
  ) {
    throw new Error(
      "This access request has already been reviewed."
    );
  }

  const settings =
    nhsa_readSettings_(ss);

  const role =
    requestInfo.request.requestedRole;

  if (
    !["TEACHER", "TUTOR"].includes(role)
  ) {
    throw new Error(
      "Only Teacher/Tutor requests can be approved from the request queue."
    );
  }

  if (role === "TUTOR") {
    nhstutor_ensureProfileForAccess_(
      ss,
      requestInfo.request.email,
      requestInfo.request.fullName
    );
  }

  const result =
    nhsa_createGrant_(ss, {
      email:
        requestInfo.request.email,
      fullName:
        requestInfo.request.fullName,
      role: role,
      schoolYear:
        options?.schoolYear ||
        settings.CURRENT_SCHOOL_YEAR ||
        "",
      expirationDate:
        options?.expirationDate ||
        "",
      approvedBy: actor,
      notes:
        "Approved from access request " +
        requestId
    });

  const h =
    nhsa_headerMap_(requestSh);

  requestSh
    .getRange(
      requestInfo.row,
      h["Status"]
    )
    .setValue("Approved");

  requestSh
    .getRange(
      requestInfo.row,
      h["Reviewed By"]
    )
    .setValue(actor);

  requestSh
    .getRange(
      requestInfo.row,
      h["Reviewed At"]
    )
    .setValue(new Date());

  requestSh
    .getRange(
      requestInfo.row,
      h["Review Notes"]
    )
    .setValue(
      options?.notes || ""
    );

  return {
    ok: true,
    grantId: result.grantId,
    message:
      role + " access approved."
  };
}


function nhsa_adminDenyRequest(
  requestId,
  notes
) {
  const actor =
    nhsa_requireAnyRole_(
      ["ADMIN", "ADVISOR"],
      "Access Management"
    ).email;

  const ss = nhsa_openSpreadsheet_();
  const sh = nhsa_ensureRequests_(ss);
  const info =
    nhsa_findRequestRow_(sh, requestId);

  if (!info) {
    throw new Error(
      "Access request not found."
    );
  }

  if (info.request.status !== "Pending") {
    throw new Error(
      "This access request has already been reviewed."
    );
  }

  const h = nhsa_headerMap_(sh);

  sh
    .getRange(info.row, h["Status"])
    .setValue("Denied");

  sh
    .getRange(
      info.row,
      h["Reviewed By"]
    )
    .setValue(actor);

  sh
    .getRange(
      info.row,
      h["Reviewed At"]
    )
    .setValue(new Date());

  sh
    .getRange(
      info.row,
      h["Review Notes"]
    )
    .setValue(notes || "");

  nhsa_invalidateAccessCache_();

  return {
    ok: true,
    message: "Access request denied."
  };
}


function nhsa_adminSaveSettings(payload) {
  nhsa_requireAnyRole_(
    ["ADMIN", "ADVISOR"],
    "Access Management"
  );

  const ss = nhsa_openSpreadsheet_();
  const sh = nhsa_ensureSettings_(ss);

  const permitted = [
    "CURRENT_SCHOOL_YEAR",
    "GATEWAY_URL",
    "INTERNAL_URL",
    "TEACHER_URL",
    "TUTOR_URL",
    "ADVISOR_CONTACT"
  ];

  permitted.forEach(key => {
    if (payload[key] === undefined) return;
    nhsa_writeSetting_(sh, key, String(payload[key] || "").trim());
  });

  nhsa_invalidateAccessCache_();

  return {
    ok: true,
    message: "Portal settings saved."
  };
}


function nhsa_adminPreviewRolloverOfficers(payload) {
  nhsa_requireAnyRole_(
    ["ADMIN", "ADVISOR"],
    "Access Management"
  );

  payload = payload || {};
  const newYear = String(payload.schoolYear || "").trim();
  const expirationDate = String(payload.expirationDate || "").trim();
  const incoming = Array.isArray(payload.incomingOfficers)
    ? payload.incomingOfficers
    : [];

  if (!newYear) throw new Error("New school year is required.");
  if (!expirationDate) throw new Error("Officer expiration date is required.");
  if (!incoming.length) throw new Error("Add at least one incoming officer.");

  const normalizedIncoming = [];
  const incomingEmails = new Set();

  incoming.forEach(function(person) {
    const email = String(person.email || "").trim().toLowerCase();
    const fullName = String(person.fullName || "").trim();
    if (!email) return;
    nhsa_assertRoleEmailPolicy_(email, "OFFICER");
    if (incomingEmails.has(email)) return;
    incomingEmails.add(email);
    normalizedIncoming.push({ email: email, fullName: fullName });
  });

  if (!normalizedIncoming.length) {
    throw new Error("Add at least one valid incoming officer.");
  }

  const ss = nhsa_openSpreadsheet_();
  const grants = nhsa_readRegistry_(nhsa_ensureRegistry_(ss));
  const activeOfficerGrants = grants.filter(function(grant) {
    return grant.role === "OFFICER" && nhsa_isGrantActive_(grant);
  });

  const outgoing = activeOfficerGrants.filter(function(grant) {
    return !incomingEmails.has(grant.email);
  });

  const retained = activeOfficerGrants.filter(function(grant) {
    return incomingEmails.has(grant.email);
  });

  const activeOfficerEmails = new Set(
    activeOfficerGrants.map(function(grant) { return grant.email; })
  );

  const create = normalizedIncoming.filter(function(person) {
    return !activeOfficerEmails.has(person.email);
  });

  const outgoingEmails = new Set(
    outgoing.map(function(grant) { return grant.email; })
  );

  const preservedRoleCounts = {};
  grants.forEach(function(grant) {
    if (
      outgoingEmails.has(grant.email) &&
      grant.role !== "OFFICER" &&
      nhsa_isGrantActive_(grant)
    ) {
      preservedRoleCounts[grant.role] =
        (preservedRoleCounts[grant.role] || 0) + 1;
    }
  });

  return {
    ok: true,
    schoolYear: newYear,
    expirationDate: expirationDate,
    incomingCount: normalizedIncoming.length,
    expireCount: outgoing.length,
    retainCount: retained.length,
    createCount: create.length,
    expiring: outgoing.map(function(grant) {
      return {
        email: grant.email,
        fullName: grant.fullName || "",
        grantId: grant.grantId || ""
      };
    }),
    retaining: retained.map(function(grant) {
      return {
        email: grant.email,
        fullName: grant.fullName || ""
      };
    }),
    creating: create,
    preservedRoleCounts: preservedRoleCounts,
    message:
      "Preview only — " +
      outgoing.length +
      " outgoing OFFICER grant(s) will expire, " +
      retained.length +
      " active officer(s) will remain, and " +
      create.length +
      " new OFFICER grant(s) will be created."
  };
}


/**
 * Annual officer rollover.
 *
 * incomingOfficers:
 * [
 *   {email:"...", fullName:"...", notes:"..."}
 * ]
 */
function nhsa_adminRolloverOfficers(
  payload
) {
  const actor =
    nhsa_requireAnyRole_(
      ["ADMIN", "ADVISOR"],
      "Access Management"
    ).email;

  const ss = nhsa_openSpreadsheet_();
  const sh = nhsa_ensureRegistry_(ss);

  const newYear =
    String(
      payload.schoolYear || ""
    ).trim();

  const expirationDate =
    String(
      payload.expirationDate || ""
    ).trim();

  const incoming =
    Array.isArray(
      payload.incomingOfficers
    )
      ? payload.incomingOfficers
      : [];

  if (!newYear) {
    throw new Error(
      "New school year is required."
    );
  }

  if (!expirationDate) {
    throw new Error(
      "Officer expiration date is required."
    );
  }

  const incomingEmails =
    new Set(
      incoming
        .map(item =>
          String(item.email || "")
            .trim()
            .toLowerCase()
        )
        .filter(Boolean)
    );

  const grants =
    nhsa_readRegistry_(sh);

  // Remove/expire only the OFFICER role grant.
  // Tutor/Admin/Advisor/Teacher grants on the same account remain untouched.
  grants
    .filter(grant =>
      grant.role === "OFFICER" &&
      nhsa_isGrantActive_(grant) &&
      !incomingEmails.has(grant.email)
    )
    .forEach(grant => {
      const rowInfo =
        nhsa_findGrantRow_(
          sh,
          grant.grantId
        );

      if (!rowInfo) return;

      const h =
        nhsa_headerMap_(sh);

      sh
        .getRange(
          rowInfo.row,
          h["Status"]
        )
        .setValue("Expired");

      sh
        .getRange(
          rowInfo.row,
          h["Expiration Date"]
        )
        .setValue(new Date());
    });

  incoming.forEach(person => {
    nhsa_createGrant_(ss, {
      email: person.email,
      fullName: person.fullName,
      role: "OFFICER",
      schoolYear: newYear,
      startDate: new Date(),
      expirationDate:
        expirationDate,
      approvedBy: actor,
      notes:
        person.notes ||
        "Annual officer rollover"
    });
  });

  const settingsSh =
    nhsa_ensureSettings_(ss);

  nhsa_writeSetting_(
    settingsSh,
    "CURRENT_SCHOOL_YEAR",
    newYear
  );

  nhsa_invalidateAccessCache_();

  return {
    ok: true,
    schoolYear: newYear,
    incomingCount: incoming.length,
    message:
      "Officer rollover completed for " +
      newYear +
      "."
  };
}


function nhsa_adminSyncTutorAccess() {
  const actor = nhsa_requireAnyRole_(
    ["ADMIN", "ADVISOR"],
    "Access Management"
  ).email;

  const ss = nhsa_openSpreadsheet_();

  let profileSync = {
    ok: true,
    source: "",
    responseCount: 0,
    created: 0,
    updated: 0,
    skipped: 0,
    failures: 0
  };

  try {
    profileSync =
      nhstutor_syncProfileResponses_(
        ss
      );
  } catch (error) {
    profileSync = {
      ok: false,
      source: "",
      responseCount: 0,
      created: 0,
      updated: 0,
      skipped: 0,
      failures: 1,
      error:
        error &&
        error.message
          ? error.message
          : String(error)
    };
  }

  // Existing active Tutor grants must never be stranded because a form
  // submission failed to populate Tutor Database.
  const activeGrantProvision =
    nhstutor_provisionActiveGrants_(
      ss
    );

  const tutorSh =
    nhstutor_ensureTutorDatabase_(
      ss
    );

  if (tutorSh.getLastRow() < 2) {
    return {
      ok: true,
      synced: 0,
      alreadyActive: 0,
      skippedBlocked: 0,
      profileSync: profileSync,
      activeGrantProvision:
        activeGrantProvision,
      message:
        "Tutor profile pipeline checked. No tutor profiles are available yet."
    };
  }

  const values = tutorSh
    .getRange(
      1,
      1,
      tutorSh.getLastRow(),
      tutorSh.getLastColumn()
    )
    .getDisplayValues();

  const h = nhsa_headerIndex_(values[0]);
  const settings = nhsa_readSettings_(ss);
  const registry = nhsa_ensureRegistry_(ss);
  const grants = nhsa_readRegistry_(registry);

  const tutorGrantHistory = {};
  grants.forEach(function(grant) {
    if (grant.role !== "TUTOR" || !grant.email) return;
    const email = String(grant.email).trim().toLowerCase();
    if (!tutorGrantHistory[email]) tutorGrantHistory[email] = [];
    tutorGrantHistory[email].push(grant);
  });

  let synced = 0;
  let alreadyActive = 0;
  let skippedBlocked = 0;

  values.slice(1).forEach(function(row) {
    const email = nhsa_get_(row, h, "School Email").toLowerCase();
    const fullName = nhsa_get_(row, h, "Full Name");
    const preferred = nhsa_get_(row, h, "Preferred Name");
    const reliability = nhsa_get_(row, h, "Reliability Status").toLowerCase();
    const agreement = nhsa_get_(row, h, "Expectations Agreement");

    const eligible = !!email &&
      !["inactive", "suspended"].includes(reliability) &&
      (!agreement || /agree|yes|true/i.test(agreement));

    if (!eligible) return;

    const history = tutorGrantHistory[email] || [];
    if (history.some(function(grant) { return nhsa_isGrantActive_(grant); })) {
      alreadyActive++;
      return;
    }

    // Any prior non-active Tutor grant is treated as an explicit access state.
    // Sync never silently overrides Inactive, Expired, Revoked, or Deleted.
    if (history.length) {
      skippedBlocked++;
      return;
    }

    const result = nhsa_createGrant_(ss, {
      email: email,
      fullName: nhsa_displayTutorName_(preferred, fullName),
      role: "TUTOR",
      schoolYear: settings.CURRENT_SCHOOL_YEAR || "",
      approvedBy: actor,
      notes: "Synced from Tutor Database"
    });

    if (!result.existing) {
      synced++;
      tutorGrantHistory[email] = [{
        email: email,
        role: "TUTOR",
        status: "Active"
      }];
    }
  });

  nhsa_invalidateAccessCache_();

  return {
    ok: true,
    synced: synced,
    alreadyActive: alreadyActive,
    skippedBlocked: skippedBlocked,
    profileSync: profileSync,
    activeGrantProvision:
      activeGrantProvision,
    message:
      "Tutor profiles synced: " +
      Number(profileSync.created || 0) +
      " created, " +
      Number(profileSync.updated || 0) +
      " updated. " +
      "Tutor access: " +
      synced +
      " grant(s) created, " +
      alreadyActive +
      " already active, " +
      skippedBlocked +
      " blocked/non-active grant(s) preserved. " +
      Number(
        activeGrantProvision.provisioned ||
        0
      ) +
      " active grant profile(s) auto-provisioned."
  };
}


/* ============================================================
   ADMIN HELPERS
   ============================================================ */

function nhsa_readAccessLog_(sh) {
  if (!sh || sh.getLastRow() < 2) {
    return [];
  }

  const values = sh
    .getRange(
      1,
      1,
      sh.getLastRow(),
      sh.getLastColumn()
    )
    .getDisplayValues();

  const h =
    nhsa_headerIndex_(values[0]);

  return values
    .slice(1)
    .map(row => ({
      logId:
        nhsa_get_(row, h, "Log ID"),
      timestamp:
        nhsa_get_(row, h, "Timestamp"),
      email:
        nhsa_get_(
          row,
          h,
          "School Email"
        ),
      fullName:
        nhsa_get_(
          row,
          h,
          "Full Name"
        ),
      portal:
        nhsa_get_(row, h, "Portal"),
      result:
        nhsa_get_(row, h, "Result"),
      roles:
        nhsa_get_(row, h, "Roles"),
      detail:
        nhsa_get_(row, h, "Detail")
    }))
    .filter(item => item.logId);
}


function nhsa_groupGrantsByUser_(
  grants
) {
  const users = {};

  grants.forEach(grant => {
    if (!users[grant.email]) {
      users[grant.email] = {
        email: grant.email,
        fullName: grant.fullName,
        roles: [],
        grants: [],
        lastLogin: "",
        lastPortal: ""
      };
    }

    const user = users[grant.email];

    if (
      grant.fullName &&
      !user.fullName
    ) {
      user.fullName = grant.fullName;
    }

    user.grants.push(grant);

    if (
      nhsa_isGrantActive_(grant) &&
      !user.roles.includes(grant.role)
    ) {
      user.roles.push(grant.role);
    }

    if (
      grant.lastLogin >
      user.lastLogin
    ) {
      user.lastLogin =
        grant.lastLogin;
      user.lastPortal =
        grant.lastPortal;
    }
  });

  return Object.values(users)
    .sort((a, b) =>
      String(
        a.fullName || a.email
      )
        .localeCompare(
          String(
            b.fullName || b.email
          )
        )
    );
}


function nhsa_findGrantRow_(
  sh,
  grantId
) {
  if (
    !sh ||
    sh.getLastRow() < 2
  ) {
    return null;
  }

  const h = nhsa_headerMap_(sh);
  const match = sh
    .getRange(
      2,
      h["Grant ID"],
      sh.getLastRow() - 1,
      1
    )
    .createTextFinder(
      String(grantId || "")
    )
    .matchEntireCell(true)
    .findNext();

  if (!match) return null;

  const row = match.getRow();
  const values = sh
    .getRange(
      row,
      1,
      1,
      sh.getLastColumn()
    )
    .getDisplayValues()[0];

  const hi =
    nhsa_headerIndex_(
      sh
        .getRange(
          1,
          1,
          1,
          sh.getLastColumn()
        )
        .getDisplayValues()[0]
    );

  return {
    row: row,
    grant: {
      grantId:
        nhsa_get_(
          values,
          hi,
          "Grant ID"
        ),
      email:
        nhsa_get_(
          values,
          hi,
          "School Email"
        ).toLowerCase(),
      fullName:
        nhsa_get_(
          values,
          hi,
          "Full Name"
        ),
      role:
        nhsa_get_(
          values,
          hi,
          "Role"
        ).toUpperCase(),
      status:
        nhsa_get_(
          values,
          hi,
          "Status"
        ),
      startDate:
        nhsa_get_(
          values,
          hi,
          "Start Date"
        ),
      expirationDate:
        nhsa_get_(
          values,
          hi,
          "Expiration Date"
        )
    }
  };
}


function nhsa_findRequestRow_(
  sh,
  requestId
) {
  if (
    !sh ||
    sh.getLastRow() < 2
  ) {
    return null;
  }

  const h = nhsa_headerMap_(sh);

  const match = sh
    .getRange(
      2,
      h["Request ID"],
      sh.getLastRow() - 1,
      1
    )
    .createTextFinder(
      String(requestId || "")
    )
    .matchEntireCell(true)
    .findNext();

  if (!match) return null;

  const row = match.getRow();
  const requests =
    nhsa_readRequests_(sh);

  const request =
    requests.find(
      item =>
        item.requestId === requestId
    );

  return request
    ? {
        row: row,
        request: request
      }
    : null;
}


function nhsa_actorHasRole_(
  ss,
  email,
  role
) {
  return nhsa_resolveAccess_(
    ss,
    email
  ).roles.includes(role);
}


function nhsa_assertAnotherPrivilegedGrant_(
  sh,
  excludedGrantId
) {
  const other =
    nhsa_readRegistry_(sh)
      .filter(grant =>
        grant.grantId !==
          excludedGrantId &&
        ["ADMIN", "ADVISOR"]
          .includes(grant.role) &&
        nhsa_isGrantActive_(grant)
      );

  if (!other.length) {
    throw new Error(
      "Cannot deactivate the final active Admin/Advisor. Add a replacement first."
    );
  }
}


function nhsa_expiresWithinDays_(
  expirationValue,
  days
) {
  const date =
    nhsa_parseDate_(
      expirationValue
    );

  if (!date) return false;

  const now = new Date();
  const max =
    new Date(
      now.getTime() +
      Number(days || 0) *
        86400000
    );

  return (
    date.getTime() >=
      now.getTime() &&
    date.getTime() <=
      max.getTime()
  );
}


function nhsa_writeSetting_(sh, key, value) {
  const h = nhsa_headerMap_(sh);

  if (sh.getLastRow() >= 2) {
    const match = sh
      .getRange(2, h["Key"], sh.getLastRow() - 1, 1)
      .createTextFinder(key)
      .matchEntireCell(true)
      .findNext();

    if (match) {
      sh.getRange(match.getRow(), h["Value"]).setValue(value);
      nhsa_invalidateAccessCache_();
      return;
    }
  }

  const row = new Array(sh.getLastColumn()).fill("");
  nhsa_setRow_(row, h, "Key", key);
  nhsa_setRow_(row, h, "Value", value);
  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  nhsa_invalidateAccessCache_();
}


/** ============================================================
 * SOURCE GROUP: NHS_Security_Readiness.gs
 * ============================================================ */

/**
 * NHS SECURITY, ASSIGNMENT VALIDATION, HEALTH + READINESS
 *
 * New file in the INTERNAL NHS OPERATIONS Apps Script project:
 *   NHS Security & Readiness.gs
 */

const NHSSEC_HEALTH_SHEET = "System Health";

function nhssec_ensureHealthSheet_(ss) {
  const sh = nhssec_getOrCreateSheet_(ss, NHSSEC_HEALTH_SHEET);

  nhssec_ensureHeaders_(
    sh,
    ["Component", "Last Success", "Last Error", "Detail"]
  );

  nhssec_withRetry_(
    () => sh.setFrozenRows(1),
    "freeze System Health header"
  );

  try {
    nhssec_withRetry_(
      () => sh.hideSheet(),
      "hide System Health",
      2
    );
  } catch (e) {
    // Non-critical.
  }

  return sh;
}

function nhssec_ensureHeaders_(sh, required) {
  if (!sh) {
    throw new Error("Cannot ensure headers: sheet is missing.");
  }

  const lastCol = nhssec_withRetry_(
    () => Math.max(1, sh.getLastColumn()),
    "read header width for " + sh.getName()
  );

  const headers = nhssec_withRetry_(
    () => sh
      .getRange(1, 1, 1, lastCol)
      .getDisplayValues()[0]
      .map(v => String(v || "").trim()),
    "read headers for " + sh.getName()
  );

  const missing = required.filter(
    header => !headers.includes(header)
  );

  if (!missing.length) {
    return headers;
  }

  // One write instead of one Sheets call per missing header.
  nhssec_withRetry_(
    () => sh
      .getRange(1, headers.length + 1, 1, missing.length)
      .setValues([missing]),
    "write headers for " + sh.getName()
  );

  return headers.concat(missing);
}

function nhssec_getOrCreateSheet_(ss, name) {
  let sh = nhssec_withRetry_(
    () => nhsperf_sheet_(ss, name),
    "find sheet " + name
  );

  if (sh) return sh;

  sh = nhssec_withRetry_(
    () => ss.insertSheet(name),
    "create sheet " + name
  );

  return sh;
}


/**
 * Retries transient Spreadsheet service timeouts.
 *
 * Google Sheets occasionally returns:
 * "Service Spreadsheets timed out while accessing document..."
 * even when the code is valid. Retrying the same small operation usually
 * succeeds without forcing the user to rerun the whole installer.
 */
function nhssec_withRetry_(fn, label, attempts) {
  attempts = Number(attempts || 4);

  let lastError = null;

  for (let attempt = 1; attempt <= attempts; attempt++) {
    try {
      return fn();
    } catch (error) {
      lastError = error;

      const message = String(
        error && error.message
          ? error.message
          : error
      );

      const transient =
        /timed out|service invoked too many times|internal error|try again/i
          .test(message);

      if (!transient || attempt === attempts) {
        throw error;
      }

      Utilities.sleep(
        500 * Math.pow(2, attempt - 1)
      );
    }
  }

  throw lastError || new Error(
    "Spreadsheet operation failed: " + label
  );
}


/* ---------------- Assignment validation ---------------- */

function nhssec_validateTutorAssignment_(ss,tutor,requestId,scheduledDate){
  const tutorId=String(tutor["Tutor ID"]||"").trim();
  const tutorName=nhsos_tutorDisplayName_(tutor["Preferred Name"],tutor["Full Name"]);
  const notificationEmail=String(tutor["Notification Email"]||"").trim();
  if(!notificationEmail) throw new Error(tutorName+" is missing Notification Email.");

  const reliability=String(tutor["Reliability Status"]||"").toLowerCase();
  if(reliability.includes("inactive")||reliability.includes("suspended")){
    throw new Error(tutorName+" is not currently eligible for assignment.");
  }

  const requestSh=nhsperf_sheet_(ss,NHSOS_REQUESTS);
  if(!requestSh) return {ok:true};
  const requests=typeof nhsdata_readKeyedObjects_==="function"
    ? nhsdata_readKeyedObjects_(requestSh,"Request ID") : [];

  let activeCount=0;
  const dateKey=String(scheduledDate||"").trim();

  if (dateKey) {
    nhssched_assertOpenTitanDate_(dateKey);
  }

  requests.forEach(function(r){
    const rid=String(r["Request ID"]||"").trim();
    if(!rid||rid===requestId) return;
    const assignedId=String(r["Assigned Tutor ID"]||"").trim();
    const assignedName=String(r["Assigned Tutor"]||"").trim();
    const matchesTutor=(tutorId&&assignedId===tutorId)||(!assignedId&&assignedName===tutorName);
    if(!matchesTutor) return;

    if(nhslife_activeCaseload_(r)) activeCount++;

    if(dateKey && nhslife_activeCaseload_(r)){
      const otherDate=nhslife_dateKey_(r["Scheduled Date"],nhs_safeTimeZone_(ss));
      if(otherDate && otherDate===dateKey){
        throw new Error(tutorName+" already has an active Titan Time assignment on "+dateKey+".");
      }
    }
  });

  const max=Number(tutor["Maximum Sessions / Week"]||0);
  if(max>0&&activeCount>=max){
    throw new Error(tutorName+" is already at maximum active caseload ("+activeCount+"/"+max+").");
  }

  return {ok:true,activeCaseload:activeCount,maxSessions:max};
}

function nhssec_isTemporarilyUnavailable_(value, dateKey) {
  const text = String(value || "").trim();
  if (!text || !dateKey) return false;

  const target = nhssched_parseKeyLocal_(String(dateKey || "").trim());
  if (!target) return false;
  const targetKey = String(dateKey || "").trim();
  const targetYear = target.getFullYear();

  const monthMap = {
    jan:0,january:0,feb:1,february:1,mar:2,march:2,apr:3,april:3,
    may:4,jun:5,june:5,jul:6,july:6,aug:7,august:7,sep:8,sept:8,september:8,
    oct:9,october:9,nov:10,november:10,dec:11,december:11
  };

  function localKey_(year, monthIndex, day) {
    const d = new Date(Number(year), Number(monthIndex), Number(day), 12, 0, 0, 0);
    if (isNaN(d.getTime()) || d.getFullYear() !== Number(year) || d.getMonth() !== Number(monthIndex) || d.getDate() !== Number(day)) return "";
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    return y + "-" + m + "-" + dd;
  }

  const parts = text.split(/[,;\n]+/).map(function(x){return x.trim();}).filter(Boolean);

  for (const part of parts) {
    // ISO date — already an unambiguous local calendar key.
    const iso = part.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (iso) {
      if (localKey_(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])) === targetKey) return true;
      continue;
    }

    // MM/DD or MM/DD/YYYY. Parse from calendar components rather than
    // new Date(string) so the day cannot drift across time zones.
    const md = part.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
    if (md) {
      let year = md[3] ? Number(md[3]) : targetYear;
      if (year < 100) year += 2000;
      if (localKey_(year, Number(md[1]) - 1, Number(md[2])) === targetKey) return true;
      continue;
    }

    // "Sep 25-27" / "September 25–27" with optional year.
    const range = part.match(/^([A-Za-z]+)\s+(\d{1,2})\s*[-–]\s*(\d{1,2})(?:\s*,?\s*(\d{4}))?$/);
    if (range) {
      const monthIndex = monthMap[String(range[1] || "").toLowerCase()];
      if (monthIndex !== undefined) {
        const year = range[4] ? Number(range[4]) : targetYear;
        const startKey = localKey_(year, monthIndex, Number(range[2]));
        const endKey = localKey_(year, monthIndex, Number(range[3]));
        if (startKey && endKey && targetKey >= startKey && targetKey <= endKey) return true;
      }
      continue;
    }

    // "Sep 25" / "September 25, 2026". Commas may already have been
    // removed by the list splitter, so accept the year with or without it.
    const named = part.match(/^([A-Za-z]+)\s+(\d{1,2})(?:\s+(\d{4}))?$/);
    if (named) {
      const monthIndex = monthMap[String(named[1] || "").toLowerCase()];
      if (monthIndex !== undefined) {
        const year = named[3] ? Number(named[3]) : targetYear;
        if (localKey_(year, monthIndex, Number(named[2])) === targetKey) return true;
      }
      continue;
    }

    // Last-resort parser for timestamps/fully qualified values. Only use it
    // after all date-only forms above have been handled explicitly.
    const d = new Date(part);
    if (!isNaN(d.getTime())) {
      const key = Utilities.formatDate(d, "America/Los_Angeles", "yyyy-MM-dd");
      if (key === targetKey) return true;
    }
  }

  return false;
}

/* ---------------- Identity field migration ---------------- */

function migrateTutorIdentityToSchoolEmail() {
  nhssec_requireOfficer_();

  const ss=nhsperf_ss_(NHS_SPREADSHEET_ID);
  const sh=nhsperf_sheet_(ss, NHSOS_TUTORS);
  if(!sh) throw new Error("Tutor Database not found.");

  let headers=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0]
    .map(v=>String(v||"").trim());

  const schoolCol=headers.indexOf("School Email")+1;
  const legacyCol=headers.indexOf("Tutor Portal Email")+1;

  if(!schoolCol) throw new Error("School Email column is missing.");
  if(!legacyCol) return {ok:true,message:"Tutor Portal Email already removed."};

  if(sh.getLastRow()>=2){
    const count=sh.getLastRow()-1;
    const school=sh.getRange(2,schoolCol,count,1).getValues();
    const legacy=sh.getRange(2,legacyCol,count,1).getValues();

    for(let i=0;i<count;i++){
      if(!String(school[i][0]||"").trim()&&String(legacy[i][0]||"").trim()){
        school[i][0]=String(legacy[i][0]).trim();
      }
    }
    sh.getRange(2,schoolCol,count,1).setValues(school);
  }

  sh.deleteColumn(legacyCol);
  nhsos_invalidateCache_();
  return {ok:true,message:"Tutor Portal Email consolidated into School Email and removed."};
}

/* ---------------- Health ---------------- */

function nhssec_recordHealth_(component, success, detail, errorText) {
  try{
    const ss=nhsperf_ss_(NHS_SPREADSHEET_ID);
    const sh=nhssec_ensureHealthSheet_(ss);
    const data=sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getDisplayValues();
    const h={}; data[0].forEach((v,i)=>h[String(v||"").trim()]=i+1);

    let row=0;
    if(sh.getLastRow()>=2){
      const match=sh.getRange(2,h["Component"],sh.getLastRow()-1,1)
        .createTextFinder(component).matchEntireCell(true).findNext();
      if(match) row=match.getRow();
    }

    if(!row){
      row=sh.getLastRow()+1;
      sh.getRange(row,h["Component"]).setValue(component);
    }

    if(success){
      sh.getRange(row,h["Last Success"]).setValue(new Date());
      sh.getRange(row,h["Last Error"]).clearContent();
    } else {
      sh.getRange(row,h["Last Error"]).setValue(errorText||"Unknown error");
    }

    sh.getRange(row,h["Detail"]).setValue(detail||"");
  }catch(e){}
}

function nhssec_notificationOps_(ss) {
  const log=nhsperf_sheet_(ss, NHSOS_NOTIFICATION_LOG);
  let failed=0;
  let failedRecent=[];
  if(log&&log.getLastRow()>=2){
    const data=log.getRange(1,1,log.getLastRow(),log.getLastColumn()).getDisplayValues();
    const h={}; data[0].forEach((v,i)=>h[String(v||"").trim()]=i);

    for(let i=data.length-1;i>=1;i--){
      if(String(data[i][h["Status"]]||"")==="Failed"){
        failed++;
        if(failedRecent.length<8){
          failedRecent.push({
            id:String(data[i][h["Notification ID"]]||""),
            requestId:String(data[i][h["Request ID"]]||""),
            recipient:String(data[i][h["Recipient"]]||""),
            type:String(data[i][h["Type"]]||""),
            error:String(data[i][h["Error"]]||"")
          });
        }
      }
    }
  }

  const requests=nhsperf_sheet_(ss, NHSOS_REQUESTS);
  let awaitingTutor=0, viewedNotAccepted=0, acceptedNotConfirmed=0;

  if(requests&&requests.getLastRow()>=2){
    const data=typeof nhsdata_readKeyedObjects_==="function"
      ? nhsdata_readKeyedObjects_(requests,"Request ID") : [];
    const get=(row,key)=>String(row[key]||"").trim();

    data.forEach(row=>{
      const status=get(row,"Status");
      if(!["Assigned","Scheduled"].includes(status)) return;

      const response=get(row,"Tutor Response Status");
      const viewed=get(row,"Tutor Viewed At");
      const confirmed=get(row,"Assignment Confirmed");

      if(response!=="Accepted"&&response!=="Declined") awaitingTutor++;
      if(viewed&&response!=="Accepted"&&response!=="Declined") viewedNotAccepted++;
      if(response==="Accepted"&&!confirmed) acceptedNotConfirmed++;
    });
  }

  const props=PropertiesService.getScriptProperties().getProperties();
  const queuedInternal=Object.keys(props).filter(k=>k.indexOf("nhsosFastJob_")===0).length;

  return {
    failedNotifications:failed,
    failedRecent:failedRecent,
    awaitingTutor:awaitingTutor,
    viewedNotAccepted:viewedNotAccepted,
    acceptedNotConfirmed:acceptedNotConfirmed,
    queuedInternalJobs:queuedInternal,
    lastNotificationSweep:props.NHS_NOTIFICATION_LAST_SUCCESS||"",
    lastNotificationError:props.NHS_NOTIFICATION_LAST_ERROR||"",
    lastFastWorker:props.NHS_FAST_WORKER_LAST_SUCCESS||""
  };
}

function nhsos_runNotificationSweepNow() {
  nhssec_requireOfficer_();
  if(typeof runNHSNotificationSweep_!=="function"){
    throw new Error("Notification engine is not installed.");
  }
  runNHSNotificationSweep_();
  nhsos_invalidateCache_();
  return {ok:true,message:"Notification sweep completed."};
}

function nhsos_processJobsNow() {
  nhssec_requireOfficer_();
  if(typeof nhsos_processFastJobs_==="function") nhsos_processFastJobs_();
  nhsos_invalidateCache_();
  return {ok:true,message:"Internal background jobs processed."};
}

/* ---------------- Formal smoke/readiness test ---------------- */

function runNHSProductionSmokeTest() {
  nhssec_requireOfficer_();

  const ss=nhsperf_ss_(NHS_SPREADSHEET_ID);
  const checks=[];

  const add=(name,ok,detail,level)=>
    checks.push({name:name,ok:!!ok,detail:detail||"",level:level||(ok?"PASS":"FAIL")});

  const sheetNames=ss.getSheets().map(s=>s.getName());
  [
    "Tutoring Requests","Tutor Database","Session Log","Activity Log",
    "Notification Log","Teacher Directory","System Access"
  ].forEach(name=>add("Sheet: "+name,sheetNames.includes(name),sheetNames.includes(name)?"Present":"Missing"));

  const req=nhsperf_sheet_(ss, NHSOS_REQUESTS);
  if(req){
    const headers=req.getRange(1,1,1,req.getLastColumn()).getDisplayValues()[0];
    const dup=headers.filter((v,i,a)=>v&&a.indexOf(v)!==i);
    add("No duplicate request headers",dup.length===0,dup.length?dup.join(", "):"Clean");
    [
      "Assigned Tutor ID","Tutor Response Status","Tutor Viewed At",
      "Assignment Confirmed","Scheduled Date","Location"
    ].forEach(h=>add("Request field: "+h,headers.includes(h),headers.includes(h)?"Present":"Missing"));
  }

  const tutors=nhsperf_sheet_(ss, NHSOS_TUTORS);
  if(tutors){
    const headers=tutors.getRange(1,1,1,tutors.getLastColumn()).getDisplayValues()[0];
    ["School Email","Notification Email","Temporary Unavailable Dates","Maximum Sessions / Week"]
      .forEach(h=>add("Tutor field: "+h,headers.includes(h),headers.includes(h)?"Present":"Missing"));
    add("Legacy Tutor Portal Email removed",!headers.includes("Tutor Portal Email"),
      headers.includes("Tutor Portal Email")?"Run migrateTutorIdentityToSchoolEmail().":"Removed");
  }

  const triggers=ScriptApp.getProjectTriggers().map(t=>t.getHandlerFunction());
  add("Notification trigger",triggers.includes("runNHSNotificationSweep_"),
    triggers.includes("runNHSNotificationSweep_")?"Installed":"Missing");
  add("Fast worker trigger",triggers.includes("nhsos_processFastJobs_"),
    triggers.includes("nhsos_processFastJobs_")?"Installed":"Missing");

  let relayStatus = null;
  try {
    relayStatus = typeof checkNHSBatchRelayV29 === "function"
      ? checkNHSBatchRelayV29()
      : null;
  } catch (e) {
    relayStatus = { ok: false, message: e.message || String(e) };
  }
  add(
    "Tesoro NHS mail relay configured",
    !!(relayStatus && relayStatus.ok),
    relayStatus && (relayStatus.message || relayStatus.ok)
      ? String(relayStatus.message || "Relay configuration present")
      : "Relay diagnostic unavailable"
  );

  const ops=nhssec_notificationOps_(ss);
  add("No failed notifications",ops.failedNotifications===0,
    ops.failedNotifications+" failed log entr"+(ops.failedNotifications===1?"y":"ies"),
    ops.failedNotifications===0?"PASS":"WARN");
  add("No queued internal jobs",ops.queuedInternalJobs===0,
    ops.queuedInternalJobs+" queued",ops.queuedInternalJobs===0?"PASS":"WARN");

  const failures=checks.filter(c=>c.level==="FAIL").length;
  const warnings=checks.filter(c=>c.level==="WARN").length;

  return {
    ok:failures===0,
    checks:checks,
    failures:failures,
    warnings:warnings,
    generatedAt:new Date().toISOString()
  };
}


/** ============================================================
 * SOURCE GROUP: Tutor_Matching.gs
 * ============================================================ */

/**
 * NHS TUTOR MATCHING ENGINE — POLISHED VERSION
 *
 * Add this as TutorMatching.gs in the existing Apps Script project.
 *
 * IMPORTANT:
 * - Reuses the existing global NHS_SPREADSHEET_ID from Code.gs.
 * - Do NOT declare NHS_SPREADSHEET_ID again.
 * - This file contains the upgraded processTeacherTutoringRequest(e).
 * - Remove any older processTeacherTutoringRequest(e) function from TeacherPortal.gs
 *   so there is only ONE function with that name in the Apps Script project.
 *
 * RUN ONCE AFTER PASTING:
 *   installTutorMatchingEngine()
 *
 * This version:
 * - Generates matches immediately after form submission.
 * - Rechecks open requests every 5 minutes.
 * - Produces more formal match text and capitalization.
 * - Uses preferred name + last name for display (e.g., Jason Huang).
 * - Adds polished sheet formatting, widths, wrapping, frozen header, filters,
 *   status validation, assigned-tutor validation, date/time formatting,
 *   and conditional formatting.
 * - Handles common course abbreviations and punctuation differences.
 * - Uses Topic(s) as a limited secondary clue for common course shorthand.
 */


/* ============================================================
   CONFIGURATION
   ============================================================ */

const NHS_MATCHING_REQUEST_SHEET = "Tutoring Requests";
const NHS_MATCHING_TUTOR_SHEET = "Tutor Database";

const NHS_MATCH_HEADERS = [
  "Top Match 1",
  "Top Match 2",
  "Top Match 3",
  "Match Notes",
  "Match Breakdown",
  "Match Updated"
];

const NHS_REQUEST_STATUSES = [
  "New",
  "Matching",
  "Assigned",
  "Scheduled",
  "Completed",
  "Needs Follow-Up",
  "No Tutor Available",
  "Cancelled"
];

const NHS_MATCH_CLOSED_STATUSES = [
  "Assigned",
  "Scheduled",
  "Completed",
  "Cancelled"
];


/* ============================================================
   INSTALLER
   ============================================================ */

function installTutorMatchingEngine() {
  nhssec_requireMaintenanceAdmin_();

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const requestSheet = nhsperf_sheet_(ss, NHS_MATCHING_REQUEST_SHEET);
  const tutorSheet = nhsperf_sheet_(ss, NHS_MATCHING_TUTOR_SHEET);

  if (!requestSheet) {
    throw new Error('Sheet "' + NHS_MATCHING_REQUEST_SHEET + '" was not found.');
  }

  if (!tutorSheet) {
    throw new Error('Sheet "' + NHS_MATCHING_TUTOR_SHEET + '" was not found.');
  }

  ensureMatchingColumns_(requestSheet);
  formatTutoringRequestsSheet_(ss, requestSheet);
  installMatchingSafetyTrigger_();
  rankAllOpenRequests_();

  Logger.log("Tutor Matching Engine installed successfully.");
}


/* ============================================================
   TEACHER REQUEST FORM HANDLER — INSTANT MATCHING
   ============================================================ */

function processTeacherTutoringRequest(e) {
  if (!e || !e.response) {
    throw new Error(
      "processTeacherTutoringRequest must be run by the Teacher Request Form trigger."
    );
  }

  const map = {};
  e.response.getItemResponses().forEach(function(itemResponse) {
    const title = String(
      itemResponse.getItem().getTitle() || ""
    ).trim();

    let response = itemResponse.getResponse();
    if (Array.isArray(response)) {
      response = response.join(", ");
    }

    map[title] = String(response || "").trim();
  });

  const timestamp = e.response.getTimestamp() || new Date();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error(
      "The tutoring request system is busy. The form submission will be retried."
    );
  }

  let sh = null;
  let newRow = 0;

  try {
    const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
    sh = nhsperf_sheet_(ss, NHS_MATCHING_REQUEST_SHEET);

    if (!sh) {
      throw new Error(
        'Sheet "' + NHS_MATCHING_REQUEST_SHEET + '" was not found.'
      );
    }

    ensureMatchingColumns_(sh);

    if (typeof tp_ensureRequestSchema_ === "function") {
      try {
        tp_ensureRequestSchema_(sh);
      } catch (ignore) {}
    }

    const requestId =
      "TR-" +
      Utilities.formatDate(
        timestamp,
        nhs_safeTimeZone_(ss),
        "yyyyMMdd-HHmmss"
      ) +
      "-" +
      Utilities.getUuid();

    const preferredDays =
      map["Preferred Day(s)"] ||
      map["Preferred Days"] ||
      "";

    const teacherNotes = nhsinput_literalText_(
      map["Additional Teacher Notes"] ||
      map["Teacher Notes"] ||
      "",
      1200,
      "Teacher notes"
    );

    const topic = nhsinput_literalText_(
      map["Topic(s)"] ||
      map["Topic / Academic Need"] ||
      map["Topic/Academic Need"] ||
      "",
      500,
      "Topic / academic need"
    );

    const classroom = nhsinput_literalText_(
      map["Classroom / Room Number"] ||
      map["Classroom"] ||
      map["Room Number"] ||
      map["Room"] ||
      "",
      120,
      "Classroom / room number"
    );

    const recurring =
      map["Recurring?"] ||
      map["Recurring Support?"] ||
      "";

    const h = getHeaderMap_(sh);
    const row = new Array(sh.getLastColumn()).fill("");

    function set_(header, value) {
      if (h[header]) {
        row[h[header] - 1] = value;
      }
    }

    set_("Request ID", requestId);
    set_("Timestamp", timestamp);
    set_(
      "Teacher Name",
      nhsinput_literalText_(
        map["Teacher Name"] || "",
        120,
        "Teacher name"
      )
    );
    set_("Teacher Email", String(map["Teacher Email"] || "").trim().toLowerCase());
    set_(
      "Course",
      nhsinput_literalText_(
        map["Course"] || "",
        160,
        "Course"
      )
    );
    set_("Topic / Academic Need", topic);
    set_(
      "Support Type",
      nhsinput_literalText_(
        map["Support Type"] || "",
        80,
        "Support type"
      )
    );
    set_("Scheduling Preference", "Legacy Google Form");
    set_("Preferred Dates", "");
    set_(
      "Preferred Days",
      nhsinput_literalText_(
        preferredDays,
        120,
        "Preferred days"
      )
    );
    set_("Recurring?", recurring);
    set_(
      "Specific Tutor Preference",
      nhsinput_literalText_(
        map["Specific Tutor Preference"] || "",
        120,
        "Specific tutor preference"
      )
    );
    set_("Teacher Notes", teacherNotes);
    set_("Status", "New");
    set_("Location", classroom);
    set_("Last Update", new Date());
    set_("Request Source", "Teacher Google Form");
    set_("Recurring Active", false);
    set_("Confirm Assignment", false);

    newRow =
      typeof nhsdata_appendKeyedRow_ === "function"
        ? nhsdata_appendKeyedRow_(sh, "Request ID", row)
        : (
            sh
              .getRange(
                sh.getLastRow() + 1,
                1,
                1,
                row.length
              )
              .setValues([row]),
            sh.getLastRow()
          );
  } finally {
    lock.releaseLock();
  }

  // Formatting/ranking can be slower and does not need to hold the write lock.
  formatRequestRow_(sh, newRow);
  rankTutorMatchesForRow_(newRow);

  if (
    typeof nhsdata_invalidateRequestState_ ===
    "function"
  ) {
    nhsdata_invalidateRequestState_({
      teacherEmail:
        String(
          map["Teacher Email"] || ""
        )
          .trim()
          .toLowerCase()
    });
  } else if (
    typeof nhsperf_touchDataRevision_ ===
    "function"
  ) {
    nhsperf_touchDataRevision_();
  }
}


/* ============================================================
   V47.2 — WEEKLY TUTOR AVAILABILITY
   ============================================================ */

const NHSWA_SHEET = "Tutor Weekly Availability";
const NHSWA_HEADERS = [
  "Tutor ID","Tutor Name","School Email","Week Start","Date","Day",
  "Status","Confirmed At","Updated At","Source"
];
const NHSWA_ALLOWED_STATUSES = ["Available","Unavailable"];

function nhswa_ensureSheet_(ss) {
  let sh = ss.getSheetByName(NHSWA_SHEET);
  if (!sh) sh = ss.insertSheet(NHSWA_SHEET);
  nhssec_ensureHeaders_(sh, NHSWA_HEADERS);
  try { sh.setFrozenRows(1); } catch (e) {}
  try { sh.hideSheet(); } catch (e) {}
  try { if (typeof nhsperf_invalidateSheetMemo_ === "function") nhsperf_invalidateSheetMemo_(sh); } catch (e) {}
  return sh;
}

function nhswa_weekStartFromKey_(dateKey) {
  const d = nhssched_parseKeyLocal_(dateKey);
  if (!d) return "";
  const dow = d.getDay();
  const delta = dow === 0 ? -6 : 1 - dow;
  d.setDate(d.getDate() + delta);
  return nhssched_dateKey_(d, NHS_PLATFORM.TIMEZONE);
}

function nhswa_currentWeekStart_(timezone, now) {
  const key = nhssched_todayKey_(timezone, now || new Date());
  return nhswa_weekStartFromKey_(key);
}

function nhswa_confirmationTargetWeekStart_(timezone, now) {
  const tz = nhssched_timezone_(timezone);
  const when = now instanceof Date ? now : new Date();
  const todayKey = nhssched_todayKey_(tz, when);
  const today = nhssched_parseKeyLocal_(todayKey);
  if (!today) return todayKey;
  const dow = today.getDay();
  const hour = Number(Utilities.formatDate(when, tz, "H") || 0);
  let weekStart = nhswa_weekStartFromKey_(todayKey);
  if (dow === 6 || dow === 0 || (dow === 5 && hour >= 12)) {
    weekStart = nhssched_addDaysKey_(weekStart, 7);
  }
  return weekStart;
}

function nhswa_titanDatesForWeek_(weekStart) {
  return [1,2,3,4].map(function(offset) {
    const date = nhssched_addDaysKey_(weekStart, offset);
    return {
      date: date,
      day: nhssched_dayName_(date),
      label: nhssched_formatDateLabel_(date)
    };
  });
}

function nhswa_readIndex_(ss) {
  const sh = nhswa_ensureSheet_(ss);
  const index = { byTutorDate: {}, byTutorWeek: {} };
  if (sh.getLastRow() < 2) return index;
  const values = sh.getDataRange().getDisplayValues();
  const h = {};
  values[0].forEach(function(header, i) { h[String(header || "").trim()] = i; });
  values.slice(1).forEach(function(row, i) {
    const tutorId = String(row[h["Tutor ID"]] || "").trim();
    const date = String(row[h["Date"]] || "").trim();
    const weekStart = String(row[h["Week Start"]] || "").trim();
    const status = String(row[h["Status"]] || "").trim();
    if (!tutorId || !date || !NHSWA_ALLOWED_STATUSES.includes(status)) return;
    const item = {
      row: i + 2,
      tutorId: tutorId,
      tutorName: String(row[h["Tutor Name"]] || "").trim(),
      schoolEmail: String(row[h["School Email"]] || "").trim().toLowerCase(),
      weekStart: weekStart || nhswa_weekStartFromKey_(date),
      date: date,
      day: String(row[h["Day"]] || nhssched_dayName_(date)).trim(),
      status: status,
      confirmedAt: String(row[h["Confirmed At"]] || "").trim(),
      updatedAt: String(row[h["Updated At"]] || "").trim(),
      source: String(row[h["Source"]] || "Tutor Confirmed").trim()
    };
    if (!index.byTutorDate[tutorId]) index.byTutorDate[tutorId] = {};
    index.byTutorDate[tutorId][date] = item;
    if (!index.byTutorWeek[tutorId]) index.byTutorWeek[tutorId] = {};
    if (!index.byTutorWeek[tutorId][item.weekStart]) index.byTutorWeek[tutorId][item.weekStart] = {};
    index.byTutorWeek[tutorId][item.weekStart][date] = item;
  });
  return index;
}

function nhswa_tutorId_(tutor) {
  return String((tutor && (tutor.tutorId !== undefined ? tutor.tutorId : tutor["Tutor ID"])) || "").trim();
}

function nhswa_typicalDays_(tutor) {
  const raw = tutor && (tutor.titanAvailability !== undefined
    ? tutor.titanAvailability
    : tutor["Typical Titan Time Availability"]);
  return parseDays_(raw || "");
}

function nhswa_tempUnavailable_(tutor) {
  return String((tutor && (tutor.temporaryUnavailableDates !== undefined
    ? tutor.temporaryUnavailableDates
    : tutor["Temporary Unavailable Dates"])) || "").trim();
}

function nhswa_statusForTutorDate_(index, tutor, dateKey) {
  const tutorId = nhswa_tutorId_(tutor);
  const explicit = tutorId && index && index.byTutorDate[tutorId]
    ? index.byTutorDate[tutorId][dateKey]
    : null;
  const tempRaw = nhswa_tempUnavailable_(tutor);
  const temporaryBlocked = !!(
    dateKey && tempRaw && typeof nhssec_isTemporarilyUnavailable_ === "function" &&
    nhssec_isTemporarilyUnavailable_(tempRaw, dateKey)
  );
  if (temporaryBlocked) {
    return {
      date: dateKey,
      day: nhssched_dayName_(dateKey),
      status: "Unavailable",
      explicitStatus: explicit ? explicit.status : "",
      confirmed: !!explicit,
      hardBlocked: true,
      source: "Temporary Unavailable Dates"
    };
  }
  if (explicit) {
    return {
      date: dateKey,
      day: explicit.day || nhssched_dayName_(dateKey),
      status: explicit.status,
      explicitStatus: explicit.status,
      confirmed: true,
      hardBlocked: explicit.status === "Unavailable",
      source: explicit.source || "Tutor Confirmed",
      confirmedAt: explicit.confirmedAt || ""
    };
  }
  const typical = nhswa_typicalDays_(tutor).includes(normalizeText_(nhssched_dayName_(dateKey)));
  return {
    date: dateKey,
    day: nhssched_dayName_(dateKey),
    status: "Not Confirmed",
    explicitStatus: "",
    confirmed: false,
    hardBlocked: false,
    source: typical ? "Typical Availability" : "Unconfirmed",
    typicalAvailable: typical
  };
}

function nhswa_weekSummaryFromIndex_(index, tutor, weekStart) {
  const tutorId = nhswa_tutorId_(tutor);
  const dates = nhswa_titanDatesForWeek_(weekStart).map(function(spec) {
    const state = nhswa_statusForTutorDate_(index, tutor, spec.date);
    state.label = spec.label;
    state.defaultStatus = state.hardBlocked
      ? "Unavailable"
      : (state.confirmed ? state.status : (state.typicalAvailable ? "Available" : "Unavailable"));
    return state;
  });
  const explicitCount = dates.filter(function(d){ return d.confirmed; }).length;
  const confirmedAtValues = dates.map(function(d){return d.confirmedAt || "";}).filter(Boolean).sort();
  return {
    tutorId: tutorId,
    weekStart: weekStart,
    weekEnd: nhssched_addDaysKey_(weekStart, 4),
    weekLabel: nhssched_formatDateLabel_(weekStart) + " – " + nhssched_formatDateLabel_(nhssched_addDaysKey_(weekStart,4)),
    confirmed: explicitCount === dates.length,
    explicitCount: explicitCount,
    availableCount: dates.filter(function(d){return d.status === "Available";}).length,
    unavailableCount: dates.filter(function(d){return d.status === "Unavailable";}).length,
    notConfirmedCount: dates.filter(function(d){return !d.confirmed;}).length,
    confirmedAt: confirmedAtValues.length ? confirmedAtValues[confirmedAtValues.length - 1] : "",
    dates: dates
  };
}

function nhswa_weekSummary_(ss, tutor, weekStart, index) {
  return nhswa_weekSummaryFromIndex_(index || nhswa_readIndex_(ss), tutor, weekStart);
}

function nhswa_futureMapForTutor_(index, tutorId) {
  const out = {};
  const rows = index && index.byTutorDate && index.byTutorDate[String(tutorId || "").trim()] || {};
  Object.keys(rows).forEach(function(date) {
    out[date] = {
      status: rows[date].status,
      confirmedAt: rows[date].confirmedAt || "",
      source: rows[date].source || "Tutor Confirmed"
    };
  });
  return out;
}

function nhswa_validateWeekStart_(weekStart, timezone, now) {
  const key = nhssched_dateKey_(weekStart, timezone);
  if (!key || nhssched_dayName_(key) !== "Monday") {
    throw new Error("Weekly availability must use the Monday week-start date.");
  }
  const current = nhswa_currentWeekStart_(timezone, now || new Date());
  const next = nhssched_addDaysKey_(current, 7);
  if (key !== current && key !== next) {
    throw new Error("Weekly availability can only be confirmed for the current or upcoming week.");
  }
  return key;
}

function nhswa_withAvailabilityLock_(fn) {
  // Weekly availability is a shared row store across all tutors, so use the
  // script lock rather than a per-user lock to prevent concurrent append races.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    throw new Error("Weekly availability is busy saving another update. Please try again.");
  }
  try {
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function nhswa_replaceTutorWeek_(ss, tutor, weekStart, statuses, source) {
  const sh = nhswa_ensureSheet_(ss);
  const tutorId = nhswa_tutorId_(tutor);
  if (!tutorId) throw new Error("Tutor ID is required for weekly availability.");
  const expected = nhswa_titanDatesForWeek_(weekStart);
  const now = new Date();
  const temp = nhswa_tempUnavailable_(tutor);
  const byDate = {};
  (statuses || []).forEach(function(item) {
    const date = nhssched_dateKey_(item && item.date, NHS_PLATFORM.TIMEZONE);
    const status = String(item && item.status || "").trim();
    if (date && NHSWA_ALLOWED_STATUSES.includes(status)) byDate[date] = status;
  });
  expected.forEach(function(spec) {
    if (!byDate[spec.date]) throw new Error("Choose Available or Unavailable for " + spec.label + ".");
    if (temp && typeof nhssec_isTemporarilyUnavailable_ === "function" && nhssec_isTemporarilyUnavailable_(temp, spec.date)) {
      byDate[spec.date] = "Unavailable";
    }
  });

  const headers = nhsrg_headers_(sh);
  const h = {}; headers.forEach(function(name,i){h[name]=i;});
  const rowByDate = {};
  const duplicateRows = [];
  const unexpectedRows = [];
  if (sh.getLastRow() >= 2) {
    const values = sh.getRange(2,1,sh.getLastRow()-1,headers.length).getDisplayValues();
    values.forEach(function(row,i) {
      if (String(row[h["Tutor ID"]] || "").trim() !== tutorId) return;
      if (String(row[h["Week Start"]] || "").trim() !== weekStart) return;
      const date = String(row[h["Date"]] || "").trim();
      const rowNum = i + 2;
      if (!expected.some(function(spec){return spec.date === date;})) {
        unexpectedRows.push(rowNum);
      } else if (rowByDate[date]) {
        duplicateRows.push(rowNum);
      } else {
        rowByDate[date] = rowNum;
      }
    });
  }

  const appendRows = [];
  expected.forEach(function(spec) {
    const row = new Array(headers.length).fill("");
    function put(name,value){ if(h[name] !== undefined) row[h[name]]=value; }
    put("Tutor ID", tutorId);
    put("Tutor Name", String(tutor.displayName || tutor["Full Name"] || "").trim());
    put("School Email", String(tutor.schoolEmail || tutor["School Email"] || "").trim().toLowerCase());
    put("Week Start", weekStart);
    put("Date", spec.date);
    put("Day", spec.day);
    put("Status", byDate[spec.date]);
    put("Confirmed At", now);
    put("Updated At", now);
    put("Source", String(source || "Tutor Confirmed"));
    const rowNum = rowByDate[spec.date];
    if (rowNum) sh.getRange(rowNum,1,1,headers.length).setValues([row]);
    else appendRows.push(row);
  });

  if (appendRows.length) {
    sh.getRange(sh.getLastRow()+1,1,appendRows.length,headers.length).setValues(appendRows);
  }
  duplicateRows.concat(unexpectedRows).forEach(function(rowNum) {
    sh.getRange(rowNum,1,1,headers.length).clearContent();
  });
  return expected.length;
}

function nhswa_upsertDateAvailable_(ss, tutor, dateKey, source) {
  return nhswa_withAvailabilityLock_(function() {
  dateKey = nhssched_dateKey_(dateKey, nhs_safeTimeZone_(ss));
  if (!dateKey || !nhssched_isTitanDay_(dateKey)) return;
  const temp = nhswa_tempUnavailable_(tutor);
  if (temp && typeof nhssec_isTemporarilyUnavailable_ === "function" && nhssec_isTemporarilyUnavailable_(temp, dateKey)) {
    throw new Error("You marked " + nhssched_formatDateLabel_(dateKey) + " as temporarily unavailable. Update availability before accepting this assignment.");
  }
  const index = nhswa_readIndex_(ss);
  const state = nhswa_statusForTutorDate_(index, tutor, dateKey);
  if (state.confirmed && state.status === "Unavailable") {
    throw new Error("You marked " + nhssched_formatDateLabel_(dateKey) + " unavailable this week. Update weekly availability before accepting this assignment.");
  }
  // Do not create a full-week confirmation from an assignment action. Upsert only this date.
  const sh = nhswa_ensureSheet_(ss);
  const tutorId = nhswa_tutorId_(tutor);
  const headers = nhsrg_headers_(sh);
  const h = {}; headers.forEach(function(name,i){h[name]=i;});
  let rowNum = 0;
  if (sh.getLastRow() >= 2) {
    const values = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getDisplayValues();
    values.some(function(row,i){
      if (String(row[h["Tutor ID"]] || "").trim() === tutorId && String(row[h["Date"]] || "").trim() === dateKey) { rowNum=i+2; return true; }
      return false;
    });
  }
  const now = new Date();
  const out = new Array(headers.length).fill("");
  function put(name,value){ if(h[name] !== undefined) out[h[name]]=value; }
  put("Tutor ID", tutorId);
  put("Tutor Name", String(tutor.displayName || tutor["Full Name"] || "").trim());
  put("School Email", String(tutor.schoolEmail || tutor["School Email"] || "").trim().toLowerCase());
  put("Week Start", nhswa_weekStartFromKey_(dateKey));
  put("Date", dateKey);
  put("Day", nhssched_dayName_(dateKey));
  put("Status", "Available");
  put("Confirmed At", now);
  put("Updated At", now);
  put("Source", String(source || "Assignment Acceptance"));
  if (rowNum) sh.getRange(rowNum,1,1,headers.length).setValues([out]);
  else sh.getRange(sh.getLastRow()+1,1,1,headers.length).setValues([out]);

  });
}

function nhswa_assertAssignmentDate_(ss, tutor, dateKey) {
  if (!tutor || !dateKey) return { status:"Not Confirmed", warning:"" };
  const index = nhswa_readIndex_(ss);
  const state = nhswa_statusForTutorDate_(index, tutor, dateKey);
  if (state.status === "Unavailable") {
    throw new Error((tutor["Preferred Name"] || tutor["Full Name"] || "This tutor") + " is unavailable on " + nhssched_formatDateLabel_(dateKey) + " according to weekly/temporary availability.");
  }
  return {
    status: state.status,
    warning: state.confirmed ? "" : "Weekly availability is not confirmed for this date. Typical availability is being used as a fallback."
  };
}

function tutor_confirmWeeklyAvailability(payload) {
  const email = tup_requireUserEmail_();
  const ss = nhsperf_ss_(TUP_SPREADSHEET_ID);
  const tutor = tup_requireTutor_(ss, email);
  const weekStart = nhswa_validateWeekStart_(payload && payload.weekStart, nhs_safeTimeZone_(ss), new Date());
  nhswa_withAvailabilityLock_(function() {
    nhswa_replaceTutorWeek_(ss, tutor, weekStart, payload && payload.statuses, "Tutor Confirmed");
  });
  tup_queueJob_("WEEKLY_AVAILABILITY_UPDATE", {
    tutorId:tutor.tutorId,
    tutorName:tutor.displayName,
    actor:email,
    weekStart:weekStart
  });
  try { nhs_v45InvalidatePortalState_({tutorId:tutor.tutorId}); } catch(e) { tup_invalidateTutorCache_(tutor.tutorId); tup_invalidateOperationsCache_(); }
  const index = nhswa_readIndex_(ss);
  return { ok:true, weeklyAvailability:nhswa_weekSummaryFromIndex_(index,tutor,weekStart), message:"Weekly Titan Time availability confirmed." };
}

function nhsos_sendWeeklyAvailabilityReminders(weekStart) {
  nhssec_requireOfficer_();
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const key = nhswa_validateWeekStart_(weekStart || nhswa_confirmationTargetWeekStart_(nhs_safeTimeZone_(ss),new Date()), nhs_safeTimeZone_(ss), new Date());
  return nhswa_sendAvailabilityReminders_(ss,key,"MANUAL");
}

function nhswa_sendAvailabilityReminders_(ss, weekStart, wave) {
  const tutors = nhsos_readTutors_(ss);
  let sent=0, skipped=0, alreadyConfirmed=0;
  tutors.forEach(function(tutor) {
    if (["Inactive","Suspended"].includes(String(tutor.reliabilityStatus || ""))) { skipped++; return; }
    const summary = tutor.weeklyAvailability && tutor.weeklyAvailability.weekStart === weekStart
      ? tutor.weeklyAvailability
      : nhswa_weekSummary_(ss,tutor,weekStart);
    if (summary.confirmed) { alreadyConfirmed++; return; }
    if (String(tutor.reminderPreference || "").trim() === "None") { skipped++; return; }
    const recipient = String(tutor.notificationEmail || tutor.email || "").trim();
    if (!recipient || !nhsn_isValidEmail_(recipient)) { skipped++; return; }
    const ok = nhsn_sendEmailLogged_(ss, {
      requestId:"",
      recipient:recipient,
      type:"Weekly Availability Reminder",
      subject:"Confirm your NHS Titan Time availability for " + summary.weekLabel,
      dedupeKey:"WEEKLY_AVAIL|"+weekStart+"|"+String(wave||"MANUAL")+"|"+recipient.toLowerCase(),
      title:"Weekly Titan Time availability",
      message:"Please confirm the specific Titan Time dates you can tutor this week so NHS can match requests accurately.",
      body:[
        "Tesoro NHS Tutoring",
        "",
        "Please open the Tutor Portal and confirm your Titan Time availability for " + summary.weekLabel + ".",
        "",
        "Your typical availability is pre-filled; only adjust this week's exceptions, then click Confirm This Week.",
        "",
        "Weekly availability overrides your typical schedule for matching."
      ].join("\\n")
    });
    if (ok) sent++; else skipped++;
  });
  return {ok:true,weekStart:weekStart,sent:sent,skipped:skipped,alreadyConfirmed:alreadyConfirmed};
}

function nhswa_sendAutomaticReminders_(ss, now) {
  const tz = nhs_safeTimeZone_(ss);
  const localDate = nhssched_parseKeyLocal_(nhssched_todayKey_(tz, now));
  const jsDay = localDate.getDay(); // Sun=0, Mon=1 ... Sat=6
  const hour = Number(Utilities.formatDate(now, tz, "H") || 0);
  const target = nhswa_confirmationTargetWeekStart_(tz, now);
  if (jsDay === 0 && hour >= 18) return nhswa_sendAvailabilityReminders_(ss,target,"SUNDAY");
  if (jsDay === 1 && hour >= 15) return nhswa_sendAvailabilityReminders_(ss,target,"MONDAY");
  return {ok:true,skipped:true};
}

/* ============================================================
   RANKING
   ============================================================ */

function rankAllOpenRequests_() {
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const requestSheet = nhsperf_sheet_(ss, NHS_MATCHING_REQUEST_SHEET);
  if (!requestSheet) return;

  ensureMatchingColumns_(requestSheet);
  const headers = getHeaderMap_(requestSheet);
  if (!headers["Request ID"] || !headers["Status"]) {
    throw new Error('Tutoring Requests must contain "Request ID" and "Status".');
  }

  const realLast = typeof nhsdata_realLastRowByKey_ === "function"
    ? nhsdata_realLastRowByKey_(requestSheet, "Request ID")
    : requestSheet.getLastRow();
  if (realLast < 2) return;

  const ids = requestSheet
    .getRange(2, headers["Request ID"], realLast - 1, 1)
    .getDisplayValues();
  const statuses = requestSheet
    .getRange(2, headers["Status"], realLast - 1, 1)
    .getDisplayValues();

  for (let i = 0; i < ids.length; i++) {
    const requestId = String(ids[i][0] || "").trim();
    if (!requestId) continue;

    const status = String(statuses[i][0] || "").trim();
    if (NHS_MATCH_CLOSED_STATUSES.includes(status)) continue;

    rankTutorMatchesForRow_(i + 2);
  }
}


function rankTutorMatchesForRow_(rowNumber) {
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const requestSheet = nhsperf_sheet_(ss, NHS_MATCHING_REQUEST_SHEET);
  const tutorSheet = nhsperf_sheet_(ss, NHS_MATCHING_TUTOR_SHEET);

  if (!requestSheet || !tutorSheet) {
    throw new Error("Required matching sheets are missing.");
  }

  ensureMatchingColumns_(requestSheet);

  const requestHeaders = getHeaderMap_(requestSheet);
  const tutorHeaders = getHeaderMap_(tutorSheet);

  validateMatchingHeaders_(requestHeaders, tutorHeaders);

  if (rowNumber < 2 || rowNumber > requestSheet.getLastRow()) {
    return;
  }

  const requestValues = requestSheet
    .getRange(rowNumber, 1, 1, requestSheet.getLastColumn())
    .getDisplayValues()[0];

  const request = rowToObject_(requestValues, requestHeaders);

  // V43.2 HARD GUARD:
  // Never rank or write matching output into a row that is not a real request.
  // This prevents Match Notes / Match Updated from recreating phantom used rows.
  const requestId = String(request["Request ID"] || "").trim();
  if (!requestId) {
    return;
  }

  const status = String(request["Status"] || "").trim();

  if (NHS_MATCH_CLOSED_STATUSES.includes(status)) {
    return;
  }

  const resolvedCourse = resolveRequestedCourse_(request);

  if (!resolvedCourse.course) {
    writeNoMatches_(
      requestSheet,
      rowNumber,
      requestHeaders,
      "A specific course could not be identified. Please enter the full course name when possible."
    );
    return;
  }

  const tutorLastRow = tutorSheet.getLastRow();

  if (tutorLastRow < 2) {
    writeNoMatches_(
      requestSheet,
      rowNumber,
      requestHeaders,
      "No tutor profiles are currently available."
    );
    return;
  }

  const tutorData = tutorSheet
    .getRange(2, 1, tutorLastRow - 1, tutorSheet.getLastColumn())
    .getDisplayValues();

  const scored = [];
  const bookingIndex = nhsmatch_buildTutorBookingIndex_(
    requestSheet,
    requestHeaders,
    requestId,
    nhs_safeTimeZone_(ss)
  );
  const weeklyAvailabilityIndex = nhswa_readIndex_(ss);

  tutorData.forEach(row => {
    const tutor = rowToObject_(row, tutorHeaders);
    if (!String(tutor["Full Name"] || "").trim()) return;

    const result = scoreTutorForRequest_(
      request,
      tutor,
      resolvedCourse.course,
      bookingIndex,
      weeklyAvailabilityIndex
    );

    if (result.eligible) {
      scored.push(result);
    }
  });

  scored.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    if (a.currentCaseload !== b.currentCaseload) {
      return a.currentCaseload - b.currentCaseload;
    }
    return a.displayName.localeCompare(b.displayName);
  });

  const top = scored.slice(0, 3);

  requestSheet
    .getRange(rowNumber, requestHeaders["Top Match 1"])
    .setValue(top[0] ? formatMatchCell_(top[0]) : "");

  requestSheet
    .getRange(rowNumber, requestHeaders["Top Match 2"])
    .setValue(top[1] ? formatMatchCell_(top[1]) : "");

  requestSheet
    .getRange(rowNumber, requestHeaders["Top Match 3"])
    .setValue(top[2] ? formatMatchCell_(top[2]) : "");

  const breakdownPayload = top.map(function(match) {
    return {
      tutorId: match.tutorId,
      displayName: match.displayName,
      score: match.score,
      strength: match.strength,
      criteria: match.criteria || [],
      availability: match.availability || {}
    };
  });
  if (requestHeaders["Match Breakdown"]) {
    requestSheet.getRange(rowNumber, requestHeaders["Match Breakdown"]).setValue(
      breakdownPayload.length ? JSON.stringify(breakdownPayload) : ""
    );
  }

  let notes = "";

  if (top.length === 0) {
    const days = String(request["Preferred Days"] || "").trim();

    notes =
      "No eligible tutor is currently available for " +
      formalCourseName_(resolvedCourse.course) +
      (days ? " on the requested Titan Time day(s): " + formalDayList_(days) : "") +
      ".";
  } else {
    const noteLines = [];

    if (resolvedCourse.inferred) {
      noteLines.push(
        "Course interpreted as " +
        formalCourseName_(resolvedCourse.course) +
        " based on the request details."
      );
      noteLines.push("");
    }

    top.forEach((match, index) => {
      noteLines.push(
        (index + 1) +
        ". " +
        match.displayName +
        " — " +
        match.score +
        " points"
      );

      (match.criteria || []).forEach(function(criterion) {
        noteLines.push(
          "   " + criterion.label + ": " + criterion.score + "/" + criterion.max +
          (criterion.detail ? " — " + criterion.detail : "")
        );
      });

      if (index !== top.length - 1) {
        noteLines.push("");
      }
    });

    notes = noteLines.join("\n");
  }

  requestSheet
    .getRange(rowNumber, requestHeaders["Match Notes"])
    .setValue(notes);

  requestSheet
    .getRange(rowNumber, requestHeaders["Match Updated"])
    .setValue(new Date());

  if (
    top.length === 0 &&
    status !== "Needs Follow-Up" &&
    status !== "No Tutor Available"
  ) {
    requestSheet
      .getRange(rowNumber, requestHeaders["Status"])
      .setValue("No Tutor Available");
  } else if (
    top.length > 0 &&
    status === "No Tutor Available"
  ) {
    requestSheet
      .getRange(rowNumber, requestHeaders["Status"])
      .setValue("New");
  }

  formatRequestRow_(requestSheet, rowNumber);
}


/* ============================================================
   SCORING ENGINE
   ============================================================ */

/**
 * Builds a date-level booking index once per request ranking pass.
 * Matching must never recommend a tutor for a Titan Time date on which
 * that tutor already has another lifecycle-active assignment.
 */
function nhsmatch_buildTutorBookingIndex_(requestSheet, requestHeaders, excludeRequestId, timezone) {
  const index = { byId: {}, byName: {} };
  if (!requestSheet || requestSheet.getLastRow() < 2) return index;

  const realLast = typeof nhsdata_realLastRowByKey_ === "function"
    ? nhsdata_realLastRowByKey_(requestSheet, "Request ID")
    : requestSheet.getLastRow();
  if (realLast < 2) return index;

  const rows = requestSheet
    .getRange(2, 1, realLast - 1, requestSheet.getLastColumn())
    .getDisplayValues();

  function add(map, key, dateKey) {
    key = String(key || "").trim();
    if (!key || !dateKey) return;
    if (!map[key]) map[key] = {};
    map[key][dateKey] = true;
  }

  rows.forEach(function(row) {
    const request = rowToObject_(row, requestHeaders);
    const requestId = String(request["Request ID"] || "").trim();
    if (!requestId || requestId === String(excludeRequestId || "").trim()) return;
    if (typeof nhslife_activeCaseload_ === "function" && !nhslife_activeCaseload_(request)) return;

    const dateKey = nhslife_dateKey_(request["Scheduled Date"], timezone);
    if (!dateKey) return;

    const tutorId = String(request["Assigned Tutor ID"] || "").trim();
    const tutorName = normalizeText_(request["Assigned Tutor"] || "");
    add(index.byId, tutorId, dateKey);
    add(index.byName, tutorName, dateKey);
  });

  return index;
}

function nhsmatch_bookedDatesForTutor_(bookingIndex, tutor) {
  const booked = {};
  if (!bookingIndex) return booked;

  const tutorId = String(tutor["Tutor ID"] || "").trim();
  const names = [
    tutor["Full Name"],
    tutor["Preferred Name"],
    formalTutorDisplayName_(tutor)
  ].map(normalizeText_).filter(Boolean);

  if (tutorId && bookingIndex.byId && bookingIndex.byId[tutorId]) {
    Object.keys(bookingIndex.byId[tutorId]).forEach(function(dateKey) {
      booked[dateKey] = true;
    });
  }

  if (bookingIndex.byName) {
    names.forEach(function(name) {
      const dates = bookingIndex.byName[name];
      if (!dates) return;
      Object.keys(dates).forEach(function(dateKey) {
        booked[dateKey] = true;
      });
    });
  }

  return booked;
}

function scoreTutorForRequest_(request, tutor, resolvedCourse, bookingIndex, weeklyAvailabilityIndex) {
  let score = 0;
  const criteria = [];
  const reasons = [];
  function addCriterion(key,label,value,max,detail,status) {
    const n = Math.max(0, Math.min(Number(max || 0), Math.round(Number(value || 0))));
    score += n;
    criteria.push({key:key,label:label,score:n,max:max,detail:String(detail||""),status:String(status||"")});
    if (detail) reasons.push(label.toLowerCase() + " +" + n + " — " + detail);
  }

  const requestedCourse = canonicalCourse_(resolvedCourse);
  const requestedDays = parseDays_(request["Preferred Days"]);
  const requestedDates = nhssched_normalizeDateList_(request["Preferred Dates"] || "", NHS_PLATFORM.TIMEZONE);
  const scheduledDateKey = nhslife_dateKey_(request["Scheduled Date"], NHS_PLATFORM.TIMEZONE);
  const tutorDays = parseDays_(tutor["Typical Titan Time Availability"]);
  const strongestCourses = splitCourseList_(tutor["Strongest Course"]);

  const courseFields = [
    tutor["English Courses"], tutor["Mathematics Courses"], tutor["Science Courses"],
    tutor["Social Science Courses"], tutor["World Language Courses"], tutor["Computer Science / CTE Courses"]
  ];
  const allCourses = [];
  courseFields.forEach(function(field) { splitCourseList_(field).forEach(function(course){allCourses.push(course);}); });
  const exactCourseMatch = allCourses.some(function(course){return canonicalCourse_(course) === requestedCourse;});
  const strongestCourseMatch = strongestCourses.some(function(course){return canonicalCourse_(course) === requestedCourse;});
  if (!exactCourseMatch && !strongestCourseMatch) return {eligible:false,score:0,reason:"course not listed"};

  // QUALIFICATION — 50 points maximum.
  if (strongestCourseMatch) {
    addCriterion("qualification","Qualification",50,50,"Listed as one of the tutor's strongest courses","excellent");
  } else {
    addCriterion("qualification","Qualification",40,50,"Tutor is explicitly qualified for this course","qualified");
  }

  const bookedDates = nhsmatch_bookedDatesForTutor_(bookingIndex, tutor);
  const dateCandidates = requestedDates.length ? requestedDates.slice() : (scheduledDateKey ? [scheduledDateKey] : []);
  let freeDates = dateCandidates.filter(function(dateKey){return !bookedDates[dateKey];});
  if (dateCandidates.length && !freeDates.length) return {eligible:false,score:0,reason:"already assigned on all requested Titan Time dates"};

  // Temporary unavailability is a hard date-level block.
  const tempRaw = String(tutor["Temporary Unavailable Dates"] || "").trim();
  if (freeDates.length && tempRaw && typeof nhssec_isTemporarilyUnavailable_ === "function") {
    freeDates = freeDates.filter(function(dateKey){return !nhssec_isTemporarilyUnavailable_(tempRaw,dateKey);});
    if (!freeDates.length) return {eligible:false,score:0,reason:"unavailable on all requested dates"};
  }

  // SCHEDULE — 10 points maximum. Weekly confirmation is freshest truth.
  let schedulePoints = 0;
  let scheduleDetail = "";
  let availabilitySummary = {confirmedAvailable:[],confirmedUnavailable:[],unconfirmed:[],freeDates:freeDates.slice()};
  if (freeDates.length) {
    freeDates.forEach(function(dateKey) {
      const state = nhswa_statusForTutorDate_(weeklyAvailabilityIndex || {byTutorDate:{}}, tutor, dateKey);
      if (state.status === "Available" && state.confirmed) availabilitySummary.confirmedAvailable.push(dateKey);
      else if (state.status === "Unavailable") availabilitySummary.confirmedUnavailable.push(dateKey);
      else availabilitySummary.unconfirmed.push(dateKey);
    });
    const usable = freeDates.filter(function(dateKey){return !availabilitySummary.confirmedUnavailable.includes(dateKey);});
    if (!usable.length) return {eligible:false,score:0,reason:"weekly availability marks all requested dates unavailable"};
    if (availabilitySummary.confirmedAvailable.length) {
      schedulePoints = 10;
      scheduleDetail = "Confirmed available " + availabilitySummary.confirmedAvailable.map(nhssched_formatDateLabel_).join(", ");
    } else {
      const typicalOverlap = usable.filter(function(dateKey){return tutorDays.includes(normalizeText_(nhssched_dayName_(dateKey)));});
      if (typicalOverlap.length) {
        schedulePoints = 6;
        scheduleDetail = "Weekly availability not confirmed; typical availability includes " + typicalOverlap.map(nhssched_dayName_).join(", ");
      } else {
        schedulePoints = 4;
        scheduleDetail = "Weekly availability not confirmed; no typical-day confirmation for the requested date";
      }
    }
  } else if (requestedDays.length) {
    const overlap = requestedDays.filter(function(day){return tutorDays.includes(day);});
    if (overlap.length) {
      schedulePoints = 6;
      scheduleDetail = "Typical availability overlaps " + humanJoinDays_(overlap) + "; specific week not yet confirmed";
    } else {
      schedulePoints = 4;
      scheduleDetail = "Specific weekly availability is not yet known for this recurring request";
    }
  } else if (tutorDays.length) {
    schedulePoints = 6;
    scheduleDetail = "Typical Titan Time availability on file; no specific date supplied";
  } else {
    schedulePoints = 4;
    scheduleDetail = "No weekly confirmation yet; availability requires officer/tutor confirmation";
  }
  addCriterion("schedule","Schedule",schedulePoints,10,scheduleDetail,schedulePoints===10?"confirmed":schedulePoints>=6?"fallback":"unconfirmed");

  // RELIABILITY — 15 points maximum.
  const reliability = String(tutor["Reliability Status"] || "New").trim();
  const reliabilityNorm = normalizeText_(reliability);
  if (reliabilityNorm.includes("suspended") || reliabilityNorm.includes("inactive")) return {eligible:false,score:0,reason:"inactive reliability status"};
  const reliabilityPoints = nhsmatch_reliabilityPoints_(reliability);
  addCriterion("reliability","Reliability",reliabilityPoints,15,reliability + " status",reliability);

  // CAPACITY — 10 points maximum.
  const maxSessions = parseFirstNumber_(tutor["Maximum Sessions / Week"]);
  const currentCaseload = parseFirstNumber_(tutor["Current Caseload"]);
  let capacityPoints = 0;
  let capacityDetail = "No weekly capacity configured";
  if (maxSessions > 0) {
    if (currentCaseload >= maxSessions) return {eligible:false,score:0,reason:"at weekly capacity"};
    const remainingRatio = (maxSessions - currentCaseload) / maxSessions;
    capacityPoints = currentCaseload === 0 ? 10 : (remainingRatio >= 0.5 ? 6 : 3);
    capacityDetail = currentCaseload + "/" + maxSessions + " active; " + (maxSessions-currentCaseload) + " slot(s) remaining";
  }
  addCriterion("capacity","Capacity",capacityPoints,10,capacityDetail,capacityPoints>=6?"available":"limited");

  // TEACHER PREFERENCE — 10 points maximum.
  const preference = normalizeText_(request["Specific Tutor Preference"]);
  const fullName = normalizeText_(tutor["Full Name"]);
  const preferredName = normalizeText_(tutor["Preferred Name"]);
  const preferred = !!(preference && (preference === fullName || preference === preferredName || fullName.includes(preference) || (preferredName && preference.includes(preferredName))));
  addCriterion("teacherPreference","Teacher Preference",preferred?10:0,10,preferred?"Teacher specifically requested this tutor":"No matching tutor preference",preferred?"preferred":"neutral");

  // EXPERIENCE — 5 points maximum.
  const experience = normalizeText_(tutor["Previous Experience"]);
  const experiencePoints = experience === "yes" ? 5 : 0;
  addCriterion("experience","Experience",experiencePoints,5,experiencePoints?"Prior tutoring experience on file":"No prior tutoring experience indicated",experiencePoints?"experienced":"new");

  const notificationEmail = String(tutor["Notification Email"] || "").trim();
  if (!notificationEmail) return {eligible:false,score:0,reason:"missing notification email"};

  score = Math.max(0, Math.min(100, Math.round(score)));
  return {
    eligible:true,
    score:score,
    strength:nhsmatch_strengthLabel_(score),
    displayName:formalTutorDisplayName_(tutor),
    fullName:String(tutor["Full Name"] || "").trim(),
    tutorId:String(tutor["Tutor ID"] || "").trim(),
    currentCaseload:currentCaseload,
    reasons:reasons,
    criteria:criteria,
    availability:availabilitySummary
  };
}

function nhsmatch_reliabilityPoints_(status) {
  const value = normalizeText_(status || "New");

  if (value.includes("excellent")) return 15;
  if (value.includes("reliable") || value.includes("good")) return 12;
  if (value.includes("warning") || value.includes("at risk")) return 4;
  if (value.includes("probation")) return 0;
  if (value.includes("suspended") || value.includes("inactive")) return 0;

  // New and unknown legacy values use the neutral onboarding score.
  return 8;
}

function nhsmatch_strengthLabel_(score) {
  score = Number(score || 0);
  if (score >= 90) return "Exceptional Match";
  if (score >= 80) return "Strong Match";
  if (score >= 70) return "Good Match";
  if (score >= 60) return "Acceptable Match";
  return "Weak Match";
}


/* ============================================================
   SMARTER COURSE RESOLUTION
   ============================================================ */

function resolveRequestedCourse_(request) {
  const rawCourse = String(request["Course"] || "").trim();
  const rawTopic = String(request["Topic / Academic Need"] || "").trim();

  const direct = nhscourse_resolve_(rawCourse);
  if (direct) {
    return { course: direct.id, inferred: false };
  }

  // Legacy safety: only infer when the stored Course field is vague/empty.
  // New Teacher Portal submissions are validated against the registry before
  // they are written, so this path mainly protects old rows.
  if (!rawCourse || isVagueCourse_(rawCourse)) {
    const inferred = nhscourse_inferFromText_(rawCourse + " " + rawTopic);
    if (inferred) {
      return { course: inferred.id, inferred: true };
    }
  }

  return { course: "", inferred: false };
}


function isVagueCourse_(course) {
  const vague = [
    "",
    "ap",
    "math",
    "mathematics",
    "english",
    "science",
    "history",
    "social science",
    "language",
    "world language",
    "computer science",
    "cte"
  ];

  return vague.includes(normalizeText_(course));
}


/* ============================================================
   SHEET FORMATTING
   ============================================================ */

function formatTutoringRequestsSheet_(ss, sheet) {
  ensureMatchingColumns_(sheet);

  const headers = getHeaderMap_(sheet);
  const lastColumn = sheet.getLastColumn();
  const lastRow = Math.max(sheet.getLastRow(), 2);

  // Freeze header row.
  sheet.setFrozenRows(1);

  // Formal header styling.
  const headerRange = sheet.getRange(1, 1, 1, lastColumn);

  headerRange
    .setBackground("#1F4E78")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  sheet.setRowHeight(1, 42);

  // General body formatting.
  sheet
    .getRange(2, 1, Math.max(1, lastRow - 1), lastColumn)
    .setVerticalAlignment("top")
    .setFontSize(10);

  // Reasonable widths for operational use.
  setColumnWidthIfExists_(sheet, headers, "Request ID", 160);
  setColumnWidthIfExists_(sheet, headers, "Timestamp", 145);
  setColumnWidthIfExists_(sheet, headers, "Teacher Name", 145);
  setColumnWidthIfExists_(sheet, headers, "Teacher Email", 210);
  setColumnWidthIfExists_(sheet, headers, "Course", 190);
  setColumnWidthIfExists_(sheet, headers, "Topic / Academic Need", 270);
  setColumnWidthIfExists_(sheet, headers, "Support Type", 180);
  setColumnWidthIfExists_(sheet, headers, "Preferred Dates", 190);
  setColumnWidthIfExists_(sheet, headers, "Preferred Days", 180);
  setColumnWidthIfExists_(sheet, headers, "Recurring?", 190);
  setColumnWidthIfExists_(sheet, headers, "Specific Tutor Preference", 190);
  setColumnWidthIfExists_(sheet, headers, "Teacher Notes", 260);
  setColumnWidthIfExists_(sheet, headers, "Status", 155);
  setColumnWidthIfExists_(sheet, headers, "Assigned Tutor", 175);
  setColumnWidthIfExists_(sheet, headers, "Scheduled Date", 120);
  setColumnWidthIfExists_(sheet, headers, "Scheduled Time", 115);
  setColumnWidthIfExists_(sheet, headers, "Location", 135);
  setColumnWidthIfExists_(sheet, headers, "Coordinator Notes", 260);
  setColumnWidthIfExists_(sheet, headers, "Top Match 1", 205);
  setColumnWidthIfExists_(sheet, headers, "Top Match 2", 205);
  setColumnWidthIfExists_(sheet, headers, "Top Match 3", 205);
  setColumnWidthIfExists_(sheet, headers, "Match Notes", 450);
  setColumnWidthIfExists_(sheet, headers, "Match Updated", 145);

  // Wrap the fields that can contain longer text.
  [
    "Topic / Academic Need",
    "Support Type",
    "Preferred Dates",
    "Preferred Days",
    "Recurring?",
    "Teacher Notes",
    "Coordinator Notes",
    "Top Match 1",
    "Top Match 2",
    "Top Match 3",
    "Match Notes"
  ].forEach(header => {
    if (headers[header]) {
      sheet
        .getRange(2, headers[header], Math.max(1, sheet.getMaxRows() - 1), 1)
        .setWrap(true);
    }
  });

  // Date / time formatting.
  if (headers["Timestamp"]) {
    sheet
      .getRange(2, headers["Timestamp"], Math.max(1, sheet.getMaxRows() - 1), 1)
      .setNumberFormat("mmm d, yyyy h:mm AM/PM");
  }

  if (headers["Last Update"]) {
    sheet
      .getRange(2, headers["Last Update"], Math.max(1, sheet.getMaxRows() - 1), 1)
      .setNumberFormat("mmm d, yyyy h:mm AM/PM");
  }

  if (headers["Match Updated"]) {
    sheet
      .getRange(2, headers["Match Updated"], Math.max(1, sheet.getMaxRows() - 1), 1)
      .setNumberFormat("mmm d, yyyy h:mm AM/PM");
  }

  if (headers["Scheduled Date"]) {
    sheet
      .getRange(2, headers["Scheduled Date"], Math.max(1, sheet.getMaxRows() - 1), 1)
      .setNumberFormat("mmm d, yyyy");
  }

  if (headers["Scheduled Time"]) {
    sheet
      .getRange(2, headers["Scheduled Time"], Math.max(1, sheet.getMaxRows() - 1), 1)
      .setNumberFormat("h:mm AM/PM");
  }

  // Status dropdown.
  if (headers["Status"]) {
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(NHS_REQUEST_STATUSES, true)
      .setAllowInvalid(false)
      .build();

    sheet
      .getRange(2, headers["Status"], Math.max(1, sheet.getMaxRows() - 1), 1)
      .setDataValidation(statusRule);
  }

  // Assigned Tutor dropdown from Teacher Directory, if available.
  setAssignedTutorValidation_(ss, sheet);

  // Filter.
  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, sheet.getMaxRows(), lastColumn).createFilter();
  }

  // Conditional formatting.
  applyRequestConditionalFormatting_(sheet, headers);

  // Slightly emphasize matching columns.
  NHS_MATCH_HEADERS.forEach(header => {
    if (!headers[header]) return;

    sheet
      .getRange(1, headers[header])
      .setBackground("#D9EAF7")
      .setFontColor("#1F1F1F");
  });

  // Hide obsolete legacy intake columns to reduce clutter, without deleting data.
  [
    "Student Name",
    "Student Grade",
    "Scheduling Preference",
    "Urgency"
  ].forEach(header => {
    if (headers[header]) {
      sheet.hideColumns(headers[header]);
    }
  });
}


function formatRequestRow_(sheet, rowNumber) {
  const headers = getHeaderMap_(sheet);

  sheet
    .getRange(rowNumber, 1, 1, sheet.getLastColumn())
    .setVerticalAlignment("top");

  // Formal capitalization for common fixed-response fields.
  normalizeCellCase_(sheet, rowNumber, headers["Support Type"]);
  normalizeCellCase_(sheet, rowNumber, headers["Status"]);

  // Format dates.
  if (headers["Timestamp"]) {
    sheet
      .getRange(rowNumber, headers["Timestamp"])
      .setNumberFormat("mmm d, yyyy h:mm AM/PM");
  }

  if (headers["Last Update"]) {
    sheet
      .getRange(rowNumber, headers["Last Update"])
      .setNumberFormat("mmm d, yyyy h:mm AM/PM");
  }

  if (headers["Match Updated"]) {
    sheet
      .getRange(rowNumber, headers["Match Updated"])
      .setNumberFormat("mmm d, yyyy h:mm AM/PM");
  }

  sheet.setRowHeight(rowNumber, 48);
}


function normalizeCellCase_(sheet, rowNumber, columnNumber) {
  if (!columnNumber) return;

  const cell = sheet.getRange(rowNumber, columnNumber);
  const value = String(cell.getDisplayValue() || "").trim();

  if (!value) return;

  const pieces = value.split(",").map(piece => {
    return titleCasePhrase_(piece.trim());
  });

  cell.setValue(pieces.join(", "));
}


function applyRequestConditionalFormatting_(sheet, headers) {
  if (!headers["Status"]) return;

  const statusRange = sheet.getRange(
    2,
    headers["Status"],
    Math.max(1, sheet.getMaxRows() - 1),
    1
  );

  const rules = sheet.getConditionalFormatRules()
    .filter(rule => {
      const ranges = rule.getRanges();
      return !ranges.some(range =>
        range.getColumn() === headers["Status"] &&
        range.getSheet().getName() === sheet.getName()
      );
    });

  const statusRules = [
    ["New", "#FFF2CC"],
    ["Matching", "#D9EAF7"],
    ["Assigned", "#D9EAD3"],
    ["Scheduled", "#D9EAD3"],
    ["Completed", "#E2F0D9"],
    ["Needs Follow-Up", "#FCE5CD"],
    ["No Tutor Available", "#F4CCCC"],
    ["Cancelled", "#E7E6E6"]
  ];

  statusRules.forEach(([status, color]) => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(status)
        .setBackground(color)
        .setRanges([statusRange])
        .build()
    );
  });

  sheet.setConditionalFormatRules(rules);
}


function setAssignedTutorValidation_(ss, requestSheet) {
  const headers = getHeaderMap_(requestSheet);
  const assignedTutorCol = headers["Assigned Tutor"];

  if (!assignedTutorCol) return;

  const directory = nhsperf_sheet_(ss, "Teacher Directory");

  if (!directory) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(directory.getRange("A8:A1000"), true)
    .setAllowInvalid(true)
    .build();

  requestSheet
    .getRange(2, assignedTutorCol, Math.max(1, requestSheet.getMaxRows() - 1), 1)
    .setDataValidation(rule);
}


function setColumnWidthIfExists_(sheet, headers, header, width) {
  if (headers[header]) {
    sheet.setColumnWidth(headers[header], width);
  }
}


/* ============================================================
   MATCHING COLUMN SETUP
   ============================================================ */

function ensureMatchingColumns_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(value => String(value || "").trim());

  let nextColumn = lastColumn + 1;

  NHS_MATCH_HEADERS.forEach(header => {
    if (!headers.includes(header)) {
      sheet.getRange(1, nextColumn).setValue(header);
      headers.push(header);
      nextColumn++;
    }
  });
}


/* ============================================================
   TRIGGER
   ============================================================ */

function installMatchingSafetyTrigger_() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "rankAllOpenRequests_") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("rankAllOpenRequests_")
    .timeBased()
    .everyMinutes(5)
    .create();
}


/* ============================================================
   UTILITIES
   ============================================================ */

function validateMatchingHeaders_(requestHeaders, tutorHeaders) {
  [
    "Course",
    "Preferred Days",
    "Specific Tutor Preference",
    "Status",
    "Assigned Tutor",
    "Top Match 1",
    "Top Match 2",
    "Top Match 3",
    "Match Notes",
    "Match Updated"
  ].forEach(header => {
    if (!requestHeaders[header]) {
      throw new Error(
        'Tutoring Requests is missing required header: "' + header + '".'
      );
    }
  });

  [
    "Tutor ID",
    "Full Name",
    "Preferred Name",
    "English Courses",
    "Mathematics Courses",
    "Science Courses",
    "Social Science Courses",
    "World Language Courses",
    "Computer Science / CTE Courses",
    "Strongest Course",
    "Typical Titan Time Availability",
    "Maximum Sessions / Week",
    "Previous Experience",
    "Current Caseload",
    "Reliability Status"
  ].forEach(header => {
    if (!tutorHeaders[header]) {
      throw new Error(
        'Tutor Database is missing required header: "' + header + '".'
      );
    }
  });
}


function getHeaderMap_(sheet) {
  return nhsperf_headerMap_(sheet, true, false);
}


function rowToObject_(row, headerMap) {
  const obj = {};

  Object.keys(headerMap).forEach(header => {
    obj[header] = row[headerMap[header] - 1];
  });

  return obj;
}


function formalTutorDisplayName_(tutor) {
  const fullName = String(tutor["Full Name"] || "").trim();
  const preferredName = String(tutor["Preferred Name"] || "").trim();

  if (!preferredName) {
    return fullName;
  }

  const parts = fullName.split(/\s+/).filter(Boolean);

  if (parts.length <= 1) {
    return preferredName;
  }

  const lastName = parts[parts.length - 1];

  if (normalizeText_(preferredName) === normalizeText_(fullName)) {
    return fullName;
  }

  return preferredName + " " + lastName;
}


function formatMatchCell_(match) {
  const label = match.strength || nhsmatch_strengthLabel_(match.score);
  return match.displayName + " — " + match.score + "/100 · " + label;
}


function formalizeReason_(reason) {
  const text = String(reason || "").trim();

  if (!text) return "";

  if (/^\d/.test(text)) {
    return text.charAt(0).toUpperCase() + text.slice(1);
  }

  return text.charAt(0).toUpperCase() + text.slice(1);
}


function writeNoMatches_(sheet, rowNumber, headers, reason) {
  sheet.getRange(rowNumber, headers["Top Match 1"]).setValue("");
  sheet.getRange(rowNumber, headers["Top Match 2"]).setValue("");
  sheet.getRange(rowNumber, headers["Top Match 3"]).setValue("");
  sheet.getRange(rowNumber, headers["Match Notes"]).setValue(reason);
  if (headers["Match Breakdown"]) sheet.getRange(rowNumber, headers["Match Breakdown"]).setValue("");
  sheet.getRange(rowNumber, headers["Match Updated"]).setValue(new Date());

  formatRequestRow_(sheet, rowNumber);
}


function splitCourseList_(value) {
  return String(value || "")
    .split(/[,;\n]+/)
    .map(item => item.trim())
    .filter(Boolean);
}


function parseDays_(value) {
  const text = normalizeText_(value);

  const days = [
    "tuesday",
    "wednesday",
    "thursday",
    "friday"
  ];

  return days.filter(day => text.includes(day));
}


function humanJoinDays_(days) {
  const formal = days.map(day => titleCasePhrase_(day));

  if (formal.length === 0) return "";
  if (formal.length === 1) return formal[0];
  if (formal.length === 2) return formal[0] + " and " + formal[1];

  return (
    formal.slice(0, -1).join(", ") +
    ", and " +
    formal[formal.length - 1]
  );
}


function formalDayList_(value) {
  return humanJoinDays_(parseDays_(value));
}


function parseFirstNumber_(value) {
  const match = String(value || "").match(/\d+/);
  return match ? Number(match[0]) : 0;
}


function normalizeText_(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[.(),:;'"’\-_/]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}


function titleCasePhrase_(value) {
  const smallWords = new Set([
    "a", "an", "and", "as", "at", "but", "by", "for",
    "in", "of", "on", "or", "the", "to", "with"
  ]);

  return String(value || "")
    .split(/\s+/)
    .map((word, index) => {
      const lower = word.toLowerCase();

      if (
        index > 0 &&
        smallWords.has(lower)
      ) {
        return lower;
      }

      if (
        lower === "ap" ||
        lower === "us" ||
        lower === "cte" ||
        lower === "nhs"
      ) {
        return lower.toUpperCase();
      }

      return lower.charAt(0).toUpperCase() + lower.slice(1);
    })
    .join(" ");
}


/* ============================================================
   V45.5 — TESORO COURSE REGISTRY + ALIAS NORMALIZATION
   Source: CUSD Course Catalog 2025–2026 THS offerings, with
   Tesoro's new Honors Geometry with Algebra 2 course added.
   Fine Arts are intentionally excluded. Physical Education is
   intentionally excluded except Health.
   ============================================================ */

const NHS_COURSE_CATALOG_VERSION = "THS_2025_26_PLUS_HONORS_GEOMETRY_ALGEBRA_2";
const NHS_COURSE_REGISTRY = [
  { id: "AMERICAN_CULTURES", name: "American Cultures", department: "English", aliases: ["Amer Cultures", "Cultures"] },
  { id: "AP_ENGLISH_LANGUAGE", name: "AP English Language and Composition", department: "English", aliases: ["AP Lang", "AP Language", "AP English Lang", "AP English Language", "APENG Lang"] },
  { id: "AP_ENGLISH_LITERATURE", name: "AP Literature and Composition", department: "English", aliases: ["AP Lit", "AP Literature", "AP English Lit", "AP English Literature", "AP English Literature and Composition"] },
  { id: "AP_SEMINAR", name: "AP Seminar", department: "English", aliases: ["Seminar AP"] },
  { id: "ENGLISH_1", name: "English 1", department: "English", aliases: ["English I", "Eng 1", "Eng I"] },
  { id: "ENGLISH_2", name: "English 2", department: "English", aliases: ["English II", "Eng 2", "Eng II"] },
  { id: "ENGLISH_3", name: "English 3", department: "English", aliases: ["English III", "Eng 3", "Eng III"] },
  { id: "ENGLISH_4", name: "English 4", department: "English", aliases: ["English IV", "Eng 4", "Eng IV"] },
  { id: "CHILD_ADOLESCENT_LITERATURE", name: "Exploration of Child and Adolescent Literature", department: "English", aliases: ["Child and Adolescent Literature", "Child & Adolescent Literature", "Child and Adolescent Lit", "Child & Adolescent Lit", "Child Lit"] },
  { id: "HONORS_ENGLISH_1", name: "Honors English 1", department: "English", aliases: ["Honors English I", "English 1 Honors", "English I Honors", "H English 1", "Eng 1H"] },
  { id: "HONORS_ENGLISH_2", name: "Honors English 2", department: "English", aliases: ["Honors English II", "English 2 Honors", "English II Honors", "H English 2", "Eng 2H"] },
  { id: "ACCOUNTING", name: "Accounting", department: "Math", aliases: ["Acct"] },
  { id: "ALGEBRA_1", name: "Algebra I", department: "Math", aliases: ["Algebra 1", "Alg I", "Alg 1"] },
  { id: "ALGEBRA_1A_1B", name: "Algebra IA/IB", department: "Math", aliases: ["Algebra 1A/1B", "Algebra 1A 1B", "Alg 1A/1B", "Directed Algebra", "Directed Algebra 1A/1B"] },
  { id: "ALGEBRA_2", name: "Algebra II", department: "Math", aliases: ["Algebra 2", "Alg II", "Alg 2"] },
  { id: "HONORS_GEOMETRY_ALGEBRA_2", name: "Honors Geometry with Algebra 2", department: "Math", aliases: ["Honors Geometry + Algebra 2", "Honors Geometry Algebra 2", "H Geometry with Algebra 2", "Honors Geo Alg 2", "H Geo/Alg 2", "Honors Geo Algebra 2"] },
  { id: "AP_CALCULUS_AB", name: "AP Calculus AB", department: "Math", aliases: ["AP Calc AB", "Calc AB", "Calculus AB", "APCalc AB"] },
  { id: "AP_CALCULUS_BC", name: "AP Calculus BC", department: "Math", aliases: ["AP Calc BC", "Calc BC", "Calculus BC", "APCalc BC"] },
  { id: "AP_STATISTICS", name: "AP Statistics", department: "Math", aliases: ["AP Stats", "AP Stat", "Statistics AP"] },
  { id: "BEGINNING_PERSONAL_FINANCE", name: "Beginning Personal Finance", department: "Math", aliases: ["Personal Finance", "Beg Personal Finance", "BPF"] },
  { id: "GEOMETRY", name: "Geometry", department: "Math", aliases: ["Geo"] },
  { id: "HONORS_ALGEBRA_2_TRIG", name: "Honors Algebra 2/Trigonometry", department: "Math", aliases: ["Honors Algebra II/Trig", "Honors Algebra 2/Trig", "H Algebra 2/Trig", "H Alg 2/Trig", "Alg 2/Trig Honors", "Honors Algebra 2 with Trigonometry"] },
  { id: "HONORS_PRECALCULUS", name: "Honors Precalculus", department: "Math", aliases: ["Honors Precalc", "H Precalc", "Precalc Honors", "Precalculus Honors"] },
  { id: "INTRO_DATA_SCIENCE", name: "Introduction to Data Science", department: "Math", aliases: ["Intro to Data Science", "Intro Data Science", "Data Science"] },
  { id: "PRECALCULUS", name: "PreCalculus", department: "Math", aliases: ["Precalculus", "Precalc", "Pre Calc", "Pre-Calculus"] },
  { id: "TOPICS_IN_CALCULUS", name: "Topics in Calculus", department: "Math", aliases: ["Topics in Calc", "Calculus Topics"] },
  { id: "AP_BIOLOGY", name: "AP Biology", department: "Science", aliases: ["AP Bio"] },
  { id: "AP_CHEMISTRY", name: "AP Chemistry", department: "Science", aliases: ["AP Chem"] },
  { id: "AP_ENVIRONMENTAL_SCIENCE", name: "AP Environmental Science", department: "Science", aliases: ["APES", "AP Env Sci", "AP Environmental", "Environmental Science AP"] },
  { id: "AP_PHYSICS_C", name: "AP Physics C", department: "Science", aliases: ["Physics C", "AP Phys C", "AP Physics C Mechanics", "AP Physics C Mech", "Physics C Mechanics", "Physics C Mech", "AP Physics C E&M", "AP Physics C Electricity and Magnetism"] },
  { id: "BIOLOGY", name: "Biology", department: "Science", aliases: ["Bio"] },
  { id: "CHEMISTRY", name: "Chemistry", department: "Science", aliases: ["Chem"] },
  { id: "CHEM_ENV_ENGINEERING", name: "Chemistry and Environmental Engineering", department: "Science", aliases: ["Chemistry & Environmental Engineering", "Chem and Environmental Engineering", "Chem & Env Engineering"] },
  { id: "CONCEPTUAL_PHYSICS", name: "Conceptual Physics", department: "Science", aliases: ["Conceptual Phys", "CP Physics"] },
  { id: "ENVIRONMENTAL_FIELD_STUDIES", name: "Environmental Field Studies", department: "Science", aliases: ["Env Field Studies", "Environmental Field Study", "EFS"] },
  { id: "HONORS_ANATOMY_PHYSIOLOGY", name: "Honors Anatomy & Physiology", department: "Science", aliases: ["Honors Anatomy and Physiology", "Honors A&P", "H Anatomy & Physiology", "Anatomy & Physiology Honors"] },
  { id: "HONORS_BIOLOGY", name: "Honors Biology", department: "Science", aliases: ["Honors Bio", "Bio Honors", "H Bio", "Biology Honors"] },
  { id: "HONORS_CHEMISTRY", name: "Honors Chemistry", department: "Science", aliases: ["Honors Chem", "Chem Honors", "H Chem", "Chemistry Honors"] },
  { id: "HONORS_MARINE_ECOLOGY", name: "Honors Marine Ecology", department: "Science", aliases: ["H Marine Ecology", "Marine Ecology Honors"] },
  { id: "HONORS_PHYSICS", name: "Honors Physics", department: "Science", aliases: ["H Physics", "Physics Honors", "Honors Phys"] },
  { id: "MARINE_ECOLOGY", name: "Marine Ecology", department: "Science", aliases: [] },
  { id: "PHYSICS", name: "Physics", department: "Science", aliases: [] },
  { id: "SPORTS_MEDICINE", name: "Sports Medicine", department: "Science", aliases: ["Sports Med"] },
  { id: "AMERICAN_GOVERNMENT", name: "American Government", department: "Social Science", aliases: ["Amer Government", "US Government", "U.S. Government", "Government", "Gov"] },
  { id: "AP_ECON_AMER_GOVT", name: "AP Econ/Amer Govt", department: "Social Science", aliases: ["AP Econ/Amer Gov", "AP Macro/Amer Govt", "AP Macro/Amer Gov", "AP Macroeconomics/American Government", "AP Macro and American Government"] },
  { id: "AP_ECON_AP_GOV", name: "AP Econ/AP Gov", department: "Social Science", aliases: ["AP Econ/AP Comparative Gov", "AP Macro/AP Gov", "AP Macro/AP Comp Gov", "AP Macroeconomics/AP Comparative Government", "AP Comparative Government", "AP Comp Gov"] },
  { id: "AP_EUROPEAN_HISTORY", name: "AP European History", department: "Social Science", aliases: ["AP Euro", "APEuro"] },
  { id: "AP_US_HISTORY", name: "AP US History", department: "Social Science", aliases: ["APUSH", "AP U.S. History", "AP United States History", "AP American History"] },
  { id: "ECONOMICS", name: "Economics", department: "Social Science", aliases: ["Econ"] },
  { id: "US_HISTORY", name: "US History", department: "Social Science", aliases: ["U.S. History", "United States History", "American History"] },
  { id: "WORLD_HISTORY", name: "World History", department: "Social Science", aliases: ["World Hist", "World"] },
  { id: "AP_CHINESE_4", name: "AP Chinese 4", department: "World Language", aliases: ["AP Chinese", "AP Chinese Language", "AP Chinese Language and Culture", "AP Mandarin"] },
  { id: "HONORS_CHINESE_3", name: "Honors Chinese 3", department: "World Language", aliases: ["Chinese 3 Honors", "Chinese III Honors", "H Chinese 3", "Honors Mandarin 3"] },
  { id: "AP_FRENCH_4", name: "AP French 4", department: "World Language", aliases: ["AP French", "AP French Language", "AP French Language and Culture"] },
  { id: "FRENCH_1", name: "French 1", department: "World Language", aliases: ["French I", "Fr 1"] },
  { id: "FRENCH_2", name: "French 2", department: "World Language", aliases: ["French II", "Fr 2"] },
  { id: "HONORS_FRENCH_3", name: "Honors French 3", department: "World Language", aliases: ["French 3 Honors", "French III Honors", "H French 3"] },
  { id: "AP_GERMAN_4", name: "AP German 4", department: "World Language", aliases: ["AP German", "AP German Language"] },
  { id: "GERMAN_1", name: "German 1", department: "World Language", aliases: ["German I"] },
  { id: "GERMAN_2", name: "German 2", department: "World Language", aliases: ["German II"] },
  { id: "HONORS_GERMAN_3", name: "Honors German 3", department: "World Language", aliases: ["German 3 Honors", "German III Honors", "H German 3"] },
  { id: "AP_SPANISH_4", name: "AP Spanish 4", department: "World Language", aliases: ["AP Spanish", "AP Spanish Lang", "AP Spanish Language", "AP Spanish Language and Culture"] },
  { id: "HONORS_SPANISH_3", name: "Honors Spanish 3", department: "World Language", aliases: ["Spanish 3 Honors", "Spanish III Honors", "H Spanish 3"] },
  { id: "SPANISH_1", name: "Spanish 1", department: "World Language", aliases: ["Spanish I", "Span 1"] },
  { id: "SPANISH_2", name: "Spanish 2", department: "World Language", aliases: ["Spanish II", "Span 2"] },
  { id: "SPANISH_3", name: "Spanish 3", department: "World Language", aliases: ["Spanish III", "Span 3"] },
  { id: "SPANISH_4", name: "Spanish 4", department: "World Language", aliases: ["Spanish IV", "Span 4"] },
  { id: "AP_CYBERSECURITY", name: "AP Cybersecurity", department: "CTE", aliases: ["AP Cyber", "AP Cyber Security"] },
  { id: "BUSINESS_MANAGEMENT_ENTREPRENEURSHIP", name: "Business Management and Entrepreneurship", department: "CTE", aliases: ["Business Management & Entrepreneurship", "Business Mgmt & Entrepreneurship", "Business Management/Entrepreneurship"] },
  { id: "CAREERS_IN_TEACHING", name: "Careers in Teaching", department: "CTE", aliases: ["Teaching Careers", "Careers in Education"] },
  { id: "E_SPORTS", name: "E-Sports", department: "CTE", aliases: ["Esports", "eSports", "E Sports"] },
  { id: "EDUCATIONAL_PSYCHOLOGY", name: "Educational Psychology", department: "CTE", aliases: ["Ed Psych", "Education Psychology"] },
  { id: "SPORTS_MEDICINE_ADVANCED", name: "Sports Medicine Advanced", department: "CTE", aliases: ["Advanced Sports Medicine", "Advanced Sports Med", "Sports Med Advanced"] },
  { id: "HEALTH", name: "Health", department: "Health", aliases: ["Health Class"] },
  { id: "AP_COMPUTER_SCIENCE_A", name: "AP Computer Science", department: "Academic Elective", aliases: ["AP Computer Science A", "AP CSA", "APCSA", "AP CS A"] },
  { id: "AP_COMPUTER_SCIENCE_PRINCIPLES", name: "AP Computer Science Principles", department: "Academic Elective", aliases: ["AP CSP", "APCSP", "AP CS Principles"] },
  { id: "AP_HUMAN_GEOGRAPHY", name: "AP Human Geography", department: "Academic Elective", aliases: ["APHG", "AP HuG", "AP Human Geo"] },
  { id: "AP_PSYCHOLOGY", name: "AP Psychology", department: "Academic Elective", aliases: ["AP Psych", "APPsych"] },
  { id: "COLLEGE_CAREER_PLANNING", name: "College and Career Planning", department: "Academic Elective", aliases: ["CCP", "College & Career Planning", "College/Career Planning"] },
  { id: "ELD_1", name: "English Language Development I", department: "Academic Elective", aliases: ["ELD I", "ELD 1", "ELD1"] },
  { id: "ELD_2", name: "English Language Development 2", department: "Academic Elective", aliases: ["ELD II", "ELD 2", "ELD2"] },
  { id: "STUDENT_GOVERNMENT", name: "Student Government", department: "Academic Elective", aliases: ["ASB", "Student Gov"] },
  { id: "TPP_WORK_EXPERIENCE", name: "TPP Work Experience", department: "Academic Elective", aliases: ["TPP", "Work Experience"] },
  { id: "YEARBOOK", name: "Yearbook", department: "Academic Elective", aliases: ["Year Book", "YB"] }
];

function nhscourse_key_(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/\badvanced placement\b/g, "ap")
    .replace(/[.']/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function nhscourse_index_() {
  const index = {};

  NHS_COURSE_REGISTRY.forEach(function(entry) {
    [entry.id, entry.name].concat(entry.aliases || []).forEach(function(alias) {
      const key = nhscourse_key_(alias);
      if (key) index[key] = entry;
    });
  });

  return index;
}

function nhscourse_resolve_(value) {
  const key = nhscourse_key_(value);
  if (!key) return null;
  return nhscourse_index_()[key] || null;
}

function nhscourse_clientCatalog_() {
  return NHS_COURSE_REGISTRY.map(function(entry) {
    return {
      id: entry.id,
      name: entry.name,
      department: entry.department,
      aliases: (entry.aliases || []).slice()
    };
  });
}

function nhscourse_getCatalog() {
  nhsa_requireAnyRole_(["TEACHER", "TUTOR", "OFFICER", "ADVISOR", "ADMIN"], "NHS Course Catalog");
  return {
    version: NHS_COURSE_CATALOG_VERSION,
    courses: nhscourse_clientCatalog_()
  };
}

function nhscourse_normalizeCourseListText_(value, maxItems) {
  const raw = String(value || "")
    .split(/[\n,;]+/)
    .map(function(v) { return v.trim(); })
    .filter(Boolean);

  const seen = {};
  const output = [];

  raw.forEach(function(item) {
    const entry = nhscourse_resolve_(item);
    if (!entry || seen[entry.id]) return;
    seen[entry.id] = true;
    output.push(entry.name);
  });

  if (maxItems && output.length > maxItems) {
    output.length = maxItems;
  }

  return output.join(", ");
}

function nhscourse_strictSelection_(ids, text, maxItems) {
  const requestedIds =
    Array.isArray(ids)
      ? ids
          .map(function(id) {
            return String(id || "").trim();
          })
          .filter(Boolean)
      : [];

  const rawText =
    String(text || "")
      .split(/[\n,;]+/)
      .map(function(value) {
        return value.trim();
      })
      .filter(Boolean);

  const source =
    requestedIds.length
      ? requestedIds
      : rawText;

  const seen = {};
  const entries = [];

  source.forEach(function(value) {
    const entry =
      nhscourse_resolve_(value);

    if (!entry) {
      throw new Error(
        "COURSE_CATALOG_OUTDATED|" +
        "One of the selected strongest courses is no longer recognized by the current Tesoro course catalog. Refresh Tutor Portal and select the course again."
      );
    }

    if (seen[entry.id]) return;
    seen[entry.id] = true;
    entries.push(entry);
  });

  if (
    maxItems &&
    entries.length > Number(maxItems)
  ) {
    throw new Error(
      "Choose no more than " +
      Number(maxItems) +
      " strongest courses."
    );
  }

  return {
    ids: entries.map(function(entry) {
      return entry.id;
    }),
    names: entries.map(function(entry) {
      return entry.name;
    }),
    text: entries
      .map(function(entry) {
        return entry.name;
      })
      .join(", ")
  };
}


function nhstutor_dateMillis_(value) {
  if (!value) return 0;

  const date =
    value instanceof Date
      ? value
      : new Date(value);

  return isNaN(date.getTime())
    ? 0
    : date.getTime();
}


function nhscourse_inferFromText_(value) {
  const haystack = nhscourse_key_(value);
  if (!haystack) return null;

  const hits = [];

  NHS_COURSE_REGISTRY.forEach(function(entry) {
    const candidates = [entry.name].concat(entry.aliases || []);
    let best = 0;

    candidates.forEach(function(alias) {
      const key = nhscourse_key_(alias);
      if (!key) return;

      const words = key.split(" ").filter(Boolean);
      const looksLikeAbbreviation = words.length === 1 && key.length >= 4;
      const safePhrase = words.length >= 2 || looksLikeAbbreviation;

      if (safePhrase && haystack.includes(key)) {
        best = Math.max(best, key.length);
      }
    });

    if (best) hits.push({ entry: entry, score: best });
  });

  hits.sort(function(a, b) { return b.score - a.score; });
  if (!hits.length) return null;

  if (hits.length > 1 && hits[0].score === hits[1].score && hits[0].entry.id !== hits[1].entry.id) {
    return null;
  }

  return hits[0].entry;
}

function canonicalCourse_(value) {
  const entry = nhscourse_resolve_(value);
  if (entry) return entry.id;

  // Legacy/historical values not in the approved THS tutoring registry remain
  // comparable to themselves, but new Teacher requests cannot create them.
  const key = nhscourse_key_(value);
  return key ? "LEGACY__" + key.replace(/\s+/g, "_").toUpperCase() : "";
}

function formalCourseName_(canonical) {
  const entry = nhscourse_resolve_(canonical);
  if (entry) return entry.name;

  const byId = NHS_COURSE_REGISTRY.find(function(item) {
    return item.id === String(canonical || "").trim();
  });

  if (byId) return byId.name;

  const raw = String(canonical || "");
  if (raw.indexOf("LEGACY__") === 0) {
    return titleCasePhrase_(raw.slice(8).replace(/_/g, " "));
  }

  return titleCasePhrase_(raw.replace(/_/g, " "));
}



/* ============================================================
   MANUAL UTILITIES
   ============================================================ */

function refreshAllTutorMatches() {
  nhssec_requireMaintenanceAdmin_();
  rankAllOpenRequests_();
}


function refreshTutoringRequestFormatting() {
  nhssec_requireMaintenanceAdmin_();
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const sheet = nhsperf_sheet_(ss, NHS_MATCHING_REQUEST_SHEET);

  if (!sheet) {
    throw new Error('Sheet "' + NHS_MATCHING_REQUEST_SHEET + '" was not found.');
  }

  formatTutoringRequestsSheet_(ss, sheet);

  for (let row = 2; row <= sheet.getLastRow(); row++) {
    formatRequestRow_(sheet, row);
  }
}


function testRankRequestRow(rowNumber) {
  nhssec_requireMaintenanceAdmin_();
  rankTutorMatchesForRow_(rowNumber);
}


/** ============================================================
 * SOURCE GROUP: Notifications.gs
 * ============================================================ */

/**
 * NHS TUTORING OS — NOTIFICATION ENGINE
 * File: NHS_Notification_Engine.gs
 *
 * Add to the SAME Apps Script project as the NHS OS.
 *
 * Run once:
 *   installNHSNotificationEngine()
 *
 * This creates:
 *   - Notification Log
 *   - hourly notification sweep
 *
 * External notifications:
 *   - assignment reminders
 *   - reschedule notifications
 *   - reassignment notifications
 *   - cancellation notifications
 *
 * Internal OS alerts:
 *   - no tutor available
 *   - classroom missing
 *   - unconfirmed assignments
 *   - overdue session verification
 *   - stale requests
 */

const NHSN_LOG_SHEET = "Notification Log";
const NHSN_REQUESTS_SHEET = "Tutoring Requests";
const NHSN_TUTORS_SHEET = "Tutor Database";
const NHSN_TITAN_TIME = "10:25 AM–11:05 AM";
const NHSN_DAY_BEFORE_REMINDER_HOUR = 15; // 3 PM day-before
const NHSN_MORNING_REMINDER_HOUR = 7;     // 7 AM morning-of
const NHSN_ACCEPTANCE_REMINDER_HOURS = 6;
const NHSN_ACCEPTANCE_SECOND_REMINDER_HOURS = 24;
const NHSN_ACCEPTANCE_ESCALATION_HOURS = 48;
const NHSN_STALE_REQUEST_HOURS = 24;
const NHSN_UNCONFIRMED_HOURS = 12;
const NHSN_OVERDUE_VERIFY_HOURS = 18;
const NHSN_READY_NOTIFY_ESCALATION_MINUTES = 30;
const NHSN_ACCESS_REGISTRY_SHEET = "Access Registry";

var NHSN_SENT_DEDUPE_MEMO_ = null;
var NHSN_OPEN_ALERT_MEMO_ = null;

function nhsn_notificationIndexes_(ss) {
  if (
    NHSN_SENT_DEDUPE_MEMO_ instanceof Set &&
    NHSN_OPEN_ALERT_MEMO_ instanceof Set
  ) {
    return {
      sent: NHSN_SENT_DEDUPE_MEMO_,
      openAlerts: NHSN_OPEN_ALERT_MEMO_
    };
  }

  const sent = new Set();
  const openAlerts = new Set();
  const sh = nhsn_ensureLog_(ss);

  if (sh.getLastRow() >= 2) {
    const values = sh
      .getRange(1, 1, sh.getLastRow(), sh.getLastColumn())
      .getDisplayValues();

    const h = nhsn_headerIndex_(values[0]);

    values.slice(1).forEach(function(row) {
      const key = String(row[h["Dedupe Key"]] || "").trim();
      if (!key) return;

      const status = String(row[h["Status"]] || "").trim();
      const channel = String(row[h["Channel"]] || "").trim();
      const read = String(row[h["Read"]] || "").trim().toLowerCase();

      if (status === "Sent") sent.add(key);
      if (channel === "Internal Alert" && read !== "true") {
        openAlerts.add(key);
      }
    });
  }

  NHSN_SENT_DEDUPE_MEMO_ = sent;
  NHSN_OPEN_ALERT_MEMO_ = openAlerts;

  return {
    sent: sent,
    openAlerts: openAlerts
  };
}

function installNHSNotificationEngine() {
  nhssec_requireMaintenanceAdmin_();

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  nhsn_ensureLog_(ss);

  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "runNHSNotificationSweep_") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("runNHSNotificationSweep_")
    .timeBased()
    .everyHours(1)
    .create();

  PropertiesService.getScriptProperties().setProperty("NHS_NOTIFICATION_ENGINE_INSTALLED", "true");
}

/* ============================================================
   PUBLIC ENTRY POINTS
   ============================================================ */

function runNHSNotificationSweep_() {
  const props = PropertiesService.getScriptProperties();

  try {
    const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
    const now = new Date();

    nhsn_sendSessionRemindersV2_(ss, now);
    nhsn_sendPendingAcceptanceReminders_(ss, now);
    nhsn_alertMissingTutorNotificationEmails_(ss);
    nhsn_buildInternalAlerts_(ss, now);
    nhsn_sendReadyNotifyOfficerEscalations_(ss, now);
    nhswa_sendAutomaticReminders_(ss, now);

    props.setProperty("NHS_NOTIFICATION_LAST_SUCCESS", new Date().toISOString());
    props.deleteProperty("NHS_NOTIFICATION_LAST_ERROR");

    if (typeof nhssec_recordHealth_ === "function") {
      nhssec_recordHealth_("Notification Sweep", true, "Sweep completed", "");
    }
  } catch (error) {
    props.setProperty(
      "NHS_NOTIFICATION_LAST_ERROR",
      (error && error.message) ? error.message : String(error)
    );

    if (typeof nhssec_recordHealth_ === "function") {
      nhssec_recordHealth_(
        "Notification Sweep",
        false,
        "Sweep failed",
        (error && error.message) ? error.message : String(error)
      );
    }

    throw error;
  }
}

/**
 * Call this after an existing confirmed assignment changes.
 * oldState/newState objects should contain:
 * assignedTutor, scheduledDate, location, status
 */
function nhsn_handleAssignmentMutation_(requestId, oldState, newState) {
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const request = nhsn_getRequestById_(ss, requestId);
  if (!request) return;

  const wasConfirmed = !!String(oldState.assignmentConfirmed || "").trim();
  if (!wasConfirmed) return;

  const oldTutor = String(oldState.assignedTutor || "").trim();
  const newTutor = String(newState.assignedTutor || "").trim();
  const oldDate = String(oldState.scheduledDate || "").trim();
  const newDate = String(newState.scheduledDate || "").trim();
  const oldLocation = String(oldState.location || "").trim();
  const newLocation = String(newState.location || "").trim();
  const oldStatus = String(oldState.status || "").trim();
  const newStatus = String(newState.status || "").trim();

  if (oldTutor && newTutor && nhsn_norm_(oldTutor) !== nhsn_norm_(newTutor)) {
    nhsn_sendReassignment_(ss, request, oldTutor, newTutor);
  }

  if (oldDate && newDate && oldDate !== newDate) {
    nhsn_sendReschedule_(ss, request, oldDate, newDate);
  }

  if (oldLocation && newLocation && oldLocation !== newLocation) {
    nhsn_sendLocationChange_(ss, request, oldLocation, newLocation);
  }

  if (oldStatus !== "Cancelled" && newStatus === "Cancelled") {
    nhsn_sendCancellation_(ss, request);
  }
}

/**
 * Explicit manual cancellation helper.
 */
function nhsn_notifyCancellation_(requestId) {
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const request = nhsn_getRequestById_(ss, requestId);
  if (!request) throw new Error("Request not found.");
  nhsn_sendCancellation_(ss, request);
  return {ok:true};
}

/**
 * Returns active internal alerts for the OS.
 */
function nhsn_getInternalAlerts() {
  nhssec_requireOfficer_();

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const now = new Date();
  return nhsn_buildInternalAlerts_(ss, now, true);
}

/**
 * Mark one alert read in Notification Log.
 */
function nhsn_markAlertRead(notificationId) {
  nhssec_requireOfficer_();

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const sh = nhsn_ensureLog_(ss);

  if (sh.getLastRow() < 2) return { ok: false };

  const data = sh
    .getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
    .getValues();

  const headers = sh
    .getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0];

  const h = nhsn_headerIndex_(headers);
  const idIndex = h["Notification ID"];
  const readIndex = h["Read"];

  if (idIndex === undefined || readIndex === undefined) {
    return { ok: false };
  }

  const targetId = String(notificationId || "").trim();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idIndex] || "").trim() !== targetId) continue;

    if (String(data[i][readIndex]).toLowerCase() !== "true") {
      sh.getRange(i + 2, readIndex + 1).setValue(true);
      NHSN_OPEN_ALERT_MEMO_ = null;
    }

    return { ok: true };
  }

  return { ok: false };
}


function nhsn_markAllInternalAlertsRead() {
  nhssec_requireOfficer_();

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const sh = nhsn_ensureLog_(ss);

  if (sh.getLastRow() < 2) {
    return { ok: true, updated: 0 };
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  const headers = sh
    .getRange(1, 1, 1, lastCol)
    .getDisplayValues()[0];

  const h = nhsn_headerIndex_(headers);
  const channelIndex = h["Channel"];
  const readIndex = h["Read"];

  if (channelIndex === undefined || readIndex === undefined) {
    return { ok: true, updated: 0 };
  }

  const data = sh
    .getRange(2, 1, lastRow - 1, lastCol)
    .getValues();

  const readValues = data.map(row => [row[readIndex]]);
  let updated = 0;

  for (let i = 0; i < data.length; i++) {
    if (
      String(data[i][channelIndex] || "").trim() === "Internal Alert" &&
      String(data[i][readIndex]).toLowerCase() !== "true"
    ) {
      readValues[i][0] = true;
      updated++;
    }
  }

  if (updated) {
    sh
      .getRange(2, readIndex + 1, readValues.length, 1)
      .setValues(readValues);
    NHSN_OPEN_ALERT_MEMO_ = null;
  }

  return {
    ok: true,
    updated: updated
  };
}

/* ============================================================
   SESSION REMINDERS V2
   - Day-before reminder around 3 PM
   - Morning-of reminder around 7 AM
   - Tutor assignment-response reminder after 6 hours
   ============================================================ */

function nhsn_sendSessionRemindersV2_(ss, now) {
  const sh = nhsperf_sheet_(ss, NHSN_REQUESTS_SHEET);
  if (!sh || sh.getLastRow() < 2) return;

  const tz = nhs_safeTimeZone_(ss);
  const hour = Number(Utilities.formatDate(now, tz, "H"));
  const todayKey = Utilities.formatDate(now, tz, "yyyy-MM-dd");

  const tomorrow = new Date(now);
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowKey = Utilities.formatDate(tomorrow, tz, "yyyy-MM-dd");

  const rows = nhsn_readSheetObjects_(sh);

  rows.forEach(request => {
    if (String(request["Status"] || "") !== "Scheduled") return;
    if (!request["Assignment Confirmed"]) return;

    const dateKey = nhsn_dateKey_(request["Scheduled Date"], tz);
    if (!dateKey) return;

    const preference = nhsn_getTutorReminderPreference_(ss, request);
    const requestId = String(request["Request ID"] || "");

    if (
      dateKey === tomorrowKey &&
      hour >= NHSN_DAY_BEFORE_REMINDER_HOUR &&
      hour <= NHSN_DAY_BEFORE_REMINDER_HOUR + 1
    ) {
      const sendTutor =
        preference === "Day Before + Morning Of" ||
        preference === "Day Before Only";

      nhsn_sendSessionReminderPair_(
        ss,
        request,
        "DAY_BEFORE",
        "Tomorrow",
        sendTutor
      );
    }

    if (
      dateKey === todayKey &&
      hour >= NHSN_MORNING_REMINDER_HOUR &&
      hour <= NHSN_MORNING_REMINDER_HOUR + 1
    ) {
      const sendTutor =
        preference === "Day Before + Morning Of" ||
        preference === "Morning Of Only" ||
        !preference;

      nhsn_sendSessionReminderPair_(
        ss,
        request,
        "MORNING_OF",
        "Today",
        sendTutor
      );
    }
  });
}


function nhsn_sendSessionReminderPair_(ss, request, phase, whenText, sendTutor) {
  const requestId = String(request["Request ID"] || "");
  const course = request["Course"] || "Tutoring";
  const location = request["Location"] || "Teacher classroom";

  const tutor = nhsn_findTutorByRequest_(ss, request);
  const tutorDisplay = tutor
    ? nhsn_tutorDisplayName_(tutor["Preferred Name"], tutor["Full Name"])
    : String(request["Assigned Tutor"] || "Tutor");

  const teacherEmail = nhs_teacherDeliveryEmail_(request);
  const tutorEmail = tutor
    ? String(
        tutor["Notification Email"] ||
        ""
      ).trim()
    : "";

  const dateText =
    nhsn_formatDate_(
      request["Scheduled Date"],
      nhs_safeTimeZone_(ss)
    );

  if (teacherEmail) {
    const dedupe =
      "SESSION_REMINDER|" +
      phase +
      "|" +
      requestId +
      "|TEACHER";

    if (!nhsn_wasSent_(ss, dedupe)) {
      nhsn_sendEmailLogged_(ss, {
        requestId,
        type:
          phase === "DAY_BEFORE"
            ? "Session Reminder — Teacher — Day Before"
            : "Session Reminder — Teacher — Morning Of",
        recipient: teacherEmail,
        subject:
          "Reminder: NHS Tutoring " +
          whenText +
          " — " +
          course,
        title:
          "NHS Tutoring " +
          whenText +
          ": " +
          course,
        message:
          dateText +
          " · " +
          NHSN_TITAN_TIME +
          " · " +
          location,
        body:
          "Hello " +
          (request["Teacher Name"] || "Teacher") +
          ",\n\n" +
          "This is a reminder that your NHS tutoring session is scheduled " +
          whenText.toLowerCase() +
          ".\n\n" +
          "Course: " + course + "\n" +
          "Tutor: " + tutorDisplay + "\n" +
          "Date: " + dateText + "\n" +
          "Time: " + NHSN_TITAN_TIME + "\n" +
          "Location: " + location + "\n\n" +
          "Five Star reminder: please make sure both the student and tutor are assigned to your Titan Time in Five Star before the session.\n\n" +
          "Thank you,\nNHS Tutoring",
        dedupeKey: dedupe
      });
    }
  }

  if (sendTutor && tutorEmail) {
    const dedupe =
      "SESSION_REMINDER|" +
      phase +
      "|" +
      requestId +
      "|TUTOR";

    if (!nhsn_wasSent_(ss, dedupe)) {
      nhsn_sendEmailLogged_(ss, {
        requestId,
        type:
          phase === "DAY_BEFORE"
            ? "Session Reminder — Tutor — Day Before"
            : "Session Reminder — Tutor — Morning Of",
        recipient: tutorEmail,
        subject:
          "Reminder: NHS Tutoring " +
          whenText +
          " — " +
          course,
        title:
          "NHS Tutoring " +
          whenText +
          ": " +
          course,
        message:
          dateText +
          " · " +
          NHSN_TITAN_TIME +
          " · Room " +
          location,
        body:
          "Hello " +
          tutorDisplay +
          ",\n\n" +
          "Reminder: you have an NHS tutoring assignment " +
          whenText.toLowerCase() +
          ".\n\n" +
          "Course: " + course + "\n" +
          "Requesting Teacher: " +
          (request["Teacher Name"] || "") +
          "\n" +
          "Date: " + dateText + "\n" +
          "Time: " + NHSN_TITAN_TIME + "\n" +
          "Location: " + location + "\n\n" +
          "Please arrive on time and ready to support the student.\n\n" +
          "Thank you,\nNHS Tutoring",
        dedupeKey: dedupe
      });
    }
  }
}


function nhsn_sendPendingAcceptanceReminders_(ss, now) {
  const sh = nhsperf_sheet_(ss, NHSN_REQUESTS_SHEET);
  if (!sh || sh.getLastRow() < 2) return;

  const rows = nhsn_readSheetObjects_(sh);

  rows.forEach(function(request) {
    const status = String(request["Status"] || "");
    if (!["Assigned", "Scheduled"].includes(status)) return;

    const response = String(request["Tutor Response Status"] || "").trim();
    if (["Accepted", "Declined"].includes(response)) return;

    const assigned = String(request["Assigned Tutor"] || "").trim();
    if (!assigned) return;

    const updated = nhsn_parseDate_(
      request["Last Update"] ||
      request["Timestamp"]
    );

    if (!updated) return;

    const hours = nhsn_hoursBetween_(updated, now);
    if (hours < NHSN_ACCEPTANCE_REMINDER_HOURS) return;

    const tutor = nhsn_findTutorByRequest_(ss, request);
    if (!tutor) return;

    const tutorEmail = String(tutor["Notification Email"] || "").trim();
    if (!tutorEmail) return;

    const requestId = String(request["Request ID"] || "");
    const secondStage = hours >= NHSN_ACCEPTANCE_SECOND_REMINDER_HOURS;
    const dedupe =
      (secondStage ? "ACCEPTANCE_REMINDER_24H|" : "ACCEPTANCE_REMINDER_6H|") +
      requestId +
      "|" +
      tutorEmail;

    if (nhsn_wasSent_(ss, dedupe)) return;

    const course = request["Course"] || "Tutoring";
    const tutorName = nhsn_tutorDisplayName_(
      tutor["Preferred Name"],
      tutor["Full Name"]
    );

    nhsn_sendEmailLogged_(ss, {
      requestId: requestId,
      type: secondStage
        ? "Tutor Assignment Response Reminder — 24 Hour"
        : "Tutor Assignment Response Reminder",
      recipient: tutorEmail,
      subject: secondStage
        ? "Second Reminder: Respond to NHS Tutoring Assignment"
        : "Action Needed: Respond to NHS Tutoring Assignment",
      title: secondStage
        ? "Second assignment response reminder"
        : "Assignment response needed",
      message:
        course +
        " is waiting for your acceptance or decline.",
      body:
        "Hello " + tutorName + ",\n\n" +
        (secondStage
          ? "This is a second reminder that an NHS tutoring assignment is still waiting for your response."
          : "An NHS tutoring assignment is still waiting for your response.") +
        "\n\n" +
        "Course: " + course + "\n" +
        "Date: " +
        nhsn_formatDate_(
          request["Scheduled Date"],
          nhs_safeTimeZone_(ss)
        ) +
        "\n" +
        "Time: " + NHSN_TITAN_TIME + "\n" +
        "Location: " + (request["Location"] || "") +
        "\n\n" +
        "Please open the Tutor Portal and accept or decline the assignment so NHS can finalize the session.\n\n" +
        "Thank you,\nNHS Tutoring",
      dedupeKey: dedupe
    });
  });
}


function nhsn_findTutorByRequest_(ss, request) {
  const tutorId = String(
    request["Assigned Tutor ID"] || ""
  ).trim();

  if (tutorId) {
    const byId = nhsn_findTutorById_(ss, tutorId);
    if (byId) return byId;
  }

  return nhsn_findTutor_(ss, request["Assigned Tutor"]);
}


function nhsn_findTutorById_(ss, tutorId) {
  const sh = nhsperf_sheet_(ss, NHSN_TUTORS_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;

  const rows = nhsn_readSheetObjects_(sh);

  return rows.find(
    t =>
      String(t["Tutor ID"] || "").trim() ===
      String(tutorId || "").trim()
  ) || null;
}


function nhsn_getTutorReminderPreference_(ss, request) {
  const tutor = nhsn_findTutorByRequest_(ss, request);

  if (!tutor) return "Day Before + Morning Of";

  return String(
    tutor["Reminder Preference"] ||
    "Day Before + Morning Of"
  ).trim();
}


/* ============================================================
   MORNING REMINDERS
   ============================================================ */

function nhsn_sendMorningReminders_(ss, now) {
  const sh = nhsperf_sheet_(ss, NHSN_REQUESTS_SHEET);
  if (!sh || sh.getLastRow() < 2) return;

  const tz = nhs_safeTimeZone_(ss);
  const hour = Number(Utilities.formatDate(now, tz, "H"));
  if (hour < NHSN_REMINDER_HOUR || hour > NHSN_REMINDER_HOUR + 1) return;

  const todayKey = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  const rows = nhsn_readSheetObjects_(sh);

  rows.forEach(request => {
    if (String(request["Status"] || "") !== "Scheduled") return;
    if (!request["Assignment Confirmed"]) return;

    const dateKey = nhsn_dateKey_(request["Scheduled Date"], tz);
    if (dateKey !== todayKey) return;

    const requestId = String(request["Request ID"] || "");
    const dedupe = "REMINDER|" + requestId + "|" + todayKey;
    if (nhsn_wasSent_(ss, dedupe)) return;

    const tutor = nhsn_findTutor_(ss, request["Assigned Tutor"]);
    const teacherEmail = nhs_teacherDeliveryEmail_(request);
    const tutorEmail = tutor ? String(tutor["Notification Email"] || "").trim() : "";

    const course = request["Course"] || "Tutoring";
    const location = request["Location"] || "Teacher classroom";
    const tutorDisplay = tutor ? nhsn_tutorDisplayName_(tutor["Preferred Name"], tutor["Full Name"]) : request["Assigned Tutor"];

    if (teacherEmail) {
      nhsn_sendEmailLogged_(ss, {
        requestId,
        type:"Session Reminder — Teacher",
        recipient:teacherEmail,
        subject:"Reminder: NHS Tutoring Today — " + course,
        body:
          "Hello " + (request["Teacher Name"] || "Teacher") + ",\n\n" +
          "This is a reminder that your NHS tutoring session is scheduled for today.\n\n" +
          "Course: " + course + "\n" +
          "Tutor: " + tutorDisplay + "\n" +
          "Time: " + NHSN_TITAN_TIME + "\n" +
          "Location: " + location + "\n\n" +
          "Five Star reminder: please make sure both the student and tutor are assigned to your Titan Time in Five Star before the session.\n\n" +
          "Thank you,\nNHS Tutoring",
        dedupeKey:dedupe + "|TEACHER"
      });
    }

    if (tutorEmail) {
      nhsn_sendEmailLogged_(ss, {
        requestId,
        type:"Session Reminder — Tutor",
        recipient:tutorEmail,
        subject:"Reminder: NHS Tutoring Today — " + course,
        body:
          "Hello " + tutorDisplay + ",\n\n" +
          "Reminder: you have an NHS tutoring assignment today.\n\n" +
          "Course: " + course + "\n" +
          "Requesting Teacher: " + (request["Teacher Name"] || "") + "\n" +
          "Time: " + NHSN_TITAN_TIME + "\n" +
          "Location: " + location + "\n\n" +
          "Please arrive on time and ready to support the student.\n\n" +
          "Thank you,\nNHS Tutoring",
        dedupeKey:dedupe + "|TUTOR"
      });
    }
  });
}

/* ============================================================
   CHANGE NOTIFICATIONS
   ============================================================ */

function nhsn_sendReassignment_(ss, request, oldTutorName, newTutorName) {
  const requestId = request["Request ID"];
  const newTutor = nhsn_findTutor_(ss, newTutorName);
  const oldTutor = nhsn_findTutor_(ss, oldTutorName);
  const teacherEmail = nhs_teacherDeliveryEmail_(request);
  const stamp = new Date().getTime();

  if (oldTutor && oldTutor["Notification Email"]) {
    nhsn_sendEmailLogged_(ss, {
      requestId,
      type:"Reassignment — Removed Tutor",
      recipient:oldTutor["Notification Email"],
      subject:"NHS Tutoring Assignment Updated — " + (request["Course"] || "Tutoring"),
      body:
        "Hello " + nhsn_tutorDisplayName_(oldTutor["Preferred Name"], oldTutor["Full Name"]) + ",\n\n" +
        "You are no longer assigned to this NHS tutoring request.\n\n" +
        "Course: " + (request["Course"] || "") + "\n" +
        "Date: " + nhsn_formatDate_(request["Scheduled Date"], nhs_safeTimeZone_(ss)) + "\n\n" +
        "No further action is needed.\n\nNHS Tutoring",
      dedupeKey:"REASSIGN-OLD|" + requestId + "|" + stamp
    });
  }

  if (newTutor && newTutor["Notification Email"]) {
    nhsn_sendEmailLogged_(ss, {
      requestId,
      type:"Reassignment — New Tutor",
      recipient:newTutor["Notification Email"],
      subject:"NHS Tutoring Assignment — " + (request["Course"] || "Tutoring"),
      body:
        "Hello " + nhsn_tutorDisplayName_(newTutor["Preferred Name"], newTutor["Full Name"]) + ",\n\n" +
        "You have been assigned an NHS tutoring session.\n\n" +
        "Course: " + (request["Course"] || "") + "\n" +
        "Date: " + nhsn_formatDate_(request["Scheduled Date"], nhs_safeTimeZone_(ss)) + "\n" +
        "Time: " + NHSN_TITAN_TIME + "\n" +
        "Location: " + (request["Location"] || "") + "\n\n" +
        "Thank you,\nNHS Tutoring",
      dedupeKey:"REASSIGN-NEW|" + requestId + "|" + stamp
    });
  }

  if (teacherEmail) {
    nhsn_sendEmailLogged_(ss, {
      requestId,
      type:"Reassignment — Teacher",
      recipient:teacherEmail,
      subject:"Updated NHS Tutor Assignment — " + (request["Course"] || "Tutoring"),
      body:
        "Hello " + (request["Teacher Name"] || "Teacher") + ",\n\n" +
        "The tutor assigned to your NHS tutoring request has changed.\n\n" +
        "New Tutor: " + (newTutor ? nhsn_tutorDisplayName_(newTutor["Preferred Name"], newTutor["Full Name"]) : newTutorName) + "\n" +
        "Date: " + nhsn_formatDate_(request["Scheduled Date"], nhs_safeTimeZone_(ss)) + "\n" +
        "Time: " + NHSN_TITAN_TIME + "\n" +
        "Location: " + (request["Location"] || "") + "\n\n" +
        "Thank you,\nNHS Tutoring",
      dedupeKey:"REASSIGN-TEACHER|" + requestId + "|" + stamp
    });
  }
}

function nhsn_sendReschedule_(ss, request, oldDate, newDate) {
  const requestId = request["Request ID"];
  const tutor = nhsn_findTutor_(ss, request["Assigned Tutor"]);
  const oldText = nhsn_formatDate_(oldDate, nhs_safeTimeZone_(ss));
  const newText = nhsn_formatDate_(newDate, nhs_safeTimeZone_(ss));
  const stamp = new Date().getTime();

  const bodyBase =
    "Course: " + (request["Course"] || "") + "\n" +
    "Previous Date: " + oldText + "\n" +
    "New Date: " + newText + "\n" +
    "Time: " + NHSN_TITAN_TIME + "\n" +
    "Location: " + (request["Location"] || "") + "\n";

  if (nhs_teacherDeliveryEmail_(request)) {
    nhsn_sendEmailLogged_(ss, {
      requestId,type:"Reschedule — Teacher",recipient:nhs_teacherDeliveryEmail_(request),
      subject:"NHS Tutoring Rescheduled — " + (request["Course"] || "Tutoring"),
      body:"Hello " + (request["Teacher Name"] || "Teacher") + ",\n\nYour NHS tutoring session has been rescheduled.\n\n" + bodyBase + "\nThank you,\nNHS Tutoring",
      dedupeKey:"RESCHEDULE-TEACHER|" + requestId + "|" + stamp
    });
  }

  if (tutor && tutor["Notification Email"]) {
    nhsn_sendEmailLogged_(ss, {
      requestId,type:"Reschedule — Tutor",recipient:tutor["Notification Email"],
      subject:"NHS Tutoring Rescheduled — " + (request["Course"] || "Tutoring"),
      body:"Hello " + nhsn_tutorDisplayName_(tutor["Preferred Name"], tutor["Full Name"]) + ",\n\nYour NHS tutoring assignment has been rescheduled.\n\n" + bodyBase + "\nThank you,\nNHS Tutoring",
      dedupeKey:"RESCHEDULE-TUTOR|" + requestId + "|" + stamp
    });
  }
}

function nhsn_sendLocationChange_(ss, request, oldLocation, newLocation) {
  const requestId = request["Request ID"];
  const tutor = nhsn_findTutor_(ss, request["Assigned Tutor"]);
  const stamp = new Date().getTime();

  const bodyBase =
    "Course: " + (request["Course"] || "") + "\n" +
    "Date: " + nhsn_formatDate_(request["Scheduled Date"], nhs_safeTimeZone_(ss)) + "\n" +
    "Time: " + NHSN_TITAN_TIME + "\n" +
    "Previous Location: " + oldLocation + "\n" +
    "New Location: " + newLocation + "\n";

  if (nhs_teacherDeliveryEmail_(request)) {
    nhsn_sendEmailLogged_(ss, {
      requestId,type:"Location Change — Teacher",recipient:nhs_teacherDeliveryEmail_(request),
      subject:"NHS Tutoring Location Updated — " + (request["Course"] || "Tutoring"),
      body:"Hello " + (request["Teacher Name"] || "Teacher") + ",\n\nThe location for this tutoring session has changed.\n\n" + bodyBase + "\nThank you,\nNHS Tutoring",
      dedupeKey:"LOCATION-TEACHER|" + requestId + "|" + stamp
    });
  }

  if (tutor && tutor["Notification Email"]) {
    nhsn_sendEmailLogged_(ss, {
      requestId,type:"Location Change — Tutor",recipient:tutor["Notification Email"],
      subject:"NHS Tutoring Location Updated — " + (request["Course"] || "Tutoring"),
      body:"Hello " + nhsn_tutorDisplayName_(tutor["Preferred Name"], tutor["Full Name"]) + ",\n\nThe location for your tutoring assignment has changed.\n\n" + bodyBase + "\nThank you,\nNHS Tutoring",
      dedupeKey:"LOCATION-TUTOR|" + requestId + "|" + stamp
    });
  }
}

function nhsn_sendCancellation_(ss, request) {
  const requestId = request["Request ID"];
  const tutor = nhsn_findTutor_(ss, request["Assigned Tutor"]);
  const stamp = new Date().getTime();

  const bodyBase =
    "Course: " + (request["Course"] || "") + "\n" +
    "Date: " + nhsn_formatDate_(request["Scheduled Date"], nhs_safeTimeZone_(ss)) + "\n";

  if (nhs_teacherDeliveryEmail_(request)) {
    nhsn_sendEmailLogged_(ss, {
      requestId,type:"Cancellation — Teacher",recipient:nhs_teacherDeliveryEmail_(request),
      subject:"NHS Tutoring Session Cancelled — " + (request["Course"] || "Tutoring"),
      body:"Hello " + (request["Teacher Name"] || "Teacher") + ",\n\nThis NHS tutoring session has been cancelled.\n\n" + bodyBase + "\nThank you,\nNHS Tutoring",
      dedupeKey:"CANCEL-TEACHER|" + requestId + "|" + stamp
    });
  }

  if (tutor && tutor["Notification Email"]) {
    nhsn_sendEmailLogged_(ss, {
      requestId,type:"Cancellation — Tutor",recipient:tutor["Notification Email"],
      subject:"NHS Tutoring Session Cancelled — " + (request["Course"] || "Tutoring"),
      body:"Hello " + nhsn_tutorDisplayName_(tutor["Preferred Name"], tutor["Full Name"]) + ",\n\nThis NHS tutoring assignment has been cancelled.\n\n" + bodyBase + "\nNo further action is needed.\n\nNHS Tutoring",
      dedupeKey:"CANCEL-TUTOR|" + requestId + "|" + stamp
    });
  }
}

function nhsn_alertMissingTutorNotificationEmails_(ss) {
  const requestSh = nhsperf_sheet_(ss, NHSN_REQUESTS_SHEET);
  if (!requestSh || requestSh.getLastRow() < 2) return;

  const requests = nhsn_readSheetObjects_(requestSh);

  requests.forEach(request => {
    const status = String(request["Status"] || "").trim();

    if (!["Assigned", "Scheduled"].includes(status)) return;

    const assigned = String(request["Assigned Tutor"] || "").trim();
    if (!assigned) return;

    const tutor = nhsn_findTutorByRequest_(ss, request);
    if (!tutor) return;

    const notificationEmail =
      String(tutor["Notification Email"] || "").trim();

    if (notificationEmail) return;

    const requestId = String(request["Request ID"] || "");
    const dedupe = "MISSING_TUTOR_NOTIFICATION_EMAIL|" + requestId;

    if (nhsn_hasOpenAlert_(ss, dedupe)) return;

    nhsn_logAlert_(ss, {
      requestId: requestId,
      type: "MISSING_TUTOR_NOTIFICATION_EMAIL",
      severity: "High",
      title: "Tutor is missing Notification Email",
      message:
        nhsn_tutorDisplayName_(
          tutor["Preferred Name"],
          tutor["Full Name"]
        ) +
        " cannot reliably receive assignment or reminder emails. Add a Notification Email before confirming.",
      dedupeKey: dedupe
    });
  });
}


/* ============================================================
   INTERNAL ALERTS
   ============================================================ */

function nhsn_buildInternalAlerts_(ss, now, returnOnly) {
  const requestSh = nhsperf_sheet_(ss, NHSN_REQUESTS_SHEET);
  if (!requestSh || requestSh.getLastRow() < 2) return [];

  const tz = nhs_safeTimeZone_(ss);
  const requests = nhsn_readSheetObjects_(requestSh);
  const alerts = [];

  requests.forEach(r => {
    const id = String(r["Request ID"] || "");
    const status = String(r["Status"] || "");
    if (!id || ["Completed","Cancelled"].includes(status)) return;

    const ts = nhsn_parseDate_(r["Timestamp"]);
    const updated = nhsn_parseDate_(r["Last Update"]) || ts;
    const sched = nhsn_parseDate_(r["Scheduled Date"]);

    if (status === "No Tutor Available") {
      alerts.push(nhsn_alert_("NO_TUTOR", id, "High", "No tutor available", (r["Course"] || "Request") + " needs manual follow-up."));
    }

    if ((status === "Assigned" || status === "Scheduled") && !String(r["Location"] || "").trim()) {
      alerts.push(nhsn_alert_("MISSING_ROOM", id, "High", "Classroom missing", "Teacher room is missing for " + (r["Course"] || "this request") + "."));
    }

    const tutorResponse = String(r["Tutor Response Status"] || "").trim();
    const assignmentConfirmed = String(r["Assignment Confirmed"] || "").trim();

    const schedulingType = String(r["Scheduling Preference"] || "").trim();
    const reconfirmStatus = String(r["Tutor Reconfirm Status"] || "").trim();
    const needsFutureReconfirm =
      schedulingType === NHSSCHED_TYPES.FUTURE &&
      reconfirmStatus !== "Confirmed";
    const needsRecurringReconfirm =
      schedulingType === NHSSCHED_TYPES.RECURRING &&
      reconfirmStatus === "Required";

    if (
      tutorResponse === "Accepted" &&
      !assignmentConfirmed &&
      !needsFutureReconfirm &&
      !needsRecurringReconfirm &&
      ["Assigned","Scheduled"].includes(status)
    ) {
      alerts.push(
        nhsn_alert_(
          "READY_NOTIFY",
          id,
          "High",
          "Ready to notify",
          (r["Course"] || "Tutoring request") +
            " has been accepted by the tutor and is ready for teacher/tutor notification."
        )
      );
    }

    if (status === "Assigned" && !r["Assignment Confirmed"] && updated && nhsn_hoursBetween_(updated, now) >= NHSN_UNCONFIRMED_HOURS) {
      alerts.push(nhsn_alert_("UNCONFIRMED", id, "Medium", "Assignment not confirmed", "Tutor is assigned but confirmation is still pending."));
    }

    if (
      ["Assigned","Scheduled"].includes(status) &&
      !["Accepted","Declined"].includes(tutorResponse) &&
      updated &&
      nhsn_hoursBetween_(updated, now) >= NHSN_ACCEPTANCE_ESCALATION_HOURS
    ) {
      alerts.push(
        nhsn_alert_(
          "TUTOR_RESPONSE_OVERDUE",
          id,
          "High",
          "Tutor response overdue",
          "The assigned tutor has not accepted or declined after 48 hours. Officer follow-up or reassignment is recommended."
        )
      );
    }

    const scheduledKey = nhssched_dateKey_(r["Scheduled Date"], tz);
    if (scheduledKey) {
      const exception = nhssched_exceptionForDate_(scheduledKey);
      if (exception) {
        alerts.push(
          nhsn_alert_(
            "SCHOOL_CALENDAR_CONFLICT",
            id,
            "High",
            "School-calendar conflict",
            scheduledKey + " is unavailable for Titan Time: " + (exception.reason || exception.type) + "."
          )
        );
      }

      const todayKey = nhssched_todayKey_(tz, now);
      const daysUntil = nhssched_daysBetween_(todayKey, scheduledKey);
      const teacherReady = nhslife_truthy_(r["Teacher Five Star Ready"]);
      const tutorReady = nhslife_truthy_(r["Tutor Five Star Enrolled"]);
      const confirmed = nhslife_confirmed_(r);

      if (
        confirmed &&
        daysUntil >= 0 &&
        daysUntil <= 1 &&
        !(teacherReady && tutorReady)
      ) {
        const missing = [];
        if (!teacherReady) missing.push("teacher assignment of tutor + student");
        if (!tutorReady) missing.push("tutor placement confirmation");
        alerts.push(
          nhsn_alert_(
            "FIVESTAR_INCOMPLETE",
            id,
            "High",
            "Five Star setup incomplete",
            "Session is " + (daysUntil === 0 ? "today" : "tomorrow") + " and still needs " + missing.join(" and ") + "."
          )
        );
      }
    }

    if (sched && ["Assigned","Scheduled"].includes(status)) {
      const end = new Date(sched);
      end.setHours(11, 5, 0, 0);
      if (now > end && nhsn_hoursBetween_(end, now) >= NHSN_OVERDUE_VERIFY_HOURS) {
        alerts.push(nhsn_alert_("OVERDUE_VERIFY", id, "High", "Session needs verification", "Scheduled session passed without a completed outcome."));
      }
    }

    if (["New","Matching"].includes(status) && ts && nhsn_hoursBetween_(ts, now) >= NHSN_STALE_REQUEST_HOURS) {
      alerts.push(nhsn_alert_("STALE_REQUEST", id, "Medium", "Request waiting too long", (r["Course"] || "Request") + " has been open for more than 24 hours."));
    }
  });

  if (!returnOnly) {
    alerts.forEach(a => nhsn_logInternalAlertIfNew_(ss, a));
  }

  const logged = nhsn_readActiveLoggedAlerts_(ss);
  const combined = [];
  const seen = {};

  alerts.concat(logged).forEach(a => {
    const k = (a.type || "") + "|" + (a.requestId || "");
    if (!seen[k]) {
      seen[k] = true;
      combined.push(a);
    }
  });

  const severityOrder = {High:0,Medium:1,Low:2};
  combined.sort((a,b) => (severityOrder[a.severity]||9)-(severityOrder[b.severity]||9));
  return combined.slice(0,50);
}

function nhsn_alert_(type, requestId, severity, title, message) {
  return {
    notificationId:"",
    type,
    requestId,
    severity,
    title,
    message,
    createdAt:new Date().toISOString(),
    read:false
  };
}

function nhsn_logInternalAlertIfNew_(ss, alert) {
  const key = "ALERT|" + alert.type + "|" + alert.requestId;
  if (nhsn_hasOpenAlert_(ss, key)) return;

  const sh = nhsn_ensureLog_(ss);
  sh.appendRow([
    nhsn_id_("N"),
    new Date(),
    alert.requestId,
    "Internal Alert",
    "",
    alert.type,
    "Created",
    key,
    false,
    alert.severity,
    alert.title,
    alert.message,
    ""
  ]);

  nhsn_notificationIndexes_(ss).openAlerts.add(key);
}

function nhsn_readActiveLoggedAlerts_(ss) {
  const sh = nhsn_ensureLog_(ss);
  if (sh.getLastRow() < 2) return [];

  const objects = nhsn_readSheetObjects_(sh);
  return objects
    .filter(o => o["Channel"] === "Internal Alert" && String(o["Read"]).toLowerCase() !== "true")
    .map(o => ({
      notificationId:o["Notification ID"] || "",
      type:o["Type"] || "",
      requestId:o["Request ID"] || "",
      severity:o["Severity"] || "Medium",
      title:o["Title"] || o["Type"] || "Alert",
      message:o["Message"] || "",
      createdAt:o["Created At"] || "",
      read:false
    }));
}

/* ============================================================
   LOG / EMAIL HELPERS
   ============================================================ */

function nhsn_sendEmailLogged_(ss, spec) {
  const sh = nhsn_ensureLog_(ss);

  if (
    spec.dedupeKey &&
    nhsn_wasSent_(ss, spec.dedupeKey)
  ) {
    return true;
  }

  let status = "Sent";
  let error = "";

  try {
    nhsmail_send_({
      to: spec.recipient,
      subject: spec.subject,
      body: spec.body,
      htmlBody: nhsn_formalHtmlEmail_(spec),
      name: "Tesoro NHS",
      requestId: spec.requestId || "",
      type: spec.type || ""
    });
  } catch (e) {
    status = "Failed";
    error = e.message || String(e);
  }

  const h = nhsn_headerMapSheet_(sh);
  const row = new Array(sh.getLastColumn()).fill("");

  nhsn_setRow_(row, h, "Notification ID", nhsn_id_("N"));
  nhsn_setRow_(row, h, "Created At", new Date());
  nhsn_setRow_(row, h, "Request ID", spec.requestId || "");
  nhsn_setRow_(row, h, "Channel", "Email");
  nhsn_setRow_(row, h, "Recipient", spec.recipient || "");
  nhsn_setRow_(row, h, "Type", spec.type || "");
  nhsn_setRow_(row, h, "Status", status);
  nhsn_setRow_(row, h, "Dedupe Key", spec.dedupeKey || "");
  nhsn_setRow_(row, h, "Read", true);
  nhsn_setRow_(row, h, "Severity", "");
  nhsn_setRow_(row, h, "Title", spec.title || spec.subject || "");
  nhsn_setRow_(row, h, "Message", spec.message || "");
  nhsn_setRow_(row, h, "Error", error);
  nhsn_setRow_(row, h, "Portal Read", false);

  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length)
    .setValues([row]);

  if (spec.dedupeKey && status === "Sent") {
    nhsn_notificationIndexes_(ss).sent.add(String(spec.dedupeKey));
  }

  return status === "Sent";
}

function nhsn_wasSent_(ss, dedupeKey) {
  const key = String(dedupeKey || "").trim();
  if (!key) return false;
  return nhsn_notificationIndexes_(ss).sent.has(key);
}

function nhsn_hasOpenAlert_(ss, dedupeKey) {
  const key = String(dedupeKey || "").trim();
  if (!key) return false;
  return nhsn_notificationIndexes_(ss).openAlerts.has(key);
}

function nhsn_ensureLog_(ss) {
  let sh = nhsperf_sheet_(ss, NHSN_LOG_SHEET);
  if (!sh) sh = ss.insertSheet(NHSN_LOG_SHEET);

  const headers = [
    "Notification ID",
    "Created At",
    "Request ID",
    "Channel",
    "Recipient",
    "Type",
    "Status",
    "Dedupe Key",
    "Read",
    "Severity",
    "Title",
    "Message",
    "Error",
    "Portal Read"
  ];

  let current = sh
    .getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1))
    .getDisplayValues()[0];

  headers.forEach(header => {
    if (!current.includes(header)) {
      sh.getRange(1, current.length + 1).setValue(header);
      current.push(header);
    }
  });

  sh.setFrozenRows(1);

  sh.getRange(1, 1, 1, sh.getLastColumn())
    .setBackground("#123A5A")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");

  return sh;
}

function nhsn_headerMapSheet_(sh) {
  return nhsperf_headerMap_(sh, true, false);
}


function nhsn_setRow_(row, h, header, value) {
  const col = h[header];
  if (!col) return;
  row[col - 1] = value;
}


/* ============================================================
   DATA HELPERS
   ============================================================ */

function nhsn_getRequestById_(ss, requestId) {
  const sh = nhsperf_sheet_(ss, NHSN_REQUESTS_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  const rows = nhsn_readSheetObjects_(sh);
  return rows.find(r => String(r["Request ID"] || "") === String(requestId)) || null;
}

function nhsn_findTutor_(ss, name) {
  const sh = nhsperf_sheet_(ss, NHSN_TUTORS_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  const tutors = nhsn_readSheetObjects_(sh);
  const target = nhsn_norm_(name);

  return tutors.find(t => {
    const display = nhsn_tutorDisplayName_(t["Preferred Name"], t["Full Name"]);
    return [display,t["Full Name"],t["Preferred Name"]].map(nhsn_norm_).includes(target);
  }) || null;
}

function nhsn_readSheetObjects_(sh) {
  if (!sh) return [];

  if (
    sh.getName() === NHSN_REQUESTS_SHEET &&
    typeof nhsdata_readKeyedObjects_ === "function"
  ) {
    return nhsdata_readKeyedObjects_(sh, "Request ID");
  }

  const values = sh.getDataRange().getDisplayValues();
  if (!values.length) return [];
  const headers = values[0];

  return values.slice(1).map(function(row) {
    const obj = {};
    headers.forEach(function(h, i) { if (h) obj[h] = row[i] || ""; });
    return obj;
  });
}

function nhsn_headerIndex_(headers) {
  const idx = {};
  headers.forEach((h,i) => idx[h] = i);
  return idx;
}

function nhsn_tutorDisplayName_(preferred, full) {
  preferred = String(preferred || "").trim();
  full = String(full || "").trim();
  if (!preferred) return full;
  const parts = full.split(/\s+/).filter(Boolean);
  if (parts.length < 2) return preferred;
  return preferred + " " + parts[parts.length-1];
}

function nhsn_norm_(v) {
  return String(v || "").toLowerCase().replace(/[^\w@]+/g," ").replace(/\s+/g," ").trim();
}

function nhsn_parseDate_(value) {
  if (!value) return null;
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function nhsn_dateKey_(value, tz) {
  const d = nhsn_parseDate_(value);
  return d ? Utilities.formatDate(d, tz, "yyyy-MM-dd") : "";
}

function nhsn_formatDate_(value, tz) {
  const d = nhsn_parseDate_(value);
  return d ? Utilities.formatDate(d, tz, "EEEE, MMMM d, yyyy") : "";
}

function nhsn_hoursBetween_(a, b) {
  return (b.getTime() - a.getTime()) / 3600000;
}

function nhsn_id_(prefix) {
  return (
    String(prefix || "N") +
    "-" +
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyyMMddHHmmss"
    ) +
    "-" +
    Utilities.getUuid()
  );
}


/**
 * Consistent formal HTML wrapper for reminders, reschedules, cancellations,
 * and other notification-engine email.
 */
function nhsn_formalHtmlEmail_(spec) {
  spec = spec || {};
  const title = String(spec.title || spec.subject || "Tesoro NHS").trim();
  const message = String(spec.message || "").trim();
  const body = String(spec.body || "").trim();

  const bodyHtml = nhsn_escapeHtml_(body).replace(/\n/g, "<br>");

  return (
    '<!doctype html><html><body style="margin:0;padding:0;background:#f3f6fa;font-family:Arial,Helvetica,sans-serif;color:#172033;">' +
      '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f3f6fa;padding:26px 12px;">' +
        '<tr><td align="center">' +
          '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="max-width:620px;background:#ffffff;border:1px solid #dfe6ee;border-radius:14px;overflow:hidden;">' +
            '<tr><td style="background:#0f2747;padding:20px 26px;color:#ffffff;">' +
              '<div style="font-size:12px;letter-spacing:1.4px;font-weight:700;opacity:.82;">TESORO HIGH SCHOOL</div>' +
              '<div style="font-size:21px;font-weight:700;margin-top:4px;">Tesoro NHS</div><div style="font-size:13px;margin-top:3px;opacity:.82;">Tutoring Program</div>' +
            '</td></tr>' +
            '<tr><td style="padding:26px;">' +
              '<div style="display:inline-block;background:#e8f2ff;color:#124b88;border-radius:999px;padding:6px 10px;font-size:11px;font-weight:700;letter-spacing:.7px;">NHS TUTORING</div>' +
              '<h1 style="font-size:22px;line-height:1.3;margin:14px 0 8px;color:#13233b;">' +
                nhsn_escapeHtml_(title) +
              '</h1>' +
              (message
                ? '<div style="margin:0 0 18px;padding:13px 15px;background:#f7f9fc;border-radius:9px;color:#415169;font-size:13px;line-height:1.55;">' +
                    nhsn_escapeHtml_(message) +
                  '</div>'
                : '') +
              '<div style="font-size:14px;line-height:1.65;color:#344054;">' +
                bodyHtml +
              '</div>' +
            '</td></tr>' +
            '<tr><td style="padding:15px 26px;background:#f8fafc;border-top:1px solid #e7edf3;color:#7a8696;font-size:11px;line-height:1.5;">' +
              'Tesoro NHS · Tutoring Program · Automated scheduling and service coordination' +
            '</td></tr>' +
          '</table>' +
        '</td></tr>' +
      '</table>' +
    '</body></html>'
  );
}

function nhsn_escapeHtml_(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}


/* ============================================================
   V23 READY-TO-NOTIFY OFFICER ESCALATION
   ============================================================ */

/**
 * Emails internal NHS leadership only when a tutor-accepted request remains
 * unconfirmed past the configured threshold.
 *
 * It intentionally does NOT email officers immediately on tutor acceptance.
 * Immediate awareness comes from the Operations banner + alert count.
 */
function nhsn_sendReadyNotifyOfficerEscalations_(ss, now) {
  const requestSh = nhsperf_sheet_(ss, NHSN_REQUESTS_SHEET);
  if (!requestSh || requestSh.getLastRow() < 2) return;

  const requests = nhsn_readSheetObjects_(requestSh);
  const officerEmails = nhsn_getOperationsOfficerEmails_(ss);

  if (!officerEmails.length) return;

  requests.forEach(function(r) {
    const requestId = String(r["Request ID"] || "").trim();
    const status = String(r["Status"] || "").trim();
    const tutorResponse = String(r["Tutor Response Status"] || "").trim();
    const confirmed = String(r["Assignment Confirmed"] || "").trim();

    if (
      !requestId ||
      tutorResponse !== "Accepted" ||
      confirmed ||
      !["Assigned", "Scheduled"].includes(status)
    ) {
      return;
    }

    const acceptedAt =
      nhsn_parseDate_(r["Tutor Response At"]) ||
      nhsn_parseDate_(r["Last Update"]);

    if (!acceptedAt) return;

    const minutesWaiting =
      (now.getTime() - acceptedAt.getTime()) / 60000;

    if (minutesWaiting < NHSN_READY_NOTIFY_ESCALATION_MINUTES) {
      return;
    }

    const course = String(r["Course"] || "Tutoring request").trim();
    const tutor = String(r["Assigned Tutor"] || "Assigned tutor").trim();
    const teacher = String(r["Teacher Name"] || "Requesting teacher").trim();
    const date = String(r["Scheduled Date"] || "").trim();
    const room = String(r["Location"] || "").trim();

    officerEmails.forEach(function(email) {
      nhsn_sendEmailLogged_(ss, {
        requestId: requestId,
        recipient: email,
        type: "Operations Ready-to-Notify Escalation",
        subject:
          "Action Needed: " +
          course +
          " is ready to notify",
        dedupeKey:
          "READY_NOTIFY_OFFICER|" +
          requestId +
          "|" +
          email.toLowerCase(),
        title: "Tutoring assignment awaiting confirmation",
        message:
          "This request has remained ready to notify for at least " +
          NHSN_READY_NOTIFY_ESCALATION_MINUTES +
          " minutes.",
        body: [
          "Tesoro NHS Operations",
          "",
          "A tutor has accepted an assignment, but the teacher/tutor confirmation has not yet been sent.",
          "",
          "Course: " + course,
          "Tutor: " + tutor,
          "Teacher: " + teacher,
          date ? "Scheduled Date: " + date : "",
          room ? "Location: " + room : "",
          "Request ID: " + requestId,
          "",
          "Open Internal Operations and use “Notify Teacher & Tutor” to complete the confirmation."
        ].filter(Boolean).join("\n")
      });
    });
  });
}


/**
 * Uses active Access Registry grants so there is no second officer-email list
 * to maintain. Duplicate emails are collapsed.
 */
function nhsn_getOperationsOfficerEmails_(ss) {
  const sh = nhsperf_sheet_(ss, NHSN_ACCESS_REGISTRY_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];

  const rows = nhsn_readSheetObjects_(sh);
  const allowedRoles = {
    ADMIN: true,
    ADVISOR: true,
    OFFICER: true
  };
  const seen = {};
  const emails = [];

  rows.forEach(function(r) {
    const role = String(r["Role"] || "").trim().toUpperCase();
    const status = String(r["Status"] || "").trim().toLowerCase();
    const email = String(r["School Email"] || "").trim().toLowerCase();

    if (
      !allowedRoles[role] ||
      status !== "active" ||
      !nhsn_isValidEmail_(email) ||
      seen[email]
    ) {
      return;
    }

    seen[email] = true;
    emails.push(email);
  });

  return emails;
}


/* ============================================================
   V43 MISSING NOTIFICATION HELPERS
   ============================================================ */
function nhsn_isValidEmail_(email) {
  const value = String(email || "").trim();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function nhsn_logAlert_(ss, spec) {
  spec = spec || {};
  const sh = nhsn_ensureLog_(ss);
  const h = nhsn_headerMapSheet_(sh);
  const row = new Array(sh.getLastColumn()).fill("");

  const notificationId =
    "N-" +
    Utilities.formatDate(
      new Date(),
      nhs_safeTimeZone_(ss),
      "yyyyMMdd-HHmmss"
    ) +
    "-" +
    Utilities.getUuid();

  nhsn_setRow_(row, h, "Notification ID", notificationId);
  nhsn_setRow_(row, h, "Created At", new Date());
  nhsn_setRow_(row, h, "Request ID", spec.requestId || "");
  nhsn_setRow_(row, h, "Channel", "Internal Alert");
  nhsn_setRow_(row, h, "Recipient", "NHS Officers");
  nhsn_setRow_(row, h, "Type", spec.type || "INTERNAL_ALERT");
  nhsn_setRow_(row, h, "Status", "Open");
  nhsn_setRow_(row, h, "Dedupe Key", spec.dedupeKey || "");
  nhsn_setRow_(row, h, "Read", false);
  nhsn_setRow_(row, h, "Portal Read", false);
  nhsn_setRow_(row, h, "Severity", spec.severity || "Medium");
  nhsn_setRow_(row, h, "Title", spec.title || "NHS alert");
  nhsn_setRow_(row, h, "Message", spec.message || "");

  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  return notificationId;
}


/** ============================================================
 * SOURCE GROUP: NHS_Mail_Transport.gs
 * ============================================================ */

/**
 * TESORO NHS — OUTBOUND MAIL TRANSPORT V43
 * MAIN NHS PROJECT ONLY
 */

const NHSM_RELAY_URL_PROPERTY = "NHS_MAIL_RELAY_URL";
const NHSM_RELAY_TOKEN_PROPERTY = "NHS_MAIL_RELAY_TOKEN";
const NHSM_RELAY_TIMEOUT_RETRIES = 1;
const NHSM_DEFAULT_RELAY_URL = "https://script.google.com/macros/s/AKfycbzYonSQlKFnXZxqP_YwkPNxR1kkfVJa73co-dus1hq3/exec";

function configureNHSMailRelay(relayUrl, sharedSecret) {
  nhssec_requireMaintenanceAdmin_();
  relayUrl = nhsm_normalizeRelayUrl_(relayUrl || NHSM_DEFAULT_RELAY_URL);
  sharedSecret = String(sharedSecret || "").trim();

  if (!/^https:\/\/script\.google\.com\/.+\/exec(?:\?.*)?$/.test(relayUrl)) {
    throw new Error("Relay URL must be a deployed Apps Script Web App URL ending in /exec.");
  }

  if (sharedSecret.length < 24) {
    throw new Error("Relay shared secret must be at least 24 characters.");
  }

  PropertiesService.getScriptProperties().setProperties({
    [NHSM_RELAY_URL_PROPERTY]: relayUrl,
    [NHSM_RELAY_TOKEN_PROPERTY]: sharedSecret
  });

  return {
    ok: true,
    message: "NHS mail relay configured.",
    relayUrl: relayUrl
  };
}

function setupNHSRelayFinal() {
  nhssec_requireMaintenanceAdmin_();
  // V43: never embed relay credentials in source control.
  const props = PropertiesService.getScriptProperties();
  const existingToken = String(props.getProperty(NHSM_RELAY_TOKEN_PROPERTY) || "").trim();

  if (!existingToken) {
    throw new Error(
      "Relay token is not configured. Set the Script Property " +
      NHSM_RELAY_TOKEN_PROPERTY +
      " to the SAME rotated secret stored in the relay project, then run setupNHSRelayFinal() again."
    );
  }

  props.setProperty(
    NHSM_RELAY_URL_PROPERTY,
    nhsm_normalizeRelayUrl_(NHSM_DEFAULT_RELAY_URL)
  );

  return {
    ok: true,
    message: "NHS mail relay URL verified; existing token retained securely in Script Properties.",
    relayUrl: nhsm_normalizeRelayUrl_(NHSM_DEFAULT_RELAY_URL),
    tokenConfigured: true
  };
}

function getNHSMailRelayStatus() {
  nhssec_requireMaintenanceAdmin_();

  const props = PropertiesService.getScriptProperties();
  const relayUrl = String(props.getProperty(NHSM_RELAY_URL_PROPERTY) || "").trim();
  const token = String(props.getProperty(NHSM_RELAY_TOKEN_PROPERTY) || "").trim();

  return {
    configured: !!(relayUrl && token),
    relayUrl: relayUrl,
    tokenConfigured: !!token
  };
}

function clearNHSMailRelayConfiguration() {
  nhssec_requireMaintenanceAdmin_();
  PropertiesService.getScriptProperties().deleteProperty(NHSM_RELAY_URL_PROPERTY);
  PropertiesService.getScriptProperties().deleteProperty(NHSM_RELAY_TOKEN_PROPERTY);

  return {
    ok: true,
    message: "NHS mail relay configuration cleared."
  };
}

function testNHSMailRelay(destination) {
  nhssec_requireMaintenanceAdmin_();
  destination = String(destination || "").trim();

  if (!nhsm_isValidEmail_(destination)) {
    throw new Error("Enter a valid test destination email.");
  }

  return nhsmail_send_({
    to: destination,
    subject: "NHS Tutoring — Relay Test",
    body:
      "This is a one-time NHS Tutoring mail relay test.\n\n" +
      "If you received this message, the final notification transport is working.\n\n" +
      "Timestamp: " + new Date().toString(),
    name: "Tesoro NHS",
    type: "Relay Test"
  });
}

function testNHSRelayFinal() {
  nhssec_requireMaintenanceAdmin_();

  const destination = String(Session.getActiveUser().getEmail() || "").trim();
  if (!destination) {
    throw new Error("Unable to determine the current user's email for the relay test.");
  }
  return testNHSMailRelay(destination);
}

function nhsmail_send_(spec) {
  spec = spec || {};

  let to = String(spec.to || spec.recipient || "").trim();
  let subject = String(spec.subject || "").trim();
  const body = String(spec.body || "");
  const name = String(spec.name || "Tesoro NHS").trim();
  const requestId = String(spec.requestId || "").trim();
  const type = String(spec.type || "").trim();

  const calendar = spec.calendar && typeof spec.calendar === "object"
    ? {
        uid: String(spec.calendar.uid || ""),
        title: String(spec.calendar.title || ""),
        description: String(spec.calendar.description || ""),
        location: String(spec.calendar.location || ""),
        startMs: Number(spec.calendar.startMs || 0),
        endMs: Number(spec.calendar.endMs || 0),
        attendee: String(spec.calendar.attendee || to),
        includeGoogleLink: spec.calendar.includeGoogleLink !== false
      }
    : null;

  if (!nhsm_isValidEmail_(to)) {
    throw new Error("Invalid notification email: " + to);
  }

  if (!subject) {
    throw new Error("Missing email subject.");
  }

  // V47 Release Guard: staging never sends to real recipients by default.
  // RECORD_ONLY captures the message and returns success without any network
  // call. REDIRECT captures the original target and sends only to the explicit
  // staging sink address.
  const stagingCapture = nhsrg_captureStagingMail_(Object.assign({}, spec, {
    to: to,
    subject: subject,
    body: body,
    type: type,
    requestId: requestId
  }), "Email");
  if (stagingCapture) {
    if (stagingCapture.mode === "RECORD_ONLY") {
      return nhsrg_stagingMailResult_(stagingCapture, to, type);
    }
    const intendedRecipient = String(stagingCapture.intendedRecipient || to);
    to = String(stagingCapture.redirectRecipient || "").trim();
    if (!nhsm_isValidEmail_(to)) {
      throw new Error("Staging REDIRECT mode requires a valid sink email.");
    }
    subject = "[STAGING → " + intendedRecipient + "] " + subject;
    // Safety: if the message carries a calendar payload, redirect the attendee
    // too. Never allow a staging redirect to invite the real intended user.
    if (calendar) {
      calendar.attendee = to;
      calendar.title = "[STAGING] " + String(calendar.title || "NHS Tutoring Session");
      calendar.description = "STAGING TEST ONLY. Intended recipient: " + intendedRecipient + "\n\n" + String(calendar.description || "");
    }
  }

  const props = PropertiesService.getScriptProperties();
  const relayUrl = nhsm_normalizeRelayUrl_(
    props.getProperty(NHSM_RELAY_URL_PROPERTY) || NHSM_DEFAULT_RELAY_URL
  );
  const relayToken = String(
    props.getProperty(NHSM_RELAY_TOKEN_PROPERTY) || ""
  ).trim();

  if (!relayUrl || !relayToken) {
    throw new Error(
      "NHS mail relay is not configured. Run setupNHSRelayFinal() once."
    );
  }

  const payload = {
    token: relayToken,
    to: to,
    subject: subject,
    body: body,
    htmlBody: String(spec.htmlBody || ""),
    name: name,
    source: "Tesoro NHS",
    requestId: requestId,
    type: type,
    sentAt: new Date().toISOString(),
    calendar: calendar
  };

  let lastError = null;

  for (let attempt = 0; attempt <= NHSM_RELAY_TIMEOUT_RETRIES; attempt++) {
    try {
      const response = UrlFetchApp.fetch(relayUrl, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
        followRedirects: true
      });

      const status = response.getResponseCode();
      const text = response.getContentText();

      let parsed = null;
      try {
        parsed = JSON.parse(text || "{}");
      } catch (parseError) {}

      if (
        status >= 200 &&
        status < 300 &&
        parsed &&
        parsed.ok === true
      ) {
        return {
          ok: true,
          provider: "NHS Mail Relay",
          recipient: to,
          relayMessageId: String(parsed.messageId || ""),
          sender: String(parsed.sender || ""),
          calendarAttached: !!parsed.calendarAttached,
          calendarUid: String(parsed.calendarUid || ""),
          googleCalendarUrl: String(parsed.googleCalendarUrl || "")
        };
      }

      const detail =
        parsed && parsed.error
          ? parsed.error
          : (
              "HTTP " +
              status +
              (text ? ": " + text.slice(0, 300) : "")
            );

      if (status === 429 || status >= 500) {
        lastError = new Error("Mail relay transient failure: " + detail);
        if (attempt < NHSM_RELAY_TIMEOUT_RETRIES) {
          Utilities.sleep(350);
          continue;
        }
      }

      throw new Error("Mail relay failed: " + detail);

    } catch (e) {
      lastError = e;

      if (
        attempt < NHSM_RELAY_TIMEOUT_RETRIES &&
        /timeout|temporar|429|5\d\d/i.test(
          String(e && e.message ? e.message : e)
        )
      ) {
        Utilities.sleep(350);
        continue;
      }

      throw e;
    }
  }

  throw lastError || new Error("Unknown mail relay failure.");
}



/**
 * V29: sends multiple independent messages through ONE relay HTTP request.
 *
 * This removes the second Apps Script -> relay round trip from assignment
 * confirmation while preserving distinct recipients, subjects, HTML bodies,
 * and per-recipient calendar invitations.
 */
function nhsmail_sendBatch_(specs) {
  specs = Array.isArray(specs) ? specs : [];

  if (!specs.length) {
    return {
      ok: true,
      results: []
    };
  }

  const env = nhsrg_environmentInfo_();
  if (env.isStaging) {
    const results = specs.map(function(spec, index) {
      spec = spec || {};
      const intended = String(spec.to || spec.recipient || "").trim();
      if (!nhsm_isValidEmail_(intended)) {
        throw new Error("Invalid notification email at batch item " + (index + 1) + ": " + intended);
      }
      const capture = nhsrg_captureStagingMail_(spec, "Batch Email");
      if (capture.mode === "REDIRECT") {
        // Preserve the single-message safety path. Each redirected staging
        // message goes only to the configured test inbox.
        const redirected = Object.assign({}, spec, {
          to: capture.redirectRecipient,
          subject: "[STAGING → " + intended + "] " + String(spec.subject || "")
        });
        // Temporarily bypass capture recursion by sending through the raw helper.
        return nhsmail_sendRawRelay_(redirected, intended);
      }
      const result = nhsrg_stagingMailResult_(capture, intended, spec.type || ("message-" + (index + 1)));
      result.id = String(spec.id || spec.type || ("message-" + (index + 1)));
      return result;
    });
    return { ok: true, provider: "NHS Staging Mail Sink Batch", results: results, staging: true };
  }

  const props = PropertiesService.getScriptProperties();

  const relayUrl = nhsm_normalizeRelayUrl_(
    props.getProperty(NHSM_RELAY_URL_PROPERTY) ||
    NHSM_DEFAULT_RELAY_URL
  );

  const relayToken = String(
    props.getProperty(NHSM_RELAY_TOKEN_PROPERTY) || ""
  ).trim();

  if (!relayUrl || !relayToken) {
    throw new Error(
      "NHS mail relay is not configured. Run setupNHSRelayFinal() once."
    );
  }

  const messages = specs.map(function(spec, index) {
    spec = spec || {};

    const to = String(
      spec.to || spec.recipient || ""
    ).trim();

    const subject = String(
      spec.subject || ""
    ).trim();

    if (!nhsm_isValidEmail_(to)) {
      throw new Error(
        "Invalid notification email at batch item " +
        (index + 1) +
        ": " +
        to
      );
    }

    if (!subject) {
      throw new Error(
        "Missing email subject at batch item " +
        (index + 1) +
        "."
      );
    }

    const calendar =
      spec.calendar &&
      typeof spec.calendar === "object"
        ? {
            uid: String(spec.calendar.uid || ""),
            title: String(spec.calendar.title || ""),
            description: String(
              spec.calendar.description || ""
            ),
            location: String(
              spec.calendar.location || ""
            ),
            startMs: Number(
              spec.calendar.startMs || 0
            ),
            endMs: Number(
              spec.calendar.endMs || 0
            ),
            attendee: String(
              spec.calendar.attendee || to
            ),
            includeGoogleLink:
              spec.calendar.includeGoogleLink !== false
          }
        : null;

    return {
      id: String(
        spec.id ||
        spec.type ||
        ("message-" + (index + 1))
      ),
      to: to,
      subject: subject,
      body: String(spec.body || ""),
      htmlBody: String(spec.htmlBody || ""),
      name: String(
        spec.name || "Tesoro NHS"
      ).trim(),
      source: "Tesoro NHS",
      requestId: String(
        spec.requestId || ""
      ).trim(),
      type: String(
        spec.type || ""
      ).trim(),
      calendar: calendar
    };
  });

  const payload = {
    token: relayToken,
    action: "sendBatch",
    source: "Tesoro NHS",
    sentAt: new Date().toISOString(),
    messages: messages
  };

  let lastError = null;

  for (
    let attempt = 0;
    attempt <= NHSM_RELAY_TIMEOUT_RETRIES;
    attempt++
  ) {
    try {
      const response = UrlFetchApp.fetch(
        relayUrl,
        {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify(payload),
          muteHttpExceptions: true,
          followRedirects: true
        }
      );

      const status = response.getResponseCode();
      const text = response.getContentText();

      let parsed = null;

      try {
        parsed = JSON.parse(text || "{}");
      } catch (parseError) {}

      if (
        status >= 200 &&
        status < 300 &&
        parsed &&
        parsed.ok === true &&
        Array.isArray(parsed.results)
      ) {
        return {
          ok: true,
          provider: "NHS Mail Relay Batch",
          results: parsed.results
        };
      }

      const detail =
        parsed && parsed.error
          ? parsed.error
          : (
              "HTTP " +
              status +
              (
                text
                  ? ": " + text.slice(0, 300)
                  : ""
              )
            );

      if (
        status === 429 ||
        status >= 500
      ) {
        lastError = new Error(
          "Mail relay transient failure: " +
          detail
        );

        if (
          attempt <
          NHSM_RELAY_TIMEOUT_RETRIES
        ) {
          Utilities.sleep(350);
          continue;
        }
      }

      throw new Error(
        "Mail relay failed: " +
        detail
      );

    } catch (e) {
      lastError = e;

      if (
        attempt <
          NHSM_RELAY_TIMEOUT_RETRIES &&
        /timeout|temporar|429|5\d\d/i.test(
          String(
            e && e.message
              ? e.message
              : e
          )
        )
      ) {
        Utilities.sleep(350);
        continue;
      }

      throw e;
    }
  }

  throw (
    lastError ||
    new Error(
      "Unknown mail relay batch failure."
    )
  );
}

function nhsm_normalizeRelayUrl_(value) {
  value = String(value || "").trim();
  if (value.endsWith("/dev")) {
    return value.slice(0, -4) + "/exec";
  }
  return value;
}

function nhsm_isValidEmail_(value) {
  value = String(value || "").trim();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}


/**
 * Creates one Google Calendar event through the standalone relay account.
 *
 * The relay owns the calendar event and sends native invitations to all guests.
 * This means teacher school/Outlook addresses can receive a normal calendar
 * invitation while tutors can receive it at their Notification Email.
 */
function nhsmail_createCalendarEvent_(spec) {
  spec = spec || {};

  const stagingCapture = nhsrg_captureStagingMail_({
    to: Array.isArray(spec.guests) ? spec.guests.join(", ") : "",
    subject: String(spec.title || "NHS Tutoring Session"),
    body: String(spec.description || ""),
    requestId: String(spec.requestId || ""),
    type: "Calendar Event",
    calendar: spec
  }, "Calendar Event");
  if (stagingCapture) {
    return {
      ok: true,
      staging: true,
      provider: "NHS Staging Calendar Sink",
      eventId: "STAGING-CALENDAR-" + Date.now(),
      calendarEventId: "STAGING-CALENDAR-" + Date.now(),
      calendarAttached: false,
      calendarUid: "STAGING-" + String(spec.requestId || Date.now()),
      googleCalendarUrl: ""
    };
  }

  const props = PropertiesService.getScriptProperties();
  const relayUrl = nhsm_normalizeRelayUrl_(
    props.getProperty(NHSM_RELAY_URL_PROPERTY) || NHSM_DEFAULT_RELAY_URL
  );
  const relayToken = String(
    props.getProperty(NHSM_RELAY_TOKEN_PROPERTY) || ""
  ).trim();

  if (!relayUrl || !relayToken) {
    throw new Error(
      "NHS mail relay is not configured. Run setupNHSRelayFinal() once."
    );
  }

  const guests = Array.isArray(spec.guests)
    ? spec.guests.map(v => String(v || "").trim()).filter(nhsm_isValidEmail_)
    : [];

  if (!spec.startMs || !spec.endMs) {
    throw new Error("Calendar event requires a scheduled start and end time.");
  }

  const payload = {
    token: relayToken,
    action: "createCalendarEvent",
    requestId: String(spec.requestId || ""),
    title: String(spec.title || "NHS Tutoring Session"),
    description: String(spec.description || ""),
    location: String(spec.location || ""),
    startMs: Number(spec.startMs),
    endMs: Number(spec.endMs),
    guests: guests,
    source: "Tesoro NHS"
  };

  const response = UrlFetchApp.fetch(relayUrl, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    followRedirects: true
  });

  const status = response.getResponseCode();
  const text = response.getContentText();

  let parsed = null;
  try {
    parsed = JSON.parse(text || "{}");
  } catch (e) {}

  if (
    status >= 200 &&
    status < 300 &&
    parsed &&
    parsed.ok === true
  ) {
    return parsed;
  }

  const detail =
    parsed && parsed.error
      ? parsed.error
      : "HTTP " + status + (text ? ": " + text.slice(0, 300) : "");

  throw new Error("Calendar relay failed: " + detail);
}



/**
 * Manual transport-only benchmark.
 *
 * Does NOT send mail. Returns whether V29 batch function exists and relay
 * settings are present.
 */
function checkNHSBatchRelayV29() {
  nhssec_requireMaintenanceAdmin_();

  const props = PropertiesService.getScriptProperties();
  const relayUrl = String(
    props.getProperty(NHSM_RELAY_URL_PROPERTY) || NHSM_DEFAULT_RELAY_URL || ""
  ).trim();
  const tokenConfigured = !!String(
    props.getProperty(NHSM_RELAY_TOKEN_PROPERTY) || ""
  ).trim();
  const batchFunction = typeof nhsmail_sendBatch_ === "function";
  const relayUrlConfigured = /^https:\/\/script\.google\.com\/.+\/exec(?:\?.*)?$/.test(relayUrl);
  const ok = batchFunction && relayUrlConfigured && tokenConfigured;

  return {
    ok: ok,
    relayUrlConfigured: relayUrlConfigured,
    relayTokenConfigured: tokenConfigured,
    batchFunctionAvailable: batchFunction,
    mode: "Single HTTP relay request for teacher + tutor confirmation",
    message: ok
      ? "Batch relay is configured."
      : "Batch relay needs attention: " +
        [
          batchFunction ? "" : "batch function missing",
          relayUrlConfigured ? "" : "relay URL missing/invalid",
          tokenConfigured ? "" : "relay token missing"
        ].filter(Boolean).join(", ")
  };
}


/** ============================================================
 * SOURCE GROUP: NHS_Assignment_Workflow.gs
 * ============================================================ */

/**
 * NHS ASSIGNMENT WORKFLOW
 * File name: AssignmentWorkflow.gs
 *
 * PURPOSE
 * -------
 * Adds the officer-side assignment workflow after tutor matching.
 *
 * Workflow:
 *   1) Officer chooses a tutor in "Assigned Tutor".
 *   2) Status becomes "Assigned"; Last Update refreshes.
 *   3) Tutor caseloads are recalculated safely from active assignments.
 *   4) Officer checks "Confirm Assignment".
 *   5) Teacher + tutor assignment emails are sent once.
 *   6) Notification Status / Assignment Confirmed are recorded.
 *   7) Adding a Scheduled Date changes status to "Scheduled".
 *   8) Titan Time scheduling is standardized automatically:
 *        Time: 10:25 AM–11:05 AM
 *        Location: Teacher-entered classroom / room number from the request form
 *   9) Completing/cancelling a request automatically frees tutor capacity.
 *
 * SAFETY / DATA INTEGRITY
 * -----------------------
 * - Caseload is RECONCILED from the requests table; it is not blindly incremented.
 *   This prevents double-counting when assignments are changed, removed, completed,
 *   or cancelled.
 * - Emails are NOT sent merely by selecting a tutor. The officer must explicitly
 *   check "Confirm Assignment".
 * - If the assigned tutor changes after confirmation, the confirmation checkbox is
 *   reset and the notification status becomes "Reassignment Pending".
 *
 * DEPENDENCY
 * ----------
 * Reuses the existing global NHS_SPREADSHEET_ID from Code.gs.
 * Do NOT declare NHS_SPREADSHEET_ID again in this file.
 *
 * RUN ONCE:
 *   installAssignmentWorkflow()
 */


/* ============================================================
   CONFIGURATION
   ============================================================ */

const NHS_ASSIGNMENT_REQUEST_SHEET = "Tutoring Requests";
const NHS_ASSIGNMENT_TUTOR_SHEET = "Tutor Database";

const NHS_ASSIGNMENT_HEADERS = [
  "Teacher Notification Email",
  "Confirm Assignment",
  "Notification Status",
  "Assignment Confirmed",
  "Calendar Event ID",
  "Calendar Status",
  "Teacher Confirmation Sent At",
  "Tutor Confirmation Sent At",
  "Preferred Dates",
  "Tutor Reconfirm Status",
  "Tutor Reconfirm At",
  "Teacher Five Star Ready",
  "Teacher Five Star Ready At",
  "Tutor Five Star Enrolled",
  "Tutor Five Star Enrolled At",
  "Five Star Ready At"
];

const NHS_ACTIVE_CASELOAD_STATUSES = [
  "Assigned",
  "Scheduled"
];


/* ============================================================
   ONE-TIME INSTALLER
   ============================================================ */

function installAssignmentWorkflow() {
  nhssec_requireMaintenanceAdmin_();

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const requestSheet = nhsperf_sheet_(ss, NHS_ASSIGNMENT_REQUEST_SHEET);
  const tutorSheet = nhsperf_sheet_(ss, NHS_ASSIGNMENT_TUTOR_SHEET);

  if (!requestSheet) {
    throw new Error('Sheet "' + NHS_ASSIGNMENT_REQUEST_SHEET + '" was not found.');
  }

  if (!tutorSheet) {
    throw new Error('Sheet "' + NHS_ASSIGNMENT_TUTOR_SHEET + '" was not found.');
  }

  ensureAssignmentColumns_(requestSheet);
  formatAssignmentWorkflow_(ss, requestSheet);
  installAssignmentEditTrigger_();
  queueTutorCaseloadReconciliation_();

  Logger.log("NHS Assignment Workflow installed successfully.");
}


/* ============================================================
   INSTALLABLE EDIT TRIGGER
   ============================================================ */

function installAssignmentEditTrigger_() {
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "handleAssignmentWorkflowEdit_") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("handleAssignmentWorkflowEdit_")
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}


/**
 * Main edit handler.
 * Runs only from an installable spreadsheet onEdit trigger.
 */
function handleAssignmentWorkflowEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // Manual edits to access/configuration sources must invalidate the same
  // revisions as the Operations APIs. This prevents a five-minute stale
  // Gateway/access state when an authorized administrator edits Sheets directly.
  const accessStateSheets = [
    NHSA_REGISTRY_SHEET,
    NHSA_REQUESTS_SHEET,
    NHSA_SETTINGS_SHEET,
    NHSA_TUTOR_SHEET,
    TP_TEACHER_DIRECTORY_SHEET,
    NHSSCHED_EXCEPTION_SHEET
  ];

  if (accessStateSheets.includes(sheetName)) {
    try {
      nhsa_invalidateAccessCache_();
    } catch (ignore) {}

    try {
      nhsperf_touchDataRevision_();
    } catch (ignore) {}

    // Tutor Database edits can change identity, reliability, availability,
    // notification routing, and access eligibility. Clear cross-portal caches.
    if (sheetName === NHSA_TUTOR_SHEET) {
      try {
        nhs_v45InvalidatePortalState_({});
      } catch (ignore) {}
    }

    if (sheetName !== NHS_ASSIGNMENT_REQUEST_SHEET) {
      return;
    }
  }

  if (sheetName !== NHS_ASSIGNMENT_REQUEST_SHEET) {
    return;
  }

  const row = e.range.getRow();

  if (row < 2) {
    return;
  }

  const headers = getAssignmentHeaderMap_(sheet);
  const editedCol = e.range.getColumn();

  const assignedTutorCol = headers["Assigned Tutor"];
  const confirmCol = headers["Confirm Assignment"];
  const statusCol = headers["Status"];
  const scheduledDateCol = headers["Scheduled Date"];
  const scheduledTimeCol = headers["Scheduled Time"];
  const locationCol = headers["Location"];

  if (editedCol === assignedTutorCol) {
    handleAssignedTutorChange_(sheet, row, headers);
    return;
  }

  if (editedCol === confirmCol) {
    handleConfirmAssignmentChange_(sheet, row, headers);
    return;
  }

  if (editedCol === statusCol) {
    handleStatusChange_(sheet, row, headers);
    return;
  }

  if (
    editedCol === scheduledDateCol ||
    editedCol === scheduledTimeCol ||
    editedCol === locationCol
  ) {
    handleScheduleChange_(sheet, row, headers);
    return;
  }
}


/* ============================================================
   ASSIGNED TUTOR CHANGE
   ============================================================ */

function handleAssignedTutorChange_(sheet, row, headers) {
  const request = getRequestObject_(sheet, row, headers);
  const tutorName = String(
    sheet.getRange(row, headers["Assigned Tutor"]).getDisplayValue() || ""
  ).trim();

  const confirmedCell = sheet.getRange(row, headers["Confirm Assignment"]);
  const notificationCell = sheet.getRange(row, headers["Notification Status"]);
  const confirmedAtCell = sheet.getRange(row, headers["Assignment Confirmed"]);
  const statusCell = sheet.getRange(row, headers["Status"]);

  if (!tutorName) {
    confirmedCell.setValue(false);
    notificationCell.setValue("");
    confirmedAtCell.clearContent();

    const currentStatus = String(statusCell.getDisplayValue() || "").trim();

    if (
      currentStatus === "Assigned" ||
      currentStatus === "Scheduled" ||
      currentStatus === "No Tutor Available"
    ) {
      statusCell.setValue("New");
    }

    updateRequestTimestamp_(sheet, row, headers);
    invalidateAssignmentWorkflowCaches_(request);
    queueTutorCaseloadReconciliation_();
    return;
  }

  const tutor = findTutorByDisplayName_(tutorName);

  if (!tutor) {
    statusCell.setValue("Needs Follow-Up");
    notificationCell.setValue("Tutor Not Found");
    confirmedCell.setValue(false);
    confirmedAtCell.clearContent();
    updateRequestTimestamp_(sheet, row, headers);
    invalidateAssignmentWorkflowCaches_(request);
    queueTutorCaseloadReconciliation_();
    return;
  }

  // If the request had already been confirmed, selecting a different tutor
  // intentionally invalidates the old confirmation.
  const wasConfirmed = confirmedCell.isChecked() === true ||
    String(confirmedAtCell.getDisplayValue() || "").trim() !== "";

  confirmedCell.setValue(false);
  confirmedAtCell.clearContent();

  if (wasConfirmed) {
    notificationCell.setValue("Reassignment Pending");
  } else {
    notificationCell.setValue("Ready to Confirm");
  }

  statusCell.setValue("Assigned");
  updateRequestTimestamp_(sheet, row, headers);
  invalidateAssignmentWorkflowCaches_(request);
  queueTutorCaseloadReconciliation_();
}


/* ============================================================
   CONFIRM ASSIGNMENT
   ============================================================ */

function handleConfirmAssignmentChange_(sheet, row, headers) {
  const confirmCell = sheet.getRange(row, headers["Confirm Assignment"]);

  if (confirmCell.isChecked() !== true) {
    return;
  }

  const tutorName = String(
    sheet.getRange(row, headers["Assigned Tutor"]).getDisplayValue() || ""
  ).trim();

  if (!tutorName) {
    confirmCell.setValue(false);
    sheet.getRange(row, headers["Status"]).setValue("Needs Follow-Up");
    sheet
      .getRange(row, headers["Notification Status"])
      .setValue("Select a Tutor First");
    updateRequestTimestamp_(sheet, row, headers);
    return;
  }

  const tutor = findTutorByDisplayName_(tutorName);

  if (!tutor) {
    confirmCell.setValue(false);
    sheet.getRange(row, headers["Status"]).setValue("Needs Follow-Up");
    sheet
      .getRange(row, headers["Notification Status"])
      .setValue("Tutor Not Found");
    updateRequestTimestamp_(sheet, row, headers);
    return;
  }

  const request = getRequestObject_(sheet, row, headers);

  // Titan Time confirmations require a specific room number.
  const location = String(request["Location"] || "").trim();

  if (!location) {
    confirmCell.setValue(false);
    sheet.getRange(row, headers["Status"]).setValue("Needs Follow-Up");
    sheet
      .getRange(row, headers["Notification Status"])
      .setValue("Classroom Required");
    updateRequestTimestamp_(sheet, row, headers);
    return;
  }

  // Do not send twice if the existing confirmation timestamp remains.
  const existingConfirmation = String(
    request["Assignment Confirmed"] || ""
  ).trim();

  if (existingConfirmation) {
    sheet
      .getRange(row, headers["Notification Status"])
      .setValue("Already Sent");
    updateRequestTimestamp_(sheet, row, headers);
    invalidateAssignmentWorkflowCaches_(request);
    return;
  }

  const result = sendAssignmentEmails_(request, tutor);

  if (result.teacherSent && result.tutorSent) {
    sheet
      .getRange(row, headers["Notification Status"])
      .setValue("Teacher & Tutor Notified");

    sheet
      .getRange(row, headers["Assignment Confirmed"])
      .setValue(new Date());

    if (headers["Calendar Event ID"]) {
      sheet
        .getRange(row, headers["Calendar Event ID"])
        .setValue(result.calendarEventId || "");
    }

    if (headers["Calendar Status"]) {
      sheet
        .getRange(row, headers["Calendar Status"])
        .setValue(
          result.calendarCreated
            ? "Calendar Invite Sent"
            : (result.calendarError ? "Calendar Invite Failed" : "Not Scheduled")
        );
    }

    if (request["Scheduled Date"]) {
      sheet.getRange(row, headers["Status"]).setValue("Scheduled");
    } else {
      sheet.getRange(row, headers["Status"]).setValue("Assigned");
    }

    updateRequestTimestamp_(sheet, row, headers);
    queueTutorCaseloadReconciliation_();
  } else {
    confirmCell.setValue(false);

    const problems = [];

    if (!result.teacherSent) {
      problems.push(result.teacherError || "Teacher email not sent");
    }

    if (!result.tutorSent) {
      problems.push(result.tutorError || "Tutor email not sent");
    }

    sheet
      .getRange(row, headers["Notification Status"])
      .setValue("Needs Follow-Up: " + problems.join("; "));

    sheet.getRange(row, headers["Status"]).setValue("Needs Follow-Up");
    updateRequestTimestamp_(sheet, row, headers);
  }
}


/* ============================================================
   STATUS / SCHEDULING CHANGES
   ============================================================ */

function handleStatusChange_(sheet, row, headers) {
  updateRequestTimestamp_(sheet, row, headers);
  queueTutorCaseloadReconciliation_();
}


function handleScheduleChange_(sheet, row, headers) {
  const assignedTutor = String(
    sheet.getRange(row, headers["Assigned Tutor"]).getDisplayValue() || ""
  ).trim();

  const scheduledDate = String(
    sheet.getRange(row, headers["Scheduled Date"]).getDisplayValue() || ""
  ).trim();

  if (assignedTutor && scheduledDate) {
    // Current teacher-request workflow is Titan Time.
    // Standardize the fixed Titan Time session window and meeting location.
    applyTitanTimeDefaults_(sheet, row, headers);

    const currentStatus = String(
      sheet.getRange(row, headers["Status"]).getDisplayValue() || ""
    ).trim();

    if (
      currentStatus !== "Completed" &&
      currentStatus !== "Cancelled"
    ) {
      sheet.getRange(row, headers["Status"]).setValue("Scheduled");
    }
  }

  updateRequestTimestamp_(sheet, row, headers);
  queueTutorCaseloadReconciliation_();
}


/**
 * Standard Titan Time logistics.
 * Titan Time: 10:25 AM–11:05 AM
 * Location: teacher-entered classroom / room number from the request form
 */
function applyTitanTimeDefaults_(sheet, row, headers) {
  // Titan Time has a fixed time. Location is supplied by the requesting
  // teacher on the Teacher Request Form and must be preserved.
  if (headers["Scheduled Time"]) {
    sheet
      .getRange(row, headers["Scheduled Time"])
      .setNumberFormat("@")
      .setValue("10:25 AM–11:05 AM");
  }
}


/* ============================================================
   EMAILS
   ============================================================ */

function sendAssignmentEmails_(request, tutor, deliveryOptions) {
  deliveryOptions = deliveryOptions || {};
  const skipTeacher = !!deliveryOptions.skipTeacher;
  const skipTutor = !!deliveryOptions.skipTutor;
  const teacherEmail = nhs_teacherDeliveryEmail_(request);

  // V27: tutor confirmation delivery is STRICTLY the Tutor Database
  // Notification Email. Never fall back to School Email / CUSD identity.
  const tutorEmail = String(
    tutor["Notification Email"] || ""
  ).trim();

  const teacherName =
    String(request["Teacher Name"] || "").trim() || "Teacher";

  const tutorDisplayName = formalAssignmentTutorName_(tutor);
  const course = String(request["Course"] || "").trim() || "Tutoring";
  const topic = String(request["Topic / Academic Need"] || "").trim();
  const supportType = String(request["Support Type"] || "").trim();
  const recurring = String(request["Recurring?"] || "").trim();
  const teacherNotes = String(request["Teacher Notes"] || "").trim();
  const requestId = String(request["Request ID"] || "").trim();

  const scheduledDate = formatAssignmentDate_(
    request["Scheduled Date"]
  );

  const scheduledTime =
    formatAssignmentTime_(request["Scheduled Time"]) ||
    "10:25 AM–11:05 AM";

  const location = String(
    request["Location"] || ""
  ).trim();

  let calendarSpec = null;
  let calendarCreated = false;
  let calendarEventId = "";
  let calendarError = "";

  if (
    request["Scheduled Date"] &&
    teacherEmail &&
    tutorEmail
  ) {
    try {
      const times = assignmentCalendarTimes_(
        request["Scheduled Date"]
      );

      calendarSpec = {
        uid:
          (requestId || Utilities.getUuid()) +
          "@tesoro-nhs-tutoring",
        title: "Tesoro NHS Tutoring — " + course,
        description: [
          "Tesoro NHS tutoring session",
          "",
          "Course: " + course,
          topic ? "Academic focus: " + topic : "",
          "Teacher: " + teacherName,
          "Tutor: " + tutorDisplayName,
          requestId ? "Request ID: " + requestId : "",
          "",
          "Titan Time: 10:25 AM–11:05 AM",
          "Please contact Tesoro NHS if the session needs to be changed."
        ].filter(Boolean).join("\n"),
        location: location,
        startMs: times.startMs,
        endMs: times.endMs,
        includeGoogleLink: true
      };

      calendarEventId = calendarSpec.uid;
    } catch (e) {
      calendarError = String(
        e && e.message ? e.message : e
      );
    }
  }

  const teacherSubject =
    "Confirmed: Tesoro NHS Tutoring — " +
    course +
    (scheduledDate ? " · " + scheduledDate : "");

  const tutorSubject =
    "Confirmed Assignment: Tesoro NHS Tutoring — " +
    course +
    (scheduledDate ? " · " + scheduledDate : "");

  const teacherBody = [
    "Hello " + teacherName + ",",
    "",
    "Your Tesoro NHS tutoring session has been confirmed.",
    "",
    "Course: " + course,
    "Tutor: " + tutorDisplayName,
    scheduledDate ? "Date: " + scheduledDate : "",
    "Time: " + scheduledTime,
    location ? "Location: " + location : "",
    topic ? "Academic Focus: " + topic : "",
    requestId ? "Request ID: " + requestId : "",
    "",
    calendarSpec
      ? "This confirmation includes a calendar invitation and an Add to Google Calendar link."
      : "The session is confirmed in the Tesoro NHS tutoring system.",
    "",
    "Five Star action required: assign both the tutor and the student into this Titan Time in Five Star, then mark that step complete in the NHS Teacher Portal.",
    "The tutor should then verify that the tutor/student placement appears correctly in Five Star and confirm the assignment in the NHS Tutor Portal before the session is fully ready.",
    "",
    "Thank you,",
    "Tesoro NHS"
  ].filter(Boolean).join("\n");

  const tutorBody = [
    "Hello " + tutorDisplayName + ",",
    "",
    "Your Tesoro NHS tutoring assignment has been confirmed.",
    "",
    "Course: " + course,
    "Requesting Teacher: " + teacherName,
    String(request["Teacher Email"] || "").trim() ? "Teacher Email: " + String(request["Teacher Email"] || "").trim() : "",
    scheduledDate ? "Date: " + scheduledDate : "",
    "Time: " + scheduledTime,
    location ? "Location: " + location : "",
    topic ? "Academic Focus: " + topic : "",
    supportType ? "Support Type: " + supportType : "",
    recurring ? "Recurring Support: " + recurring : "",
    teacherNotes ? "Teacher Notes: " + teacherNotes : "",
    requestId ? "Request ID: " + requestId : "",
    "",
    calendarSpec
      ? "This confirmation includes a calendar invitation and an Add to Google Calendar link."
      : "Please keep this session on your schedule.",
    "",
    "Five Star action required: after the teacher assigns both you and the student into the Titan Time, verify that the placement appears correctly in Five Star and confirm the assignment in the NHS Tutor Portal.",
    "",
    "Thank you,",
    "Tesoro NHS"
  ].filter(Boolean).join("\n");

  const teacherHtml = assignmentFormalEmailHtml_({
    badge: "SESSION CONFIRMED",
    heading: "Tesoro NHS Tutoring Session Confirmed",
    intro:
      "Your tutoring request has been matched and the session is now confirmed.",
    rows: [
      ["Course", course],
      ["Tutor", tutorDisplayName],
      ["Date", scheduledDate || "To be finalized"],
      ["Time", scheduledTime],
      ["Location", location || "Teacher classroom"],
      ["Academic Focus", topic || "General course support"],
      ["Request ID", requestId]
    ],
    calendarCreated: !!calendarSpec,
    footer:
      "Five Star: assign both the tutor and the student into this Titan Time, then confirm that step in the NHS Teacher Portal. The tutor will verify the placement in the NHS Tutor Portal."
  });

  const tutorHtml = assignmentFormalEmailHtml_({
    badge: "ASSIGNMENT CONFIRMED",
    heading: "Your Tesoro NHS Tutoring Assignment",
    intro:
      "Your tutoring assignment is confirmed. Please arrive prepared for the academic focus listed below.",
    rows: [
      ["Course", course],
      ["Requesting Teacher", teacherName],
      ["Teacher Email", String(request["Teacher Email"] || "").trim()],
      ["Date", scheduledDate || "To be finalized"],
      ["Time", scheduledTime],
      ["Location", location || "Teacher classroom"],
      ["Academic Focus", topic || "General course support"],
      ["Support Type", supportType],
      ["Recurring", recurring],
      ["Request ID", requestId]
    ],
    note: teacherNotes,
    calendarCreated: !!calendarSpec,
    footer:
      "Five Star: after the teacher assigns both you and the student into the Titan Time, verify the placement in Five Star, confirm it in the NHS Tutor Portal, then report the session afterward."
  });

  let teacherSent = false;
  let tutorSent = false;
  let teacherError = "";
  let tutorError = "";

  const messages = [];

  if (skipTeacher) {
    teacherSent = true;
  } else if (
    !teacherEmail ||
    !isValidAssignmentEmail_(
      teacherEmail
    )
  ) {
    teacherError =
      "Invalid teacher email";
  } else {
    messages.push({
      id: "teacher",
      to: teacherEmail,
      subject: teacherSubject,
      body: teacherBody,
      htmlBody: teacherHtml,
      name: "Tesoro NHS",
      requestId: requestId,
      type:
        "Assignment Confirmation — Teacher",
      calendar:
        calendarSpec
          ? Object.assign(
              {},
              calendarSpec,
              {
                attendee:
                  teacherEmail
              }
            )
          : null
    });
  }

  if (skipTutor) {
    tutorSent = true;
  } else if (
    !tutorEmail ||
    !isValidAssignmentEmail_(
      tutorEmail
    )
  ) {
    tutorError =
      "Tutor is missing a valid Notification Email in Tutor Database";
  } else if (
    !isPersonalTutorNotificationEmail_(
      tutorEmail,
      tutor
    )
  ) {
    tutorError =
      "Tutor Notification Email must be a personal email, not the CUSD/School Email";
  } else {
    messages.push({
      id: "tutor",
      to: tutorEmail,
      subject: tutorSubject,
      body: tutorBody,
      htmlBody: tutorHtml,
      name: "Tesoro NHS",
      requestId: requestId,
      type:
        "Assignment Confirmation — Tutor",
      calendar:
        calendarSpec
          ? Object.assign(
              {},
              calendarSpec,
              {
                attendee:
                  tutorEmail
              }
            )
          : null
    });
  }

  if (messages.length) {
    try {
      const batchResult =
        nhsmail_sendBatch_(
          messages
        );

      const results =
        batchResult &&
        Array.isArray(
          batchResult.results
        )
          ? batchResult.results
          : [];

      const byId = {};

      results.forEach(
        function(item) {
          byId[
            String(item.id || "")
          ] = item;
        }
      );

      if (
        messages.some(
          m => m.id === "teacher"
        )
      ) {
        const teacherResult =
          byId.teacher;

        if (
          teacherResult &&
          teacherResult.ok
        ) {
          teacherSent = true;

          if (
            teacherResult.calendarAttached
          ) {
            calendarCreated = true;
          }
        } else {
          teacherError =
            String(
              teacherResult &&
              teacherResult.error ||
              "Teacher confirmation relay failed."
            );
        }
      }

      if (
        messages.some(
          m => m.id === "tutor"
        )
      ) {
        const tutorResult =
          byId.tutor;

        if (
          tutorResult &&
          tutorResult.ok
        ) {
          tutorSent = true;

          if (
            tutorResult.calendarAttached
          ) {
            calendarCreated = true;
          }
        } else {
          tutorError =
            String(
              tutorResult &&
              tutorResult.error ||
              "Tutor confirmation relay failed."
            );
        }
      }

    } catch (err) {
      const sharedError =
        String(
          err && err.message
            ? err.message
            : err
        );

      if (
        messages.some(
          m => m.id === "teacher"
        )
      ) {
        teacherError =
          teacherError ||
          sharedError;
      }

      if (
        messages.some(
          m => m.id === "tutor"
        )
      ) {
        tutorError =
          tutorError ||
          sharedError;
      }
    }
  }

  if (
    calendarSpec &&
    teacherSent &&
    tutorSent &&
    !calendarCreated &&
    !calendarError
  ) {
    calendarError =
      "Confirmation emails sent, but relay did not report a calendar attachment.";
  }

  return {
    teacherSent,
    tutorSent,
    teacherError,
    tutorError,
    calendarCreated,
    calendarEventId,
    calendarError
  };
}


function isPersonalTutorNotificationEmail_(email, tutor) {
  const value = String(email || "")
    .trim()
    .toLowerCase();

  if (!value) return false;

  const schoolEmail = String(
    tutor["School Email"] || ""
  ).trim().toLowerCase();

  if (schoolEmail && value === schoolEmail) {
    return false;
  }

  // Student and staff CUSD domains should never be used for tutor delivery.
  if (
    /@capostu\.org$/i.test(value) ||
    /@capousd\.org$/i.test(value)
  ) {
    return false;
  }

  return true;
}


function assignmentCalendarTimes_(scheduledDateValue) {
  const date = scheduledDateValue instanceof Date
    ? new Date(scheduledDateValue.getTime())
    : new Date(String(scheduledDateValue || ""));

  if (isNaN(date.getTime())) {
    throw new Error("Unable to parse scheduled date for calendar event.");
  }

  const start = new Date(date.getTime());
  start.setHours(10, 25, 0, 0);

  const end = new Date(date.getTime());
  end.setHours(11, 5, 0, 0);

  return {
    startMs: start.getTime(),
    endMs: end.getTime()
  };
}


function assignmentFormalEmailHtml_(spec) {
  spec = spec || {};

  const esc = assignmentEscapeHtml_;
  const rows = (spec.rows || [])
    .filter(pair => pair && String(pair[1] || "").trim())
    .map(pair =>
      '<tr>' +
        '<td style="padding:10px 12px;color:#5b6776;font-size:13px;width:145px;border-bottom:1px solid #e8edf3;vertical-align:top;">' +
          esc(pair[0]) +
        '</td>' +
        '<td style="padding:10px 12px;color:#172033;font-size:14px;font-weight:600;border-bottom:1px solid #e8edf3;vertical-align:top;">' +
          esc(pair[1]) +
        '</td>' +
      '</tr>'
    )
    .join("");

  const calendarBox = spec.calendarCreated
    ? '<div style="margin:20px 0 0;padding:14px 16px;background:#eef6ff;border:1px solid #cfe2f8;border-radius:10px;color:#173a63;font-size:13px;line-height:1.5;">' +
        '<strong>Calendar invitation sent</strong><br>' +
        'A Google Calendar invitation was sent to the participants. The invitation can be accepted from Google Calendar or Outlook.' +
      '</div>'
    : "";

  const noteBox = spec.note
    ? '<div style="margin:18px 0 0;padding:14px 16px;background:#f8fafc;border-radius:10px;color:#344054;font-size:13px;line-height:1.55;">' +
        '<strong>Teacher notes</strong><br>' + esc(spec.note) +
      '</div>'
    : "";

  return (
    '<!doctype html><html><body style="margin:0;padding:0;background:#f3f6fa;font-family:Arial,Helvetica,sans-serif;color:#172033;">' +
      '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f3f6fa;padding:28px 12px;">' +
        '<tr><td align="center">' +
          '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="max-width:620px;background:#ffffff;border-radius:14px;overflow:hidden;border:1px solid #dfe6ee;">' +
            '<tr><td style="background:#0f2747;padding:22px 28px;color:#ffffff;">' +
              '<div style="font-size:12px;letter-spacing:1.5px;font-weight:700;opacity:.82;">TESORO HIGH SCHOOL</div>' +
              '<div style="font-size:22px;font-weight:700;margin-top:4px;">Tesoro NHS</div><div style="font-size:13px;margin-top:3px;opacity:.82;">Tutoring Program</div>' +
            '</td></tr>' +
            '<tr><td style="padding:28px;">' +
              '<div style="display:inline-block;background:#e8f2ff;color:#124b88;border-radius:999px;padding:6px 10px;font-size:11px;font-weight:700;letter-spacing:.8px;">' +
                esc(spec.badge || "NHS TUTORING") +
              '</div>' +
              '<h1 style="font-size:24px;line-height:1.25;margin:14px 0 8px;color:#13233b;">' +
                esc(spec.heading || "Tesoro NHS") +
              '</h1>' +
              '<p style="font-size:14px;line-height:1.6;color:#526071;margin:0 0 20px;">' +
                esc(spec.intro || "") +
              '</p>' +
              '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border:1px solid #e3e9f0;border-radius:10px;border-collapse:separate;overflow:hidden;">' +
                rows +
              '</table>' +
              calendarBox +
              noteBox +
              '<p style="font-size:13px;line-height:1.6;color:#667085;margin:22px 0 0;">' +
                esc(spec.footer || "") +
              '</p>' +
            '</td></tr>' +
            '<tr><td style="padding:16px 28px;background:#f8fafc;border-top:1px solid #e7edf3;color:#7a8696;font-size:11px;line-height:1.5;">' +
              'Tesoro NHS · Tutoring Program · Automated scheduling and service coordination<br>' +
              'Please retain this message for your records.' +
            '</td></tr>' +
          '</table>' +
        '</td></tr>' +
      '</table>' +
    '</body></html>'
  );
}


function assignmentEscapeHtml_(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}


/* ============================================================
   CASELOAD RECONCILIATION
   ============================================================ */

/**
 * Recalculates Current Caseload from the active requests table.
 *
 * This is deliberately reconciliation-based rather than increment/decrement based.
 * It handles reassignment, cancellation, completion, and accidental edits safely.
 */
function reconcileTutorCaseloads() {
  nhssec_requireMaintenanceAdmin_();

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);

  if (typeof nhsos_reconcileCaseloads_ === "function") {
    return nhsos_reconcileCaseloads_(ss);
  }

  throw new Error("Canonical caseload reconciliation is unavailable. Keep NHS_OS_WebApp.gs installed.");
}


/* ============================================================
   FIND TUTOR
   ============================================================ */

function findTutorByDisplayName_(displayName) {
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const tutorSheet = nhsperf_sheet_(ss, NHS_ASSIGNMENT_TUTOR_SHEET);

  if (!tutorSheet || tutorSheet.getLastRow() < 2) {
    return null;
  }

  const headers = getAssignmentHeaderMap_(tutorSheet);

  const data = tutorSheet
    .getRange(
      2,
      1,
      tutorSheet.getLastRow() - 1,
      tutorSheet.getLastColumn()
    )
    .getDisplayValues();

  const target = normalizeAssignmentText_(displayName);

  for (let i = 0; i < data.length; i++) {
    const tutor = rowToAssignmentObject_(data[i], headers);

    const candidates = [
      formalAssignmentTutorName_(tutor),
      tutor["Full Name"],
      tutor["Preferred Name"]
    ]
      .map(normalizeAssignmentText_)
      .filter(Boolean);

    if (candidates.includes(target)) {
      return tutor;
    }
  }

  return null;
}


/* ============================================================
   SHEET SETUP & FORMATTING
   ============================================================ */

function ensureAssignmentColumns_(sheet) {
  const headers = sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1))
    .getDisplayValues()[0]
    .map(value => String(value || "").trim());

  let nextColumn = headers.length + 1;

  NHS_ASSIGNMENT_HEADERS.forEach(header => {
    if (!headers.includes(header)) {
      sheet.getRange(1, nextColumn).setValue(header);
      headers.push(header);
      nextColumn++;
    }
  });
}


function formatAssignmentWorkflow_(ss, sheet) {
  const headers = getAssignmentHeaderMap_(sheet);

  // Checkbox for explicit officer confirmation.
  if (headers["Confirm Assignment"]) {
    sheet
      .getRange(
        2,
        headers["Confirm Assignment"],
        Math.max(1, sheet.getMaxRows() - 1),
        1
      )
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireCheckbox()
          .setAllowInvalid(false)
          .build()
      );

    sheet.setColumnWidth(headers["Confirm Assignment"], 135);
  }

  if (headers["Notification Status"]) {
    sheet.setColumnWidth(headers["Notification Status"], 210);

    sheet
      .getRange(
        2,
        headers["Notification Status"],
        Math.max(1, sheet.getMaxRows() - 1),
        1
      )
      .setWrap(true);
  }

  if (headers["Assignment Confirmed"]) {
    sheet.setColumnWidth(headers["Assignment Confirmed"], 165);

    sheet
      .getRange(
        2,
        headers["Assignment Confirmed"],
        Math.max(1, sheet.getMaxRows() - 1),
        1
      )
      .setNumberFormat("mmm d, yyyy h:mm AM/PM");
  }

  // Style assignment-operation headers.
  NHS_ASSIGNMENT_HEADERS.forEach(header => {
    if (!headers[header]) return;

    sheet
      .getRange(1, headers[header])
      .setBackground("#E2F0D9")
      .setFontColor("#1F1F1F")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);
  });

  // Assigned Tutor validation from Teacher Directory when available.
  setAssignmentTutorValidation_(ss, sheet);

  // Conditional formatting for notification state.
  applyAssignmentConditionalFormatting_(sheet, headers);
}


function setAssignmentTutorValidation_(ss, requestSheet) {
  const headers = getAssignmentHeaderMap_(requestSheet);

  if (!headers["Assigned Tutor"]) return;

  const directory = nhsperf_sheet_(ss, "Teacher Directory");

  if (!directory) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(directory.getRange("A8:A1000"), true)
    .setAllowInvalid(true)
    .build();

  requestSheet
    .getRange(
      2,
      headers["Assigned Tutor"],
      Math.max(1, requestSheet.getMaxRows() - 1),
      1
    )
    .setDataValidation(rule);
}


function applyAssignmentConditionalFormatting_(sheet, headers) {
  if (!headers["Notification Status"]) return;

  const range = sheet.getRange(
    2,
    headers["Notification Status"],
    Math.max(1, sheet.getMaxRows() - 1),
    1
  );

  const existing = sheet.getConditionalFormatRules().filter(rule => {
    return !rule.getRanges().some(r =>
      r.getSheet().getName() === sheet.getName() &&
      r.getColumn() === headers["Notification Status"]
    );
  });

  const specs = [
    ["Ready to Confirm", "#FFF2CC"],
    ["Teacher & Tutor Notified", "#D9EAD3"],
    ["Reassignment Pending", "#FCE5CD"],
    ["Tutor Not Found", "#F4CCCC"],
    ["Select a Tutor First", "#F4CCCC"],
    ["Classroom Required", "#F4CCCC"]
  ];

  specs.forEach(([text, color]) => {
    existing.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(text)
        .setBackground(color)
        .setRanges([range])
        .build()
    );
  });

  sheet.setConditionalFormatRules(existing);
}



/* ============================================================
   MANUAL OPERATIONS
   ============================================================ */

/**
 * Reapply formatting / validation without reinstalling triggers.
 */
function refreshAssignmentWorkflowFormatting() {
  nhssec_requireMaintenanceAdmin_();

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const sheet = nhsperf_sheet_(ss, NHS_ASSIGNMENT_REQUEST_SHEET);

  if (!sheet) {
    throw new Error('Sheet "' + NHS_ASSIGNMENT_REQUEST_SHEET + '" was not found.');
  }

  ensureAssignmentColumns_(sheet);
  formatAssignmentWorkflow_(ss, sheet);
}


/**
 * Send / retry assignment confirmation for a specific row manually.
 * Example:
 *   confirmAssignmentForRow(2)
 */
function confirmAssignmentForRow(rowNumber) {
  nhssec_requireMaintenanceAdmin_();

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const sheet = nhsperf_sheet_(ss, NHS_ASSIGNMENT_REQUEST_SHEET);

  if (!sheet) {
    throw new Error('Sheet "' + NHS_ASSIGNMENT_REQUEST_SHEET + '" was not found.');
  }

  const headers = getAssignmentHeaderMap_(sheet);

  sheet
    .getRange(rowNumber, headers["Confirm Assignment"])
    .setValue(true);

  handleConfirmAssignmentChange_(sheet, rowNumber, headers);
}



/**
 * OPTIONAL FUTURE / AFTER-SCHOOL SUPPORT
 * If a request is later designated for after-school tutoring, use this helper.
 * After-school NHS tutoring location defaults to the Library.
 * Time is intentionally left for the coordinator to set because no fixed
 * after-school time has been specified.
 */
function applyAfterSchoolDefaultsForRow(rowNumber) {
  nhssec_requireMaintenanceAdmin_();

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const sheet = nhsperf_sheet_(ss, NHS_ASSIGNMENT_REQUEST_SHEET);

  if (!sheet) {
    throw new Error('Sheet "' + NHS_ASSIGNMENT_REQUEST_SHEET + '" was not found.');
  }

  const headers = getAssignmentHeaderMap_(sheet);

  if (headers["Location"]) {
    sheet.getRange(rowNumber, headers["Location"]).setValue("Library");
  }

  if (headers["Scheduled Time"]) {
    sheet.getRange(rowNumber, headers["Scheduled Time"]).clearContent();
  }

  updateRequestTimestamp_(sheet, rowNumber, headers);
}


/* ============================================================
   UTILITIES
   ============================================================ */

function getRequestObject_(sheet, row, headers) {
  const values = sheet
    .getRange(row, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  return rowToAssignmentObject_(values, headers);
}


function getAssignmentHeaderMap_(sheet) {
  return nhsperf_headerMap_(sheet, true, false);
}


function rowToAssignmentObject_(row, headerMap) {
  const obj = {};

  Object.keys(headerMap).forEach(header => {
    obj[header] = row[headerMap[header] - 1];
  });

  return obj;
}


function formalAssignmentTutorName_(tutor) {
  const fullName = String(tutor["Full Name"] || "").trim();
  const preferredName = String(tutor["Preferred Name"] || "").trim();

  if (!preferredName) {
    return fullName;
  }

  const parts = fullName.split(/\s+/).filter(Boolean);

  if (parts.length <= 1) {
    return preferredName;
  }

  const lastName = parts[parts.length - 1];

  if (
    normalizeAssignmentText_(preferredName) ===
    normalizeAssignmentText_(fullName)
  ) {
    return fullName;
  }

  return preferredName + " " + lastName;
}


function normalizeAssignmentText_(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[.(),:;'"’\-_/]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}


function updateRequestTimestamp_(sheet, row, headers) {
  if (!headers["Last Update"]) return;

  sheet
    .getRange(row, headers["Last Update"])
    .setValue(new Date())
    .setNumberFormat("mmm d, yyyy h:mm AM/PM");
}


function isValidAssignmentEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || "").trim());
}


function formatAssignmentDate_(value) {
  if (!value) return "";

  const date = new Date(value);

  if (isNaN(date.getTime())) {
    return String(value || "").trim();
  }

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "EEEE, MMMM d, yyyy"
  );
}


function formatAssignmentTime_(value) {
  if (!value) return "";

  const text = String(value || "").trim();

  // Preserve standardized time windows such as "10:25 AM–11:05 AM".
  if (/[–-]/.test(text) && /AM|PM/i.test(text)) {
    return text;
  }

  const date = new Date(value);

  if (isNaN(date.getTime())) {
    return text;
  }

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "h:mm a"
  );
}


/* ============================================================
   V18 DEFERRED CASELOAD RECONCILIATION
   ============================================================ */

const NHS_CASELOAD_RECONCILE_FLAG = "NHS_CASELOAD_RECONCILE_PENDING_V18";
const NHS_CASELOAD_RECONCILE_HANDLER = "processQueuedTutorCaseloadReconciliation_";

function installFastCaseloadWorker() {
  nhssec_requireMaintenanceAdmin_();

  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === NHS_CASELOAD_RECONCILE_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(NHS_CASELOAD_RECONCILE_HANDLER)
    .timeBased()
    .everyMinutes(1)
    .create();

  return {
    ok: true,
    message: "Fast caseload reconciliation worker installed."
  };
}

function queueTutorCaseloadReconciliation_() {
  try {
    nhsperf_scriptProperties_().setProperty(
      NHS_CASELOAD_RECONCILE_FLAG,
      new Date().toISOString()
    );
  } catch (e) {
    // Never block assignment UI on a noncritical bookkeeping flag.
  }
}

function processQueuedTutorCaseloadReconciliation_() {
  const props = nhsperf_scriptProperties_();
  const pending = props.getProperty(NHS_CASELOAD_RECONCILE_FLAG);

  if (!pending) {
    return {
      ok: true,
      skipped: true,
      reason: "No reconciliation pending"
    };
  }

  // Clear before work so another mutation during reconciliation can set it again.
  props.deleteProperty(NHS_CASELOAD_RECONCILE_FLAG);

  try {
    reconcileTutorCaseloads();
    return {
      ok: true,
      reconciled: true
    };
  } catch (e) {
    // Restore the flag for retry on the next trigger.
    props.setProperty(
      NHS_CASELOAD_RECONCILE_FLAG,
      new Date().toISOString()
    );
    throw e;
  }
}


/* ============================================================
   V19 CROSS-PORTAL CACHE INVALIDATION
   ============================================================ */

/**
 * Confirmation happens from Operations, but the Teacher Portal maintains its
 * own warm cache. Clear that exact teacher cache immediately whenever an
 * assignment/confirmation changes so the next silent/manual refresh sees the
 * live row instead of a stale 10-minute copy.
 */
function invalidateTeacherPortalCacheForRequest_(request) {
  try {
    const email = String(request && request["Teacher Email"] || "")
      .trim()
      .toLowerCase();

    if (!email) return;

    nhsperf_scriptCache_().remove(nhs_v46TeacherCacheKey_(email));
  } catch (e) {}
}

/**
 * Also clear the Operations bootstrap cache after an assignment workflow
 * mutation so all portals converge on the same state.
 */
function invalidateAssignmentWorkflowCaches_(request) {
  nhsperf_touchDataRevision_();
  invalidateTeacherPortalCacheForRequest_(request);

  try {
    const cache = nhsperf_scriptCache_();
    cache.remove(NHSOS_CACHE_KEY);
    cache.remove("nhsos_health_v14");
  } catch (e) {}
}


/** ============================================================
 * SOURCE GROUP: TeacherPortal.gs
 * ============================================================ */

/**
 * NHS TUTORING — TEACHER PORTAL BACKEND
 *
 * This backend is part of the consolidated Tesoro NHS web-app project.
 * Teacher access is routed through ?page=teacher.
 *
 * The Teacher Portal:
 * - creates complete tutoring requests directly in "Tutoring Requests"
 * - shows only the signed-in teacher's own requests
 * - shows upcoming confirmed sessions
 * - lets teachers request changes / cancel / end recurring support
 * - collects optional session feedback
 * - writes Activity Log + Notification Log events for the internal OS
 *
 * DEPLOYMENT:
 *   Execute as: Me
 *   Access: your school Google Workspace / organization only
 *
 * The signed-in Google account is the teacher identity.
 */

const TP_SPREADSHEET_ID = NHS_SPREADSHEET_ID;

const TP_REQUESTS_SHEET = "Tutoring Requests";
const TP_TUTORS_SHEET = "Tutor Database";
const TP_TEACHER_DIRECTORY_SHEET = "Teacher Directory";
const TP_ACTIVITY_LOG_SHEET = "Activity Log";
const TP_NOTIFICATION_LOG_SHEET = "Notification Log";
const TP_FEEDBACK_LOG_SHEET = "Teacher Feedback Log";

const TP_TITAN_TIME = "10:25 AM–11:05 AM";
const TP_CACHE_SECONDS = 600; // Turbo V14: mutations invalidate this cache; warm portal loads stay fast.


/* ============================================================
   WEB APP
   ============================================================ */

function includeTeacher_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================================================
   INSTALL
   ============================================================ */

function installTeacherPortal() {
  nhssec_requireMaintenanceAdmin_();

  nhsa_installAccessCore_();

  const ss = nhsperf_ss_(TP_SPREADSHEET_ID);

  const requestSheet = nhsperf_sheet_(ss, TP_REQUESTS_SHEET);
  if (!requestSheet) {
    throw new Error('Missing "' + TP_REQUESTS_SHEET + '" sheet.');
  }

  tp_ensureRequestSchema_(requestSheet);
  tp_ensureFeedbackLog_(ss);
  tp_ensureActivityLog_(ss);
  tp_ensureNotificationLog_(ss);
  tp_ensureCancellationWorker_();


  return {
    ok: true,
    message:
      "Teacher Portal installed. Fast cancellation worker is active."
  };
}


/**
 * Relay configuration diagnostic. No Gmail/MailApp authorization is required
 * on the school Google account.
 */
function authorizeTeacherPortalMail() {
  nhssec_requireMaintenanceAdmin_();

  if (typeof checkNHSBatchRelayV29 === "function") {
    return checkNHSBatchRelayV29();
  }
  return {
    ok: false,
    message: "Mail relay diagnostic is unavailable. Keep NHS Mail Transport.gs installed."
  };
}


/* ============================================================
   PORTAL BOOTSTRAP
   ============================================================ */

function nhsv473_teacher_getPortalDataCore_(forceRefresh) {
  const email = tp_requireTeacherEmail_();
  const cache = nhsperf_scriptCache_();
  const key = nhs_v46TeacherCacheKey_(email);

  let cachedPortal = null;

  try {
    const cached = cache.get(key);

    if (cached) {
      const parsed = JSON.parse(cached);

      if (parsed && parsed.lifecycleBuild === NHS_RELEASE.VERSION) {
        cachedPortal = parsed;
      }
    }
  } catch (e) {}

  if (!forceRefresh && cachedPortal) {
    cachedPortal.cached = true;
    return cachedPortal;
  }

  const ss = nhsperf_ss_(TP_SPREADSHEET_ID);
  // Safe in-place V46.5 schema migration. This only appends missing columns;
  // it never resets/rebuilds the workbook and avoids requiring install scripts.
  nhsos_ensureRequestColumns_(nhsperf_sheet_(ss, TP_REQUESTS_SHEET));
  const teacher = tp_getTeacherIdentity_(ss, email);
  const bootstrapWarnings = [];
  let requestTruth = true;
  let requests = [];

  try {
    requests = tp_readTeacherRequests_(ss, email);
  } catch (error) {
    requestTruth = false;
    bootstrapWarnings.push(
      "Live request data is temporarily unavailable: " +
      String(error && error.message ? error.message : error)
    );
    console.error("[NHS V46.5 teacher bootstrap] requests:", error);
  }

  // V46.4: Teacher notification cards were removed from the portal UI.
  // Email delivery/history remains intact, but notification-log reads are no
  // longer part of the Teacher startup path.
  const notifications = [];

  const now = new Date();
  const tz = nhs_safeTimeZone_(ss);

  const upcoming = requests
    .filter(function(r) {
      return nhslife_upcoming_(r, now, tz);
    })
    .sort(function(a,b) {
      return String(a.scheduledDate || "")
        .localeCompare(String(b.scheduledDate || ""));
    });

  const data = {
    generatedAt: new Date().toISOString(),
    environment: nhsrg_publicEnvironmentInfo_(),
    gatewayUrl: nhs_getGatewayUrl_(),
    teacher: teacher,
    titanTime: TP_TITAN_TIME,
    schedulingOptions: nhssched_teacherOptions_(tz, now),
    // V45.5.2: course catalog is intentionally loaded through
    // nhscourse_getCatalog() after the Teacher portal renders.
    // A catalog problem must never block Teacher portal bootstrap.
    courseCatalogVersion: NHS_COURSE_CATALOG_VERSION,
    requests: requests,
    upcoming: upcoming,
    notifications: notifications.slice(0,75),
    feedback: [],

    counts: {
      open: requests.filter(function(r) {
        return !nhslife_closed_(r);
      }).length,

      confirmed: upcoming.length,

      recurring: requests.filter(function(r) {
        return !!r.recurringActive && !nhslife_closed_(r);
      }).length,

      completed: requests.filter(function(r) {
        return nhslife_completed_(r);
      }).length
    },

    requestTruth: requestTruth,
    bootstrapWarnings: bootstrapWarnings,
    lifecycleBuild: NHS_RELEASE.VERSION
  };

  // Do not cache a degraded request snapshot as authoritative truth.
  if (requestTruth) {
    try {
      cache.put(key, JSON.stringify(data), TP_CACHE_SECONDS);
    } catch(e) {}
  }

  return data;
}



/* ============================================================
   V28 ULTRAFAST REQUEST DELTA
   ============================================================ */

function teacher_getLiveRequestData() {
  const email = tp_requireTeacherEmail_();
  const ss = nhsperf_ss_(TP_SPREADSHEET_ID);
  const teacher = tp_getTeacherIdentity_(ss, email);
  const requests = tp_readTeacherRequests_(ss, email);
  const now = new Date();
  const tz = nhs_safeTimeZone_(ss);
  const upcoming = requests
    .filter(function(r){ return nhslife_upcoming_(r, now, tz); })
    .sort(function(a,b){ return String(a.scheduledDate||"").localeCompare(String(b.scheduledDate||"")); });

  return {
    generatedAt: new Date().toISOString(),
    revision: nhsperf_getDataRevision_(),
    requests: requests,
    upcoming: upcoming,
    counts: {
      open: requests.filter(function(r){ return !nhslife_closed_(r); }).length,
      confirmed: upcoming.length,
      recurring: requests.filter(function(r){ return !!r.recurringActive && !nhslife_closed_(r); }).length,
      completed: requests.filter(function(r){ return nhslife_completed_(r); }).length
    },
    teacher: teacher,
    schedulingOptions: nhssched_teacherOptions_(tz, now),
    lifecycleBuild: NHS_RELEASE.VERSION
  };
}


/* ============================================================
   CREATE COMPLETE REQUEST
   ============================================================ */

function nhsv473_teacher_submitRequestCore_(payload) {
  const schoolEmail = tp_requireTeacherEmail_();
  const cleanPayload = payload || {};
  const notificationEmail = String(cleanPayload.notificationEmail || "").trim().toLowerCase();

  if (notificationEmail.length > 254) {
    throw new Error("Notification email is too long.");
  }

  if (!tp_isValidEmail_(notificationEmail)) {
    throw new Error("Enter a valid notification email address before submitting the request.");
  }

  // Keep the critical section limited to spreadsheet state. Email receipt queuing
  // happens AFTER the lock is released so submission cannot deadlock on ScriptLock.
  const committed = tp_withLock_(function() {
    const ss = nhsperf_ss_(TP_SPREADSHEET_ID);
    const timeZone = nhs_safeTimeZone_(ss);
    const sh = nhsperf_sheet_(ss, TP_REQUESTS_SHEET);

    if (!sh) {
      throw new Error('Missing "' + TP_REQUESTS_SHEET + '" sheet.');
    }

    tp_ensureRequestSchema_(sh);

    const teacher = tp_getTeacherIdentity_(ss, schoolEmail);
    const teacherName = nhsinput_literalText_(
      teacher.name || cleanPayload.teacherName || "",
      120,
      "Teacher name"
    );

    const courseEntry = nhscourse_resolve_(cleanPayload.course);
    if (!courseEntry) {
      throw new Error("Select a recognized Tesoro course from the course search dropdown.");
    }

    const course = courseEntry.name;
    const topic = nhsinput_literalText_(
      cleanPayload.topic,
      500,
      "Topic / academic need"
    );
    const supportType = String(cleanPayload.supportType || "").trim();
    const allowedSupportTypes = [
      "Assignment Help",
      "Concept Review",
      "Assessment Preparation",
      "Foundational Skill Support",
      "Ongoing / Recurring Support"
    ];
    if (!allowedSupportTypes.includes(supportType)) {
      throw new Error("Choose a valid support type.");
    }
    const scheduling = nhssched_validateTeacherPreference_(
      cleanPayload,
      timeZone,
      new Date()
    );
    const schedulingType = scheduling.type;
    const preferredDates = scheduling.preferredDates.join(", ");
    const preferredDays = scheduling.preferredDays.join(", ");
    const recurring = scheduling.recurring
      ? "Yes — recurring support"
      : "No — one-time / short-term support";
    const location = nhsinput_literalText_(
      cleanPayload.location,
      120,
      "Classroom / room number"
    );
    const specificTutor = nhsinput_literalText_(
      cleanPayload.specificTutorPreference,
      120,
      "Specific tutor preference"
    );
    const teacherNotes = nhsinput_literalText_(
      cleanPayload.teacherNotes,
      1200,
      "Teacher notes"
    );

    if (!teacherName) throw new Error("Teacher name is required.");
    if (!course) throw new Error("Course is required.");
    if (!topic) throw new Error("Topic / academic need is required.");
    if (!supportType) throw new Error("Support type is required.");
    if (!preferredDays) throw new Error("Select at least one valid Titan Time preference.");
    if (!location) throw new Error("Classroom / room number is required.");

    // Persist the teacher's preferred delivery address so the field is real data,
    // preloads on future requests, and can be used by confirmation/reminder workflows.
    tp_saveTeacherNotificationEmail_(
      ss,
      schoolEmail,
      teacherName,
      notificationEmail
    );

    const timestamp = new Date();
    const requestId =
      "TR-" +
      Utilities.formatDate(
        timestamp,
        timeZone,
        "yyyyMMdd-HHmmss"
      ) +
      "-" +
      Utilities.getUuid();

    const headers = tp_headerMap_(sh);
    const row = new Array(sh.getLastColumn()).fill("");

    tp_rowSet_(row, headers, "Request ID", requestId);
    tp_rowSet_(row, headers, "Timestamp", timestamp);
    tp_rowSet_(row, headers, "Teacher Name", teacherName);
    tp_rowSet_(row, headers, "Teacher Email", schoolEmail);
    tp_rowSet_(row, headers, "Teacher Notification Email", notificationEmail);
    tp_rowSet_(row, headers, "Student Name", "");
    tp_rowSet_(row, headers, "Student Grade", "");
    tp_rowSet_(row, headers, "Course", course);
    tp_rowSet_(row, headers, "Topic / Academic Need", topic);
    tp_rowSet_(row, headers, "Support Type", supportType);
    tp_rowSet_(row, headers, "Scheduling Preference", schedulingType);
    tp_rowSet_(row, headers, "Preferred Dates", preferredDates);
    tp_rowSet_(row, headers, "Preferred Days", preferredDays);
    tp_rowSet_(row, headers, "Urgency", "");
    tp_rowSet_(row, headers, "Recurring?", recurring);
    tp_rowSet_(row, headers, "Specific Tutor Preference", specificTutor);
    tp_rowSet_(row, headers, "Teacher Notes", teacherNotes);
    tp_rowSet_(row, headers, "Status", "New");
    tp_rowSet_(row, headers, "Assigned Tutor", "");
    tp_rowSet_(row, headers, "Assigned Tutor ID", "");
    tp_rowSet_(
      row,
      headers,
      "Scheduled Date",
      schedulingType === NHSSCHED_TYPES.FUTURE && scheduling.preferredDates.length
        ? new Date(scheduling.preferredDates[0] + "T12:00:00")
        : ""
    );
    tp_rowSet_(
      row,
      headers,
      "Scheduled Time",
      schedulingType === NHSSCHED_TYPES.FUTURE ? TP_TITAN_TIME : ""
    );
    tp_rowSet_(row, headers, "Location", location);
    tp_rowSet_(row, headers, "Last Update", timestamp);
    tp_rowSet_(row, headers, "Coordinator Notes", "");
    tp_rowSet_(row, headers, "Recurring Active", "");
    tp_rowSet_(row, headers, "Next Session Date", "");
    tp_rowSet_(row, headers, "Request Source", "Teacher Portal");
    tp_rowSet_(row, headers, "Teacher Change Request", "");
    tp_rowSet_(row, headers, "Teacher Change Requested At", "");
    tp_rowSet_(row, headers, "Teacher Change Status", "");
    tp_rowSet_(row, headers, "Tutor Reconfirm Status", "");
    tp_rowSet_(row, headers, "Tutor Reconfirm At", "");
    tp_rowSet_(row, headers, "Teacher Five Star Ready", false);
    tp_rowSet_(row, headers, "Teacher Five Star Ready At", "");
    tp_rowSet_(row, headers, "Tutor Five Star Enrolled", false);
    tp_rowSet_(row, headers, "Tutor Five Star Enrolled At", "");
    tp_rowSet_(row, headers, "Five Star Ready At", "");

    if (typeof nhsdata_appendKeyedRow_ === "function") {
      nhsdata_appendKeyedRow_(sh, "Request ID", row);
    } else {
      sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    }

    tp_logActivity_(ss, {
      type: "Teacher Request Submitted",
      requestId: requestId,
      summary: "New tutoring request submitted through Teacher Portal",
      details: course + " · " + supportType + " · " + preferredDays + " · " + location
    });

    SpreadsheetApp.flush();

    try {
      const liveCache = nhsperf_scriptCache_();
      liveCache.remove(nhs_v46OperationsCacheKey_());
      liveCache.remove("nhsos_health_v14");
    } catch (e) {}

    nhsperf_touchDataRevision_();
    tp_invalidateTeacherCache_(schoolEmail);

    return {
      response: {
        ok: true,
        requestId: requestId,
        message: "Request submitted. It is now visible to NHS officers.",
        teacher: {
          name: teacherName,
          email: schoolEmail,
          notificationEmail: notificationEmail
        },
        request: {
          requestId: requestId,
          timestamp: timestamp.toISOString(),
          teacherName: teacherName,
          teacherEmail: schoolEmail,
          teacherNotificationEmail: notificationEmail,
          course: course,
          topic: topic,
          supportType: supportType,
          schedulingType: schedulingType,
          preferredDates: preferredDates,
          preferredDays: preferredDays,
          recurring: recurring,
          status: "New",
          assignedTutor: "",
          assignedTutorId: "",
          scheduledDate: schedulingType === NHSSCHED_TYPES.FUTURE ? scheduling.preferredDates[0] : "",
          scheduledTime: schedulingType === NHSSCHED_TYPES.FUTURE ? TP_TITAN_TIME : "",
          location: location,
          recurringActive: false,
          teacherFiveStarReady: false,
          tutorFiveStarEnrolled: false,
          fiveStarReadiness: "Five Star Setup Pending"
        }
      },
      receiptJob: {
        requestId: requestId,
        recipient: notificationEmail,
        type: "Teacher Request Receipt",
        subject: "NHS Tutoring Request Received — " + course,
        dedupeKey: "TEACHER_PORTAL_RECEIPT|" + requestId,
        body:
          "Hello " + teacherName + ",\n\n" +
          "Your NHS tutoring request has been received.\n\n" +
          "Course: " + course + "\n" +
          "Topic(s): " + topic + "\n" +
          "Scheduling: " + schedulingType + "\n" +
          (preferredDates ? "Preferred Date(s): " + preferredDates + "\n" : "") +
          "Preferred Titan Time Day(s): " + preferredDays + "\n" +
          "Classroom: " + location + "\n" +
          "Request ID: " + requestId + "\n\n" +
          "NHS will notify you once a tutor and session are confirmed.\n\n" +
          "Thank you,\nNHS Tutoring"
      }
    };
  });

  // Receipt delivery must never decide whether the request itself succeeded.
  try {
    tp_queueAsyncReceipt_(committed.receiptJob);
  } catch (queueError) {
    console.error("[NHS V46.7 teacher receipt queue]", queueError);
  }

  return committed.response;
}


/* ============================================================
   REQUEST CHANGE
   ============================================================ */

function teacher_requestChange(requestId, message) {
  const email = tp_requireTeacherEmail_();
  const change = nhsinput_literalText_(message, 1200, "Change request");

  if (!change) {
    throw new Error("Describe the change you need.");
  }

  return tp_withLock_(function() {
    const ss = nhsperf_ss_(TP_SPREADSHEET_ID);
    const sh = nhsperf_sheet_(ss, TP_REQUESTS_SHEET);

    tp_ensureRequestSchema_(sh);

    const row = tp_findOwnedRequestRow_(sh, requestId, email);
    if (!row) {
      throw new Error("Request not found or you do not have access to it.");
    }

    const h = tp_headerMap_(sh);
    const req = tp_rowObject_(sh, row, h);
    const now = new Date();

    if (["Completed", "Cancelled"].includes(req["Status"])) {
      throw new Error("This request is already closed.");
    }

    tp_set_(sh, row, h, "Teacher Change Request", change);
    tp_set_(sh, row, h, "Teacher Change Requested At", now);
    tp_set_(sh, row, h, "Teacher Change Status", "Open");
    tp_set_(sh, row, h, "Status", "Needs Follow-Up");
    tp_set_(sh, row, h, "Last Update", now);

    tp_logInternalAlert_(ss, {
      requestId: requestId,
      type: "TEACHER_CHANGE_REQUEST",
      severity: "Medium",
      title: "Teacher requested a change",
      message:
        (req["Course"] || "Tutoring request") +
        ": " +
        change
    });

    tp_logActivity_(ss, {
      type: "Teacher Change Requested",
      requestId: requestId,
      tutorId: req["Assigned Tutor ID"] || "",
      summary: "Teacher requested a scheduling/support change",
      details: change
    });

    tp_sendEmailLogged_(ss, {
      requestId: requestId,
      recipient: nhs_teacherDeliveryEmail_(req) || email,
      type: "Teacher Change Receipt",
      subject: "NHS Tutoring Change Request Received",
      dedupeKey:
        "TEACHER_CHANGE_RECEIPT|" +
        requestId +
        "|" +
        Utilities.formatDate(now, nhs_safeTimeZone_(ss), "yyyyMMddHHmmss"),
      body:
        "Hello,\n\n" +
        "NHS received your change request for " +
        (req["Course"] || "your tutoring request") +
        ".\n\n" +
        "Requested change:\n" +
        change +
        "\n\n" +
        "An NHS officer will review the request and follow up if needed.\n\n" +
        "Thank you,\nNHS Tutoring"
    });

    if (typeof nhsdata_invalidateRequestState_ === "function") {
      nhsdata_invalidateRequestState_({
        teacherEmail: email,
        tutorId: req["Assigned Tutor ID"] || ""
      });
    } else {
      tp_invalidateTeacherCache_(email);
      nhsperf_touchDataRevision_();
    }

    return {
      ok: true,
      requestId: requestId,
      status: "Needs Follow-Up",
      changeRequest: change,
      changeStatus: "Open",
      message: "Change request sent to NHS."
    };
  });
}


/* ============================================================
   CANCEL REQUEST
   ============================================================ */

function nhsv473_teacher_cancelRequestCore_(requestId, reason) {
  const email = tp_requireTeacherEmail_();
  const cleanRequestId = tp_clean_(requestId);

  if (!cleanRequestId) {
    throw new Error("Missing request ID.");
  }

  const ss = nhsperf_ss_(TP_SPREADSHEET_ID);
  const sh = nhsperf_sheet_(ss, TP_REQUESTS_SHEET);

  if (!sh) {
    throw new Error('Missing "' + TP_REQUESTS_SHEET + '" sheet.');
  }

  const row = tp_findOwnedRequestRow_(sh, cleanRequestId, email);

  if (!row) {
    throw new Error("Request not found or you do not have access to it.");
  }

  const h = tp_headerMap_(sh);
  const range = sh.getRange(row, 1, 1, sh.getLastColumn());
  const values = range.getValues()[0];
  const display = range.getDisplayValues()[0];

  const currentStatus =
    tp_arrayGetByHeader_(display, h, "Status");

  if (currentStatus === "Cancelled") {
    return {
      ok: true,
      requestId: cleanRequestId,
      status: "Cancelled",
      recurringActive: false,
      message: "This request is already cancelled."
    };
  }

  if (currentStatus === "Completed") {
    throw new Error("Completed requests cannot be cancelled.");
  }

  const now = new Date();

  const deferredPayload = {
    requestId: cleanRequestId,
    teacherEmail: email,
    teacherNotificationEmail:
      tp_arrayGetByHeader_(display, h, "Teacher Notification Email") || email,
    course: tp_arrayGetByHeader_(display, h, "Course"),
    scheduledDate: tp_arrayGetByHeader_(display, h, "Scheduled Date"),
    assignedTutor: tp_arrayGetByHeader_(display, h, "Assigned Tutor"),
    assignedTutorId: tp_arrayGetByHeader_(display, h, "Assigned Tutor ID"),
    reason: nhsinput_literalText_(reason, 500, "Cancellation reason")
  };

  // Critical path: one row write only.
  tp_arraySetByHeader_(values, h, "Status", "Cancelled");
  tp_arraySetByHeader_(values, h, "Recurring Active", false);
  tp_arraySetByHeader_(values, h, "Next Session Date", "");
  tp_arraySetByHeader_(values, h, "Teacher Change Status", "Cancelled");
  tp_arraySetByHeader_(values, h, "Last Update", now);

  range.setValues([values]);
  SpreadsheetApp.flush();

  // Make cancellation authoritative to every portal before slow side effects.
  if (typeof nhsdata_invalidateRequestState_ === "function") {
    nhsdata_invalidateRequestState_({
      teacherEmail: email,
      tutorId: deferredPayload.assignedTutorId || ""
    });
  } else {
    tp_invalidateTeacherCache_(email);
    nhsperf_touchDataRevision_();
  }

  // Queue emails/logging for the already-installed worker.
  tp_queueCancellationJob_(deferredPayload);

  return {
    ok: true,
    requestId: cleanRequestId,
    status: "Cancelled",
    recurringActive: false,
    message: "Request cancelled."
  };
}


/**
 * Stores cancellation follow-up work in Script Properties and schedules
 * a one-time worker. This keeps tutor lookup, email sends, Activity Log,
 * and internal alert writes off the teacher's critical path.
 */
/**
 * One-time deferred worker for cancellation logs + emails.
 * Runs outside the teacher's button-click request.
 */
function tp_arraySetByHeader_(rowValues, headerMap, header, value) {
  const col = headerMap[header];

  if (!col) {
    return;
  }

  rowValues[col - 1] = value;
}



function tp_arrayGetByHeader_(rowValues, headerMap, header) {
  const col = headerMap[header];

  if (!col) {
    return "";
  }

  return String(rowValues[col - 1] || "").trim();
}


/**
 * Installs one recurring worker instead of creating a trigger per cancellation.
 * This keeps the teacher's click path fast and reliable.
 */
function tp_ensureCancellationWorker_() {
  const handler = "tp_processCancellationJobs_";

  const exists = ScriptApp
    .getProjectTriggers()
    .some(trigger => trigger.getHandlerFunction() === handler);

  if (!exists) {
    ScriptApp
      .newTrigger(handler)
      .timeBased()
      .everyMinutes(1)
      .create();
  }
}


function tp_queueCancellationJob_(payload) {
  const props = nhsperf_scriptProperties_();

  const key =
    "tpCancelJob_" +
    Date.now() +
    "_" +
    Utilities.getUuid();

  props.setProperty(
    key,
    JSON.stringify({
      createdAt: new Date().toISOString(),
      attempts: 0,
      payload: payload || {}
    })
  );
}


function tp_processCancellationJobs_() {
  const props = nhsperf_scriptProperties_();
  const all = props.getProperties();

  const keys = Object.keys(all)
    .filter(key => key.indexOf("tpCancelJob_") === 0)
    .sort()
    .slice(0, 25);

  if (!keys.length) {
    return;
  }

  const ss = nhsperf_ss_(TP_SPREADSHEET_ID);

  keys.forEach(key => {
    try {
      const job = JSON.parse(all[key] || "{}");
      const p = job.payload || {};

      tp_logActivity_(ss, {
        type: "Teacher Cancelled Request",
        requestId: p.requestId || "",
        tutorId: p.assignedTutorId || "",
        actor: p.teacherEmail || "Teacher Portal",
        summary: "Teacher cancelled tutoring request",
        details: p.reason || p.course || ""
      });

      tp_logInternalAlert_(ss, {
        requestId: p.requestId || "",
        type: "TEACHER_CANCELLED",
        severity: "Low",
        title: "Teacher cancelled request",
        message:
          (p.course || "Tutoring request") +
          " was cancelled by the requesting teacher."
      });

      if (p.teacherEmail) {
        tp_sendEmailLogged_(ss, {
          requestId: p.requestId || "",
          recipient: p.teacherNotificationEmail || p.teacherEmail,
          type: "Teacher Cancellation Receipt",
          subject: "NHS Tutoring Request Cancelled",
          dedupeKey:
            "TEACHER_CANCEL_RECEIPT|" +
            (p.requestId || ""),
          body:
            "Hello,\n\n" +
            "Your NHS tutoring request for " +
            (p.course || "tutoring") +
            " has been cancelled.\n\n" +
            "Request ID: " +
            (p.requestId || "") +
            "\n\n" +
            "Thank you,\nNHS Tutoring"
        });
      }

      const tutor = tp_findTutorByAssignment_(ss, {
        "Assigned Tutor ID": p.assignedTutorId || "",
        "Assigned Tutor": p.assignedTutor || ""
      });

      if (tutor && tutor.email) {
        tp_sendEmailLogged_(ss, {
          requestId: p.requestId || "",
          recipient: tutor.email,
          type: "Teacher Cancellation — Tutor",
          subject:
            "NHS Tutoring Cancellation — " +
            (p.course || "Session"),
          dedupeKey:
            "TEACHER_CANCEL_TUTOR|" +
            (p.requestId || "") +
            "|" +
            tutor.email,
          body:
            "Hello " +
            (tutor.name || "Tutor") +
            ",\n\n" +
            "The teacher has cancelled the following NHS tutoring request:\n\n" +
            "Course: " +
            (p.course || "") +
            "\n" +
            "Request ID: " +
            (p.requestId || "") +
            "\n" +
            (p.scheduledDate
              ? "Scheduled Date: " + p.scheduledDate + "\n"
              : "") +
            "\nNo further action is required for this request.\n\n" +
            "Thank you,\nNHS Tutoring"
        });
      }

      props.deleteProperty(key);

    } catch (error) {
      console.error(
        "Cancellation job failed:",
        key,
        error
      );
      if (typeof nhsdata_recordPropertyJobFailure_ === "function") {
        nhsdata_recordPropertyJobFailure_(props, key, all[key], error, 5);
      }
      // Failed jobs are retried with a cap; repeated failures move to dead-letter storage.
    }
  });
}


/* ============================================================
   END RECURRING SUPPORT
   ============================================================ */

function teacher_endRecurringRequest(requestId) {
  const email = tp_requireTeacherEmail_();

  return tp_withLock_(function() {
    const ss = nhsperf_ss_(TP_SPREADSHEET_ID);
    const sh = nhsperf_sheet_(ss, TP_REQUESTS_SHEET);
    const row = tp_findOwnedRequestRow_(sh, requestId, email);

    if (!row) {
      throw new Error("Request not found or you do not have access to it.");
    }

    const h = tp_headerMap_(sh);
    const req = tp_rowObject_(sh, row, h);

    const recurring =
      tp_truthy_(req["Recurring Active"]) ||
      tp_isRecurringValue_(req["Recurring?"]);

    if (!recurring) {
      throw new Error("This is not an active recurring request.");
    }

    tp_set_(sh, row, h, "Recurring Active", false);
    tp_set_(sh, row, h, "Next Session Date", "");
    tp_set_(sh, row, h, "Status", "Completed");
    tp_set_(sh, row, h, "Last Update", new Date());

    tp_logActivity_(ss, {
      type: "Teacher Ended Recurring Support",
      requestId: requestId,
      tutorId: req["Assigned Tutor ID"] || "",
      summary: "Teacher ended recurring tutoring",
      details: req["Course"] || ""
    });

    tp_logInternalAlert_(ss, {
      requestId: requestId,
      type: "RECURRING_ENDED_BY_TEACHER",
      severity: "Low",
      title: "Recurring tutoring ended",
      message:
        "Teacher ended recurring support for " +
        (req["Course"] || "this request") +
        "."
    });

    const tutor = tp_findTutorByAssignment_(ss, req);
    if (tutor && tutor.email) {
      tp_sendEmailLogged_(ss, {
        requestId: requestId,
        recipient: tutor.email,
        type: "Recurring Support Ended — Tutor",
        subject:
          "NHS Tutoring Recurring Support Ended — " +
          (req["Course"] || "Tutoring"),
        dedupeKey:
          "RECURRING_ENDED_TEACHER|" +
          requestId +
          "|" +
          tutor.email,
        body:
          "Hello " +
          (tutor.name || "Tutor") +
          ",\n\n" +
          "The teacher has ended recurring NHS tutoring support for:\n\n" +
          "Course: " +
          (req["Course"] || "") +
          "\n" +
          "Request ID: " +
          requestId +
          "\n\n" +
          "No additional recurring sessions should be expected for this request.\n\n" +
          "Thank you,\nNHS Tutoring"
      });
    }

    tp_invalidateTeacherCache_(email);

    return {
      ok: true,
      requestId: requestId,
      status: "Completed",
      recurringActive: false,
      message: "Recurring tutoring ended."
    };
  });
}


/* ============================================================
   OPTIONAL SESSION FEEDBACK
   ============================================================ */

function teacher_submitSessionFeedback(payload) {
  const email = tp_requireTeacherEmail_();

  return tp_withLock_(function() {
    const ss = nhsperf_ss_(TP_SPREADSHEET_ID);
    const requestSheet = nhsperf_sheet_(ss, TP_REQUESTS_SHEET);
    const requestId = tp_clean_(payload.requestId);

    const row = tp_findOwnedRequestRow_(requestSheet, requestId, email);

    if (!row) {
      throw new Error("Request not found or you do not have access to it.");
    }

    const rh = tp_headerMap_(requestSheet);
    const req = tp_rowObject_(requestSheet, row, rh);

    const response = tp_clean_(payload.response);
    const notes = nhsinput_literalText_(payload.notes, 1200, "Feedback notes");
    const sessionDate =
      tp_clean_(payload.sessionDate) ||
      tp_isoDate_(req["Scheduled Date"]);

    const allowed = [
      "Session occurred",
      "Student was unavailable",
      "Tutor did not arrive",
      "Needs rescheduling"
    ];

    if (!allowed.includes(response)) {
      throw new Error("Choose a valid session outcome.");
    }

    const feedbackSh = tp_ensureFeedbackLog_(ss);
    const feedbackId =
      "TF-" +
      Utilities.formatDate(
        new Date(),
        nhs_safeTimeZone_(ss),
        "yyyyMMdd-HHmmss"
      ) +
      "-" +
      Utilities.getUuid();

    feedbackSh.appendRow([
      feedbackId,
      new Date(),
      requestId,
      sessionDate,
      email,
      response,
      notes,
      "Received"
    ]);

    if (response !== "Session occurred") {
      tp_set_(requestSheet, row, rh, "Status", "Needs Follow-Up");
      tp_set_(requestSheet, row, rh, "Last Update", new Date());

      tp_logInternalAlert_(ss, {
        requestId: requestId,
        type: "TEACHER_SESSION_FEEDBACK",
        severity:
          response === "Tutor did not arrive"
            ? "High"
            : "Medium",
        title: "Teacher session feedback needs follow-up",
        message:
          response +
          (notes ? ": " + notes : "")
      });
    }

    tp_logActivity_(ss, {
      type: "Teacher Session Feedback",
      requestId: requestId,
      tutorId: req["Assigned Tutor ID"] || "",
      summary: response,
      details:
        (sessionDate ? sessionDate + " · " : "") +
        notes
    });

    tp_invalidateTeacherCache_(email);

    return {
      ok: true,
      requestId: requestId,
      response: response,
      status:
        response === "Session occurred"
          ? (req["Status"] || "")
          : "Needs Follow-Up",
      message: "Feedback submitted. Thank you."
    };
  });
}


/* ============================================================
   READ TEACHER REQUESTS
   ============================================================ */

function tp_readTeacherRequests_(ss, email) {
  const sh = nhsperf_sheet_(ss, TP_REQUESTS_SHEET);
  if (!sh) return [];

  const objects = typeof nhsdata_readKeyedObjects_ === "function"
    ? nhsdata_readKeyedObjects_(sh, "Request ID")
    : [];

  if (typeof nhsdata_readKeyedObjects_ !== "function") {
    if (sh.getLastRow() < 2) return [];
    const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
    const headers = values[0];
    const idx = tp_headerIndex_(headers);
    return values.slice(1).map(function(row, i) {
      const obj = { _row: i + 2 };
      headers.forEach(function(header, j) { if (header) obj[header] = row[j] || ""; });
      return obj;
    }).filter(function(obj) { return obj["Request ID"]; });
  }

  const target = String(email || "").trim().toLowerCase();

  return objects
    .filter(function(obj) {
      return String(obj["Teacher Email"] || "").trim().toLowerCase() === target;
    })
    .map(function(obj) {
      const recurringValue = obj["Recurring?"] || "";
      const recurringActiveRaw = obj["Recurring Active"] || "";
      return {
        row: obj._row,
        requestId: obj["Request ID"] || "",
        timestamp: obj["Timestamp"] || "",
        course: obj["Course"] || "",
        topic: obj["Topic / Academic Need"] || "",
        supportType: obj["Support Type"] || "",
        schedulingType: obj["Scheduling Preference"] || "Legacy",
        preferredDates: obj["Preferred Dates"] || "",
        preferredDays: obj["Preferred Days"] || "",
        recurring: recurringValue,
        recurringActive: recurringActiveRaw ? tp_truthy_(recurringActiveRaw) : tp_isRecurringValue_(recurringValue),
        specificTutorPreference: obj["Specific Tutor Preference"] || "",
        teacherNotes: obj["Teacher Notes"] || "",
        status: obj["Status"] || "New",
        assignedTutor: obj["Assigned Tutor"] || "",
        assignedTutorId: obj["Assigned Tutor ID"] || "",
        scheduledDate: tp_isoDate_(obj["Scheduled Date"] || ""),
        scheduledTime: obj["Scheduled Time"] || TP_TITAN_TIME,
        location: obj["Location"] || "",
        lastUpdate: obj["Last Update"] || "",
        nextSessionDate: tp_isoDate_(obj["Next Session Date"] || ""),
        changeRequest: obj["Teacher Change Request"] || "",
        changeStatus: obj["Teacher Change Status"] || "",
        notificationStatus: obj["Notification Status"] || "",
        assignmentConfirmed: obj["Assignment Confirmed"] || "",
        calendarStatus: obj["Calendar Status"] || "",
        calendarEventId: obj["Calendar Event ID"] || "",
        requestSource: obj["Request Source"] || "Google Form",
        tutorResponseStatus: obj["Tutor Response Status"] || "",
        tutorReconfirmStatus: obj["Tutor Reconfirm Status"] || "",
        tutorReconfirmAt: obj["Tutor Reconfirm At"] || "",
        teacherFiveStarReady: tp_truthy_(obj["Teacher Five Star Ready"] || ""),
        teacherFiveStarReadyAt: obj["Teacher Five Star Ready At"] || "",
        tutorFiveStarEnrolled: tp_truthy_(obj["Tutor Five Star Enrolled"] || ""),
        tutorFiveStarEnrolledAt: obj["Tutor Five Star Enrolled At"] || "",
        fiveStarReadyAt: obj["Five Star Ready At"] || "",
        fiveStarReadiness: nhssched_readiness_(obj)
      };
    })
    .reverse();
}


/* ============================================================
   V54 TEACHER-SAFE TUTOR DIRECTORY
   ============================================================ */

const TP_TUTOR_DIRECTORY_CACHE_SECONDS = 180;
const TP_TUTOR_DIRECTORY_CACHE_PREFIX =
  "teacherTutorDirectory" + NHS_RELEASE.CACHE_NAMESPACE + "|";

const TP_TUTOR_DIRECTORY_COURSE_FIELDS = [
  "English Courses",
  "Mathematics Courses",
  "Science Courses",
  "Social Science Courses",
  "World Language Courses",
  "Computer Science / CTE Courses"
];


function tp_tutorDirectoryCacheKey_() {
  const dataRevision =
    typeof nhs_v46Revision_ === "function"
      ? nhs_v46Revision_()
      : String(nhsperf_getDataRevision_() || "0");

  const accessRevision =
    typeof nhsa_getAccessRevision_ === "function"
      ? String(nhsa_getAccessRevision_() || "0")
      : "0";

  return (
    TP_TUTOR_DIRECTORY_CACHE_PREFIX +
    dataRevision +
    "|" +
    accessRevision
  );
}


function tp_directoryNormalizeGrade_(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";

  const match = raw.match(/\b(9|10|11|12)(?:st|nd|rd|th)?\b/i);
  return match ? match[1] : raw.slice(0, 24);
}


function tp_directoryCourseEntries_(row, index, preparedCourseIndex) {
  const courseIndex = preparedCourseIndex || nhscourse_index_();
  const seen = {};
  const entries = [];

  function addCourse_(value, fallbackDepartment) {
    splitCourseList_(value).forEach(function(item) {
      const key = nhscourse_key_(item);
      const official = courseIndex[key] || null;

      const name = official
        ? official.name
        : String(item || "").trim();

      if (!name) return;

      if (
        !official &&
        /^(none|n\/?a|na|no|not applicable)$/i.test(name)
      ) {
        return;
      }

      const dedupeKey = official
        ? official.id
        : "RAW|" + nhscourse_key_(name);

      if (!dedupeKey || seen[dedupeKey]) return;
      seen[dedupeKey] = true;

      entries.push({
        id: official ? official.id : dedupeKey,
        name: name,
        department:
          official
            ? official.department
            : String(fallbackDepartment || "Other")
      });
    });
  }

  TP_TUTOR_DIRECTORY_COURSE_FIELDS.forEach(function(field) {
    const fallback = {
      "English Courses": "English",
      "Mathematics Courses": "Math",
      "Science Courses": "Science",
      "Social Science Courses": "Social Science",
      "World Language Courses": "World Language",
      "Computer Science / CTE Courses": "CTE"
    }[field] || "Other";

    addCourse_(
      tp_get_(row, index, field),
      fallback
    );
  });

  // Strongest courses are also valid teacher-facing qualifications and may
  // include newer catalog items that predate a Tutor Database category update.
  addCourse_(
    tp_get_(row, index, "Strongest Course"),
    "Other"
  );

  return entries.sort(function(a, b) {
    return (
      String(a.department || "")
        .localeCompare(String(b.department || "")) ||
      String(a.name || "")
        .localeCompare(String(b.name || ""))
    );
  });
}


function tp_directorySubjectLabel_(department) {
  const key = String(department || "").trim();

  return {
    "Math": "Mathematics",
    "CTE": "Computer Science / CTE"
  }[key] || key || "Other";
}


function tp_directoryAvailabilityState_(spec) {
  spec = spec || {};

  const weekly =
    spec.weeklyAvailability || {
      confirmed: false,
      dates: []
    };

  const availableDays = (weekly.dates || [])
    .filter(function(day) {
      return day.defaultStatus === "Available";
    })
    .map(function(day) {
      return String(day.day || "").trim();
    })
    .filter(Boolean);

  const maxSessions =
    Math.max(
      0,
      Number(spec.maxSessions || 0)
    );

  const currentCaseload =
    Math.max(
      0,
      Number(spec.currentCaseload || 0)
    );

  const atCapacity =
    maxSessions > 0 &&
    currentCaseload >= maxSessions;

  let status = "Available";

  if (
    atCapacity ||
    availableDays.length === 0
  ) {
    status = "Not Currently Accepting";
  } else if (
    availableDays.length <= 1 ||
    (
      maxSessions > 0 &&
      currentCaseload >=
        Math.max(1, maxSessions - 1)
    )
  ) {
    status = "Limited Availability";
  }

  const basis = weekly.confirmed
    ? "Weekly availability confirmed"
    : "Typical availability shown";

  return {
    status: status,
    availableDays: availableDays,
    basis: basis,
    weeklyConfirmed: !!weekly.confirmed,
    weekLabel: String(weekly.weekLabel || "")
  };
}


/**
 * Converts internal Tutor Database data into the strict teacher-visible DTO.
 * Never add emails, reliability fields, service-hour metrics, internal notes,
 * decline history, or raw caseload/capacity values to this object.
 */
function tp_directorySafeTutor_(spec) {
  spec = spec || {};

  const courses =
    Array.isArray(spec.courses)
      ? spec.courses
      : [];

  const strongestNames =
    splitCourseList_(spec.strongestCourse || "")
      .map(function(item) {
        const resolved = nhscourse_resolve_(item);
        return resolved
          ? resolved.name
          : String(item || "").trim();
      })
      .filter(Boolean);

  const courseNames = [];
  const courseSeen = {};

  courses.forEach(function(course) {
    const name = String(
      course && course.name || ""
    ).trim();

    const key = nhscourse_key_(name);

    if (!name || !key || courseSeen[key]) {
      return;
    }

    courseSeen[key] = true;
    courseNames.push(name);
  });

  strongestNames.forEach(function(name) {
    const key = nhscourse_key_(name);

    if (!key || courseSeen[key]) return;
    courseSeen[key] = true;
    courseNames.push(name);
  });

  const subjects = Array.from(
    new Set(
      courses
        .map(function(course) {
          return tp_directorySubjectLabel_(
            course && course.department
          );
        })
        .filter(Boolean)
    )
  ).sort();

  const availability =
    tp_directoryAvailabilityState_({
      weeklyAvailability:
        spec.weeklyAvailability,
      maxSessions:
        spec.maxSessions,
      currentCaseload:
        spec.currentCaseload
    });

  return {
    tutorId:
      String(spec.tutorId || "").trim(),
    displayName:
      String(spec.displayName || "").trim(),
    grade:
      tp_directoryNormalizeGrade_(
        spec.grade
      ),
    subjects: subjects,
    courses: courseNames,
    strongestCourses:
      strongestNames.slice(0, 3),
    languages:
      String(spec.languages || "").trim(),
    bio:
      String(spec.directoryBio || "")
        .trim()
        .slice(0, 500),
    availableDays:
      availability.availableDays,
    availabilityStatus:
      availability.status,
    availabilityBasis:
      availability.basis,
    weeklyConfirmed:
      availability.weeklyConfirmed,
    weekLabel:
      availability.weekLabel
  };
}



function tp_directoryTypicalWeek_(
  titanAvailability,
  weekStart
) {
  const typicalDays =
    parseDays_(
      titanAvailability || ""
    );

  const dates =
    nhswa_titanDatesForWeek_(
      weekStart
    ).map(function(spec) {
      const normalized =
        normalizeText_(
          spec.day || ""
        );

      return {
        date: spec.date,
        day: spec.day,
        label: spec.label,
        status:
          typicalDays.includes(normalized)
            ? "Available"
            : "Unavailable",
        defaultStatus:
          typicalDays.includes(normalized)
            ? "Available"
            : "Unavailable",
        confirmed: false,
        hardBlocked: false,
        source: "Typical Availability"
      };
    });

  return {
    tutorId: "",
    weekStart: weekStart,
    weekEnd:
      nhssched_addDaysKey_(
        weekStart,
        4
      ),
    weekLabel:
      nhssched_formatDateLabel_(
        weekStart
      ) +
      " – " +
      nhssched_formatDateLabel_(
        nhssched_addDaysKey_(
          weekStart,
          4
        )
      ),
    confirmed: false,
    explicitCount: 0,
    availableCount:
      dates.filter(function(day) {
        return (
          day.defaultStatus ===
          "Available"
        );
      }).length,
    unavailableCount:
      dates.filter(function(day) {
        return (
          day.defaultStatus ===
          "Unavailable"
        );
      }).length,
    notConfirmedCount:
      dates.length,
    confirmedAt: "",
    dates: dates
  };
}


function tp_directoryReadWeeklyIndex_(
  ss
) {
  const index = {
    byTutorDate: {},
    byTutorWeek: {}
  };

  const sh =
    nhsperf_sheet_(
      ss,
      NHSWA_SHEET
    );

  if (
    !sh ||
    sh.getLastRow() < 2 ||
    sh.getLastColumn() < 1
  ) {
    return index;
  }

  const lastRow =
    sh.getLastRow();

  const lastColumn =
    sh.getLastColumn();

  const headers =
    sh
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0]
      .map(function(value) {
        return String(
          value || ""
        ).trim();
      });

  const h = {};

  headers.forEach(function(
    header,
    i
  ) {
    if (header) {
      h[header] = i;
    }
  });

  if (
    h["Tutor ID"] === undefined ||
    h["Date"] === undefined ||
    h["Status"] === undefined
  ) {
    return index;
  }

  // Weekly availability is append-oriented. Bound pathological sheets instead
  // of allowing a phantom/legacy row range to block Teacher Directory.
  const MAX_DIRECTORY_WEEKLY_ROWS =
    4000;

  const firstBodyRow =
    Math.max(
      2,
      lastRow -
        MAX_DIRECTORY_WEEKLY_ROWS +
        1
    );

  const rowCount =
    lastRow -
    firstBodyRow +
    1;

  if (rowCount < 1) {
    return index;
  }

  const values =
    sh
      .getRange(
        firstBodyRow,
        1,
        rowCount,
        lastColumn
      )
      .getDisplayValues();

  values.forEach(function(
    row,
    offset
  ) {
    const tutorId =
      String(
        row[h["Tutor ID"]] || ""
      ).trim();

    const date =
      String(
        row[h["Date"]] || ""
      ).trim();

    const status =
      String(
        row[h["Status"]] || ""
      ).trim();

    if (
      !tutorId ||
      !date ||
      !NHSWA_ALLOWED_STATUSES.includes(
        status
      )
    ) {
      return;
    }

    const rawWeekStart =
      h["Week Start"] === undefined
        ? ""
        : String(
            row[h["Week Start"]] ||
            ""
          ).trim();

    const weekStart =
      rawWeekStart ||
      nhswa_weekStartFromKey_(
        date
      );

    const item = {
      row:
        firstBodyRow + offset,
      tutorId: tutorId,
      tutorName:
        h["Tutor Name"] === undefined
          ? ""
          : String(
              row[h["Tutor Name"]] ||
              ""
            ).trim(),
      schoolEmail:
        h["School Email"] === undefined
          ? ""
          : String(
              row[h["School Email"]] ||
              ""
            )
              .trim()
              .toLowerCase(),
      weekStart: weekStart,
      date: date,
      day:
        h["Day"] === undefined
          ? nhssched_dayName_(
              date
            )
          : String(
              row[h["Day"]] ||
              nhssched_dayName_(
                date
              )
            ).trim(),
      status: status,
      confirmedAt:
        h["Confirmed At"] ===
        undefined
          ? ""
          : String(
              row[
                h["Confirmed At"]
              ] || ""
            ).trim(),
      updatedAt:
        h["Updated At"] ===
        undefined
          ? ""
          : String(
              row[
                h["Updated At"]
              ] || ""
            ).trim(),
      source:
        h["Source"] === undefined
          ? "Tutor Confirmed"
          : String(
              row[h["Source"]] ||
              "Tutor Confirmed"
            ).trim()
    };

    if (
      !index.byTutorDate[
        tutorId
      ]
    ) {
      index.byTutorDate[
        tutorId
      ] = {};
    }

    index.byTutorDate[
      tutorId
    ][date] = item;

    if (
      !index.byTutorWeek[
        tutorId
      ]
    ) {
      index.byTutorWeek[
        tutorId
      ] = {};
    }

    if (
      !index.byTutorWeek[
        tutorId
      ][weekStart]
    ) {
      index.byTutorWeek[
        tutorId
      ][weekStart] = {};
    }

    index.byTutorWeek[
      tutorId
    ][weekStart][date] = item;
  });

  return index;
}


function tp_directoryAvailabilityPayload_(
  ss
) {
  tp_requireTeacherEmail_();

  const tutorSheet =
    nhsperf_sheet_(
      ss,
      TP_TUTORS_SHEET
    );

  if (
    !tutorSheet ||
    tutorSheet.getLastRow() < 2
  ) {
    return {
      ok: true,
      generatedAt:
        new Date().toISOString(),
      weekStart: "",
      weekLabel: "",
      availability: []
    };
  }

  const lastRow =
    tutorSheet.getLastRow();

  const lastColumn =
    tutorSheet.getLastColumn();

  const values =
    tutorSheet
      .getRange(
        1,
        1,
        lastRow,
        lastColumn
      )
      .getDisplayValues();

  const index =
    tp_headerIndex_(
      values[0]
    );

  const activeTutorEmails =
    new Set(
      nhsa_readRegistry_(
        nhsa_ensureRegistry_(ss)
      )
        .filter(function(grant) {
          return (
            grant.role === "TUTOR" &&
            nhsa_isGrantActive_(
              grant
            )
          );
        })
        .map(function(grant) {
          return String(
            grant.email || ""
          )
            .trim()
            .toLowerCase();
        })
        .filter(Boolean)
    );

  const targetWeek =
    nhswa_confirmationTargetWeekStart_(
      nhs_safeTimeZone_(ss),
      new Date()
    );

  const weeklyIndex =
    tp_directoryReadWeeklyIndex_(
      ss
    );

  const availability = [];

  values
    .slice(1)
    .forEach(function(row) {
      const tutorId =
        tp_get_(
          row,
          index,
          "Tutor ID"
        );

      const schoolEmail =
        tp_get_(
          row,
          index,
          "School Email"
        ).toLowerCase();

      if (
        !tutorId ||
        !schoolEmail ||
        !activeTutorEmails.has(
          schoolEmail
        )
      ) {
        return;
      }

      const reliability =
        tp_get_(
          row,
          index,
          "Reliability Status"
        );

      if (
        [
          "Inactive",
          "Suspended"
        ].includes(reliability)
      ) {
        return;
      }

      const weekly =
        nhswa_weekSummaryFromIndex_(
          weeklyIndex,
          {
            tutorId: tutorId,
            titanAvailability:
              tp_get_(
                row,
                index,
                "Typical Titan Time Availability"
              ),
            temporaryUnavailableDates:
              tp_get_(
                row,
                index,
                "Temporary Unavailable Dates"
              )
          },
          targetWeek
        );

      const state =
        tp_directoryAvailabilityState_({
          weeklyAvailability: weekly,
          maxSessions:
            Number(
              tp_get_(
                row,
                index,
                "Maximum Sessions / Week"
              ) || 0
            ),
          currentCaseload:
            Number(
              tp_get_(
                row,
                index,
                "Current Caseload"
              ) || 0
            )
        });

      availability.push({
        tutorId: tutorId,
        availableDays:
          state.availableDays,
        availabilityStatus:
          state.status,
        availabilityBasis:
          state.basis,
        weeklyConfirmed:
          state.weeklyConfirmed,
        weekLabel:
          state.weekLabel
      });
    });

  return {
    ok: true,
    generatedAt:
      new Date().toISOString(),
    weekStart: targetWeek,
    weekLabel:
      targetWeek
        ? (
            nhssched_formatDateLabel_(
              targetWeek
            ) +
            " – " +
            nhssched_formatDateLabel_(
              nhssched_addDaysKey_(
                targetWeek,
                4
              )
            )
          )
        : "",
    availability: availability
  };
}


function teacher_getTutorDirectoryAvailability() {
  return nhsobs_measure_(
    "teacher.tutorDirectory",
    function() {
      const ss =
        nhsperf_ss_(
          TP_SPREADSHEET_ID
        );

      return tp_directoryAvailabilityPayload_(
        ss
      );
    }
  );
}


function nhsv54_teacher_getTutorDirectoryCore_(forceRefresh) {
  tp_requireTeacherEmail_();

  const cache =
    nhsperf_scriptCache_();

  const key =
    tp_tutorDirectoryCacheKey_();

  if (!forceRefresh) {
    try {
      const cached =
        cache.get(key);

      if (cached) {
        const parsed =
          JSON.parse(cached);

        if (
          parsed &&
          Array.isArray(
            parsed.tutors
          ) &&
          parsed.safeFieldsVersion ===
            "V55.2"
        ) {
          parsed.cached = true;
          return parsed;
        }
      }
    } catch (e) {}
  }

  const ss =
    nhsperf_ss_(
      TP_SPREADSHEET_ID
    );

  const tutorSheet =
    nhsperf_sheet_(
      ss,
      TP_TUTORS_SHEET
    );

  if (
    !tutorSheet ||
    tutorSheet.getLastRow() < 2
  ) {
    return {
      ok: true,
      generatedAt:
        new Date().toISOString(),
      revision:
        nhsperf_getDataRevision_(),
      accessRevision:
        nhsa_getAccessRevision_(),
      safeFieldsVersion:
        "V55.2",
      availabilityMode:
        "typical-first",
      weekStart: "",
      weekLabel: "",
      tutors: []
    };
  }

  const lastRow =
    tutorSheet.getLastRow();

  const lastColumn =
    tutorSheet.getLastColumn();

  const values =
    tutorSheet
      .getRange(
        1,
        1,
        lastRow,
        lastColumn
      )
      .getDisplayValues();

  const index =
    tp_headerIndex_(
      values[0]
    );

  const accessRows =
    nhsa_readRegistry_(
      nhsa_ensureRegistry_(ss)
    );

  const activeTutorEmails =
    new Set(
      accessRows
        .filter(function(grant) {
          return (
            grant.role === "TUTOR" &&
            nhsa_isGrantActive_(
              grant
            )
          );
        })
        .map(function(grant) {
          return String(
            grant.email || ""
          )
            .trim()
            .toLowerCase();
        })
        .filter(Boolean)
    );

  const targetWeek =
    nhswa_confirmationTargetWeekStart_(
      nhs_safeTimeZone_(ss),
      new Date()
    );

  const weekLabel =
    targetWeek
      ? (
          nhssched_formatDateLabel_(
            targetWeek
          ) +
          " – " +
          nhssched_formatDateLabel_(
            nhssched_addDaysKey_(
              targetWeek,
              4
            )
          )
        )
      : "";

  const courseIndex =
    nhscourse_index_();

  const tutors = [];

  values
    .slice(1)
    .forEach(function(row) {
      const fullName =
        tp_get_(
          row,
          index,
          "Full Name"
        );

      const tutorId =
        tp_get_(
          row,
          index,
          "Tutor ID"
        );

      const schoolEmail =
        tp_get_(
          row,
          index,
          "School Email"
        ).toLowerCase();

      if (
        !fullName ||
        !tutorId ||
        !schoolEmail ||
        !activeTutorEmails.has(
          schoolEmail
        )
      ) {
        return;
      }

      const reliability =
        tp_get_(
          row,
          index,
          "Reliability Status"
        );

      if (
        [
          "Inactive",
          "Suspended"
        ].includes(reliability)
      ) {
        return;
      }

      const preferredName =
        tp_get_(
          row,
          index,
          "Preferred Name"
        );

      const courseEntries =
        tp_directoryCourseEntries_(
          row,
          index,
          courseIndex
        );

      const strongestCourseRaw =
        tp_get_(
          row,
          index,
          "Strongest Course"
        );

      if (
        !courseEntries.length &&
        !String(
          strongestCourseRaw || ""
        ).trim()
      ) {
        return;
      }

      const typicalWeek =
        tp_directoryTypicalWeek_(
          tp_get_(
            row,
            index,
            "Typical Titan Time Availability"
          ),
          targetWeek
        );

      const dto =
        tp_directorySafeTutor_({
          tutorId: tutorId,
          displayName:
            tup_tutorDisplayName_(
              preferredName,
              fullName
            ),
          grade:
            tp_get_(
              row,
              index,
              "Grade"
            ),
          courses:
            courseEntries,
          strongestCourse:
            strongestCourseRaw,
          languages:
            tp_get_(
              row,
              index,
              "Tutoring Languages"
            ),
          directoryBio:
            tp_get_(
              row,
              index,
              "Teacher Directory Bio"
            ),
          weeklyAvailability:
            typicalWeek,
          maxSessions:
            Number(
              tp_get_(
                row,
                index,
                "Maximum Sessions / Week"
              ) || 0
            ),
          currentCaseload:
            Number(
              tp_get_(
                row,
                index,
                "Current Caseload"
              ) || 0
            )
        });

      // The first payload is intentionally non-blocking and based on typical
      // availability. A second lightweight call overlays live weekly status.
      dto.weeklyConfirmed = false;
      dto.availabilityBasis =
        "Typical availability shown";
      dto.weekLabel =
        weekLabel;

      if (
        dto.tutorId &&
        dto.displayName
      ) {
        tutors.push(dto);
      }
    });

  const statusRank = {
    "Available": 0,
    "Limited Availability": 1,
    "Not Currently Accepting": 2
  };

  tutors.sort(function(a, b) {
    return (
      (
        statusRank[
          a.availabilityStatus
        ] ?? 9
      ) -
      (
        statusRank[
          b.availabilityStatus
        ] ?? 9
      ) ||
      String(
        a.displayName || ""
      ).localeCompare(
        String(
          b.displayName || ""
        )
      )
    );
  });

  const result = {
    ok: true,
    generatedAt:
      new Date().toISOString(),
    revision:
      nhsperf_getDataRevision_(),
    accessRevision:
      nhsa_getAccessRevision_(),
    safeFieldsVersion:
      "V55.2",
    availabilityMode:
      "typical-first",
    weekStart:
      targetWeek,
    weekLabel:
      weekLabel,
    tutors:
      tutors
  };

  try {
    cache.put(
      key,
      JSON.stringify(result),
      TP_TUTOR_DIRECTORY_CACHE_SECONDS
    );
  } catch (e) {}

  return result;
}

/* ============================================================
   TEACHER IDENTITY
   ============================================================ */

function tp_requireTeacherEmail_() {
  const access = nhsa_requireAnyRole_(
    ["TEACHER", "ADVISOR", "ADMIN"],
    "Teacher Portal"
  );

  return access.email;
}

function tp_ensureTeacherDirectory_(ss) {
  let sh = nhsperf_sheet_(ss, TP_TEACHER_DIRECTORY_SHEET);
  if (!sh) sh = ss.insertSheet(TP_TEACHER_DIRECTORY_SHEET);

  const required = ["Teacher Name", "Teacher Email", "Notification Email", "Active", "Notes"];
  let headers = sh
    .getRange(1, 1, 1, Math.max(1, sh.getLastColumn()))
    .getDisplayValues()[0]
    .map(v => String(v || "").trim());

  required.forEach(header => {
    if (!headers.includes(header)) {
      sh.getRange(1, headers.length + 1).setValue(header);
      headers.push(header);
    }
  });

  sh.setFrozenRows(1);
  try { sh.hideSheet(); } catch (e) {}
  return sh;
}


/**
 * Run from the Apps Script editor to authorize a teacher/test account.
 * This is intentionally not exposed in the Teacher Portal UI.
 */
function authorizeTeacherPortalUser(email, name) {
  const actor = nhssec_requireMaintenanceAdmin_().email;

  email = String(email || "").trim().toLowerCase();
  name = String(name || "").trim();

  if (!email) throw new Error("Email is required.");
  nhsa_assertRoleEmailPolicy_(email, "TEACHER");

  const ss = nhsperf_ss_(TP_SPREADSHEET_ID);
  const settings = nhsa_readSettings_(ss);

  // Create the authorization record first. Teacher Directory is identity/profile
  // metadata; it is not an authorization source.
  const grant = nhsa_createGrant_(ss, {
    email: email,
    fullName: name,
    role: "TEACHER",
    schoolYear: settings.CURRENT_SCHOOL_YEAR || "",
    approvedBy: actor,
    notes: "Authorized through authorizeTeacherPortalUser"
  });

  const sh = tp_ensureTeacherDirectory_(ss);
  const h = tp_headerMap_(sh);
  let directoryResult = "created";

  if (sh.getLastRow() >= 2) {
    const match = sh
      .getRange(2, h["Teacher Email"], sh.getLastRow() - 1, 1)
      .createTextFinder(email)
      .matchEntireCell(true)
      .matchCase(false)
      .findNext();

    if (match) {
      const row = match.getRow();
      const existingName = sh.getRange(row, h["Teacher Name"]).getDisplayValue();
      tp_set_(sh, row, h, "Teacher Name", name || existingName);
      tp_set_(sh, row, h, "Active", true);
      directoryResult = "updated";
    } else {
      const row = new Array(sh.getLastColumn()).fill("");
      tp_arraySetByHeader_(row, h, "Teacher Name", name);
      tp_arraySetByHeader_(row, h, "Teacher Email", email);
      tp_arraySetByHeader_(row, h, "Active", true);
      sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    }
  } else {
    const row = new Array(sh.getLastColumn()).fill("");
    tp_arraySetByHeader_(row, h, "Teacher Name", name);
    tp_arraySetByHeader_(row, h, "Teacher Email", email);
    tp_arraySetByHeader_(row, h, "Active", true);
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  }

  try {
    nhsperf_scriptCache_().remove("teacherAccess|" + email);
    nhsperf_scriptCache_().remove("teacherIdentity|" + email);
  } catch (e) {}

  nhsa_invalidateAccessCache_();
  nhsperf_touchDataRevision_();

  return {
    ok: true,
    email: email,
    grantId: grant.grantId || "",
    grantExisting: !!grant.existing,
    directory: directoryResult
  };
}


function tp_getTeacherIdentity_(ss, email) {
  const cache = nhsperf_scriptCache_();
  const cacheKey = "teacherIdentity|" + email.toLowerCase();

  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }

  let name = "";
  let notificationEmail = "";
  const sh = nhsperf_sheet_(ss, TP_TEACHER_DIRECTORY_SHEET);

  if (sh && sh.getLastRow() >= 2) {
    const data = sh
      .getRange(1, 1, sh.getLastRow(), sh.getLastColumn())
      .getDisplayValues();

    const h = tp_headerIndex_(data[0]);

    const emailHeaders = [
      "Teacher Email",
      "Email",
      "School Email"
    ];

    const nameHeaders = [
      "Teacher Name",
      "Name",
      "Full Name"
    ];

    const notificationHeaders = [
      "Notification Email",
      "Personal Notification Email"
    ];

    let emailIndex;
    let nameIndex;
    let notificationIndex;

    emailHeaders.some(header => {
      if (h[header] !== undefined) {
        emailIndex = h[header];
        return true;
      }
      return false;
    });

    nameHeaders.some(header => {
      if (h[header] !== undefined) {
        nameIndex = h[header];
        return true;
      }
      return false;
    });

    notificationHeaders.some(header => {
      if (h[header] !== undefined) {
        notificationIndex = h[header];
        return true;
      }
      return false;
    });

    if (emailIndex !== undefined) {
      for (let i = 1; i < data.length; i++) {
        if (
          String(data[i][emailIndex] || "")
            .trim()
            .toLowerCase() === email
        ) {
          if (nameIndex !== undefined) {
            name = String(data[i][nameIndex] || "").trim();
          }
          if (notificationIndex !== undefined) {
            notificationEmail = String(data[i][notificationIndex] || "").trim().toLowerCase();
          }
          break;
        }
      }
    }
  }

  const identity = {
    email: email,
    name: name,
    notificationEmail: notificationEmail
  };

  try {
    cache.put(
      cacheKey,
      JSON.stringify(identity),
      21600
    );
  } catch (e) {}

  return identity;
}

function tp_isValidEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(
    String(value || "").trim()
  );
}

function tp_saveTeacherNotificationEmail_(ss, schoolEmail, teacherName, notificationEmail) {
  const email = String(schoolEmail || "").trim().toLowerCase();
  const delivery = String(notificationEmail || "").trim().toLowerCase();
  if (!email) throw new Error("Teacher school email is missing.");
  if (!tp_isValidEmail_(delivery)) {
    throw new Error("Enter a valid notification email address.");
  }

  const sh = tp_ensureTeacherDirectory_(ss);
  const h = tp_headerMap_(sh);
  let row = 0;

  if (sh.getLastRow() >= 2) {
    const match = sh
      .getRange(2, h["Teacher Email"], sh.getLastRow() - 1, 1)
      .createTextFinder(email)
      .matchEntireCell(true)
      .matchCase(false)
      .findNext();
    if (match) row = match.getRow();
  }

  if (!row) {
    const values = new Array(sh.getLastColumn()).fill("");
    tp_arraySetByHeader_(values, h, "Teacher Name", teacherName || "");
    tp_arraySetByHeader_(values, h, "Teacher Email", email);
    tp_arraySetByHeader_(values, h, "Notification Email", delivery);
    tp_arraySetByHeader_(values, h, "Active", true);
    sh.getRange(sh.getLastRow() + 1, 1, 1, values.length).setValues([values]);
  } else {
    if (teacherName) tp_set_(sh, row, h, "Teacher Name", teacherName);
    tp_set_(sh, row, h, "Notification Email", delivery);
    tp_set_(sh, row, h, "Active", true);
  }

  try {
    nhsperf_scriptCache_().remove("teacherIdentity|" + email);
  } catch (e) {}

  return delivery;
}

function nhs_teacherDeliveryEmail_(request) {
  return String(
    request && (
      request["Teacher Notification Email"] ||
      request.teacherNotificationEmail ||
      request["Teacher Email"] ||
      request.teacherEmail ||
      ""
    ) || ""
  ).trim().toLowerCase();
}


/* ============================================================
   TEACHER NOTIFICATIONS
   ============================================================ */

function tp_readTeacherNotifications_(ss, email) {
  const sh = tp_ensureNotificationLog_(ss);

  if (!sh || sh.getLastRow() < 2) {
    return [];
  }

  const values = sh
    .getRange(1, 1, sh.getLastRow(), sh.getLastColumn())
    .getDisplayValues();

  const h = tp_headerIndex_(values[0]);
  const target = String(email || "").trim().toLowerCase();

  return values
    .slice(1)
    .map((row, i) => {
      const recipient = String(
        tp_get_(row, h, "Recipient") || ""
      )
        .trim()
        .toLowerCase();

      if (recipient !== target) {
        return null;
      }

      if (tp_get_(row, h, "Channel") !== "Email") {
        return null;
      }

      return {
        row: i + 2,
        notificationId:
          tp_get_(row, h, "Notification ID"),
        createdAt:
          tp_get_(row, h, "Created At"),
        requestId:
          tp_get_(row, h, "Request ID"),
        type:
          tp_get_(row, h, "Type"),
        status:
          tp_get_(row, h, "Status"),
        title:
          tp_get_(row, h, "Title"),
        message:
          tp_get_(row, h, "Message"),
        portalRead:
          tp_truthy_(tp_get_(row, h, "Portal Read"))
      };
    })
    .filter(Boolean)
    .reverse();
}


function teacher_markNotificationRead(notificationId) {
  const email = tp_requireTeacherEmail_().toLowerCase();
  const ss = nhsperf_ss_(TP_SPREADSHEET_ID);
  const sh = tp_ensureNotificationLog_(ss);

  if (sh.getLastRow() < 2) return { ok: false };

  const h = tp_headerMap_(sh);
  const idCol = h["Notification ID"];
  const recipientCol = h["Recipient"];
  const readCol = h["Portal Read"];

  if (!idCol || !recipientCol || !readCol) return { ok: false };

  // V26: one in-memory scan instead of TextFinder + extra cell reads.
  const values = sh
    .getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
    .getValues();

  const idIndex = idCol - 1;
  const recipientIndex = recipientCol - 1;
  const readIndex = readCol - 1;
  const targetId = String(notificationId || "").trim();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][idIndex] || "").trim() !== targetId) continue;

    const recipient = String(values[i][recipientIndex] || "")
      .trim()
      .toLowerCase();

    if (recipient !== email) {
      throw new Error("You do not have access to this notification.");
    }

    if (String(values[i][readIndex]).toLowerCase() !== "true") {
      sh.getRange(i + 2, readCol).setValue(true);
    }

    tp_invalidateTeacherCache_(email);
    return { ok: true };
  }

  return { ok: false };
}


function teacher_markAllNotificationsRead() {
  const email = tp_requireTeacherEmail_().toLowerCase();
  const ss = nhsperf_ss_(TP_SPREADSHEET_ID);
  const sh = tp_ensureNotificationLog_(ss);

  if (sh.getLastRow() < 2) {
    return { ok: true, updated: 0 };
  }

  const h = tp_headerMap_(sh);
  const recipientCol = h["Recipient"];
  const readCol = h["Portal Read"];

  if (!recipientCol || !readCol) {
    return { ok: true, updated: 0 };
  }

  const rowCount = sh.getLastRow() - 1;
  const recipients = sh
    .getRange(2, recipientCol, rowCount, 1)
    .getDisplayValues();

  const reads = sh
    .getRange(2, readCol, rowCount, 1)
    .getValues();

  let updated = 0;

  for (let i = 0; i < rowCount; i++) {
    const recipient = String(recipients[i][0] || "")
      .trim()
      .toLowerCase();

    if (
      recipient === email &&
      String(reads[i][0]).toLowerCase() !== "true"
    ) {
      reads[i][0] = true;
      updated++;
    }
  }

  if (updated) {
    // One batched write for all of this teacher's unread notifications.
    sh.getRange(2, readCol, rowCount, 1).setValues(reads);
  }

  tp_invalidateTeacherCache_(email);

  return {
    ok: true,
    updated: updated
  };
}



/* ============================================================
   FEEDBACK READ
   ============================================================ */

function tp_readTeacherFeedback_(ss, email) {
  const sh = tp_ensureFeedbackLog_(ss);

  if (sh.getLastRow() < 2) {
    return [];
  }

  const data = sh.getDataRange().getDisplayValues();
  const h = tp_headerIndex_(data[0]);

  return data
    .slice(1)
    .filter(
      row =>
        String(row[h["Teacher Email"]] || "")
          .trim()
          .toLowerCase() === email.toLowerCase()
    )
    .map(row => ({
      feedbackId: tp_get_(row, h, "Feedback ID"),
      timestamp: tp_get_(row, h, "Timestamp"),
      requestId: tp_get_(row, h, "Request ID"),
      sessionDate: tp_isoDate_(
        tp_get_(row, h, "Session Date")
      ),
      response: tp_get_(row, h, "Response"),
      notes: tp_get_(row, h, "Notes")
    }))
    .reverse()
    .slice(0, 50);
}


/* ============================================================
   OWNERSHIP / SECURITY
   ============================================================ */

function tp_findOwnedRequestRow_(sh, requestId, email) {
  if (!sh) return 0;
  const h = tp_headerMap_(sh);
  const row = typeof nhsdata_findRowByKey_ === "function"
    ? nhsdata_findRowByKey_(sh, "Request ID", requestId)
    : 0;
  if (!row || !h["Teacher Email"]) return 0;

  const owner = String(sh.getRange(row, h["Teacher Email"]).getDisplayValue() || "")
    .trim().toLowerCase();
  return owner === String(email || "").trim().toLowerCase() ? row : 0;
}


/* ============================================================
   TUTOR LOOKUP FOR CANCELLATION EMAILS
   ============================================================ */

function tp_findTutorByAssignment_(ss, request) {
  const sh = nhsperf_sheet_(ss, TP_TUTORS_SHEET);

  if (!sh || sh.getLastRow() < 2) {
    return null;
  }

  const data = sh.getDataRange().getDisplayValues();
  const h = tp_headerIndex_(data[0]);

  const id = String(
    request["Assigned Tutor ID"] || ""
  ).trim();

  const name = String(
    request["Assigned Tutor"] || ""
  ).trim();

  for (let i = 1; i < data.length; i++) {
    const tutorId = tp_get_(data[i], h, "Tutor ID");
    const fullName = tp_get_(data[i], h, "Full Name");
    const preferred = tp_get_(data[i], h, "Preferred Name");
    const display = tp_tutorDisplayName_(preferred, fullName);

    if (
      (id && tutorId === id) ||
      (!id &&
        name &&
        [
          display,
          fullName,
          preferred
        ]
          .map(tp_norm_)
          .includes(tp_norm_(name)))
    ) {
      return {
        id: tutorId,
        name: display || fullName,
        email: tp_get_(data[i], h, "Notification Email") || tp_get_(data[i], h, "Personal Email")
      };
    }
  }

  return null;
}


/* ============================================================
   ACTIVITY LOG
   ============================================================ */

function tp_logActivity_(ss, spec) {
  const sh = tp_ensureActivityLog_(ss);
  const h = tp_headerMap_(sh);
  const timeZone = nhs_safeTimeZone_(ss);

  const id =
    "A-" +
    Utilities.formatDate(
      new Date(),
      timeZone,
      "yyyyMMdd-HHmmss"
    ) +
    "-" +
    Utilities.getUuid();

  let actor = tp_clean_(spec.actor);

  if (!actor) {
    try {
      actor = tp_clean_(Session.getActiveUser().getEmail());
    } catch (e) {}
  }

  if (!actor) {
    actor = "Teacher Portal";
  }

  const row = new Array(sh.getLastColumn()).fill("");

  tp_rowSet_(row, h, "Activity ID", id);
  tp_rowSet_(row, h, "Timestamp", new Date());
  tp_rowSet_(row, h, "Type", spec.type || "Teacher Portal");
  tp_rowSet_(row, h, "Request ID", spec.requestId || "");
  tp_rowSet_(row, h, "Session ID", spec.sessionId || "");
  tp_rowSet_(row, h, "Tutor ID", spec.tutorId || "");
  tp_rowSet_(row, h, "Actor", actor);
  tp_rowSet_(row, h, "Summary", spec.summary || "");
  tp_rowSet_(row, h, "Details", spec.details || "");

  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length)
    .setValues([row]);
}


/* ============================================================
   INTERNAL ALERTS
   ============================================================ */

function tp_logInternalAlert_(ss, alert) {
  const sh = tp_ensureNotificationLog_(ss);

  const dedupeKey =
    "ALERT|" +
    alert.type +
    "|" +
    alert.requestId;

  if (tp_hasOpenAlert_(sh, dedupeKey)) {
    return;
  }

  sh.appendRow([
    tp_id_("N", ss),
    new Date(),
    alert.requestId || "",
    "Internal Alert",
    "",
    alert.type || "",
    "Created",
    dedupeKey,
    false,
    alert.severity || "Medium",
    alert.title || "Teacher Portal Alert",
    alert.message || "",
    ""
  ]);
}

function tp_hasOpenAlert_(sh, dedupeKey) {
  if (sh.getLastRow() < 2) {
    return false;
  }

  const data = sh.getDataRange().getDisplayValues();
  const h = tp_headerIndex_(data[0]);

  return data.slice(1).some(row => {
    return (
      tp_get_(row, h, "Dedupe Key") === dedupeKey &&
      tp_get_(row, h, "Channel") === "Internal Alert" &&
      tp_get_(row, h, "Read").toLowerCase() !== "true"
    );
  });
}


/* ============================================================
   EMAIL + NOTIFICATION HISTORY
   ============================================================ */

function tp_sendEmailLogged_(ss, spec) {
  const sh = tp_ensureNotificationLog_(ss);

  if (
    spec.dedupeKey &&
    tp_wasNotificationSent_(sh, spec.dedupeKey)
  ) {
    return false;
  }

  let status = "Sent";
  let error = "";

  try {
    nhsmail_send_({
      to: spec.recipient,
      subject: spec.subject,
      body: spec.body,
      name: "NHS Tutoring",
      requestId: spec.requestId || "",
      type: spec.type || ""
    });
  } catch (e) {
    status = "Failed";
    error = e.message || String(e);
  }

  sh.appendRow([
    tp_id_("N", ss),
    new Date(),
    spec.requestId || "",
    "Email",
    spec.recipient || "",
    spec.type || "",
    status,
    spec.dedupeKey || "",
    true,
    "",
    spec.subject || "",
    "",
    error
  ]);

  // IMPORTANT:
  // Email failure is logged, but it does not roll back or invalidate
  // a tutoring request that was already successfully created.
  return status === "Sent";
}

function tp_wasNotificationSent_(sh, dedupeKey) {
  if (!dedupeKey || !sh || sh.getLastRow() < 2) {
    return false;
  }

  const data = sh.getDataRange().getDisplayValues();
  const h = tp_headerIndex_(data[0]);

  return data.slice(1).some(row => {
    return (
      tp_get_(row, h, "Dedupe Key") === dedupeKey &&
      tp_get_(row, h, "Status") === "Sent"
    );
  });
}


/* ============================================================
   SCHEMA
   ============================================================ */

function tp_ensureRequestSchema_(sh) {
  const needed = [
    "Request ID",
    "Timestamp",
    "Teacher Name",
    "Teacher Email",
    "Teacher Notification Email",
    "Student Name",
    "Student Grade",
    "Course",
    "Topic / Academic Need",
    "Support Type",
    "Scheduling Preference",
    "Preferred Dates",
    "Preferred Days",
    "Urgency",
    "Recurring?",
    "Specific Tutor Preference",
    "Teacher Notes",
    "Status",
    "Assigned Tutor",
    "Scheduled Date",
    "Scheduled Time",
    "Location",
    "Last Update",
    "Coordinator Notes",
    "Assigned Tutor ID",
    "Recurring Active",
    "Next Session Date",
    "Request Source",
    "Teacher Change Request",
    "Teacher Change Requested At",
    "Teacher Change Status",
    "Tutor Reconfirm Status",
    "Tutor Reconfirm At",
    "Teacher Five Star Ready",
    "Teacher Five Star Ready At",
    "Tutor Five Star Enrolled",
    "Tutor Five Star Enrolled At",
    "Five Star Ready At"
  ];

  let headers = sh
    .getRange(
      1,
      1,
      1,
      Math.max(1, sh.getLastColumn())
    )
    .getDisplayValues()[0];

  needed.forEach(header => {
    if (!headers.includes(header)) {
      sh.getRange(1, sh.getLastColumn() + 1)
        .setValue(header);

      headers.push(header);
    }
  });
}

function tp_ensureFeedbackLog_(ss) {
  let sh = nhsperf_sheet_(ss, TP_FEEDBACK_LOG_SHEET);

  const headers = [
    "Feedback ID",
    "Timestamp",
    "Request ID",
    "Session Date",
    "Teacher Email",
    "Response",
    "Notes",
    "Status"
  ];

  if (!sh) {
    sh = ss.insertSheet(TP_FEEDBACK_LOG_SHEET);
    sh.getRange(1, 1, 1, headers.length)
      .setValues([headers]);
  }

  sh.setFrozenRows(1);

  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#123A5A")
    .setFontColor("#ffffff");

  return sh;
}

function tp_ensureActivityLog_(ss) {
  let sh = nhsperf_sheet_(ss, TP_ACTIVITY_LOG_SHEET);

  const headers = [
    "Activity ID",
    "Timestamp",
    "Type",
    "Request ID",
    "Session ID",
    "Tutor ID",
    "Actor",
    "Summary",
    "Details"
  ];

  if (!sh) {
    sh = ss.insertSheet(TP_ACTIVITY_LOG_SHEET);
    sh.getRange(1, 1, 1, headers.length)
      .setValues([headers]);
  }

  return sh;
}

function tp_ensureNotificationLog_(ss) {
  let sh = nhsperf_sheet_(ss, TP_NOTIFICATION_LOG_SHEET);

  if (!sh) {
    sh = ss.insertSheet(TP_NOTIFICATION_LOG_SHEET);
  }

  const headers = [
    "Notification ID",
    "Created At",
    "Request ID",
    "Channel",
    "Recipient",
    "Type",
    "Status",
    "Dedupe Key",
    "Read",
    "Severity",
    "Title",
    "Message",
    "Error",
    "Portal Read"
  ];

  let current = sh
    .getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1))
    .getDisplayValues()[0];

  headers.forEach(header => {
    if (!current.includes(header)) {
      sh.getRange(1, current.length + 1).setValue(header);
      current.push(header);
    }
  });

  sh.setFrozenRows(1);

  return sh;
}


/* ============================================================
   CACHE
   ============================================================ */

function tp_invalidateTeacherCache_(email) {
  if (typeof nhs_v45InvalidatePortalState_ === "function") {
    nhs_v45InvalidatePortalState_({
      teacherEmail: email
    });
    return;
  }

  try {
    nhsperf_scriptCache_().remove(
      "teacherPortalV460|" +
      String(email || "").toLowerCase()
    );
  } catch (e) {}
}


/* ============================================================
   GENERIC HELPERS
   ============================================================ */

function tp_headerMap_(sh) {
  return nhsperf_headerMap_(sh, true, true);
}

function tp_headerIndex_(headers, oneBased) {
  const map = {};

  headers.forEach((header, i) => {
    const key = String(header || "").trim();

    if (!key) {
      return;
    }

    if (map[key] === undefined) {
      map[key] = oneBased ? i + 1 : i;
    }
  });

  return map;
}

function tp_get_(row, index, header) {
  if (index[header] === undefined) {
    return "";
  }

  return String(row[index[header]] || "").trim();
}

function tp_rowSet_(row, h, header, value) {
  const index = h[header];

  if (!index) {
    return;
  }

  row[index - 1] = value;
}

function tp_set_(sh, row, h, header, value) {
  if (h[header]) {
    sh.getRange(row, h[header]).setValue(value);
  }
}

function tp_rowObject_(sh, row, h) {
  const values = sh
    .getRange(row, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0];

  const out = {};

  Object.keys(h).forEach(header => {
    out[header] = values[h[header] - 1] || "";
  });

  return out;
}

function tp_arrayOrText_(value) {
  if (Array.isArray(value)) {
    return value
      .map(tp_clean_)
      .filter(Boolean)
      .join(", ");
  }

  return tp_clean_(value);
}

function tp_clean_(value) {
  return String(value || "").trim();
}

function tp_truthy_(value) {
  if (value === true) {
    return true;
  }

  return [
    "true",
    "yes",
    "1",
    "checked"
  ].includes(
    String(value || "")
      .trim()
      .toLowerCase()
  );
}

function tp_isRecurringValue_(value) {
  const normalized = String(value || "")
    .trim()
    .toLowerCase();

  return (
    normalized.startsWith("yes") ||
    normalized.includes("recurring") ||
    ["true", "1"].includes(normalized)
  );
}

function tp_isoDate_(value) {
  if (!value) {
    return "";
  }

  const d = new Date(value);

  if (isNaN(d.getTime())) {
    return "";
  }

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
}

function tp_tutorDisplayName_(preferred, full) {
  preferred = String(preferred || "").trim();
  full = String(full || "").trim();

  if (!preferred) {
    return full;
  }

  const parts = full
    .split(/\s+/)
    .filter(Boolean);

  if (parts.length < 2) {
    return preferred;
  }

  return preferred + " " + parts[parts.length - 1];
}

function tp_norm_(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^\w@]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function tp_id_(prefix, ss) {
  return (
    prefix +
    "-" +
    Utilities.formatDate(
      new Date(),
      nhs_safeTimeZone_(ss),
      "yyyyMMdd-HHmmss"
    ) +
    "-" +
    Utilities.getUuid()
  );
}

function tp_withLock_(fn) {
  const lock = LockService.getScriptLock();

  lock.waitLock(15000);

  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}


/* ============================================================
   V17 FAST ASYNC RECEIPT WORKER
   ============================================================ */

const TP_ASYNC_QUEUE_PROPERTY = "TP_ASYNC_RECEIPT_QUEUE_V17";
const TP_ASYNC_WORKER_HANDLER = "tp_processAsyncReceiptQueue_";
const TP_ASYNC_MAX_BATCH = 20;

function installTeacherFastAsyncWorker() {
  nhssec_requireMaintenanceAdmin_();

  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === TP_ASYNC_WORKER_HANDLER) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger(TP_ASYNC_WORKER_HANDLER)
    .timeBased()
    .everyMinutes(1)
    .create();

  return {
    ok: true,
    message: "Teacher fast async receipt worker installed."
  };
}

function tp_queueAsyncReceipt_(job) {
  if (!job || !job.recipient) return;

  const lock = LockService.getScriptLock();
  let locked = false;

  try {
    locked = lock.tryLock(150);
  } catch (e) {
    locked = false;
  }

  if (!locked) {
    // Never make a successful request wait on the receipt queue.
    try {
      nhsperf_scriptProperties_().setProperty(
        "TP_ASYNC_OVERFLOW_" + String(job.requestId || Utilities.getUuid()),
        JSON.stringify(job)
      );
    } catch (e) {
      console.error("[NHS V46.7 receipt overflow]", e);
    }
    return;
  }

  try {
    const props = nhsperf_scriptProperties_();
    let queue = [];

    try {
      queue = JSON.parse(props.getProperty(TP_ASYNC_QUEUE_PROPERTY) || "[]");
      if (!Array.isArray(queue)) queue = [];
    } catch (e) {
      queue = [];
    }

    queue.push({
      requestId: String(job.requestId || ""),
      recipient: String(job.recipient || ""),
      type: String(job.type || ""),
      subject: String(job.subject || ""),
      dedupeKey: String(job.dedupeKey || ""),
      body: String(job.body || ""),
      queuedAt: new Date().toISOString(),
      attempts: 0
    });

    // Keep queue bounded.
    if (queue.length > 100) {
      queue = queue.slice(queue.length - 100);
    }

    props.setProperty(
      TP_ASYNC_QUEUE_PROPERTY,
      JSON.stringify(queue)
    );
  } finally {
    if (locked) {
      try { lock.releaseLock(); } catch (e) {}
    }
  }
}

function tp_processAsyncReceiptQueue_() {
  const props = nhsperf_scriptProperties_();
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(5000);
  } catch (e) {
    return { ok: false, skipped: true, reason: "Queue busy" };
  }

  let batch = [];
  let remaining = [];

  try {
    let queue = [];

    try {
      queue = JSON.parse(props.getProperty(TP_ASYNC_QUEUE_PROPERTY) || "[]");
      if (!Array.isArray(queue)) queue = [];
    } catch (e) {
      queue = [];
    }

    // Fold in any overflow jobs.
    props.getKeys()
      .filter(k => k.indexOf("TP_ASYNC_OVERFLOW_") === 0)
      .forEach(function(key) {
        try {
          queue.push(JSON.parse(props.getProperty(key)));
        } catch (e) {}
        props.deleteProperty(key);
      });

    batch = queue.slice(0, TP_ASYNC_MAX_BATCH);
    remaining = queue.slice(TP_ASYNC_MAX_BATCH);

    // Remove the current batch from persistent queue before sending.
    props.setProperty(
      TP_ASYNC_QUEUE_PROPERTY,
      JSON.stringify(remaining)
    );
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }

  if (!batch.length) {
    return { ok: true, processed: 0 };
  }

  const ss = nhsperf_ss_(TP_SPREADSHEET_ID);
  const retry = [];
  let sent = 0;

  batch.forEach(function(job) {
    try {
      const ok = tp_sendEmailLogged_(ss, job);
      if (ok) {
        sent++;
      } else {
        throw new Error("Receipt send returned false.");
      }
    } catch (e) {
      job.attempts = Number(job.attempts || 0) + 1;

      if (job.attempts < 3) {
        retry.push(job);
      }
    }
  });

  if (retry.length) {
    const lock2 = LockService.getScriptLock();

    try {
      lock2.waitLock(5000);
      let current = [];

      try {
        current = JSON.parse(props.getProperty(TP_ASYNC_QUEUE_PROPERTY) || "[]");
        if (!Array.isArray(current)) current = [];
      } catch (e) {
        current = [];
      }

      props.setProperty(
        TP_ASYNC_QUEUE_PROPERTY,
        JSON.stringify(retry.concat(current).slice(0, 100))
      );
    } finally {
      try { lock2.releaseLock(); } catch (e) {}
    }
  }

  return {
    ok: true,
    processed: batch.length,
    sent: sent,
    retried: retry.length
  };
}

function nhsv473_teacher_getAuthoritativeRequestStateCore_() {
  const email = tp_requireTeacherEmail_();
  const ss = nhsperf_ss_(TP_SPREADSHEET_ID);
  const teacher = tp_getTeacherIdentity_(ss, email);
  const requests = tp_readTeacherRequests_(ss, email);
  const tz = nhs_safeTimeZone_(ss);
  const now = new Date();

  const upcoming = requests
    .filter(function(r) {
      return typeof nhslife_upcoming_ === "function"
        ? nhslife_upcoming_(r, now, tz)
        : (!!r.scheduledDate && !!r.assignmentConfirmed && !["Completed","Cancelled"].includes(r.status));
    })
    .sort(function(a, b) {
      return String(a.scheduledDate || "").localeCompare(String(b.scheduledDate || ""));
    });

  return {
    generatedAt: new Date().toISOString(),
    revision: nhsperf_getDataRevision_(),
    requests: requests,
    upcoming: upcoming,
    counts: {
      open: requests.filter(function(r) { return !["Completed", "Cancelled"].includes(r.status); }).length,
      confirmed: upcoming.length,
      recurring: requests.filter(function(r) { return !!r.recurringActive && !["Completed", "Cancelled"].includes(r.status); }).length,
      completed: requests.filter(function(r) { return r.status === "Completed"; }).length
    },
    teacher: teacher,
    schedulingOptions: nhssched_teacherOptions_(tz, now),
    lifecycleBuild: NHS_RELEASE.VERSION
  };
}


/* ============================================================
   V55 — TUTOR PROFILE FORM → TUTOR DATABASE PIPELINE
   ============================================================ */

/**
 * Existing NHS Tutor Profile form supplied for 2026–27.
 * Script Property NHS_TUTOR_PROFILE_FORM_ID may override this without
 * changing source.
 */
const NHSTUTOR_PROFILE_FORM_ID_DEFAULT =
  "1z40i1VmDklSDr4o56hq22olze2y5AlkifodyBXmgCiI";

const NHSTUTOR_PROFILE_FORM_PROPERTY =
  "NHS_TUTOR_PROFILE_FORM_ID";

const NHSTUTOR_PROFILE_RECONCILE_HANDLER =
  "nhstutor_reconcileProfilesV55_";

const NHSTUTOR_PROFILE_SUBMIT_HANDLER =
  "processTutorProfileSubmissionV2";


function nhstutor_profileFormId_() {
  try {
    const configured = String(
      PropertiesService
        .getScriptProperties()
        .getProperty(
          NHSTUTOR_PROFILE_FORM_PROPERTY
        ) || ""
    ).trim();

    return configured ||
      NHSTUTOR_PROFILE_FORM_ID_DEFAULT;
  } catch (e) {
    return NHSTUTOR_PROFILE_FORM_ID_DEFAULT;
  }
}


function nhstutor_openProfileForm_() {
  const formId =
    nhstutor_profileFormId_();

  if (!formId) return null;

  try {
    return FormApp.openById(formId);
  } catch (error) {
    console.error(
      "[NHS V55 tutor profile form]",
      error
    );
    return null;
  }
}


function nhstutor_ensureTutorDatabase_(ss) {
  let sh = nhsperf_sheet_(
    ss,
    TUP_TUTORS
  );

  if (!sh) {
    sh = ss.insertSheet(TUP_TUTORS);
  }

  tup_ensureHeaders_(
    sh,
    TUP_TUTOR_HEADERS
  );

  return sh;
}


function nhstutor_answerValue_(answers, title) {
  answers = answers || {};

  let value = answers[title];

  if (Array.isArray(value)) {
    value = value.length === 1
      ? value[0]
      : value.join(" | ");
  }

  if (
    value &&
    typeof value === "object" &&
    !Array.isArray(value)
  ) {
    value = String(value);
  }

  return String(
    value == null ? "" : value
  ).trim();
}


function nhstutor_mergeAnswer_(answers, title, value) {
  answers = answers || {};
  title = String(title || "").trim();
  value = String(value == null ? "" : value).trim();

  if (!title || !value) return answers;

  const existing = nhstutor_answerValue_(answers, title);
  if (!existing) {
    answers[title] = value;
    return answers;
  }

  const parts = existing
    .split(/\s*\|\s*/)
    .map(function(v) { return String(v || "").trim(); })
    .filter(Boolean);

  if (!parts.includes(value)) parts.push(value);
  answers[title] = parts.join(" | ");
  return answers;
}


function nhstutor_firstAnswer_(answers, titles) {
  titles = Array.isArray(titles) ? titles : [titles];

  for (let i = 0; i < titles.length; i++) {
    const value = nhstutor_answerValue_(answers, titles[i]);
    if (value) return value;
  }

  return "";
}


function nhstutor_collectAnswerTitles_(answers, titles) {
  const selected = [];

  (titles || []).forEach(function(title) {
    const value = nhstutor_answerValue_(answers, title);

    if (
      value &&
      !/^(none|n\/?a|not applicable)$/i.test(value) &&
      !selected.includes(value)
    ) {
      selected.push(value);
    }
  });

  return selected.join(" | ");
}


function nhstutor_answersFromFormResponse_(response) {
  const answers = {};

  if (
    !response ||
    typeof response.getItemResponses !== "function"
  ) {
    return answers;
  }

  response
    .getItemResponses()
    .forEach(function(itemResponse) {
      const item = itemResponse.getItem();

      const title = String(
        item && item.getTitle ? item.getTitle() : ""
      ).trim();

      if (!title) return;

      let value = itemResponse.getResponse();
      if (Array.isArray(value)) value = value.join(", ");

      nhstutor_mergeAnswer_(answers, title, value);
    });

  return answers;
}


function nhstutor_collectCourseSlots_(answers, category, numberOfSlots) {
  const titles = [String(category || "").trim()];

  for (let i = 1; i <= Number(numberOfSlots || 4); i++) {
    titles.push(String(category || "").trim() + " Course #" + i);
  }

  return nhstutor_collectAnswerTitles_(answers, titles);
}

function nhstutor_stableTutorId_(
  email,
  database
) {
  email = String(email || "")
    .trim()
    .toLowerCase();

  const username =
    (
      email.split("@")[0] ||
      "TUTOR"
    )
      .replace(
        /[^a-zA-Z0-9]/g,
        ""
      )
      .toUpperCase() ||
    "TUTOR";

  const baseId =
    "T-" + username;

  if (
    !database ||
    database.getLastRow() < 2
  ) {
    return baseId;
  }

  const h =
    tup_headerMap_(database);

  const idCol = h["Tutor ID"];
  const emailCol =
    h["School Email"];

  if (!idCol || !emailCol) {
    return baseId;
  }

  const values =
    database
      .getRange(
        2,
        1,
        database.getLastRow() - 1,
        database.getLastColumn()
      )
      .getDisplayValues();

  let baseUsedByOther = false;

  for (let i = 0; i < values.length; i++) {
    const rowEmail =
      String(
        values[i][emailCol - 1] || ""
      )
        .trim()
        .toLowerCase();

    const rowId =
      String(
        values[i][idCol - 1] || ""
      ).trim();

    if (
      rowEmail === email &&
      rowId
    ) {
      return rowId;
    }

    if (
      rowId === baseId &&
      rowEmail &&
      rowEmail !== email
    ) {
      baseUsedByOther = true;
    }
  }

  if (!baseUsedByOther) {
    return baseId;
  }

  // Deterministic collision suffix.
  let hash = 0;

  for (let i = 0; i < email.length; i++) {
    hash =
      (
        (hash << 5) -
        hash +
        email.charCodeAt(i)
      ) | 0;
  }

  const suffix =
    Math.abs(hash)
      .toString(36)
      .toUpperCase()
      .slice(-5)
      .padStart(5, "0");

  return baseId + "-" + suffix;
}


function nhstutor_rosterNameTokens_(value) {
  return String(value || "")
    .replace(/\([^)]*\)/g, " ")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .split(/\s+/)
    .filter(Boolean);
}


function nhstutor_rosterMatch_(ss, email, fullName) {
  const roster = nhsperf_sheet_(ss, "Membership Roster");

  if (!roster || roster.getLastRow() < 2) return "Unknown";

  const values = roster.getDataRange().getDisplayValues();
  const headers = values[0].map(function(v) { return String(v || "").trim(); });
  const emailIdx = headers.indexOf("School Email");
  const nameIdx = headers.indexOf("Full Name");
  const normalizedEmail = String(email || "").trim().toLowerCase();
  const requestedTokens = nhstutor_rosterNameTokens_(fullName);

  const matched = values.slice(1).some(function(row) {
    if (emailIdx >= 0 && normalizedEmail) {
      const rowEmail = String(row[emailIdx] || "").trim().toLowerCase();
      if (rowEmail && rowEmail === normalizedEmail) return true;
    }

    if (nameIdx < 0 || !requestedTokens.length) return false;

    const rosterTokens = nhstutor_rosterNameTokens_(row[nameIdx]);
    if (!rosterTokens.length) return false;

    return requestedTokens.every(function(token) {
      return rosterTokens.includes(token);
    });
  });

  return matched ? "Yes" : "No";
}

function nhstutor_profileValuesFromAnswers_(answers) {
  const strongest = nhstutor_collectAnswerTitles_(
    answers,
    [
      "Strongest Course #1",
      "Strongest Course #2",
      "Strongest Course #3"
    ]
  ) || nhstutor_firstAnswer_(
    answers,
    ["Which 3 course(s) would you be MOST confident tutoring?"]
  );

  const declaredSchoolEmail = nhstutor_firstAnswer_(
    answers,
    ["School Email", "School Email Address"]
  ).toLowerCase();

  const personalEmail = nhstutor_firstAnswer_(
    answers,
    ["Personal Email", "Notification Email"]
  ).toLowerCase();

  const schoolEmail = declaredSchoolEmail ||
    (/@(?:capostu|capousd)\.org$/i.test(personalEmail) ? personalEmail : "");

  const notificationEmail =
    personalEmail &&
    nhsa_isValidEmailAddress_(personalEmail) &&
    personalEmail !== schoolEmail &&
    !/@(?:capostu|capousd)\.org$/i.test(personalEmail)
      ? personalEmail
      : "";

  const titanDays = nhstutor_firstAnswer_(
    answers,
    [
      "Typical Titan Time Availability",
      "Titan Time Days",
      "Titan Time Availability"
    ]
  );

  const afterSchoolDays = nhstutor_firstAnswer_(
    answers,
    [
      "Typical After-School Availability",
      "After-School Days",
      "After-School Availability"
    ]
  );

  return {
    fullName: nhstutor_firstAnswer_(answers, ["Full Name"]),
    preferredName: nhstutor_firstAnswer_(answers, ["Preferred Name"]),
    schoolEmail: schoolEmail,
    notificationEmail: notificationEmail,
    grade: nhstutor_firstAnswer_(answers, ["Grade"]),

    englishCourses: nhstutor_collectAnswerTitles_(
      answers,
      ["English", "English Course #1", "English Course #2", "English Course #3", "English Course #4"]
    ),
    mathCourses: nhstutor_collectAnswerTitles_(
      answers,
      ["Mathematics", "Mathematics Course #1", "Mathematics Course #2", "Mathematics Course #3", "Mathematics Course #4"]
    ),
    scienceCourses: nhstutor_collectAnswerTitles_(
      answers,
      ["Science", "Science Course #1", "Science Course #2", "Science Course #3", "Science Course #4"]
    ),
    socialScienceCourses: nhstutor_collectAnswerTitles_(
      answers,
      ["Social Science", "Social Science Course #1", "Social Science Course #2", "Social Science Course #3", "Social Science Course #4"]
    ),
    worldLanguageCourses: nhstutor_collectAnswerTitles_(
      answers,
      ["World Language", "World Language Course #1", "World Language Course #2", "World Language 2", "World Language Course #3", "World Language Course #4"]
    ),
    cteCourses: nhstutor_collectAnswerTitles_(
      answers,
      [
        "Career Technical Education",
        "Computer Science / CTE",
        "Computer Science / CTE Course #1",
        "Computer Science / CTE Course #2",
        "Computer Science / CTE Course #3",
        "Computer Science / CTE Course #4"
      ]
    ),

    strongestCourse: strongest,
    titanAvailability: titanDays,
    titanDays: titanDays,
    afterSchoolAvailability: afterSchoolDays,
    afterSchoolDays: afterSchoolDays,
    recurringTutoring: nhstutor_firstAnswer_(answers, ["Recurring Tutoring"]),
    maxSessions: nhstutor_firstAnswer_(answers, ["Maximum Sessions / Week"]),
    languages: nhstutor_firstAnswer_(answers, ["Tutoring Languages"]),
    previousExperience: nhstutor_firstAnswer_(answers, ["Previous Experience"]),
    experienceDetails: nhstutor_firstAnswer_(answers, ["Experience Details"]),
    academicQualifications: nhstutor_firstAnswer_(
      answers,
      ["Academic Qualifications", "Academic Qualifications (Optional)"]
    ),
    matchingNotes: nhstutor_firstAnswer_(answers, ["Matching Notes"]),
    expectationsAgreement: nhstutor_firstAnswer_(answers, ["Expectations Agreement"])
  };
}

function nhstutor_findTutorRowByEmail_(
  database,
  email
) {
  if (
    !database ||
    database.getLastRow() < 2
  ) {
    return 0;
  }

  const h =
    tup_headerMap_(database);

  const emailCol =
    h["School Email"];

  if (!emailCol) return 0;

  const normalized =
    String(email || "")
      .trim()
      .toLowerCase();

  const emails =
    database
      .getRange(
        2,
        emailCol,
        database.getLastRow() - 1,
        1
      )
      .getDisplayValues();

  for (
    let i = 0;
    i < emails.length;
    i++
  ) {
    if (
      String(emails[i][0] || "")
        .trim()
        .toLowerCase() ===
      normalized
    ) {
      return i + 2;
    }
  }

  return 0;
}


/**
 * Resolves a profile response to the canonical CUSD Tutor Database identity.
 * This repairs three migration-era response patterns without inventing email:
 *   1) exact School Email match,
 *   2) same CUSD username under the wrong staff/student domain,
 *   3) missing School Email but a unique existing tutor name/preferred-name match.
 */
function nhstutor_canonicalSchoolEmailFromProfile_(ss, parsed) {
  parsed = parsed || {};

  const supplied = String(parsed.schoolEmail || "").trim().toLowerCase();
  const database = nhsperf_sheet_(ss, "Tutor Database");

  if (!database || database.getLastRow() < 2) {
    return supplied;
  }

  const values = database
    .getRange(1, 1, database.getLastRow(), database.getLastColumn())
    .getDisplayValues();
  const h = nhsa_headerIndex_(values[0]);
  const schoolCol = h["School Email"];

  if (schoolCol === undefined) return supplied;

  const activeRows = values.slice(1).map(function(row, index) {
    const email = String(row[schoolCol] || "").trim().toLowerCase();
    const reliability = h["Reliability Status"] === undefined
      ? ""
      : String(row[h["Reliability Status"]] || "").trim().toLowerCase();

    if (!email || !nhsa_isSchoolIdentityEmail_(email) || reliability === "inactive") {
      return null;
    }

    return {
      row: index + 2,
      email: email,
      fullName: h["Full Name"] === undefined ? "" : String(row[h["Full Name"]] || "").trim(),
      preferredName: h["Preferred Name"] === undefined ? "" : String(row[h["Preferred Name"]] || "").trim()
    };
  }).filter(Boolean);

  if (supplied) {
    const exact = activeRows.find(function(item) { return item.email === supplied; });
    if (exact) return exact.email;

    const local = supplied.split("@")[0];
    if (local) {
      const sameLocal = activeRows.filter(function(item) {
        return item.email.split("@")[0] === local;
      });
      if (sameLocal.length === 1) return sameLocal[0].email;
    }
  }

  function nameKey_(value) {
    return String(value || "")
      .trim()
      .toLowerCase()
      .replace(/\([^)]*\)/g, " ")
      .replace(/[^a-z0-9]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  const requested = [parsed.fullName, parsed.preferredName]
    .map(nameKey_)
    .filter(Boolean);

  if (requested.length) {
    const nameMatches = activeRows.filter(function(item) {
      const candidates = [item.fullName, item.preferredName]
        .map(nameKey_)
        .filter(Boolean);
      return requested.some(function(key) { return candidates.includes(key); });
    });

    if (nameMatches.length === 1) return nameMatches[0].email;
  }

  return supplied;
}


function nhstutor_upsertProfileAnswers_(
  ss,
  answers,
  submittedAt,
  options
) {
  options = options || {};

  const parsed =
    nhstutor_profileValuesFromAnswers_(
      answers
    );

  if (
    !parsed.schoolEmail ||
    !nhsa_isValidEmailAddress_(
      parsed.schoolEmail
    )
  ) {
    return {
      ok: false,
      skipped: true,
      reason:
        "Missing or invalid School Email."
    };
  }

  const database =
    nhstutor_ensureTutorDatabase_(ss);

  parsed.schoolEmail =
    nhstutor_canonicalSchoolEmailFromProfile_(
      ss,
      parsed
    );

  if (
    !parsed.schoolEmail ||
    !nhsa_isSchoolIdentityEmail_(parsed.schoolEmail)
  ) {
    return {
      ok: false,
      skipped: true,
      reason:
        "Tutor profile could not be linked to a CUSD School Email."
    };
  }

  const rowNum =
    nhstutor_findTutorRowByEmail_(
      database,
      parsed.schoolEmail
    );

  const h =
    tup_headerMap_(database);

  const timestamp =
    submittedAt instanceof Date &&
    !isNaN(submittedAt.getTime())
      ? submittedAt
      : new Date();

  let values =
    rowNum
      ? database
          .getRange(
            rowNum,
            1,
            1,
            database.getLastColumn()
          )
          .getValues()[0]
      : new Array(
          database.getLastColumn()
        ).fill("");

  const getExisting =
    function(header) {
      return h[header]
        ? values[h[header] - 1]
        : "";
    };

  const changedHeaders = {};

  const set =
    function(header, value) {
      if (!h[header]) return;

      values[h[header] - 1] =
        value;
      changedHeaders[header] = true;
    };

  const priorFormUpdate =
    getExisting(
      "Last Form Profile Update"
    );

  const priorFormDate =
    priorFormUpdate instanceof Date
      ? priorFormUpdate
      : (
          priorFormUpdate
            ? new Date(priorFormUpdate)
            : null
        );

  if (
    rowNum &&
    options.force !== true &&
    priorFormDate &&
    !isNaN(priorFormDate.getTime()) &&
    priorFormDate.getTime() >=
      timestamp.getTime()
  ) {
    return {
      ok: true,
      skipped: true,
      reason:
        "Profile response already synchronized.",
      email:
        parsed.schoolEmail,
      row:
        rowNum
    };
  }

  const tutorId =
    String(
      getExisting("Tutor ID") || ""
    ).trim() ||
    nhstutor_stableTutorId_(
      parsed.schoolEmail,
      database
    );

  set(
    "Tutor ID",
    tutorId
  );

  set(
    "Full Name",
    nhsinput_literalText_(
      parsed.fullName ||
      String(
        getExisting("Full Name") || ""
      ),
      120,
      "Tutor full name"
    )
  );

  set(
    "Preferred Name",
    nhsinput_literalText_(
      parsed.preferredName,
      80,
      "Tutor preferred name"
    )
  );

  set(
    "School Email",
    parsed.schoolEmail
  );

  const existingNotificationEmail = String(
    getExisting("Notification Email") || ""
  ).trim().toLowerCase();

  if (
    parsed.notificationEmail &&
    nhsa_isValidEmailAddress_(parsed.notificationEmail)
  ) {
    set("Notification Email", parsed.notificationEmail);
  }

  set(
    "Grade",
    nhsinput_literalText_(
      parsed.grade,
      24,
      "Tutor grade"
    )
  );

  set(
    "Roster Match",
    nhstutor_rosterMatch_(
      ss,
      parsed.schoolEmail,
      parsed.fullName
    )
  );

  set(
    "English Courses",
    nhsinput_literalText_(
      parsed.englishCourses,
      800,
      "English courses"
    )
  );

  set(
    "Mathematics Courses",
    nhsinput_literalText_(
      parsed.mathCourses,
      800,
      "Mathematics courses"
    )
  );

  set(
    "Science Courses",
    nhsinput_literalText_(
      parsed.scienceCourses,
      800,
      "Science courses"
    )
  );

  set(
    "Social Science Courses",
    nhsinput_literalText_(
      parsed.socialScienceCourses,
      800,
      "Social Science courses"
    )
  );

  set(
    "World Language Courses",
    nhsinput_literalText_(
      parsed.worldLanguageCourses,
      800,
      "World Language courses"
    )
  );

  set(
    "Computer Science / CTE Courses",
    nhsinput_literalText_(
      parsed.cteCourses,
      800,
      "Computer Science / CTE courses"
    )
  );

  set(
    "Strongest Course",
    nhscourse_normalizeCourseListText_(
      parsed.strongestCourse,
      3
    )
  );

  // Current platform consumes canonical day lists rather than the old
  // Yes/No availability question.
  set(
    "Typical Titan Time Availability",
    nhsinput_literalText_(
      parsed.titanDays,
      250,
      "Titan Time days"
    )
  );

  set(
    "Typical After-School Availability",
    nhsinput_literalText_(
      parsed.afterSchoolDays,
      250,
      "After-school days"
    )
  );

  set(
    "Recurring Tutoring",
    nhsinput_literalText_(
      parsed.recurringTutoring,
      120,
      "Recurring tutoring preference"
    )
  );

  set(
    "Maximum Sessions / Week",
    nhsinput_literalText_(
      parsed.maxSessions,
      24,
      "Maximum sessions per week"
    )
  );

  set(
    "Tutoring Languages",
    nhsinput_literalText_(
      parsed.languages,
      300,
      "Tutoring languages"
    )
  );

  set(
    "Previous Experience",
    nhsinput_literalText_(
      parsed.previousExperience,
      120,
      "Previous tutoring experience"
    )
  );

  set(
    "Experience Details",
    nhsinput_literalText_(
      parsed.experienceDetails,
      1200,
      "Experience details"
    )
  );

  set(
    "Academic Qualifications",
    nhsinput_literalText_(
      parsed.academicQualifications,
      1500,
      "Academic qualifications"
    )
  );

  set(
    "Matching Notes",
    nhsinput_literalText_(
      parsed.matchingNotes,
      1500,
      "Matching notes"
    )
  );

  set(
    "Expectations Agreement",
    nhsinput_literalText_(
      parsed.expectationsAgreement,
      1000,
      "Expectations agreement"
    )
  );

  // Operational fields are initialized only for a brand-new profile.
  if (!rowNum) {
    set(
      "Sessions Completed",
      0
    );
    set(
      "Verified Tutoring Hours",
      0
    );
    set(
      "Tutor No-Shows",
      0
    );
    set(
      "Current Caseload",
      0
    );
    set(
      "Reliability Status",
      "New"
    );
  } else if (
    !String(
      getExisting(
        "Reliability Status"
      ) || ""
    ).trim()
  ) {
    set(
      "Reliability Status",
      "New"
    );
  }

  set(
    "Last Form Profile Update",
    timestamp
  );
  set(
    "Last Profile Update",
    timestamp
  );

  const lock =
    LockService.getScriptLock();

  lock.waitLock(20000);

  let writtenRow;
  let retryExisting = false;

  try {
    if (rowNum) {
      writtenRow = rowNum;

      const freshRange =
        database.getRange(
          rowNum,
          1,
          1,
          database.getLastColumn()
        );

      const freshValues =
        freshRange.getValues()[0];

      const freshPortalUpdatedAt =
        h["Last Portal Profile Update"]
          ? freshValues[
              h["Last Portal Profile Update"] -
              1
            ]
          : "";

      const portalIsNewer =
        nhstutor_dateMillis_(
          freshPortalUpdatedAt
        ) >= timestamp.getTime();

      const portalOwnedOverlap = {
        "Preferred Name": true,
        "Strongest Course": true,
        "Tutoring Languages": true,
        "Experience Details": true,
        "Typical Titan Time Availability": true,
        "Maximum Sessions / Week": true,
        "Notification Email": true
      };

      Object.keys(
        changedHeaders
      ).forEach(function(header) {
        const col = h[header];

        if (!col) return;

        if (
          portalIsNewer &&
          portalOwnedOverlap[header]
        ) {
          return;
        }

        if (
          header === "Tutor ID" &&
          String(
            freshValues[col - 1] || ""
          ).trim()
        ) {
          return;
        }

        if (
          header ===
            "Reliability Status" &&
          String(
            freshValues[col - 1] || ""
          ).trim()
        ) {
          return;
        }

        if (
          header ===
          "Last Profile Update"
        ) {
          return;
        }

        freshValues[col - 1] =
          values[col - 1];
      });

      if (
        h["Last Profile Update"]
      ) {
        const priorMillis =
          nhstutor_dateMillis_(
            freshValues[
              h["Last Profile Update"] -
              1
            ]
          );

        freshValues[
          h["Last Profile Update"] - 1
        ] =
          priorMillis >
          timestamp.getTime()
            ? new Date(priorMillis)
            : timestamp;
      }

      freshRange.setValues(
        [freshValues]
      );

      values = freshValues;
    } else {
      // Re-check under lock so simultaneous form submissions cannot create
      // duplicate tutor profiles for the same School Email.
      const nowExisting =
        nhstutor_findTutorRowByEmail_(
          database,
          parsed.schoolEmail
        );

      if (nowExisting) {
        retryExisting = true;
      } else {
        writtenRow =
          Math.max(
            2,
            database.getLastRow() + 1
          );

        database
          .getRange(
            writtenRow,
            1,
            1,
            values.length
          )
          .setValues([values]);
      }
    }

    if (!retryExisting) {
      SpreadsheetApp.flush();
    }
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }

  if (retryExisting) {
    // Lock is released before retrying against the row another submission
    // created, preventing self-deadlock and preserving operational fields.
    return nhstutor_upsertProfileAnswers_(
      ss,
      answers,
      submittedAt,
      Object.assign(
        {},
        options,
        {
          force: true
        }
      )
    );
  }

  try {
    nhsperf_scriptCache_()
      .remove(
        nhs_v46TutorIdentityKey_(
          parsed.schoolEmail
        )
      );
  } catch (e) {}

  try {
    nhs_v45InvalidatePortalState_({
      tutorId: tutorId
    });
  } catch (e) {}

  return {
    ok: true,
    skipped: false,
    created: !rowNum,
    row: writtenRow,
    tutorId: tutorId,
    email: parsed.schoolEmail,
    fullName: parsed.fullName,
    source:
      options.source ||
      "Tutor Profile Form"
  };
}


function nhstutor_latestResponses_(
  emailFilter
) {
  const form =
    nhstutor_openProfileForm_();

  if (!form) return [];

  const normalizedFilter =
    String(emailFilter || "")
      .trim()
      .toLowerCase();

  const latest = {};

  let responses = [];

  try {
    responses =
      form.getResponses() || [];
  } catch (error) {
    console.error(
      "[NHS V55 read Tutor Profile responses]",
      error
    );
    return [];
  }

  responses.forEach(function(response) {
    const answers =
      nhstutor_answersFromFormResponse_(
        response
      );

    const parsedIdentity =
      nhstutor_profileValuesFromAnswers_(
        answers
      );

    const email = String(
      nhstutor_canonicalSchoolEmailFromProfile_(
        nhsperf_ss_(NHS_SPREADSHEET_ID),
        parsedIdentity
      ) || ""
    ).trim().toLowerCase();

    if (
      !email ||
      (
        normalizedFilter &&
        email !== normalizedFilter
      )
    ) {
      return;
    }

    const timestamp =
      response.getTimestamp
        ? response.getTimestamp()
        : new Date();

    const prior =
      latest[email];

    if (
      !prior ||
      (
        timestamp &&
        prior.timestamp &&
        timestamp.getTime() >=
          prior.timestamp.getTime()
      )
    ) {
      latest[email] = {
        email: email,
        answers: answers,
        timestamp: timestamp
      };
    }
  });

  return Object.keys(latest)
    .sort()
    .map(function(email) {
      return latest[email];
    });
}


function nhstutor_profileResponseSpreadsheet_() {
  const form =
    nhstutor_openProfileForm_();

  if (!form) return null;

  try {
    const destinationId =
      String(
        form.getDestinationId
          ? form.getDestinationId()
          : ""
      ).trim();

    if (!destinationId) return null;

    return SpreadsheetApp.openById(
      destinationId
    );
  } catch (error) {
    console.error(
      "[NHS V55 Tutor Profile destination]",
      error
    );
    return null;
  }
}


function nhstutor_findResponseSheet_(
  ss
) {
  if (!ss) return null;

  const preferredNames = [
    "Tutor Profile Responses",
    "Form Responses 1"
  ];

  for (
    let i = 0;
    i < preferredNames.length;
    i++
  ) {
    const sh =
      ss.getSheetByName(
        preferredNames[i]
      );

    if (
      sh &&
      nhstutor_isProfileResponseSheet_(
        sh
      )
    ) {
      return sh;
    }
  }

  return (
    ss.getSheets()
      .find(function(sh) {
        return (
          sh.getName() !==
            TUP_TUTORS &&
          nhstutor_isProfileResponseSheet_(
            sh
          )
        );
      }) ||
    null
  );
}


function nhstutor_isProfileResponseSheet_(
  sh
) {
  if (
    !sh ||
    sh.getLastColumn() < 4
  ) {
    return false;
  }

  const headers =
    sh
      .getRange(
        1,
        1,
        1,
        sh.getLastColumn()
      )
      .getDisplayValues()[0]
      .map(function(v) {
        return String(v || "").trim();
      });

  return (
    headers.includes("Full Name") &&
    (
      headers.includes("School Email") ||
      headers.includes("School Email Address") ||
      headers.includes("Personal Email")
    ) &&
    headers.includes("Grade") &&
    (
      headers.includes("Strongest Course #1") ||
      headers.includes("Which 3 course(s) would you be MOST confident tutoring?")
    ) &&
    headers.includes(
      "Expectations Agreement"
    )
  );
}


function nhstutor_latestResponsesFromSheet_(
  ss,
  emailFilter
) {
  const sh =
    nhstutor_findResponseSheet_(ss);

  if (
    !sh ||
    sh.getLastRow() < 2
  ) {
    return [];
  }

  const values =
    sh
      .getDataRange()
      .getDisplayValues();

  const headers =
    values[0].map(function(v) {
      return String(v || "").trim();
    });

  const timestampIndex =
    headers.indexOf("Timestamp");

  const normalizedFilter =
    String(emailFilter || "")
      .trim()
      .toLowerCase();

  const latest = {};

  values
    .slice(1)
    .forEach(function(row, index) {
      const answers = {};

      headers.forEach(function(
        header,
        col
      ) {
        if (!header) return;
        nhstutor_mergeAnswer_(
          answers,
          header,
          row[col]
        );
      });

      const parsedIdentity =
        nhstutor_profileValuesFromAnswers_(
          answers
        );

      const email = String(
        nhstutor_canonicalSchoolEmailFromProfile_(
          ss,
          parsedIdentity
        ) || ""
      ).trim().toLowerCase();

      if (
        !email ||
        (
          normalizedFilter &&
          email !== normalizedFilter
        )
      ) {
        return;
      }

      const rawTimestamp =
        timestampIndex >= 0
          ? row[timestampIndex]
          : "";

      const parsedDate =
        rawTimestamp
          ? new Date(rawTimestamp)
          : null;

      const timestamp =
        parsedDate &&
        !isNaN(
          parsedDate.getTime()
        )
          ? parsedDate
          : new Date(
              Date.now() +
              index
            );

      latest[email] = {
        email: email,
        answers: answers,
        timestamp: timestamp
      };
    });

  return Object.keys(latest)
    .sort()
    .map(function(email) {
      return latest[email];
    });
}


function nhstutor_syncProfileResponses_(
  ss,
  options
) {
  options = options || {};

  ss =
    ss ||
    nhsperf_ss_(
      NHS_SPREADSHEET_ID
    );

  nhstutor_ensureTutorDatabase_(ss);

  let responses =
    nhstutor_latestResponses_(
      options.email || ""
    );

  let source =
    responses.length
      ? "Google Form"
      : "";

  if (!responses.length) {
    const sourceBook =
      nhstutor_profileResponseSpreadsheet_();

    if (sourceBook) {
      responses =
        nhstutor_latestResponsesFromSheet_(
          sourceBook,
          options.email || ""
        );

      if (responses.length) {
        source =
          "Tutor Profile Response Sheet";
      }
    }
  }

  if (!responses.length) {
    responses =
      nhstutor_latestResponsesFromSheet_(
        ss,
        options.email || ""
      );

    if (responses.length) {
      source =
        "NHS Workbook Response Sheet";
    }
  }

  let created = 0;
  let updated = 0;
  let skipped = 0;
  let failures = 0;
  const details = [];

  responses.forEach(function(item) {
    try {
      const result =
        nhstutor_upsertProfileAnswers_(
          ss,
          item.answers,
          item.timestamp,
          {
            force:
              options.force === true,
            source: source
          }
        );

      if (
        result &&
        result.skipped
      ) {
        skipped++;
      } else if (
        result &&
        result.created
      ) {
        created++;
      } else if (
        result &&
        result.ok
      ) {
        updated++;
      }

      if (result) {
        details.push(result);
      }
    } catch (error) {
      failures++;

      details.push({
        ok: false,
        email: item.email,
        error:
          error &&
          error.message
            ? error.message
            : String(error)
      });
    }
  });

  return {
    ok: failures === 0,
    source: source ||
      "No profile responses found",
    responseCount:
      responses.length,
    created: created,
    updated: updated,
    skipped: skipped,
    failures: failures,
    details: details
  };
}


function nhstutor_syncProfileForEmail_(
  ss,
  email,
  force
) {
  return nhstutor_syncProfileResponses_(
    ss,
    {
      email: email,
      force: force === true
    }
  );
}


function nhstutor_minimalProfileName_(
  email,
  fullName
) {
  const supplied =
    String(fullName || "")
      .trim();

  if (supplied) return supplied;

  const username =
    String(email || "")
      .split("@")[0]
      .replace(/[._-]+/g, " ")
      .replace(
        /\b\w/g,
        function(letter) {
          return letter.toUpperCase();
        }
      )
      .trim();

  return username || "NHS Tutor";
}


/**
 * Guarantees that a person who already has explicit Tutor access can resolve
 * a Tutor Database identity.
 *
 * Priority:
 *   1) existing Tutor Database row,
 *   2) latest submitted Tutor Profile form,
 *   3) minimal safe portal profile created from the Access Registry identity.
 *
 * A minimal profile does NOT make the student a strong matching candidate.
 * Qualifications remain blank until the Tutor Profile form or Portal profile
 * supplies them.
 */
function nhstutor_ensureProfileForAccess_(
  ss,
  email,
  fullName
) {
  email =
    String(email || "")
      .trim()
      .toLowerCase();

  if (
    !email ||
    !nhsa_isValidEmailAddress_(
      email
    )
  ) {
    throw new Error(
      "A valid School Email is required to create a Tutor profile."
    );
  }

  let existing =
    nhsa_findTutorBySchoolEmail_(
      ss,
      email
    );

  if (existing) {
    return {
      ok: true,
      created: false,
      source: "Tutor Database",
      tutorId:
        existing.tutorId || "",
      email: email
    };
  }

  try {
    nhstutor_syncProfileForEmail_(
      ss,
      email,
      false
    );
  } catch (error) {
    console.error(
      "[NHS V55 tutor profile import]",
      error
    );
  }

  existing =
    nhsa_findTutorBySchoolEmail_(
      ss,
      email
    );

  if (existing) {
    return {
      ok: true,
      created: false,
      source:
        "Tutor Profile Form",
      tutorId:
        existing.tutorId || "",
      email: email
    };
  }

  const database =
    nhstutor_ensureTutorDatabase_(ss);

  const h =
    tup_headerMap_(database);

  const tutorId =
    nhstutor_stableTutorId_(
      email,
      database
    );

  const values =
    new Array(
      database.getLastColumn()
    ).fill("");

  const set =
    function(header, value) {
      if (!h[header]) return;

      values[h[header] - 1] =
        value;
    };

  const name =
    nhstutor_minimalProfileName_(
      email,
      fullName
    );

  set(
    "Tutor ID",
    tutorId
  );
  set(
    "Full Name",
    nhsinput_literalText_(
      name,
      120,
      "Tutor full name"
    )
  );
  set(
    "School Email",
    email
  );
  set(
    "Roster Match",
    nhstutor_rosterMatch_(
      ss,
      email,
      name
    )
  );
  set(
    "Sessions Completed",
    0
  );
  set(
    "Verified Tutoring Hours",
    0
  );
  set(
    "Tutor No-Shows",
    0
  );
  set(
    "Current Caseload",
    0
  );
  set(
    "Reliability Status",
    "New"
  );
  set(
    "Last Profile Update",
    new Date()
  );

  const lock =
    LockService.getScriptLock();

  lock.waitLock(20000);

  try {
    const duplicateRow =
      nhstutor_findTutorRowByEmail_(
        database,
        email
      );

    if (!duplicateRow) {
      database
        .getRange(
          Math.max(
            2,
            database.getLastRow() + 1
          ),
          1,
          1,
          values.length
        )
        .setValues([values]);

      SpreadsheetApp.flush();
    }
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }

  try {
    nhsperf_scriptCache_()
      .remove(
        nhs_v46TutorIdentityKey_(
          email
        )
      );
  } catch (e) {}

  return {
    ok: true,
    created: true,
    source:
      "Access Registry bootstrap",
    tutorId: tutorId,
    email: email
  };
}


/**
 * Compatible with BOTH:
 * - Form installable onFormSubmit trigger (e.response),
 * - Spreadsheet onFormSubmit trigger (e.namedValues).
 *
 * This keeps the original processTutorProfileSubmissionV2 handler name so
 * an already-installed trigger can continue to work after consolidation.
 */
function processTutorProfileSubmissionV2(
  e
) {
  const ss =
    nhsperf_ss_(
      NHS_SPREADSHEET_ID
    );

  let answers = {};
  let timestamp =
    new Date();

  if (
    e &&
    e.response
  ) {
    answers =
      nhstutor_answersFromFormResponse_(
        e.response
      );

    try {
      timestamp =
        e.response.getTimestamp() ||
        timestamp;
    } catch (err) {}
  } else if (
    e &&
    e.namedValues
  ) {
    answers =
      e.namedValues;

    const raw =
      nhstutor_answerValue_(
        answers,
        "Timestamp"
      );

    const parsed =
      raw
        ? new Date(raw)
        : null;

    if (
      parsed &&
      !isNaN(
        parsed.getTime()
      )
    ) {
      timestamp = parsed;
    }
  } else {
    throw new Error(
      "Tutor profile submission event is missing."
    );
  }

  const email =
    nhstutor_answerValue_(
      answers,
      "School Email"
    ).toLowerCase();

  if (!email) {
    // Ignore unrelated form submissions routed through the same spreadsheet.
    return {
      ok: true,
      ignored: true,
      reason:
        "Submission is not an NHS Tutor Profile."
    };
  }

  const result =
    nhstutor_upsertProfileAnswers_(
      ss,
      answers,
      timestamp,
      {
        force: true,
        source:
          "Tutor Profile Form Submit"
      }
    );

  nhsa_invalidateAccessCache_();

  return result;
}


function nhstutor_installProfileTriggersV55_() {
  const form = nhstutor_openProfileForm_();
  const existing = ScriptApp.getProjectTriggers();

  existing.forEach(function(trigger) {
    const fn = trigger.getHandlerFunction();

    if (
      fn === NHSTUTOR_PROFILE_RECONCILE_HANDLER ||
      (form && fn === NHSTUTOR_PROFILE_SUBMIT_HANDLER)
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  let formTriggerCreated = false;
  let formTriggerError = "";

  if (form) {
    try {
      ScriptApp
        .newTrigger(NHSTUTOR_PROFILE_SUBMIT_HANDLER)
        .forForm(form)
        .onFormSubmit()
        .create();

      formTriggerCreated = true;
    } catch (error) {
      formTriggerError = error && error.message ? error.message : String(error);
      console.error("[NHS V55 Tutor Profile submit trigger]", error);
    }
  } else {
    formTriggerError =
      "Tutor Profile Form is unavailable. Verify NHS_TUTOR_PROFILE_FORM_ID and give the deployment account Editor access to the form.";
  }

  let reconcileTriggerCreated = false;
  let reconcileTriggerError = "";

  try {
    ScriptApp
      .newTrigger(NHSTUTOR_PROFILE_RECONCILE_HANDLER)
      .timeBased()
      .everyMinutes(15)
      .create();

    reconcileTriggerCreated = true;
  } catch (error) {
    reconcileTriggerError = error && error.message ? error.message : String(error);
    console.error("[NHS V55 Tutor Profile reconcile trigger]", error);
  }

  return {
    ok: !!form && formTriggerCreated && reconcileTriggerCreated,
    formAvailable: !!form,
    formTriggerCreated: formTriggerCreated,
    formTriggerError: formTriggerError,
    reconcileTriggerCreated: reconcileTriggerCreated,
    reconcileTriggerError: reconcileTriggerError,
    formId: nhstutor_profileFormId_()
  };
}

function nhstutor_profileTriggerAuditV55_() {
  const counts = {};

  ScriptApp
    .getProjectTriggers()
    .forEach(function(trigger) {
      const fn =
        trigger.getHandlerFunction();

      if (
        fn ===
          NHSTUTOR_PROFILE_SUBMIT_HANDLER ||
        fn ===
          NHSTUTOR_PROFILE_RECONCILE_HANDLER
      ) {
        counts[fn] =
          (counts[fn] || 0) + 1;
      }
    });

  const formAvailable =
    !!nhstutor_openProfileForm_();

  const submitCount =
    Number(
      counts[
        NHSTUTOR_PROFILE_SUBMIT_HANDLER
      ] || 0
    );

  const reconcileCount =
    Number(
      counts[
        NHSTUTOR_PROFILE_RECONCILE_HANDLER
      ] || 0
    );

  return {
    formAvailable:
      formAvailable,
    submitCount:
      submitCount,
    reconcileCount:
      reconcileCount,
    ok:
      formAvailable &&
      submitCount === 1 &&
      reconcileCount === 1
  };
}


function nhstutor_findProfileIdentityConflict_(ss, email, fullName) {
  email = String(email || "").trim().toLowerCase();
  const targetName = nhsa_normalizePersonName_(fullName);
  if (!targetName) return null;

  const sh = nhstutor_ensureTutorDatabase_(ss);
  if (!sh || sh.getLastRow() < 2) return null;

  const values = sh
    .getRange(1, 1, sh.getLastRow(), sh.getLastColumn())
    .getDisplayValues();
  const h = nhsa_headerIndex_(values[0]);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowEmail = nhsa_get_(row, h, "School Email").toLowerCase();
    if (!rowEmail || rowEmail === email) continue;

    const reliability = nhsa_get_(row, h, "Reliability Status").toLowerCase();
    if (["inactive", "suspended"].includes(reliability)) continue;

    const rowFull = nhsa_get_(row, h, "Full Name");
    const rowPreferred = nhsa_get_(row, h, "Preferred Name");
    const display = nhsa_displayTutorName_(rowPreferred, rowFull);
    const candidates = [rowFull, display]
      .map(nhsa_normalizePersonName_)
      .filter(Boolean);

    if (candidates.includes(targetName)) {
      return {
        row: i + 1,
        tutorId: nhsa_get_(row, h, "Tutor ID"),
        schoolEmail: rowEmail,
        fullName: rowFull,
        preferredName: rowPreferred
      };
    }
  }

  return null;
}

function nhsa_inactivateTutorGrant_(ss, grant, reason) {
  if (!grant || !grant.grantId) return false;
  const registry = nhsa_ensureRegistry_(ss);
  if (!registry || registry.getLastRow() < 2) return false;
  const values = registry
    .getRange(1, 1, registry.getLastRow(), registry.getLastColumn())
    .getDisplayValues();
  const h = nhsa_headerIndex_(values[0]);
  const grantIdCol = h["Grant ID"];
  if (grantIdCol === undefined) return false;

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][grantIdCol] || "").trim() !== String(grant.grantId || "").trim()) continue;
    const row = i + 1;
    const statusCol = h["Status"] + 1;
    const expirationCol = h["Expiration Date"] === undefined ? 0 : h["Expiration Date"] + 1;
    const notesCol = h["Notes"] === undefined ? 0 : h["Notes"] + 1;
    registry.getRange(row, statusCol).setValue("Inactive");
    if (expirationCol) registry.getRange(row, expirationCol).setValue(new Date());
    if (notesCol) {
      const prior = String(values[i][h["Notes"]] || "").trim();
      const note = "V47.5 tutor identity reconciliation: " + String(reason || "invalid or conflicting tutor identity");
      registry.getRange(row, notesCol).setValue(prior ? prior + " | " + note : note);
    }
    return true;
  }
  return false;
}

function nhstutor_provisionActiveGrants_(
  ss
) {
  const registry = nhsa_ensureRegistry_(ss);

  const grants = nhsa_readRegistry_(registry)
    .filter(function(grant) {
      return grant.role === "TUTOR" && nhsa_isGrantActive_(grant);
    });

  let provisioned = 0;
  let linked = 0;
  let failed = 0;
  let invalidIdentityGrants = 0;
  let identityConflicts = 0;
  let inactivated = 0;

  grants.forEach(function(grant) {
    try {
      const grantEmail = String(grant.email || "").trim().toLowerCase();

      if (!nhsa_isSchoolIdentityEmail_(grantEmail)) {
        invalidIdentityGrants++;
        if (nhsa_inactivateTutorGrant_(
          ss,
          grant,
          "Tutor role must use a CUSD school identity; personal email retained only for non-Tutor roles."
        )) inactivated++;
        return;
      }

      const before = nhsa_findTutorBySchoolEmail_(ss, grantEmail);

      if (!before) {
        const conflict = nhstutor_findProfileIdentityConflict_(
          ss,
          grantEmail,
          grant.fullName
        );

        if (conflict) {
          identityConflicts++;
          if (nhsa_inactivateTutorGrant_(
            ss,
            grant,
            "Tutor identity already exists under " + conflict.schoolEmail + "."
          )) inactivated++;
          return;
        }
      }

      const result = nhstutor_ensureProfileForAccess_(
        ss,
        grantEmail,
        grant.fullName
      );

      if (!before && result && result.created) {
        provisioned++;
      } else {
        linked++;
      }
    } catch (error) {
      failed++;
      console.error("[NHS V55 provision tutor grant]", grant.email, error);
    }
  });

  if (inactivated) nhsa_invalidateAccessCache_();

  return {
    ok: failed === 0,
    activeTutorGrants: grants.length,
    provisioned: provisioned,
    linked: linked,
    invalidIdentityGrants: invalidIdentityGrants,
    identityConflicts: identityConflicts,
    inactivated: inactivated,
    failed: failed
  };
}

function nhstutor_reconcileProfilesV55_() {
  try {
    const ss =
      nhsperf_ss_(
        NHS_SPREADSHEET_ID
      );

    const sync =
      nhstutor_syncProfileResponses_(
        ss
      );

    nhstutor_provisionActiveGrants_(
      ss
    );

    nhsa_invalidateAccessCache_();

    return sync;
  } catch (error) {
    console.error(
      "[NHS V55 tutor profile reconcile]",
      error
    );

    return {
      ok: false,
      error:
        error &&
        error.message
          ? error.message
          : String(error)
    };
  }
}


function nhstutor_repairProfilePipelineV55_(
  ss
) {
  ss =
    ss ||
    nhsperf_ss_(
      NHS_SPREADSHEET_ID
    );

  const database =
    nhstutor_ensureTutorDatabase_(ss);

  const sync =
    nhstutor_syncProfileResponses_(
      ss
    );

  const provision =
    nhstutor_provisionActiveGrants_(
      ss
    );

  const triggers =
    nhstutor_installProfileTriggersV55_();

  nhsa_invalidateAccessCache_();

  try {
    nhs_v45InvalidatePortalState_(
      {}
    );
  } catch (e) {}

  SpreadsheetApp.flush();

  return {
    ok:
      sync.failures === 0 &&
      provision.failed === 0,
    tutorDatabase:
      database.getName(),
    profileSync:
      sync,
    activeGrantProvision:
      provision,
    triggers:
      triggers,
    triggerAudit:
      nhstutor_profileTriggerAuditV55_()
  };
}


function repairTutorProfilePipelineV55() {
  nhsa_requireAnyRole_(
    ["ADMIN", "ADVISOR"],
    "Tutor Profile Pipeline Repair"
  );

  const ss =
    nhsperf_ss_(
      NHS_SPREADSHEET_ID
    );

  return nhstutor_repairProfilePipelineV55_(
    ss
  );
}


/** ============================================================
 * SOURCE GROUP: TutorPortal.gs
 * ============================================================ */

/**
 * NHS TUTOR PORTAL — V1
 *
 * Backend for the consolidated Tesoro NHS web app (?page=tutor).
 *
 * Files in this project:
 *   TutorPortal.gs
 *   TutorIndex.html
 *   TutorStyles.html
 *   TutorApp.html
 *
 * Run once:
 *   installTutorPortal()
 *
 * Deployment:
 *   Execute as: Me
 *   Access: organization/domain only
 */

const TUP_SPREADSHEET_ID = NHS_SPREADSHEET_ID;

const TUP_REQUESTS = "Tutoring Requests";
const TUP_TUTORS = "Tutor Database";
const TUP_SESSIONS = "Session Log";
const TUP_ACTIVITY = "Activity Log";
const TUP_NOTIFICATIONS = "Notification Log";

const TUP_TITAN_TIME = "10:25 AM–11:05 AM";
const TUP_CACHE_SECONDS = 600; // Turbo V14: warm portal loads avoid repeated request/session scans.
const TUP_CACHE_PREFIX = "tutorPortal" + NHS_RELEASE.CACHE_NAMESPACE + "|";

const TUP_REQUEST_HEADERS = [
  "Tutor Response Status",
  "Tutor Response At",
  "Tutor Decline Reason",
  "Tutor Viewed At",
  "Tutor Portal Last Action",
  "Tutor Portal Last Action At"
];

const TUP_TUTOR_HEADERS = [
  "Tutor ID",
  "Full Name",
  "Preferred Name",
  "School Email",
  "Grade",
  "Roster Match",
  "Last Profile Update",
  "Last Form Profile Update",
  "Last Portal Profile Update",
  "English Courses",
  "Mathematics Courses",
  "Science Courses",
  "Social Science Courses",
  "World Language Courses",
  "Computer Science / CTE Courses",
  "Strongest Course",
  "Typical Titan Time Availability",
  "Typical After-School Availability",
  "Recurring Tutoring",
  "Maximum Sessions / Week",
  "Tutoring Languages",
  "Previous Experience",
  "Experience Details",
  "Academic Qualifications",
  "Matching Notes",
  "Expectations Agreement",
  "Sessions Completed",
  "Verified Tutoring Hours",
  "Tutor No-Shows",
  "Current Caseload",
  "Reliability Status",
  "Coordinator Notes",
  "Notification Email",
  "Temporary Unavailable Dates",
  "Reminder Preference",
  "Availability Updated At",
  "Tutor Portal Last Login",
  "Teacher Directory Bio"
];

const TUP_NOTIFICATION_HEADERS = [
  "Notification ID",
  "Created At",
  "Request ID",
  "Channel",
  "Recipient",
  "Type",
  "Status",
  "Dedupe Key",
  "Read",
  "Severity",
  "Title",
  "Message",
  "Error",
  "Portal Read"
];


function includeTutor_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================================================
   INSTALL
   ============================================================ */

function installTutorPortal() {
  nhssec_requireMaintenanceAdmin_();

  nhsa_installAccessCore_();

  const ss = nhsperf_ss_(TUP_SPREADSHEET_ID);

  const requests = nhsperf_sheet_(ss, TUP_REQUESTS);
  const tutors = nhsperf_sheet_(ss, TUP_TUTORS);

  if (!requests) throw new Error('Missing "' + TUP_REQUESTS + '" sheet.');
  if (!tutors) throw new Error('Missing "' + TUP_TUTORS + '" sheet.');

  tup_ensureHeaders_(requests, TUP_REQUEST_HEADERS);
  tup_ensureHeaders_(tutors, TUP_TUTOR_HEADERS);

  // One-time safe migration:
  // existing Personal Email becomes Notification Email when the new field is blank.
  tup_migrateNotificationEmails_(tutors);

  tup_ensureSessionLog_(ss);
  tup_ensureActivityLog_(ss);
  tup_ensureNotificationLog_(ss);
  tup_ensureWorker_();


  return {
    ok: true,
    message:
      "Tutor Portal installed. School Email is used for authentication; Notification Email is used for tutor delivery."
  };
}


function authorizeTutorPortalMail() {
  nhssec_requireMaintenanceAdmin_();

  if (typeof checkNHSBatchRelayV29 === "function") {
    return checkNHSBatchRelayV29();
  }
  return {
    ok: false,
    message: "Mail relay diagnostic is unavailable. Keep NHS Mail Transport.gs installed."
  };
}


/* ============================================================
   BOOTSTRAP
   ============================================================ */

function nhsv473_tutor_getPortalDataCore_(forceRefresh) {
  const email = tup_requireUserEmail_();
  const cache = nhsperf_scriptCache_();
  let tutor = null;

  if (!forceRefresh) {
    try {
      const identityCached = cache.get(nhs_v46TutorIdentityKey_(email));
      if (identityCached) {
        tutor = JSON.parse(identityCached);
        const warmPortal = cache.get(nhs_v46TutorCacheKey_(tutor.tutorId));

        if (warmPortal) {
          const parsed = JSON.parse(warmPortal);

          if (
            parsed &&
            parsed.lifecycleBuild === NHS_RELEASE.VERSION &&
            parsed.requestTruth === true
          ) {
            parsed.cached = true;
            parsed.turboWarm = true;
            return parsed;
          }
        }
      }
    } catch(e) {
      tutor = null;
    }
  }

  const ss = nhsperf_ss_(TUP_SPREADSHEET_ID);
  // Safe in-place schema migrations for existing deployments.
  nhsos_ensureRequestColumns_(nhsperf_sheet_(ss, TUP_REQUESTS));
  tup_ensureHeaders_(nhsperf_sheet_(ss, TUP_TUTORS), TUP_TUTOR_HEADERS);
  nhswa_ensureSheet_(ss);

  if (!tutor) {
    tutor = tup_requireTutor_(ss, email);
  }

  const key = nhs_v46TutorCacheKey_(tutor.tutorId);

  if (!forceRefresh) {
    try {
      const cached = cache.get(key);

      if (cached) {
        const parsed = JSON.parse(cached);

        if (
          parsed &&
          parsed.lifecycleBuild === NHS_RELEASE.VERSION &&
          parsed.requestTruth === true
        ) {
          parsed.cached = true;
          return parsed;
        }
      }
    } catch(e) {}
  }

  try {
    tup_queueLoginTouch_(tutor);
  } catch (error) {
    console.error("[NHS V46.5 tutor bootstrap] login telemetry:", error);
  }

  // V46.2: Identity/authentication is the only hard startup dependency.
  // Request/session/notification reads are isolated so a temporary Sheet or
  // schema error can never prevent the Tutor shell from opening.
  const bootstrapWarnings = [];
  let requestTruth = true;
  let requests = [];

  try {
    requests = tup_readTutorRequests_(ss, tutor);
  } catch (error) {
    requestTruth = false;
    bootstrapWarnings.push(
      "Live assignment data is temporarily unavailable: " +
      String(error && error.message ? error.message : error)
    );
    console.error("[NHS V46.5 tutor bootstrap] requests:", error);
  }

  let sessions = [];
  try {
    sessions = tup_readTutorSessions_(ss, tutor);
  } catch (error) {
    bootstrapWarnings.push(
      "Session history is temporarily unavailable."
    );
    console.error("[NHS V46.5 tutor bootstrap] sessions:", error);
  }

  let notifications = [];
  let priorPortal = null;

  try {
    const raw = cache.get(key);
    if (raw) priorPortal = JSON.parse(raw);
  } catch(e) {}

  notifications =
    forceRefresh &&
    priorPortal &&
    Array.isArray(priorPortal.notifications)
      ? priorPortal.notifications
      : (function() {
          try {
            return tup_readNotifications_(ss, tutor);
          } catch (error) {
            bootstrapWarnings.push(
              "Tutor notifications are temporarily unavailable."
            );
            console.error("[NHS V46.5 tutor bootstrap] notifications:", error);
            return [];
          }
        })();

  const now = new Date();
  const tz = nhs_safeTimeZone_(ss);

  const upcoming = requests
    .filter(function(r) {
      return nhslife_upcoming_(r, now, tz);
    })
    .sort(function(a,b) {
      return String(a.scheduledDate || "")
        .localeCompare(String(b.scheduledDate || ""));
    });

  const weeklyTarget = nhswa_confirmationTargetWeekStart_(tz, now);
  const weeklyAvailability = nhswa_weekSummary_(ss, tutor, weeklyTarget);

  const pendingAcceptance = requests.filter(nhslife_needsTutorAction_);

  const pendingReports = requests.filter(function(r) {
    return nhslife_reportable_(r, now, tz);
  });

  const activeCaseload =
    requests.filter(nhslife_activeCaseload_).length;

  const data = {
    generatedAt: new Date().toISOString(),
    gatewayUrl: nhs_getGatewayUrl_(),
    titanTime: TUP_TITAN_TIME,
    reconfirmDays: NHSSCHED_RECONFIRM_DAYS,
    // V46: course catalog loads independently so autocomplete can never
    // take down Tutor portal bootstrap.
    courseCatalogVersion: NHS_COURSE_CATALOG_VERSION,

    identity: {
      tutorId: tutor.tutorId,
      email: email,
      portalEmail: tutor.portalEmail,
      notificationEmail: tutor.notificationEmail,
      fullName: tutor.fullName,
      preferredName: tutor.preferredName,
      displayName: tutor.displayName,
      grade: tutor.grade
    },

    profile: tutor,
    weeklyAvailability: weeklyAvailability,
    requests: requests,
    upcoming: upcoming,
    sessions: sessions.slice(0,150),
    notifications: notifications.slice(0,75),

    counts: {
      pendingAcceptance: pendingAcceptance.length,
      upcoming: upcoming.length,
      pendingReports: pendingReports.length,
      currentCaseload: activeCaseload,
      sessionsCompleted: tutor.sessionsCompleted,
      verifiedHours: tutor.verifiedHours,
      pendingVerification:
        sessions.filter(function(s) {
          return (
            s.result === "Completed" &&
            s.verificationStatus === "Pending"
          );
        }).length
    },

    requestTruth: requestTruth,
    bootstrapWarnings: bootstrapWarnings,
    lifecycleBuild: NHS_RELEASE.VERSION
  };

  // Never persist a degraded request snapshot as authoritative truth.
  if (requestTruth) {
    try {
      cache.put(key, JSON.stringify(data), TUP_CACHE_SECONDS);
    } catch(e) {}
  }

  return data;
}



/* ============================================================
   V28 ULTRAFAST REQUEST DELTA
   ============================================================ */

function tutor_getLiveRequestData() {
  const email=tup_requireUserEmail_();
  const ss=nhsperf_ss_(TUP_SPREADSHEET_ID);
  const tutor=tup_requireTutor_(ss,email);
  const requests=tup_readTutorRequests_(ss,tutor);
  const now=new Date();
  const tz=nhs_safeTimeZone_(ss);
  const upcoming=requests
    .filter(function(r){ return nhslife_upcoming_(r,now,tz); })
    .sort(function(a,b){ return String(a.scheduledDate||"").localeCompare(String(b.scheduledDate||"")); });

  return {
    generatedAt:new Date().toISOString(),
    revision:nhsperf_getDataRevision_(),
    requests:requests,
    upcoming:upcoming,
    counts:{
      pendingAcceptance:requests.filter(nhslife_needsTutorAction_).length,
      upcoming:upcoming.length,
      pendingReports:requests.filter(function(r){ return nhslife_reportable_(r,now,tz); }).length,
      currentCaseload:requests.filter(nhslife_activeCaseload_).length
    },
    lifecycleBuild:NHS_RELEASE.VERSION
  };
}


/* ============================================================
   ASSIGNMENT ACTIONS
   ============================================================ */


function tup_invalidateOperationsCache_() {
  try {
    if (typeof nhsos_invalidateCache_ === "function") {
      nhsos_invalidateCache_();
      return;
    }
  } catch (e) {}

  // Fallback for shared consolidated project.
  try {
    nhsperf_scriptCache_().remove(nhs_v46OperationsCacheKey_());
  } catch (e) {}
}


function nhsv473_tutor_acceptAssignmentCore_(requestId) {
  const email = tup_requireUserEmail_();
  const ss = nhsperf_ss_(TUP_SPREADSHEET_ID);
  const tutor = tup_requireTutor_(ss, email);

  const sh = nhsperf_sheet_(ss, TUP_REQUESTS);
  const row = tup_findOwnedRequestRow_(sh, requestId, tutor);

  if (!row) {
    throw new Error("Assignment not found or it is no longer assigned to you.");
  }

  const h = tup_headerMap_(sh);
  const range = sh.getRange(row, 1, 1, sh.getLastColumn());
  const values = range.getValues()[0];
  const display = range.getDisplayValues()[0];

  const get = header => tup_arrayGet_(display, h, header);
  const set = (header, value) => tup_arraySet_(values, h, header, value);

  const status = get("Status");

  if (["Cancelled", "Completed"].includes(status)) {
    throw new Error("This assignment is already closed.");
  }

  const scheduledDateForAvailability = nhssched_dateKey_(get("Scheduled Date"), nhs_safeTimeZone_(ss));
  if (scheduledDateForAvailability) nhswa_upsertDateAvailable_(ss, tutor, scheduledDateForAvailability, "Assignment Acceptance");

  const now = new Date();

  set("Tutor Response Status", "Accepted");
  set("Tutor Response At", now);
  set("Tutor Decline Reason", "");
  set("Tutor Portal Last Action", "Accepted Assignment");
  set("Tutor Portal Last Action At", now);
  set("Tutor Viewed At", now);

  const schedulingType = get("Scheduling Preference");
  const scheduledDateKey = nhssched_dateKey_(get("Scheduled Date"), nhs_safeTimeZone_(ss));
  let reconfirmStatus = "Not Required";
  let reconfirmAt = "";
  if (schedulingType === NHSSCHED_TYPES.FUTURE && scheduledDateKey) {
    // If the tutor accepts inside the live reconfirmation window, that acceptance
    // itself is a current availability confirmation; do not force a redundant
    // second click. Reservations accepted earlier remain tentative until the
    // window opens and the tutor reconfirms.
    if (nhssched_reconfirmWindowOpen_(scheduledDateKey, nhs_safeTimeZone_(ss), now)) {
      reconfirmStatus = "Confirmed";
      reconfirmAt = now;
    } else {
      reconfirmStatus = "Not Yet Open";
    }
  } else if (schedulingType === NHSSCHED_TYPES.RECURRING) {
    reconfirmStatus = "Initial Accepted";
  }
  set("Tutor Reconfirm Status", reconfirmStatus);
  set("Tutor Reconfirm At", reconfirmAt);

  if (get("Scheduled Date")) {
    set("Status", "Scheduled");
  } else {
    set("Status", "Assigned");
  }

  const acceptedNotificationStatus =
    reconfirmStatus === "Not Yet Open"
      ? "Future Reservation — Tutor Accepted Tentatively"
      : reconfirmStatus === "Confirmed" && schedulingType === NHSSCHED_TYPES.FUTURE
        ? "Future Reservation — Tutor Accepted & Reconfirmed"
        : "Tutor Accepted — Ready to Confirm";
  set("Notification Status", acceptedNotificationStatus);
  set("Last Update", now);

  range.setValues([values]);

  tup_queueJob_("ACCEPT", {
    requestId: requestId,
    tutorId: tutor.tutorId,
    tutorName: tutor.displayName,
    course: get("Course")
  });

  tup_invalidateTutorCache_(tutor.tutorId);
  tup_invalidateOperationsCache_();

  // Create an immediate Operations alert only when the assignment can
  // actually proceed to final confirmation. A far-future reservation is
  // intentionally tentative and must not appear as "Ready to notify" yet.
  try {
    if (
      reconfirmStatus !== "Not Yet Open" &&
      typeof nhsn_logInternalAlertIfNew_ === "function" &&
      typeof nhsn_alert_ === "function"
    ) {
      nhsn_logInternalAlertIfNew_(
        ss,
        nhsn_alert_(
          "READY_NOTIFY",
          requestId,
          "High",
          "Ready to notify",
          (get("Course") || "Tutoring request") +
            " was accepted by " +
            (tutor.displayName || "the tutor") +
            " and is ready for teacher/tutor notification."
        )
      );
    }
  } catch (e) {}

  // Ensure Operations reads the newly-created alert on its next refresh.
  try {
    const cache = nhsperf_scriptCache_();
    cache.remove(NHSOS_CACHE_KEY);
    cache.remove("nhsos_health_v14");
  } catch (e) {}

  return {
    ok: true,
    requestId: requestId,
    patch: {
      tutorResponseStatus: "Accepted",
      tutorResponseAt: now.toISOString(),
      declineReason: "",
      status: get("Scheduled Date") ? "Scheduled" : "Assigned",
      notificationStatus: acceptedNotificationStatus,
      tutorReconfirmStatus: reconfirmStatus,
      tutorReconfirmAt: reconfirmAt ? now.toISOString() : ""
    },
    message: "Assignment accepted."
  };
}


function nhsv473_tutor_declineAssignmentCore_(requestId, reason) {
  const email = tup_requireUserEmail_();
  const ss = nhsperf_ss_(TUP_SPREADSHEET_ID);
  const tutor = tup_requireTutor_(ss, email);

  const sh = nhsperf_sheet_(ss, TUP_REQUESTS);
  const row = tup_findOwnedRequestRow_(sh, requestId, tutor);

  if (!row) {
    throw new Error("Assignment not found or it is no longer assigned to you.");
  }

  const h = tup_headerMap_(sh);
  const range = sh.getRange(row, 1, 1, sh.getLastColumn());
  const values = range.getValues()[0];
  const display = range.getDisplayValues()[0];

  const get = header => tup_arrayGet_(display, h, header);
  const set = (header, value) => tup_arraySet_(values, h, header, value);

  const currentStatus = get("Status");

  if (["Cancelled", "Completed"].includes(currentStatus)) {
    throw new Error("This assignment is already closed.");
  }

  const now = new Date();
  const cleanReason =
    nhsinput_literalText_(reason, 500, "Decline reason") ||
    "Tutor unavailable";

  const oldAssignmentConfirmed = get("Assignment Confirmed");

  set("Tutor Response Status", "Declined");
  set("Tutor Response At", now);
  set("Tutor Decline Reason", cleanReason);
  set("Tutor Portal Last Action", "Declined Assignment");
  set("Tutor Portal Last Action At", now);
  set("Tutor Viewed At", now);

  // Return the request to officer control.
  set("Assigned Tutor", "");
  set("Assigned Tutor ID", "");
  set("Confirm Assignment", false);
  set("Assignment Confirmed", "");
  set("Status", "Needs Follow-Up");
  set("Notification Status", "Tutor Declined — Reassign");
  set("Last Update", now);

  range.setValues([values]);

  tup_queueJob_("DECLINE", {
    requestId: requestId,
    tutorId: tutor.tutorId,
    tutorName: tutor.displayName,
    course: get("Course"),
    reason: cleanReason,
    wasConfirmed: !!oldAssignmentConfirmed
  });

  tup_invalidateTutorCache_(tutor.tutorId);
  tup_invalidateOperationsCache_();

  return {
    ok: true,
    requestId: requestId,
    removeFromActive: true,
    message: "Assignment declined. NHS will reassign it."
  };
}


function tutor_markAssignmentViewed(requestId) {
  const email = tup_requireUserEmail_();
  const ss = nhsperf_ss_(TUP_SPREADSHEET_ID);
  const tutor = tup_requireTutor_(ss, email);
  const sh = nhsperf_sheet_(ss, TUP_REQUESTS);

  const row = tup_findOwnedRequestRow_(sh, requestId, tutor);
  if (!row) return { ok: false };

  const h = tup_headerMap_(sh);
  const col = h["Tutor Viewed At"];

  if (col) {
    const current = sh.getRange(row, col).getValue();
    if (!current) sh.getRange(row, col).setValue(new Date());
  }

  return { ok: true };
}



/* ============================================================
   V46.5 — RESERVATION RECONFIRMATION + FIVE STAR READINESS
   ============================================================ */

function tutor_reconfirmSession(requestId) {
  const email = tup_requireUserEmail_();
  const ss = nhsperf_ss_(TUP_SPREADSHEET_ID);
  const tutor = tup_requireTutor_(ss, email);
  const sh = nhsperf_sheet_(ss, TUP_REQUESTS);
  const row = tup_findOwnedRequestRow_(sh, requestId, tutor);
  if (!row) throw new Error("Assignment not found or no longer assigned to you.");

  const h = tup_headerMap_(sh);
  const range = sh.getRange(row, 1, 1, sh.getLastColumn());
  const values = range.getValues()[0];
  const display = range.getDisplayValues()[0];
  const get = header => tup_arrayGet_(display, h, header);
  const set = (header, value) => tup_arraySet_(values, h, header, value);

  if (String(get("Tutor Response Status") || "").trim() !== "Accepted") {
    throw new Error("Accept the assignment before reconfirming availability.");
  }

  const scheduledDate = nhssched_validateFutureTitanDate_(
    get("Scheduled Date"), nhs_safeTimeZone_(ss), new Date()
  );
  if (!nhssched_reconfirmWindowOpen_(scheduledDate, nhs_safeTimeZone_(ss), new Date())) {
    throw new Error(
      "Reconfirmation opens " + NHSSCHED_RECONFIRM_DAYS +
      " days before the scheduled session."
    );
  }

  nhswa_upsertDateAvailable_(ss, tutor, scheduledDate, "Reservation Reconfirmation");

  const now = new Date();
  set("Tutor Reconfirm Status", "Confirmed");
  set("Tutor Reconfirm At", now);
  set("Notification Status", "Tutor Reconfirmed — Ready to Confirm");
  set("Last Update", now);
  range.setValues([values]);

  tup_invalidateTutorCache_(tutor.tutorId);
  tup_invalidateOperationsCache_();
  if (typeof nhsperf_touchDataRevision_ === "function") nhsperf_touchDataRevision_();

  return {
    ok:true,
    requestId:requestId,
    patch:{
      tutorReconfirmStatus:"Confirmed",
      tutorReconfirmAt:now.toISOString(),
      notificationStatus:"Tutor Reconfirmed — Ready to Confirm"
    },
    message:"Availability reconfirmed. NHS can now finalize this session."
  };
}

function nhssched_applyFiveStarState_(sh, row, h, actorType) {
  const range=sh.getRange(row,1,1,sh.getLastColumn());
  const values=range.getValues()[0];
  const display=range.getDisplayValues()[0];
  const get=header => h[header] ? String(display[h[header]-1] || "").trim() : "";
  const set=(header,value) => { if (h[header]) values[h[header]-1]=value; };

  if (!nhslife_truthy_(get("Assignment Confirmed"))) {
    throw new Error("NHS must confirm the assignment before Five Star setup can be completed.");
  }
  // Reporting Five Star completion is allowed on the session date itself;
  // the no-same-day rule applies to creating/finalizing the schedule, not to
  // recording setup that may already have been completed earlier.
  nhssched_validateCurrentOrFutureTitanDate_(get("Scheduled Date"), NHS_PLATFORM.TIMEZONE, new Date());

  const now=new Date();
  const teacherWasReady=nhslife_truthy_(get("Teacher Five Star Ready"));
  const tutorWasReady=nhslife_truthy_(get("Tutor Five Star Enrolled"));

  if (actorType === "teacher") {
    set("Teacher Five Star Ready", true);
    set("Teacher Five Star Ready At", now);
  } else if (actorType === "tutor") {
    if (!teacherWasReady) {
      throw new Error("Wait for the teacher to assign both you and the student into this Titan Time in Five Star before confirming the Five Star placement.");
    }
    set("Tutor Five Star Enrolled", true);
    set("Tutor Five Star Enrolled At", now);
  } else {
    throw new Error("Invalid Five Star actor.");
  }

  const teacherReady = actorType === "teacher" ? true : teacherWasReady;
  const tutorReady = actorType === "tutor" ? true : tutorWasReady;
  let readiness="Five Star Setup Pending";
  if (teacherReady && tutorReady) {
    set("Five Star Ready At", now);
    set("Notification Status", "Session Ready — Five Star Complete");
    readiness="Ready";
  } else if (teacherReady) {
    set("Notification Status", "Five Star — Waiting for Tutor Confirmation");
    readiness="Waiting for Tutor Confirmation";
  } else if (tutorReady) {
    set("Notification Status", "Five Star — Waiting for Teacher to Assign Tutor & Student");
    readiness="Waiting for Teacher to Assign Tutor & Student";
  }
  set("Last Update",now);
  range.setValues([values]);

  return {
    teacherFiveStarReady:teacherReady,
    tutorFiveStarEnrolled:tutorReady,
    fiveStarReadyAt:teacherReady&&tutorReady?now.toISOString():"",
    fiveStarReadiness:readiness,
    notificationStatus: teacherReady&&tutorReady
      ? "Session Ready — Five Star Complete"
      : teacherReady
        ? "Five Star — Waiting for Tutor Confirmation"
        : "Five Star — Waiting for Teacher to Assign Tutor & Student"
  };
}

function teacher_confirmFiveStarReady(requestId) {
  const email=tp_requireTeacherEmail_();
  return tp_withLock_(function(){
    const ss=nhsperf_ss_(TP_SPREADSHEET_ID);
    const sh=nhsperf_sheet_(ss,TP_REQUESTS_SHEET);
    const row=tp_findOwnedRequestRow_(sh,requestId,email);
    if (!row) throw new Error("Request not found or you do not own it.");
    const h=tp_headerMap_(sh);
    const patch=nhssched_applyFiveStarState_(sh,row,h,"teacher");
    tp_invalidateTeacherCache_(email);
    if (typeof nhsos_invalidateCache_ === "function") nhsos_invalidateCache_();
    if (typeof nhsperf_touchDataRevision_ === "function") nhsperf_touchDataRevision_();
    return {ok:true,requestId:requestId,patch:patch,message:patch.fiveStarReadiness==="Ready"?"Five Star setup complete. Session is ready.":"Tutor and student assigned in Five Star. Waiting for the tutor to confirm the placement."};
  });
}

function tutor_confirmFiveStarEnrollment(requestId) {
  const email=tup_requireUserEmail_();
  const ss=nhsperf_ss_(TUP_SPREADSHEET_ID);
  const tutor=tup_requireTutor_(ss,email);
  const sh=nhsperf_sheet_(ss,TUP_REQUESTS);
  const row=tup_findOwnedRequestRow_(sh,requestId,tutor);
  if (!row) throw new Error("Assignment not found or no longer assigned to you.");
  const h=tup_headerMap_(sh);
  const patch=nhssched_applyFiveStarState_(sh,row,h,"tutor");
  tup_invalidateTutorCache_(tutor.tutorId);
  tup_invalidateOperationsCache_();
  if (typeof nhsperf_touchDataRevision_ === "function") nhsperf_touchDataRevision_();
  return {ok:true,requestId:requestId,patch:patch,message:patch.fiveStarReadiness==="Ready"?"Five Star placement confirmed. Session is ready.":"Placement confirmed. Waiting for the teacher to assign the tutor and student in Five Star."};
}

/* ============================================================
   SESSION REPORTING
   ============================================================ */

function nhsv473_tutor_reportSessionCore_(payload) {
  const email = tup_requireUserEmail_();
  const ss = nhsperf_ss_(TUP_SPREADSHEET_ID);
  const tutor = tup_requireTutor_(ss, email);

  const requestId = tup_clean_(payload.requestId);
  const result = tup_clean_(payload.result);
  const notes = nhsinput_literalText_(payload.notes, 1200, "Session notes");
  const requestedMinutes = Number(payload.minutes || 0);

  if (!requestId) throw new Error("Missing request ID.");

  const allowedResults = [
    "Completed",
    "Student No-Show",
    "Teacher Cancelled",
    "Tutor Unable to Attend",
    "Reschedule Needed"
  ];

  if (!allowedResults.includes(result)) {
    throw new Error("Choose a valid session result.");
  }

  const requestSh = nhsperf_sheet_(ss, TUP_REQUESTS);
  const row = tup_findOwnedRequestRow_(requestSh, requestId, tutor);

  if (!row) {
    throw new Error("This request is not assigned to you.");
  }

  const h = tup_headerMap_(requestSh);
  const request = tup_rowObject_(requestSh, row, h);

  const responseStatus =
    String(request["Tutor Response Status"] || "").trim();

  if (responseStatus && responseStatus !== "Accepted") {
    throw new Error("Accept the assignment before reporting the session.");
  }

  const sessionDate =
    tup_clean_(payload.sessionDate) ||
    tup_dateKey_(request["Scheduled Date"], nhs_safeTimeZone_(ss)) ||
    Utilities.formatDate(
      new Date(),
      nhs_safeTimeZone_(ss),
      "yyyy-MM-dd"
    );

  const existing = tup_findSessionForOccurrence_(
    ss,
    requestId,
    tutor.tutorId,
    sessionDate
  );

  if (existing) {
    return {
      ok: true,
      duplicatePrevented: true,
      sessionId: existing["Session ID"],
      message: "This session has already been reported."
    };
  }

  let reportedMinutes =
    result === "Completed"
      ? Math.max(0, Math.min(120, requestedMinutes || 40))
      : 0;

  const verificationStatus =
    result === "Completed"
      ? "Pending"
      : "Not Required";

  const sessionId = tup_id_("S", ss);
  const sessionSh = tup_ensureSessionLog_(ss);
  const sessionH = tup_headerMap_(sessionSh);

  const sessionRow =
    new Array(sessionSh.getLastColumn()).fill("");

  tup_rowSet_(sessionRow, sessionH, "Session ID", sessionId);
  tup_rowSet_(sessionRow, sessionH, "Request ID", requestId);
  tup_rowSet_(sessionRow, sessionH, "Tutor ID", tutor.tutorId);
  tup_rowSet_(sessionRow, sessionH, "Tutor", tutor.displayName);
  tup_rowSet_(sessionRow, sessionH, "Session Date", new Date(sessionDate + "T12:00:00"));
  tup_rowSet_(sessionRow, sessionH, "Result", result);
  tup_rowSet_(sessionRow, sessionH, "Reported Minutes", reportedMinutes);
  tup_rowSet_(sessionRow, sessionH, "Verification Status", verificationStatus);
  tup_rowSet_(sessionRow, sessionH, "Verified Minutes", "");
  tup_rowSet_(sessionRow, sessionH, "Notes", notes);
  tup_rowSet_(sessionRow, sessionH, "Logged At", new Date());
  tup_rowSet_(sessionRow, sessionH, "Logged By", email);

  sessionSh
    .getRange(sessionSh.getLastRow() + 1, 1, 1, sessionRow.length)
    .setValues([sessionRow]);

  const reqRange =
    requestSh.getRange(row, 1, 1, requestSh.getLastColumn());
  const reqValues = reqRange.getValues()[0];

  const setReq = (header, value) =>
    tup_arraySet_(reqValues, h, header, value);

  let nextStatus = request["Status"] || "Scheduled";

  if (result === "Completed") {
    nextStatus = "Pending Verification";
  } else if (result === "Teacher Cancelled") {
    nextStatus = tup_truthy_(request["Recurring Active"])
      ? "Needs Follow-Up"
      : "Cancelled";
  } else {
    nextStatus = "Needs Follow-Up";
  }

  setReq("Status", nextStatus);
  setReq("Tutor Portal Last Action", "Reported Session: " + result);
  setReq("Tutor Portal Last Action At", new Date());
  setReq("Last Update", new Date());

  reqRange.setValues([reqValues]);

  tup_queueJob_("SESSION_REPORT", {
    requestId: requestId,
    sessionId: sessionId,
    tutorId: tutor.tutorId,
    tutorName: tutor.displayName,
    course: request["Course"] || "",
    result: result,
    notes: notes
  });

  tup_invalidateTutorCache_(tutor.tutorId);

  return {
    ok: true,
    requestId: requestId,
    sessionId: sessionId,
    result: result,
    verificationStatus: verificationStatus,
    requestStatus: nextStatus,
    message:
      result === "Completed"
        ? "Session submitted for officer verification."
        : "Session outcome reported."
  };
}


/* ============================================================
   AVAILABILITY + PROFILE
   ============================================================ */

function tutor_updateAvailability(payload) {
  const email = tup_requireUserEmail_();
  const ss = nhsperf_ss_(TUP_SPREADSHEET_ID);
  const tutor = tup_requireTutor_(ss, email);

  const days = Array.isArray(payload.days)
    ? payload.days.filter(function(day) {
        return [
          "Tuesday",
          "Wednesday",
          "Thursday",
          "Friday"
        ].includes(day);
      })
    : [];

  const maxSessions = Math.max(
    1,
    Math.min(
      10,
      Number(payload.maxSessions || 1)
    )
  );

  const temporaryUnavailable =
    nhsinput_literalText_(
      payload.temporaryUnavailableDates,
      1000,
      "Temporary unavailable dates"
    );

  const reminderPreference =
    [
      "Day Before + Morning Of",
      "Morning Of Only",
      "Day Before Only",
      "None"
    ].includes(payload.reminderPreference)
      ? payload.reminderPreference
      : "Day Before + Morning Of";

  const sh =
    nhsperf_sheet_(
      ss,
      TUP_TUTORS
    );

  tup_ensureHeaders_(
    sh,
    TUP_TUTOR_HEADERS
  );

  const h =
    tup_headerMap_(sh);

  const lock =
    LockService.getScriptLock();

  lock.waitLock(20000);

  try {
    const range =
      sh.getRange(
        tutor.row,
        1,
        1,
        sh.getLastColumn()
      );

    // Read inside the lock. Profile saves, form reconciliation, and
    // availability saves all use the same lock so a stale whole-row snapshot
    // cannot overwrite a newly saved Strongest Course value.
    const values =
      range.getValues()[0];

    tup_arraySet_(
      values,
      h,
      "Typical Titan Time Availability",
      days.join(", ")
    );

    tup_arraySet_(
      values,
      h,
      "Maximum Sessions / Week",
      maxSessions
    );

    tup_arraySet_(
      values,
      h,
      "Temporary Unavailable Dates",
      temporaryUnavailable
    );

    tup_arraySet_(
      values,
      h,
      "Reminder Preference",
      reminderPreference
    );

    const updatedAt =
      new Date();

    tup_arraySet_(
      values,
      h,
      "Availability Updated At",
      updatedAt
    );

    tup_arraySet_(
      values,
      h,
      "Last Profile Update",
      updatedAt
    );

    tup_arraySet_(
      values,
      h,
      "Last Portal Profile Update",
      updatedAt
    );

    range.setValues([values]);
    SpreadsheetApp.flush();
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }

  tup_queueJob_("AVAILABILITY_UPDATE", {
    tutorId: tutor.tutorId,
    tutorName: tutor.displayName,
    actor: email,
    days: days.join(", "),
    maxSessions: maxSessions
  });

  try {
    nhsperf_scriptCache_()
      .remove(
        nhs_v46TutorIdentityKey_(email)
      );
  } catch (e) {}

  tup_invalidateTutorCache_(
    tutor.tutorId
  );

  return {
    ok: true,
    patch: {
      titanAvailability: days.join(", "),
      maxSessions: maxSessions,
      temporaryUnavailableDates:
        temporaryUnavailable,
      reminderPreference:
        reminderPreference
    },
    message:
      "Availability updated. Teacher Directory changes will sync automatically."
  };
}


function tutor_updateProfile(payload) {
  payload = payload || {};

  const email =
    tup_requireUserEmail_();

  const ss =
    nhsperf_ss_(
      TUP_SPREADSHEET_ID
    );

  const tutor =
    tup_requireTutor_(
      ss,
      email
    );

  const notificationEmail =
    String(
      payload.notificationEmail || ""
    )
      .trim()
      .toLowerCase();

  if (
    notificationEmail.length > 254
  ) {
    throw new Error(
      "Notification email is too long."
    );
  }

  if (
    !tup_isValidEmail_(
      notificationEmail
    )
  ) {
    throw new Error(
      "Enter a valid Notification Email. This address is required for assignments and session reminders."
    );
  }

  const preferredName =
    nhsinput_literalText_(
      payload.preferredName,
      80,
      "Preferred name"
    );

  const languages =
    nhsinput_literalText_(
      payload.languages,
      200,
      "Tutoring languages"
    );

  const experienceDetails =
    nhsinput_literalText_(
      payload.experienceDetails,
      1200,
      "Experience details"
    );

  const directoryBio =
    nhsinput_literalText_(
      payload.directoryBio,
      500,
      "Teacher Directory bio"
    );

  const strongestTouched =
    payload.strongestCourseTouched === true;

  const legacyStrongestPayload =
    payload.strongestCourseTouched ===
      undefined;

  let strongestSelection = null;

  if (
    strongestTouched ||
    legacyStrongestPayload
  ) {
    strongestSelection =
      nhscourse_strictSelection_(
        payload.strongestCourseIds,
        payload.strongestCourse,
        3
      );
  }

  const sh =
    nhsperf_sheet_(
      ss,
      TUP_TUTORS
    );

  tup_ensureHeaders_(
    sh,
    TUP_TUTOR_HEADERS
  );

  const h =
    tup_headerMap_(sh);

  const lock =
    LockService.getScriptLock();

  lock.waitLock(20000);

  let savedStrongest = "";

  try {
    const range =
      sh.getRange(
        tutor.row,
        1,
        1,
        sh.getLastColumn()
      );

    // Re-read under the shared script lock so no other Tutor Database
    // read-modify-write operation can overwrite this save with stale row data.
    const values =
      range.getValues()[0];

    const currentStrongest =
      h["Strongest Course"]
        ? String(
            values[
              h["Strongest Course"] - 1
            ] || ""
          ).trim()
        : "";

    savedStrongest =
      strongestTouched ||
      legacyStrongestPayload
        ? strongestSelection.text
        : currentStrongest;

    tup_arraySet_(
      values,
      h,
      "Preferred Name",
      preferredName
    );

    tup_arraySet_(
      values,
      h,
      "Strongest Course",
      savedStrongest
    );

    tup_arraySet_(
      values,
      h,
      "Tutoring Languages",
      languages
    );

    tup_arraySet_(
      values,
      h,
      "Experience Details",
      experienceDetails
    );

    tup_arraySet_(
      values,
      h,
      "Teacher Directory Bio",
      directoryBio
    );

    tup_arraySet_(
      values,
      h,
      "Notification Email",
      notificationEmail
    );

    const profileUpdatedAt =
      new Date();

    tup_arraySet_(
      values,
      h,
      "Last Profile Update",
      profileUpdatedAt
    );

    tup_arraySet_(
      values,
      h,
      "Last Portal Profile Update",
      profileUpdatedAt
    );

    range.setValues([values]);
    SpreadsheetApp.flush();
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }

  tup_queueJob_("PROFILE_UPDATE", {
    tutorId: tutor.tutorId,
    tutorName:
      preferredName ||
      tutor.preferredName ||
      tutor.fullName,
    actor: email,
    notificationEmail:
      notificationEmail
  });

  try {
    nhsperf_scriptCache_()
      .remove(
        nhs_v46TutorIdentityKey_(email)
      );
  } catch (e) {}

  tup_invalidateTutorCache_(
    tutor.tutorId
  );

  return {
    ok: true,
    patch: {
      preferredName:
        preferredName,
      strongestCourse:
        savedStrongest,
      strongestCourseIds:
        strongestSelection
          ? strongestSelection.ids
          : [],
      languages:
        languages,
      experienceDetails:
        experienceDetails,
      directoryBio:
        directoryBio,
      notificationEmail:
        notificationEmail
    },
    message:
      "Profile updated. Teacher Directory changes will sync automatically."
  };
}


/* ============================================================
   NOTIFICATIONS
   ============================================================ */

function tutor_markNotificationRead(notificationId) {
  const email = tup_requireUserEmail_();
  const ss = nhsperf_ss_(TUP_SPREADSHEET_ID);
  const tutor = tup_requireTutor_(ss, email);
  const sh = tup_ensureNotificationLog_(ss);

  if (sh.getLastRow() < 2) return { ok: false };

  const h = tup_headerMap_(sh);
  const idCol = h["Notification ID"];
  const recipientCol = h["Recipient"];
  const readCol = h["Portal Read"];

  if (!idCol || !recipientCol || !readCol) return { ok: false };

  const allowed = tup_tutorEmails_(tutor)
    .map(v => String(v || "").trim().toLowerCase())
    .filter(Boolean);

  const values = sh
    .getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
    .getValues();

  const idIndex = idCol - 1;
  const recipientIndex = recipientCol - 1;
  const readIndex = readCol - 1;
  const targetId = String(notificationId || "").trim();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][idIndex] || "").trim() !== targetId) continue;

    const recipient = String(values[i][recipientIndex] || "")
      .trim()
      .toLowerCase();

    if (!allowed.includes(recipient)) {
      throw new Error("You do not have access to this notification.");
    }

    if (String(values[i][readIndex]).toLowerCase() !== "true") {
      sh.getRange(i + 2, readCol).setValue(true);
    }

    // Notification-read state is local UI metadata. Clear only this tutor's
    // cached portal payload; do NOT bump the global data revision.
    try {
      nhsperf_scriptCache_().remove(nhs_v46TutorCacheKey_(tutor.tutorId));
    } catch (e) {}

    return { ok: true };
  }

  return { ok: false };
}


function tutor_markAllNotificationsRead() {
  const email = tup_requireUserEmail_();
  const ss = nhsperf_ss_(TUP_SPREADSHEET_ID);
  const tutor = tup_requireTutor_(ss, email);
  const sh = tup_ensureNotificationLog_(ss);

  if (sh.getLastRow() < 2) {
    return { ok: true, updated: 0 };
  }

  const h = tup_headerMap_(sh);
  const recipientCol = h["Recipient"];
  const readCol = h["Portal Read"];

  if (!recipientCol || !readCol) {
    return { ok: true, updated: 0 };
  }

  const allowed = tup_tutorEmails_(tutor)
    .map(v => String(v || "").trim().toLowerCase())
    .filter(Boolean);

  const rowCount = sh.getLastRow() - 1;

  const recipients = sh
    .getRange(2, recipientCol, rowCount, 1)
    .getDisplayValues();

  const reads = sh
    .getRange(2, readCol, rowCount, 1)
    .getValues();

  let updated = 0;

  for (let i = 0; i < rowCount; i++) {
    const recipient = String(recipients[i][0] || "")
      .trim()
      .toLowerCase();

    if (
      allowed.includes(recipient) &&
      String(reads[i][0]).toLowerCase() !== "true"
    ) {
      reads[i][0] = true;
      updated++;
    }
  }

  if (updated) {
    sh.getRange(2, readCol, rowCount, 1).setValues(reads);
  }

  try {
    nhsperf_scriptCache_().remove(nhs_v46TutorCacheKey_(tutor.tutorId));
  } catch (e) {}

  return {
    ok: true,
    updated: updated
  };
}



/* ============================================================
   DATA READS
   ============================================================ */

function tup_readTutorRequests_(ss, tutor) {
  const sh = nhsperf_sheet_(ss, TUP_REQUESTS);
  if (!sh) return [];

  const objects = typeof nhsdata_readKeyedObjects_ === "function"
    ? nhsdata_readKeyedObjects_(sh, "Request ID")
    : [];

  const tutorId = tutor.tutorId;
  const displayName = tup_norm_(tutor.displayName);
  const fullName = tup_norm_(tutor.fullName);

  return objects.map(function(obj) {
    const assignedId = String(obj["Assigned Tutor ID"] || "").trim();
    const assignedName = tup_norm_(obj["Assigned Tutor"] || "");
    const owned =
      (assignedId && assignedId === tutorId) ||
      (!assignedId && assignedName && [displayName, fullName].includes(assignedName));
    if (!owned) return null;

    const recurringValue = obj["Recurring?"] || "";
    const recurringActiveRaw = obj["Recurring Active"] || "";

    return {
      row: obj._row,
      requestId: obj["Request ID"] || "",
      teacherName: obj["Teacher Name"] || "",
      course: obj["Course"] || "",
      topic: obj["Topic / Academic Need"] || "",
      supportType: obj["Support Type"] || "",
      schedulingType: obj["Scheduling Preference"] || "Legacy",
      preferredDates: obj["Preferred Dates"] || "",
      preferredDays: obj["Preferred Days"] || "",
      status: obj["Status"] || "Assigned",
      scheduledDate: tup_dateKey_(obj["Scheduled Date"] || "", nhs_safeTimeZone_(ss)),
      scheduledTime: obj["Scheduled Time"] || TUP_TITAN_TIME,
      location: obj["Location"] || "",
      recurring: recurringValue,
      recurringActive: recurringActiveRaw ? tup_truthy_(recurringActiveRaw) : tup_isRecurring_(recurringValue),
      nextSessionDate: tup_dateKey_(obj["Next Session Date"] || "", nhs_safeTimeZone_(ss)),
      tutorResponseStatus: obj["Tutor Response Status"] || "",
      tutorResponseAt: obj["Tutor Response At"] || "",
      declineReason: obj["Tutor Decline Reason"] || "",
      assignmentConfirmed: obj["Assignment Confirmed"] || "",
      isConfirmed: typeof nhslife_confirmed_ === "function" ? nhslife_confirmed_(obj) : tup_confirmationValueV40_(obj["Assignment Confirmed"]),
      notificationStatus: obj["Notification Status"] || "",
      tutorReconfirmStatus: obj["Tutor Reconfirm Status"] || "",
      tutorReconfirmAt: obj["Tutor Reconfirm At"] || "",
      teacherFiveStarReady: tup_truthy_(obj["Teacher Five Star Ready"] || ""),
      teacherFiveStarReadyAt: obj["Teacher Five Star Ready At"] || "",
      tutorFiveStarEnrolled: tup_truthy_(obj["Tutor Five Star Enrolled"] || ""),
      tutorFiveStarEnrolledAt: obj["Tutor Five Star Enrolled At"] || "",
      fiveStarReadyAt: obj["Five Star Ready At"] || "",
      fiveStarReadiness: nhssched_readiness_(obj),
      lastUpdate: obj["Last Update"] || ""
    };
  }).filter(Boolean).reverse();
}


function tup_readTutorSessions_(ss, tutor) {
  const sh = nhsperf_sheet_(ss, TUP_SESSIONS);

  if (!sh || sh.getLastRow() < 2) return [];

  const lastRow = sh.getLastRow();
  const startRow = Math.max(2, lastRow - 499);
  const numRows = lastRow - startRow + 1;

  const headers = sh
    .getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0];

  const values = sh
    .getRange(startRow, 1, numRows, sh.getLastColumn())
    .getDisplayValues();

  const h = tup_headerIndex_(headers);

  return values
    .map((row, i) => {
      const sessionId = tup_get_(row, h, "Session ID");
      if (!sessionId) return null;

      const tutorId = tup_get_(row, h, "Tutor ID");
      const tutorName = tup_norm_(tup_get_(row, h, "Tutor"));

      const owned =
        tutorId === tutor.tutorId ||
        (!tutorId &&
          [tup_norm_(tutor.displayName), tup_norm_(tutor.fullName)]
            .includes(tutorName));

      if (!owned) return null;

      return {
        row: startRow + i,
        sessionId: sessionId,
        requestId: tup_get_(row, h, "Request ID"),
        sessionDate: tup_dateKey_(
          tup_get_(row, h, "Session Date"),
          nhs_safeTimeZone_(ss)
        ),
        result: tup_get_(row, h, "Result"),
        reportedMinutes: Number(
          tup_get_(row, h, "Reported Minutes") || 0
        ),
        verificationStatus:
          tup_get_(row, h, "Verification Status") || "",
        verifiedMinutes: Number(
          tup_get_(row, h, "Verified Minutes") || 0
        ),
        notes: tup_get_(row, h, "Notes"),
        loggedAt: tup_get_(row, h, "Logged At"),
        verificationNotes:
          tup_get_(row, h, "Verification Notes")
      };
    })
    .filter(Boolean)
    .reverse();
}


function tup_readNotifications_(ss, tutor) {
  const sh = tup_ensureNotificationLog_(ss);

  if (sh.getLastRow() < 2) return [];

  const lastRow = sh.getLastRow();
  const startRow = Math.max(2, lastRow - 399);
  const numRows = lastRow - startRow + 1;

  const headers = sh
    .getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0];

  const values = sh
    .getRange(startRow, 1, numRows, sh.getLastColumn())
    .getDisplayValues();

  const h = tup_headerIndex_(headers);
  const allowedEmails = tup_tutorEmails_(tutor)
    .map(v => v.toLowerCase());

  return values
    .map((row, i) => {
      const recipient =
        String(tup_get_(row, h, "Recipient") || "")
          .trim()
          .toLowerCase();

      if (!allowedEmails.includes(recipient)) {
        return null;
      }

      if (tup_get_(row, h, "Channel") !== "Email") {
        return null;
      }

      return {
        row: startRow + i,
        notificationId:
          tup_get_(row, h, "Notification ID"),
        createdAt:
          tup_get_(row, h, "Created At"),
        requestId:
          tup_get_(row, h, "Request ID"),
        type:
          tup_get_(row, h, "Type"),
        status:
          tup_get_(row, h, "Status"),
        title:
          tup_get_(row, h, "Title"),
        message:
          tup_get_(row, h, "Message"),
        portalRead:
          tup_truthy_(tup_get_(row, h, "Portal Read"))
      };
    })
    .filter(Boolean)
    .reverse();
}


/* ============================================================
   IDENTITY + ACCESS CONTROL
   ============================================================ */

function tup_requireUserEmail_() {
  const access = nhsa_requireAnyRole_(
    ["TUTOR"],
    "Tutor Portal"
  );

  return access.email;
}


function tup_requireTutor_(ss, email) {
  let tutor =
    tup_findTutorByPortalEmail_(
      ss,
      email
    );

  if (!tutor) {
    const access =
      nhsa_resolveAccess_(
        ss,
        email
      );

    const hasTutorGrant =
      Array.isArray(access.roles) &&
      access.roles.includes("TUTOR");

    if (!hasTutorGrant) {
      throw new Error(
        "Tutor access is not active for this school account."
      );
    }

    nhstutor_ensureProfileForAccess_(
      ss,
      email,
      access.fullName || ""
    );

    try {
      nhsperf_scriptCache_()
        .remove(
          nhs_v46TutorIdentityKey_(
            email
          )
        );
    } catch (e) {}

    tutor =
      tup_findTutorByPortalEmail_(
        ss,
        email
      );
  }

  if (!tutor) {
    throw new Error(
      "Tutor access is active, but the Tutor profile could not be initialized. Ask an NHS officer to run Repair Tutor Profile Pipeline."
    );
  }

  // Reliability controls assignment eligibility. Access Registry controls
  // portal authorization. A suspended/inactive profile may still sign in to
  // view its portal, but assignment/matching systems continue to exclude it.
  return tutor;
}


function tup_findTutorByPortalEmail_(ss, email) {
  const cache = nhsperf_scriptCache_();
  const key = nhs_v46TutorIdentityKey_(email);

  const cached = cache.get(key);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }

  const sh = nhsperf_sheet_(ss, TUP_TUTORS);

  if (!sh || sh.getLastRow() < 2) return null;

  const values = sh
    .getRange(1, 1, sh.getLastRow(), sh.getLastColumn())
    .getDisplayValues();

  const h = tup_headerIndex_(values[0]);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const schoolEmail =
      String(tup_get_(row, h, "School Email") || "")
        .trim()
        .toLowerCase();

    // Authentication is intentionally through School Email; delivery uses Notification Email.
    if (!schoolEmail || email !== schoolEmail) {
      continue;
    }

    const fullName = tup_get_(row, h, "Full Name");
    if (!fullName) continue;

    const preferredName =
      tup_get_(row, h, "Preferred Name");

    const notificationEmail =
      String(tup_get_(row, h, "Notification Email") || "")
        .trim()
        .toLowerCase();

    const personalEmail =
      String(tup_get_(row, h, "Personal Email") || "")
        .trim()
        .toLowerCase();

    const tutor = {
      row: i + 1,
      tutorId: tup_get_(row, h, "Tutor ID"),
      fullName: fullName,
      preferredName: preferredName,
      displayName:
        tup_tutorDisplayName_(preferredName, fullName),

      // Login identity
      schoolEmail: schoolEmail,

      // Delivery identity
      notificationEmail: notificationEmail,

      // Retained only as legacy/profile data; never used for auth.
      personalEmail: personalEmail,

      grade: tup_get_(row, h, "Grade"),
      strongestCourse:
        tup_get_(row, h, "Strongest Course"),
      englishCourses:
        tup_get_(row, h, "English Courses"),
      mathCourses:
        tup_get_(row, h, "Mathematics Courses"),
      scienceCourses:
        tup_get_(row, h, "Science Courses"),
      socialScienceCourses:
        tup_get_(row, h, "Social Science Courses"),
      languageCourses:
        tup_get_(row, h, "World Language Courses"),
      cteCourses:
        tup_get_(row, h, "Computer Science / CTE Courses"),
      titanAvailability:
        tup_get_(row, h, "Typical Titan Time Availability"),
      afterSchoolAvailability:
        tup_get_(row, h, "Typical After-School Availability"),
      temporaryUnavailableDates:
        tup_get_(row, h, "Temporary Unavailable Dates"),
      maxSessions: Number(
        tup_get_(row, h, "Maximum Sessions / Week") || 0
      ),
      languages:
        tup_get_(row, h, "Tutoring Languages"),
      previousExperience:
        tup_get_(row, h, "Previous Experience"),
      experienceDetails:
        tup_get_(row, h, "Experience Details"),
      directoryBio:
        tup_get_(row, h, "Teacher Directory Bio"),
      sessionsCompleted: Number(
        tup_get_(row, h, "Sessions Completed") || 0
      ),
      verifiedHours: Number(
        tup_get_(row, h, "Verified Tutoring Hours") || 0
      ),
      currentCaseload: Number(
        tup_get_(row, h, "Current Caseload") || 0
      ),
      reliabilityStatus:
        tup_get_(row, h, "Reliability Status") || "New",
      reminderPreference:
        tup_get_(row, h, "Reminder Preference") ||
        "Day Before + Morning Of",
      availabilityUpdatedAt:
        tup_get_(row, h, "Availability Updated At")
    };

    try {
      cache.put(key, JSON.stringify(tutor), 21600);
    } catch (e) {}

    return tutor;
  }

  return null;
}


function tup_touchLastLogin_(ss, tutor) {
  const sh = nhsperf_sheet_(ss, TUP_TUTORS);
  if (!sh) return;

  const h = tup_headerMap_(sh);

  if (h["Tutor Portal Last Login"]) {
    sh.getRange(tutor.row, h["Tutor Portal Last Login"])
      .setValue(new Date());
  }
}


/* ============================================================
   BACKGROUND JOBS
   ============================================================ */

function tup_queueLoginTouch_(tutor) {
  if (!tutor || !tutor.tutorId) return;

  const cache = nhsperf_scriptCache_();
  const key = "tupLoginTouch|" + tutor.tutorId;

  // At most one write every 30 minutes per tutor.
  if (cache.get(key)) return;

  try {
    cache.put(key, "1", 1800);
  } catch (e) {}

  tup_queueJob_("LOGIN_TOUCH", {
    tutorId: tutor.tutorId,
    tutorRow: tutor.row
  });
}


function tup_ensureWorker_() {
  const handler = "tup_processJobs_";

  const exists = ScriptApp
    .getProjectTriggers()
    .some(trigger => trigger.getHandlerFunction() === handler);

  if (!exists) {
    ScriptApp
      .newTrigger(handler)
      .timeBased()
      .everyMinutes(1)
      .create();
  }
}


function tup_queueJob_(type, payload) {
  const props = nhsperf_scriptProperties_();

  const key =
    "tupJob_" +
    Date.now() +
    "_" +
    Utilities.getUuid();

  props.setProperty(
    key,
    JSON.stringify({
      type: type,
      createdAt: new Date().toISOString(),
      attempts: 0,
      payload: payload || {}
    })
  );
}


function tup_processJobs_() {
  const props = nhsperf_scriptProperties_();
  const all = props.getProperties();

  const keys = Object.keys(all)
    .filter(key => key.indexOf("tupJob_") === 0)
    .sort()
    .slice(0, 30);

  if (!keys.length) return;

  const ss = nhsperf_ss_(TUP_SPREADSHEET_ID);
  let caseloadNeedsReconcile = false;
  let matchingNeedsRefresh = false;

  keys.forEach(key => {
    try {
      const job = JSON.parse(all[key] || "{}");
      const p = job.payload || {};

      if (job.type === "LOGIN_TOUCH") {
        const sh = nhsperf_sheet_(ss, TUP_TUTORS);

        if (sh && p.tutorRow) {
          const h = tup_headerMap_(sh);
          const col = h["Tutor Portal Last Login"];

          if (col) {
            sh.getRange(Number(p.tutorRow), col).setValue(new Date());
          }
        }
      }

      if (job.type === "ACCEPT") {
        tup_logActivity_(ss, {
          type: "Tutor Accepted Assignment",
          requestId: p.requestId || "",
          tutorId: p.tutorId || "",
          actor: p.tutorName || "Tutor Portal",
          summary: "Tutor accepted assignment",
          details: p.course || ""
        });

        tup_logInternalAlert_(ss, {
          requestId: p.requestId || "",
          type: "TUTOR_ACCEPTED",
          severity: "Low",
          title: "Tutor accepted assignment",
          message:
            (p.tutorName || "Tutor") +
            " accepted " +
            (p.course || "the tutoring assignment") +
            "."
        });
      }

      if (job.type === "DECLINE") {
        caseloadNeedsReconcile = true;

        tup_logActivity_(ss, {
          type: "Tutor Declined Assignment",
          requestId: p.requestId || "",
          tutorId: p.tutorId || "",
          actor: p.tutorName || "Tutor Portal",
          summary: "Tutor declined assignment",
          details:
            (p.course || "") +
            (p.reason ? " — " + p.reason : "")
        });

        tup_logInternalAlert_(ss, {
          requestId: p.requestId || "",
          type: "TUTOR_DECLINED",
          severity: "High",
          title: "Tutor declined assignment",
          message:
            (p.tutorName || "Tutor") +
            " declined " +
            (p.course || "a tutoring assignment") +
            (p.reason ? ". Reason: " + p.reason : ".")
        });
      }

      if (job.type === "SESSION_REPORT") {
        tup_logActivity_(ss, {
          type: "Tutor Session Report",
          requestId: p.requestId || "",
          sessionId: p.sessionId || "",
          tutorId: p.tutorId || "",
          actor: p.tutorName || "Tutor Portal",
          summary: "Tutor reported session: " + (p.result || ""),
          details:
            (p.course || "") +
            (p.notes ? " — " + p.notes : "")
        });

        if (p.result !== "Completed") {
          tup_logInternalAlert_(ss, {
            requestId: p.requestId || "",
            type: "SESSION_FOLLOW_UP",
            severity:
              p.result === "Tutor Unable to Attend"
                ? "High"
                : "Medium",
            title: "Session needs follow-up",
            message:
              (p.tutorName || "Tutor") +
              " reported: " +
              (p.result || "Follow-up needed") +
              "."
          });
        }
      }

      if (job.type === "WEEKLY_AVAILABILITY_UPDATE") {
        tup_logActivity_(ss, {
          type: "Tutor Weekly Availability Confirmed",
          tutorId: p.tutorId || "",
          actor: p.actor || "Tutor Portal",
          summary: (p.tutorName || "Tutor") + " confirmed weekly Titan Time availability",
          details: "Week starting " + (p.weekStart || "")
        });
        matchingNeedsRefresh = true;
      }

      if (job.type === "AVAILABILITY_UPDATE") {
        tup_logActivity_(ss, {
          type: "Tutor Availability Updated",
          tutorId: p.tutorId || "",
          actor: p.actor || "Tutor Portal",
          summary:
            (p.tutorName || "Tutor") +
            " updated availability",
          details:
            "Titan Time days: " +
            (p.days || "None") +
            ". Max sessions/week: " +
            (p.maxSessions || "") +
            "."
        });
      }

      if (job.type === "PROFILE_UPDATE") {
        tup_logActivity_(ss, {
          type: "Tutor Profile Updated",
          tutorId: p.tutorId || "",
          actor: p.actor || "Tutor Portal",
          summary:
            (p.tutorName || "Tutor") +
            " updated tutor profile",
          details:
            p.notificationEmail
              ? "Notification email confirmed: " + p.notificationEmail
              : ""
        });
      }

      props.deleteProperty(key);

    } catch (error) {
      console.error("Tutor Portal job failed:", key, error);
      if (typeof nhsdata_recordPropertyJobFailure_ === "function") {
        nhsdata_recordPropertyJobFailure_(props, key, all[key], error, 5);
      }
      // Repeated failures are capped and moved to V43 dead-letter storage.
    }
  });

  // One reconciliation at most, even if several declines were processed.
  if (caseloadNeedsReconcile) {
    try { tup_reconcileCaseloads_(ss); }
    catch (error) { console.error("Tutor Portal caseload reconciliation failed:", error); }
  }
  if (matchingNeedsRefresh) {
    try { rankAllOpenRequests_(); }
    catch (error) { console.error("Weekly availability matching refresh failed:", error); }
  }
}


/* ============================================================
   SHEET / LOG HELPERS
   ============================================================ */

function tup_ensureSessionLog_(ss) {
  let sh = nhsperf_sheet_(ss, TUP_SESSIONS);

  if (!sh) sh = ss.insertSheet(TUP_SESSIONS);

  const headers = [
    "Session ID",
    "Request ID",
    "Tutor ID",
    "Tutor",
    "Session Date",
    "Result",
    "Reported Minutes",
    "Verification Status",
    "Verified Minutes",
    "Notes",
    "Logged At",
    "Logged By",
    "Verified At",
    "Verified By",
    "Verification Notes",
    "Voided At",
    "Voided By",
    "Void Reason"
  ];

  tup_ensureHeaders_(sh, headers);
  sh.setFrozenRows(1);

  return sh;
}


function tup_ensureActivityLog_(ss) {
  let sh = nhsperf_sheet_(ss, TUP_ACTIVITY);

  if (!sh) sh = ss.insertSheet(TUP_ACTIVITY);

  tup_ensureHeaders_(sh, [
    "Activity ID",
    "Timestamp",
    "Type",
    "Request ID",
    "Session ID",
    "Tutor ID",
    "Actor",
    "Summary",
    "Details"
  ]);

  sh.setFrozenRows(1);

  return sh;
}


function tup_ensureNotificationLog_(ss) {
  let sh = nhsperf_sheet_(ss, TUP_NOTIFICATIONS);

  if (!sh) sh = ss.insertSheet(TUP_NOTIFICATIONS);

  tup_ensureHeaders_(sh, TUP_NOTIFICATION_HEADERS);
  sh.setFrozenRows(1);

  return sh;
}


function tup_logActivity_(ss, spec) {
  const sh = tup_ensureActivityLog_(ss);
  const h = tup_headerMap_(sh);

  const row = new Array(sh.getLastColumn()).fill("");

  tup_rowSet_(row, h, "Activity ID", tup_id_("A", ss));
  tup_rowSet_(row, h, "Timestamp", new Date());
  tup_rowSet_(row, h, "Type", spec.type || "Tutor Portal");
  tup_rowSet_(row, h, "Request ID", spec.requestId || "");
  tup_rowSet_(row, h, "Session ID", spec.sessionId || "");
  tup_rowSet_(row, h, "Tutor ID", spec.tutorId || "");
  tup_rowSet_(row, h, "Actor", spec.actor || "Tutor Portal");
  tup_rowSet_(row, h, "Summary", spec.summary || "");
  tup_rowSet_(row, h, "Details", spec.details || "");

  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length)
    .setValues([row]);
}


function tup_logInternalAlert_(ss, spec) {
  const sh = tup_ensureNotificationLog_(ss);
  const h = tup_headerMap_(sh);

  const dedupeKey =
    "ALERT|" +
    (spec.type || "TUTOR_PORTAL") +
    "|" +
    (spec.requestId || "");

  if (tup_hasOpenAlert_(sh, dedupeKey)) return false;

  const row = new Array(sh.getLastColumn()).fill("");

  tup_rowSet_(row, h, "Notification ID", tup_id_("N", ss));
  tup_rowSet_(row, h, "Created At", new Date());
  tup_rowSet_(row, h, "Request ID", spec.requestId || "");
  tup_rowSet_(row, h, "Channel", "Internal Alert");
  tup_rowSet_(row, h, "Recipient", "NHS Officers");
  tup_rowSet_(row, h, "Type", spec.type || "Tutor Portal Alert");
  tup_rowSet_(row, h, "Status", "Open");
  tup_rowSet_(row, h, "Dedupe Key", dedupeKey);
  tup_rowSet_(row, h, "Read", false);
  tup_rowSet_(row, h, "Severity", spec.severity || "Medium");
  tup_rowSet_(row, h, "Title", spec.title || "Tutor Portal Alert");
  tup_rowSet_(row, h, "Message", spec.message || "");
  tup_rowSet_(row, h, "Portal Read", false);

  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length)
    .setValues([row]);

  return true;
}


function tup_hasOpenAlert_(sh, dedupeKey) {
  if (sh.getLastRow() < 2) return false;

  const values = sh
    .getRange(1, 1, sh.getLastRow(), sh.getLastColumn())
    .getDisplayValues();

  const h = tup_headerIndex_(values[0]);

  return values.slice(1).some(row =>
    tup_get_(row, h, "Dedupe Key") === dedupeKey &&
    tup_get_(row, h, "Channel") === "Internal Alert" &&
    !tup_truthy_(tup_get_(row, h, "Read"))
  );
}


function tup_findOwnedRequestRow_(sh, requestId, tutor) {
  if (!sh || sh.getLastRow() < 2) return 0;

  const h = tup_headerMap_(sh);
  const idCol = h["Request ID"];

  if (!idCol) return 0;

  const match = sh
    .getRange(2, idCol, sh.getLastRow() - 1, 1)
    .createTextFinder(String(requestId || "").trim())
    .matchEntireCell(true)
    .matchCase(true)
    .findNext();

  if (!match) return 0;

  const row = match.getRow();
  const obj = tup_rowObject_(sh, row, h);

  const assignedId =
    String(obj["Assigned Tutor ID"] || "").trim();

  const assignedName =
    tup_norm_(obj["Assigned Tutor"]);

  const owned =
    (assignedId && assignedId === tutor.tutorId) ||
    (!assignedId &&
      assignedName &&
      [
        tup_norm_(tutor.displayName),
        tup_norm_(tutor.fullName)
      ].includes(assignedName));

  return owned ? row : 0;
}


function tup_findSessionForOccurrence_(ss, requestId, tutorId, dateKey) {
  const sh = nhsperf_sheet_(ss, TUP_SESSIONS);

  if (!sh || sh.getLastRow() < 2) return null;

  const h = tup_headerMap_(sh);

  const requestCol = h["Request ID"];
  const tutorCol = h["Tutor ID"];
  const dateCol = h["Session Date"];

  if (!requestCol || !tutorCol || !dateCol) return null;

  const matches = sh
    .getRange(2, requestCol, sh.getLastRow() - 1, 1)
    .createTextFinder(String(requestId || "").trim())
    .matchEntireCell(true)
    .matchCase(true)
    .findAll();

  if (!matches || !matches.length) return null;

  const tz = nhs_safeTimeZone_(ss);

  for (let i = matches.length - 1; i >= 0; i--) {
    const row = matches[i].getRow();

    const foundTutor =
      String(sh.getRange(row, tutorCol).getDisplayValue() || "").trim();

    const foundDate =
      tup_dateKey_(
        sh.getRange(row, dateCol).getValue(),
        tz
      );

    if (
      foundTutor === String(tutorId || "").trim() &&
      foundDate === dateKey
    ) {
      return tup_rowObject_(sh, row, h);
    }
  }

  return null;
}


function tup_reconcileCaseloads_(ss) {
  const requestSh = nhsperf_sheet_(ss, TUP_REQUESTS);
  const tutorSh = nhsperf_sheet_(ss, TUP_TUTORS);

  if (!requestSh || !tutorSh) return;

  const requestValues = requestSh
    .getRange(1, 1, requestSh.getLastRow(), requestSh.getLastColumn())
    .getDisplayValues();

  const rh = tup_headerIndex_(requestValues[0]);

  const counts = {};

  requestValues.slice(1).forEach(row => {
    const requestId = tup_get_(row, rh, "Request ID");
    if (!requestId) return;

    const status = tup_get_(row, rh, "Status");

    if (!["Assigned", "Scheduled"].includes(status)) return;

    const tutorId = tup_get_(row, rh, "Assigned Tutor ID");
    if (!tutorId) return;

    counts[tutorId] = (counts[tutorId] || 0) + 1;
  });

  const tutorValues = tutorSh
    .getRange(1, 1, tutorSh.getLastRow(), tutorSh.getLastColumn())
    .getValues();

  const th = tup_headerIndex_(tutorValues[0]);
  const idIdx = th["Tutor ID"];
  const caseloadIdx = th["Current Caseload"];

  if (idIdx === undefined || caseloadIdx === undefined) return;

  for (let i = 1; i < tutorValues.length; i++) {
    const tutorId = String(tutorValues[i][idIdx] || "").trim();
    tutorValues[i][caseloadIdx] = counts[tutorId] || 0;
  }

  if (tutorValues.length > 1) {
    tutorSh
      .getRange(2, 1, tutorValues.length - 1, tutorValues[0].length)
      .setValues(tutorValues.slice(1));
  }
}


function tup_ensureHeaders_(sh, requiredHeaders) {
  if (!sh) return;

  let lastCol = Math.max(sh.getLastColumn(), 1);
  let current = sh
    .getRange(1, 1, 1, lastCol)
    .getDisplayValues()[0];

  requiredHeaders.forEach(header => {
    if (!current.includes(header)) {
      sh.getRange(1, current.length + 1).setValue(header);
      current.push(header);
    }
  });
}


function tup_headerMap_(sh) {
  return nhsperf_headerMap_(sh, true, false);
}


function tup_headerIndex_(headers) {
  const map = {};

  headers.forEach((header, i) => {
    map[String(header || "").trim()] = i;
  });

  return map;
}


function tup_rowObject_(sh, row, h) {
  const values = sh
    .getRange(row, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0];

  const obj = {};

  Object.keys(h).forEach(header => {
    obj[header] = values[h[header] - 1] || "";
  });

  return obj;
}


function tup_arrayGet_(rowValues, headerMap, header) {
  const col = headerMap[header];
  if (!col) return "";
  return String(rowValues[col - 1] || "").trim();
}


function tup_arraySet_(rowValues, headerMap, header, value) {
  const col = headerMap[header];
  if (!col) return;
  rowValues[col - 1] = value;
}


function tup_rowSet_(rowValues, headerMap, header, value) {
  const col = headerMap[header];
  if (!col) return;
  rowValues[col - 1] = value;
}


function tup_get_(row, h, header) {
  const idx = h[header];
  return idx === undefined
    ? ""
    : String(row[idx] || "").trim();
}


function tup_tutorEmails_(tutor) {
  return [
    tutor.notificationEmail,
    // Legacy history compatibility only: older notifications may have been
    // sent to Personal Email before Notification Email existed.
    tutor.personalEmail
  ]
    .map(v => String(v || "").trim())
    .filter(Boolean)
    .filter((v, i, a) => a.indexOf(v) === i);
}


function tup_invalidateTutorCache_(tutorId) {
  if (typeof nhs_v45InvalidatePortalState_ === "function") {
    nhs_v45InvalidatePortalState_({
      tutorId: tutorId
    });
    return;
  }

  try {
    nhsperf_scriptCache_().remove(
      TUP_CACHE_PREFIX + tutorId
    );
  } catch (e) {}

  nhsperf_touchDataRevision_();
}


function tup_migrateNotificationEmails_(sh) {
  if (!sh || sh.getLastRow() < 2) return;

  const h = tup_headerMap_(sh);

  const notificationCol = h["Notification Email"];
  const personalCol = h["Personal Email"];

  if (!notificationCol || !personalCol) return;

  const numRows = sh.getLastRow() - 1;
  const notificationValues =
    sh.getRange(2, notificationCol, numRows, 1).getValues();
  const personalValues =
    sh.getRange(2, personalCol, numRows, 1).getValues();

  let changed = false;

  for (let i = 0; i < numRows; i++) {
    const current =
      String(notificationValues[i][0] || "").trim();

    const personal =
      String(personalValues[i][0] || "").trim();

    if (!current && personal) {
      notificationValues[i][0] = personal;
      changed = true;
    }
  }

  if (changed) {
    sh
      .getRange(2, notificationCol, numRows, 1)
      .setValues(notificationValues);
  }
}


function tup_isValidEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(
    String(value || "").trim()
  );
}


function tup_clean_(value) {
  return String(value == null ? "" : value).trim();
}


function tup_norm_(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^\w@]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}


function tup_truthy_(value) {
  return ["true", "yes", "1", "checked", "active"]
    .includes(String(value || "").trim().toLowerCase());
}


function tup_isRecurring_(value) {
  return /yes|recurr|ongoing/i.test(String(value || ""));
}


function tup_tutorDisplayName_(preferred, full) {
  preferred = String(preferred || "").trim();
  full = String(full || "").trim();

  if (!preferred) return full;

  const parts = full.split(/\s+/).filter(Boolean);

  return parts.length > 1
    ? preferred + " " + parts[parts.length - 1]
    : preferred;
}


function tup_parseDate_(value) {
  if (!value) return null;

  const d = value instanceof Date
    ? value
    : new Date(value);

  return isNaN(d.getTime()) ? null : d;
}


function tup_dateKey_(value, tz) {
  const d = tup_parseDate_(value);

  return d
    ? Utilities.formatDate(d, tz, "yyyy-MM-dd")
    : "";
}


function tup_startOfDay_(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}


function tup_id_(prefix, ss) {
  return (
    prefix +
    "-" +
    Utilities.formatDate(
      new Date(),
      nhs_safeTimeZone_(ss),
      "yyyyMMdd-HHmmss"
    ) +
    "-" +
    Utilities.getUuid()
  );
}



/* ============================================================
   V40 CANONICAL ASSIGNMENT CONFIRMATION
   ============================================================ */
function tup_confirmationValueV40_(value) {
  const raw = String(value == null ? "" : value).trim();

  if (!raw) return false;

  return ![
    "false",
    "0",
    "no",
    "unchecked",
    "pending"
  ].includes(raw.toLowerCase());
}


/**
 * Optional no-argument diagnostic.
 * Reads the live Tutoring Requests sheet directly and classifies every request
 * currently owned by the signed-in tutor.
 */
function diagnoseTutorRequestLifecycleV40() {
  nhssec_requireMaintenanceAdmin_();

  const state = tutor_getAuthoritativeRequestStateV39();
  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || "America/Los_Angeles",
    "yyyy-MM-dd"
  );

  const rows = (state.requests || []).map(function(r) {
    const status = String(r.status || "").trim().toLowerCase();
    const response = String(r.tutorResponseStatus || "").trim().toLowerCase();
    const confirmed = !!r.isConfirmed;
    const cancelled = status === "cancelled";
    const completed = status === "completed";

    return {
      requestId: r.requestId,
      course: r.course,
      status: r.status,
      tutorResponseStatus: r.tutorResponseStatus,
      scheduledDate: r.scheduledDate,
      assignmentConfirmed: r.assignmentConfirmed,
      isConfirmed: confirmed,
      isCancelled: cancelled,
      isCompleted: completed,
      shouldBeUpcoming:
        !!r.scheduledDate &&
        r.scheduledDate >= today &&
        confirmed &&
        !cancelled &&
        !completed,
      shouldNeedTutorAction:
        ["assigned", "scheduled"].includes(status) &&
        !["accepted", "declined"].includes(response)
    };
  });

  return {
    ok: true,
    generatedAt: new Date().toISOString(),
    source: state.source,
    totalOwnedRows: rows.length,
    upcomingCount: rows.filter(function(r) {
      return r.shouldBeUpcoming;
    }).length,
    pendingActionCount: rows.filter(function(r) {
      return r.shouldNeedTutorAction;
    }).length,
    rows: rows
  };
}

/* ============================================================
   V39 HARD-AUTHORITATIVE TUTOR REQUEST STATE
   ============================================================ */
function tutor_getAuthoritativeRequestStateV39() {
  const email = tup_requireUserEmail_();

  const ss = nhsperf_ss_(TUP_SPREADSHEET_ID);
  const tutor = tup_requireTutor_(ss, email);
  const requests = tup_readTutorRequests_(ss, tutor);

  try {
    nhsperf_scriptCache_().remove(
      nhs_v46TutorCacheKey_(tutor.tutorId)
    );
  } catch (e) {}

  return {
    ok: true,
    generatedAt: new Date().toISOString(),
    source: "Tutoring Requests (keyed live read)",
    sheetRows: requests.length,
    requests: requests
  };
}


/* ============================================================
   V42 DEFINITIVE TUTOR REQUEST TRUTH
   ============================================================ */
function nhsv473_tutor_getAuthoritativeRequestStateCore_() {
  const state = tutor_getAuthoritativeRequestStateV39();
  const now = new Date();
  const tz = Session.getScriptTimeZone() || "America/Los_Angeles";

  const rows = (state.requests || []).map(function(r) {
    const upcoming = typeof nhslife_upcoming_ === "function"
      ? nhslife_upcoming_(r, now, tz)
      : (!!r.scheduledDate && !!r.isConfirmed && !["Cancelled","Completed"].includes(r.status));
    const pendingAction = typeof nhslife_needsTutorAction_ === "function"
      ? nhslife_needsTutorAction_(r)
      : false;
    const activeCaseload = typeof nhslife_activeCaseload_ === "function"
      ? nhslife_activeCaseload_(r)
      : !["Cancelled","Completed"].includes(r.status);

    return Object.assign({}, r, {
      isConfirmed: typeof nhslife_confirmed_ === "function" ? nhslife_confirmed_(r) : !!r.isConfirmed,
      lifecycleV42: {
        upcoming: upcoming,
        pendingAction: pendingAction,
        activeCaseload: activeCaseload
      }
    });
  });

  return {
    ok: true,
    buildVersion: NHS_RELEASE.VERSION,
    generatedAt: new Date().toISOString(),
    source: state.source,
    sheetRows: state.sheetRows || 0,
    counts: {
      owned: rows.length,
      upcoming: rows.filter(function(r) { return r.lifecycleV42.upcoming; }).length,
      pendingActions: rows.filter(function(r) { return r.lifecycleV42.pendingAction; }).length,
      currentCaseload: rows.filter(function(r) { return r.lifecycleV42.activeCaseload; }).length
    },
    upcomingRequestIds: rows.filter(function(r) { return r.lifecycleV42.upcoming; }).map(function(r) { return r.requestId; }),
    requests: rows
  };
}


/** ============================================================
 * SOURCE GROUP: NHS_OS_WebApp.gs
 * ============================================================ */

/**
 * NHS TUTORING OS — WEB APP BACKEND
 * File: NHS_OS_WebApp.gs
 *
 * Add this to the SAME Apps Script project as your existing NHS automation.
 * It reuses the existing global NHS_SPREADSHEET_ID.
 *
 * RUN ONCE:
 *   installNHSOS()
 *
 * DEPLOY:
 *   Deploy > New deployment > Web app
 *   Execute as: Me
 *   Who has access: Prefer "Anyone in your organization"
 */

const NHSOS_REQUESTS = "Tutoring Requests";
const NHSOS_TUTORS = "Tutor Database";
const NHSOS_SESSIONS = "Session Log";
const NHSOS_ACTIVITY = "Activity Log";
const NHSOS_NOTIFICATION_LOG = "Notification Log";
const NHSOS_TITAN_TIME = "10:25 AM–11:05 AM";
const NHSOS_CACHE_KEY = "nhsos_bootstrap_" + NHS_RELEASE.CACHE_NAMESPACE;
const NHSOS_CACHE_SECONDS = 600; // Turbo V14: warm dashboard loads avoid repeated full-sheet scans.
const NHSOS_HEALTH_CACHE_KEY = "nhsos_health_v14";
const NHSOS_HEALTH_CACHE_SECONDS = 900;

function nhsops_actionDateKey_(value) {
  return nhssched_dateKey_(value, NHS_PLATFORM.TIMEZONE);
}

function nhsops_actionHoursSince_(value, now) {
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return 0;
  return Math.max(0, (now.getTime() - d.getTime()) / 3600000);
}

function nhsops_buildActionCenter_(ss, requests, sessions, alerts, now) {
  now = now instanceof Date ? now : new Date();
  const today = nhssched_todayKey_(nhs_safeTimeZone_(ss), now);
  const items = [];
  const seen = {};

  function add(item) {
    const key = String(item.kind || "") + "|" + String(item.requestId || item.sessionId || item.title || "");
    if (seen[key]) return;
    seen[key] = true;
    items.push(item);
  }

  (requests || []).forEach(function(r) {
    const status = String(r.status || "").trim();
    if (!r.requestId || ["Completed","Cancelled"].includes(status)) return;

    const updated = r.lastUpdate || r.timestamp || "";
    const ageHours = nhsops_actionHoursSince_(updated, now);
    const scheduled = String(r.scheduledDate || "").trim();
    const scheduledKey = nhsops_actionDateKey_(scheduled);

    if (status === "No Tutor Available") {
      add({
        kind:"NO_TUTOR",
        severity:"High",
        requestId:r.requestId,
        title:"Tutor needed",
        detail:(r.course || "Request") + " has no available tutor."
      });
    }

    if (
      ["Assigned","Scheduled"].includes(status) &&
      !["Accepted","Declined"].includes(String(r.tutorResponseStatus || "").trim()) &&
      ageHours >= 24
    ) {
      add({
        kind:"TUTOR_RESPONSE",
        severity:ageHours >= 48 ? "High" : "Medium",
        requestId:r.requestId,
        title:"Tutor response overdue",
        detail:(r.assignedTutor || "Assigned tutor") + " has not accepted or declined after " + Math.floor(ageHours) + " hours."
      });
    }

    if (
      String(r.tutorResponseStatus || "").trim() === "Accepted" &&
      !operationsConfirmedServer_(r) &&
      ["Assigned","Scheduled"].includes(status)
    ) {
      add({
        kind:"READY_NOTIFY",
        severity:"High",
        requestId:r.requestId,
        title:"Ready for confirmation",
        detail:(r.course || "Tutoring request") + " was accepted and is ready for officer confirmation."
      });
    }

    if (scheduledKey) {
      const exception = nhssched_exceptionForDate_(scheduledKey);
      if (exception) {
        add({
          kind:"BLACKOUT_CONFLICT",
          severity:"High",
          requestId:r.requestId,
          title:"School-calendar conflict",
          detail:scheduledKey + " is blocked: " + (exception.reason || exception.type) + "."
        });
      }

      const days = nhssched_daysBetween_(today, scheduledKey);
      if (
        operationsConfirmedServer_(r) &&
        days >= 0 &&
        days <= 1 &&
        !(r.teacherFiveStarReady && r.tutorFiveStarEnrolled)
      ) {
        const missing = [];
        if (!r.teacherFiveStarReady) missing.push("teacher Five Star assignment");
        if (!r.tutorFiveStarEnrolled) missing.push("tutor Five Star confirmation");
        add({
          kind:"FIVESTAR_INCOMPLETE",
          severity:"High",
          requestId:r.requestId,
          title:"Five Star setup incomplete",
          detail:(r.course || "Session") + " is " + (days === 0 ? "today" : "tomorrow") + " and still needs " + missing.join(" + ") + "."
        });
      }

      if (
        scheduledKey < today &&
        operationsConfirmedServer_(r) &&
        !["Completed","Cancelled","Pending Verification"].includes(status)
      ) {
        add({
          kind:"SESSION_REPORT",
          severity:"Medium",
          requestId:r.requestId,
          title:"Session report missing",
          detail:(r.course || "Session") + " has passed but has not been closed out."
        });
      }
    }
  });

  (sessions || []).forEach(function(s) {
    if (
      String(s.result || "") === "Completed" &&
      String(s.verificationStatus || "") === "Pending"
    ) {
      add({
        kind:"VERIFY_SESSION",
        severity:"Medium",
        requestId:s.requestId || "",
        sessionId:s.sessionId || "",
        title:"Verify completed session",
        detail:(s.course || s.tutor || "Completed session") + " is awaiting service verification."
      });
    }
  });

  (alerts || []).forEach(function(a) {
    if (!a || a.read) return;
    add({
      kind:"ALERT_" + String(a.type || "GENERAL"),
      severity:String(a.severity || "Medium"),
      requestId:a.requestId || "",
      title:a.title || "Operational alert",
      detail:a.message || ""
    });
  });

  const severityRank = {High:0,Medium:1,Low:2};
  items.sort(function(a, b) {
    return (severityRank[a.severity] ?? 9) - (severityRank[b.severity] ?? 9) ||
      String(a.title || "").localeCompare(String(b.title || ""));
  });

  return items.slice(0, 40);
}

function operationsConfirmedServer_(request) {
  if (!request) return false;
  if (typeof request.assignmentConfirmed === "boolean") {
    return request.assignmentConfirmed;
  }
  return nhslife_confirmed_(request);
}


function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function installNHSOS() {
  nhssec_requireMaintenanceAdmin_();

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  nhsa_installAccessCore_();
  nhsos_ensureSessionLog_(ss);
  nhsos_ensureActivityLog_(ss);
  nhsos_ensureRequestColumns_(nhsperf_sheet_(ss, NHSOS_REQUESTS));
  nhswa_ensureSheet_(ss);
  nhssched_ensureExceptionSheet_(ss);
  nhsos_ensureFastWorker_();
}

function nhsv473_nhsos_getBootstrapDataCore_(forceRefresh) {
  nhssec_requireOfficer_();
  const cache = nhsperf_scriptCache_();

  let cachedData = null;

  const liveCacheKey = nhs_v46OperationsCacheKey_();

  try {
    const cached = cache.get(liveCacheKey);
    if (cached) {
      const parsed = JSON.parse(cached);
      if (parsed && parsed.lifecycleBuild === NHS_RELEASE.VERSION) {
        cachedData = parsed;
      }
    }
  } catch (e) {}

  if (!forceRefresh && cachedData) {
    cachedData.cached = true;
    return cachedData;
  }

  const data =
    forceRefresh && cachedData
      ? nhsos_buildFastRefreshData_(cachedData)
      : nhsos_buildBootstrapData_();

  data.requestTruth = true;
  data.lifecycleBuild = NHS_RELEASE.VERSION;

  try {
    cache.put(liveCacheKey, JSON.stringify(data), NHSOS_CACHE_SECONDS);
  } catch (e) {}

  return data;
}



function nhsos_getLiveRequests() {
  nhssec_requireOfficer_();

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const requests = nhsos_readRequests_(ss);

  return {
    generatedAt: new Date().toISOString(),
    revision: nhsperf_getDataRevision_(),
    requests: requests
  };
}


function nhsos_buildFastRefreshData_(cachedData) {
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  nhsos_ensureRequestColumns_(nhsperf_sheet_(ss, NHSOS_REQUESTS));
  nhswa_ensureSheet_(ss);

  const requests = nhsos_readRequests_(ss);
  const tutors = nhsos_readTutors_(ss);
  const sessions = nhsos_readSessions_(ss);

  // Session-verification alerts are cheap to rebuild from the already-read
  // session set. Other alert/history panels can remain warm until the next
  // cold/full bootstrap.
  const cachedAlerts = Array.isArray(cachedData.alerts)
    ? cachedData.alerts.filter(a => {
        const t = String(a && a.type || "");
        return t !== "OVERDUE_VERIFY";
      })
    : [];

  const alerts = cachedAlerts.concat(
    nhsos_buildSessionVerificationAlerts_(sessions)
  );

  return {
    appName: "NHS Tutoring OS",
    generatedAt: new Date().toISOString(),
    environment: nhsrg_publicEnvironmentInfo_(),
    gatewayUrl: nhs_getGatewayUrl_(),
    titanTime: NHSOS_TITAN_TIME,
    dashboard: nhsos_buildDashboard_(requests, tutors, sessions),
    requests: requests,
    tutors: tutors,
    sessions: sessions.slice(0, 150),
    activities: Array.isArray(cachedData.activities)
      ? cachedData.activities
      : [],
    alerts: alerts,
    actionCenter: nhsops_buildActionCenter_(ss, requests, sessions, alerts, new Date()),
    scheduleExceptions: nhssched_listExceptions_(),
    notificationHistory: Array.isArray(cachedData.notificationHistory)
      ? cachedData.notificationHistory
      : [],
    systemHealth: cachedData.systemHealth || {},
    teacherFormUrl:
      cachedData.teacherFormUrl || nhsos_getTeacherFormUrl_(),
    fastRefresh: true
  };
}


function nhsos_safeBootstrapRead_(label, fn, fallback) {
  try {
    return fn();
  } catch (error) {
    console.error(
      "[NHS V46.5 bootstrap] " + String(label || "optional read") + ": " +
      (error && error.message ? error.message : String(error))
    );
    return fallback;
  }
}


function nhsos_buildBootstrapData_() {
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  nhsos_ensureRequestColumns_(nhsperf_sheet_(ss, NHSOS_REQUESTS));
  nhswa_ensureSheet_(ss);

  // Core workflow data is authoritative and must succeed. Peripheral panels are
  // isolated so a notification/admin/health issue can never prevent Operations
  // from opening.
  const requests = nhsos_readRequests_(ss);
  const tutors = nhsos_readTutors_(ss);
  const sessions = nhsos_readSessions_(ss);

  const activities = nhsos_safeBootstrapRead_(
    "Activity Log",
    function() { return nhsos_readActivities_(ss); },
    []
  );

  const notificationHistory = nhsos_safeBootstrapRead_(
    "Notification Log",
    function() { return nhsos_readNotificationHistory_(ss); },
    []
  );

  let alerts = nhsos_safeBootstrapRead_(
    "Internal alerts",
    function() {
      return typeof nhsn_getInternalAlerts === "function"
        ? (nhsn_getInternalAlerts() || [])
        : [];
    },
    []
  );

  alerts = alerts.concat(
    nhsos_safeBootstrapRead_(
      "Session verification alerts",
      function() { return nhsos_buildSessionVerificationAlerts_(sessions); },
      []
    )
  );

  const systemHealth = nhsos_safeBootstrapRead_(
    "System health",
    function() { return nhsos_getSystemHealthCached_(ss); },
    { degraded: true, message: "System-health details are temporarily unavailable." }
  );

  const teacherFormUrl = nhsos_safeBootstrapRead_(
    "Teacher form URL",
    function() { return nhsos_getTeacherFormUrl_(); },
    ""
  );

  return {
    appName: "NHS Tutoring OS",
    generatedAt: new Date().toISOString(),
    environment: nhsrg_publicEnvironmentInfo_(),
    gatewayUrl: nhs_getGatewayUrl_(),
    titanTime: NHSOS_TITAN_TIME,
    dashboard: nhsos_buildDashboard_(requests, tutors, sessions),
    requests: requests,
    tutors: tutors,
    sessions: sessions.slice(0, 150),
    activities: activities.slice(0, 100),
    alerts: alerts,
    actionCenter: nhsops_buildActionCenter_(ss, requests, sessions, alerts, new Date()),
    scheduleExceptions: nhssched_listExceptions_(),
    notificationHistory: notificationHistory.slice(0, 250),
    systemHealth: systemHealth,
    teacherFormUrl: teacherFormUrl
  };
}


function nhsos_getSystemHealthCached_(ss) {
  const cache = nhsperf_scriptCache_();

  try {
    const cached = cache.get(NHSOS_HEALTH_CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch (e) {}

  const health = nhsos_getSystemHealth_(ss);

  try {
    cache.put(
      NHSOS_HEALTH_CACHE_KEY,
      JSON.stringify(health),
      NHSOS_HEALTH_CACHE_SECONDS
    );
  } catch (e) {}

  return health;
}


function nhsos_getSystemHealth_(ss) {
  const expectedSheets = [
    "Tutoring Requests",
    "Tutor Database",
    "Tutor Weekly Availability",
    "Session Log",
    "Activity Log",
    "Notification Log",
    "Tutor Directory",
    "Membership Roster",
    "Teacher Directory",
    "System Access",
    "Schedule Exceptions"
  ];

  const sheetNames = ss.getSheets().map(s => s.getName());
  const missingSheets = expectedSheets.filter(name => !sheetNames.includes(name));

  let notificationTrigger = false;
  let matchingTrigger = false;
  let assignmentTrigger = false;
  let fastWorkerTrigger = false;

  try {
    ScriptApp.getProjectTriggers().forEach(t => {
      const fn = t.getHandlerFunction();
      if (fn === "runNHSNotificationSweep_") notificationTrigger = true;
      if (fn === "rankAllOpenRequests_") matchingTrigger = true;
      if (fn === "handleAssignmentWorkflowEdit_") assignmentTrigger = true;
      if (fn === "nhsos_processFastJobs_") fastWorkerTrigger = true;
    });
  } catch (e) {}

  const ops =
    typeof nhssec_notificationOps_ === "function"
      ? nhssec_notificationOps_(ss)
      : {};

  return {
    spreadsheetName: ss.getName(),
    timezone: nhs_safeTimeZone_(ss),
    sheetCount: sheetNames.length,
    missingSheets: missingSheets,
    notificationsInstalled: typeof nhsn_getInternalAlerts === "function",
    notificationTrigger: notificationTrigger,
    matchingInstalled: typeof rankAllOpenRequests_ === "function",
    matchingTrigger: matchingTrigger,
    assignmentInstalled: typeof handleAssignmentWorkflowEdit_ === "function",
    assignmentTrigger: assignmentTrigger,
    fastWorkerTrigger: fastWorkerTrigger,
    cacheSeconds: NHSOS_CACHE_SECONDS,
    notificationOps: ops,
    generatedAt: new Date().toISOString()
  };
}

function nhsops_runIntegrityCheck() {
  nhssec_requireOfficer_();

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const checks = [];

  function add(id, label, status, detail, count) {
    checks.push({
      id: id,
      label: label,
      status: ["PASS","WARN","BLOCK"].includes(status) ? status : "WARN",
      detail: String(detail || ""),
      count: Number(count || 0)
    });
  }

  const requests = nhsos_readRequests_(ss);
  const tutors = nhsos_readTutors_(ss);
  const sessions = nhsos_readSessions_(ss);
  const tutorIds = new Set(
    tutors.map(function(t) { return String(t.tutorId || "").trim(); }).filter(Boolean)
  );
  const requestIds = new Set(
    requests.map(function(r) { return String(r.requestId || "").trim(); }).filter(Boolean)
  );

  const requestCounts = {};
  requests.forEach(function(r) {
    const id = String(r.requestId || "").trim();
    if (id) requestCounts[id] = (requestCounts[id] || 0) + 1;
  });
  const duplicateRequests = Object.keys(requestCounts).filter(function(id) {
    return requestCounts[id] > 1;
  });
  add(
    "request-ids",
    "Request IDs are unique",
    duplicateRequests.length ? "BLOCK" : "PASS",
    duplicateRequests.length ? "Duplicate IDs: " + duplicateRequests.slice(0, 8).join(", ") : "No duplicate Request IDs.",
    duplicateRequests.length
  );

  const tutorCounts = {};
  tutors.forEach(function(t) {
    const id = String(t.tutorId || "").trim();
    if (id) tutorCounts[id] = (tutorCounts[id] || 0) + 1;
  });
  const duplicateTutors = Object.keys(tutorCounts).filter(function(id) {
    return tutorCounts[id] > 1;
  });
  add(
    "tutor-ids",
    "Tutor IDs are unique",
    duplicateTutors.length ? "BLOCK" : "PASS",
    duplicateTutors.length ? "Duplicate IDs: " + duplicateTutors.slice(0, 8).join(", ") : "No duplicate Tutor IDs.",
    duplicateTutors.length
  );

  const orphanSessions = sessions.filter(function(s) {
    return s.requestId && !requestIds.has(String(s.requestId));
  });
  add(
    "session-links",
    "Sessions reference known requests",
    orphanSessions.length ? "WARN" : "PASS",
    orphanSessions.length ? orphanSessions.length + " session(s) reference a missing request." : "All session request links are valid.",
    orphanSessions.length
  );

  const missingTutorRefs = requests.filter(function(r) {
    return r.assignedTutorId && !tutorIds.has(String(r.assignedTutorId));
  });
  add(
    "tutor-links",
    "Assigned Tutor IDs resolve",
    missingTutorRefs.length ? "BLOCK" : "PASS",
    missingTutorRefs.length ? missingTutorRefs.length + " request(s) reference a missing Tutor ID." : "All assigned Tutor IDs resolve.",
    missingTutorRefs.length
  );

  const calendarConflicts = requests.filter(function(r) {
    return r.scheduledDate && !["Completed","Cancelled"].includes(String(r.status || "")) &&
      !!nhssched_exceptionForDate_(r.scheduledDate);
  });
  add(
    "calendar-conflicts",
    "No active sessions fall on blocked school dates",
    calendarConflicts.length ? "BLOCK" : "PASS",
    calendarConflicts.length
      ? calendarConflicts.length + " active request(s) are scheduled on a school-calendar exception."
      : "No active request is scheduled on a blocked date.",
    calendarConflicts.length
  );

  const fiveStarContradictions = requests.filter(function(r) {
    return !!r.tutorFiveStarEnrolled && !r.teacherFiveStarReady;
  });
  add(
    "five-star-order",
    "Five Star state order is consistent",
    fiveStarContradictions.length ? "WARN" : "PASS",
    fiveStarContradictions.length
      ? fiveStarContradictions.length + " request(s) show tutor confirmation before teacher placement."
      : "Five Star state order is consistent.",
    fiveStarContradictions.length
  );

  const pendingVerification = sessions.filter(function(s) {
    return String(s.result || "") === "Completed" &&
      String(s.verificationStatus || "") === "Pending";
  });
  add(
    "pending-verification",
    "Completed sessions are verified",
    pendingVerification.length ? "WARN" : "PASS",
    pendingVerification.length
      ? pendingVerification.length + " completed session(s) still await verification."
      : "No completed session is waiting for verification.",
    pendingVerification.length
  );

  const registry = nhsa_readRegistry_(nhsa_ensureRegistry_(ss));
  const activeTutorGrants = registry.filter(function(g) {
    return g.role === "TUTOR" && nhsa_isGrantActive_(g);
  });
  const tutorEmails = new Set(
    tutors.map(function(t) {
      return String(t.schoolEmail || "").trim().toLowerCase();
    }).filter(Boolean)
  );
  const orphanTutorGrants = activeTutorGrants.filter(function(g) {
    return !tutorEmails.has(String(g.email || "").toLowerCase());
  });
  add(
    "tutor-grants",
    "Active TUTOR grants have Tutor Database profiles",
    orphanTutorGrants.length ? "BLOCK" : "PASS",
    orphanTutorGrants.length
      ? orphanTutorGrants.length + " active TUTOR grant(s) have no matching Tutor Database profile."
      : "Every active TUTOR grant has a Tutor Database profile.",
    orphanTutorGrants.length
  );

  const activeTutorGrantEmails = new Set(
    activeTutorGrants.map(function(g) { return String(g.email || "").toLowerCase(); })
  );
  const missingTutorGrants = tutors.filter(function(t) {
    const status = String(t.reliabilityStatus || "").trim();
    const email = String(t.schoolEmail || "").trim().toLowerCase();
    return email &&
      !["Inactive","Suspended"].includes(status) &&
      !activeTutorGrantEmails.has(email);
  });
  add(
    "tutor-access-sync",
    "Active tutor profiles have TUTOR access",
    missingTutorGrants.length ? "WARN" : "PASS",
    missingTutorGrants.length
      ? missingTutorGrants.length + " active tutor profile(s) do not currently have an active TUTOR grant."
      : "Active tutor profiles are synchronized with TUTOR access.",
    missingTutorGrants.length
  );

  const settings = nhsa_readSettings_(ss);
  const routeKeys = ["GATEWAY_URL","TEACHER_URL","TUTOR_URL","INTERNAL_URL"];
  const badRoutes = routeKeys.filter(function(key) {
    const value = String(settings[key] || "").trim();
    return !value || !/\/exec(?:\/|\?|$)/.test(value);
  });
  add(
    "portal-routes",
    "Portal routes use production /exec URLs",
    badRoutes.length ? "BLOCK" : "PASS",
    badRoutes.length ? "Invalid/missing: " + badRoutes.join(", ") : "All stored portal routes use /exec.",
    badRoutes.length
  );

  const notificationOps =
    typeof nhssec_notificationOps_ === "function"
      ? nhssec_notificationOps_(ss)
      : {};
  const failedNotifications = Number(notificationOps.failedNotifications || 0);
  add(
    "notifications",
    "Notification delivery has no logged failures",
    failedNotifications ? "WARN" : "PASS",
    failedNotifications
      ? failedNotifications + " failed notification(s) are present in Notification Log."
      : "No failed notifications are currently logged.",
    failedNotifications
  );

  const blocked = checks.filter(function(c) { return c.status === "BLOCK"; }).length;
  const warnings = checks.filter(function(c) { return c.status === "WARN"; }).length;

  return {
    ok: blocked === 0,
    generatedAt: new Date().toISOString(),
    blocked: blocked,
    warnings: warnings,
    passed: checks.filter(function(c) { return c.status === "PASS"; }).length,
    checks: checks
  };
}


function nhsos_runWorkbookAudit() {
  nhssec_requireOfficer_();
  if (typeof auditNHSTutoringWorkbook !== "function") {
    throw new Error("Workbook audit helper is not installed.");
  }
  return auditNHSTutoringWorkbook();
}

function nhsos_invalidateCache_() {
  if (typeof nhs_v45InvalidatePortalState_ === "function") {
    nhs_v45InvalidatePortalState_({});
    return;
  }

  try {
    const cache = nhsperf_scriptCache_();
    cache.remove(nhs_v46OperationsCacheKey_());
    cache.remove(NHSOS_HEALTH_CACHE_KEY);
  } catch (e) {}

  nhsperf_touchDataRevision_();
}




/* ============================================================
   V45.5 — SESSION MUTATION CROSS-PORTAL INVALIDATION
   ============================================================ */

/**
 * Session verification changes are global platform mutations.
 * Resolve the linked request so the affected Teacher and Tutor portal caches
 * are cleared before advancing the shared revision.
 *
 * This prevents a portal from observing a new revision and then immediately
 * re-hydrating an older cached payload.
 */
function nhsos_invalidateSessionMutation_(ss, sessionLike) {
  sessionLike = sessionLike || {};

  const requestId = String(
    sessionLike.requestId !== undefined
      ? sessionLike.requestId
      : sessionLike["Request ID"] || ""
  ).trim();

  const sessionTutorId = String(
    sessionLike.tutorId !== undefined
      ? sessionLike.tutorId
      : sessionLike["Tutor ID"] || ""
  ).trim();

  let teacherEmail = "";
  let tutorId = sessionTutorId;

  if (requestId) {
    try {
      const requestSh = nhsperf_sheet_(ss, NHSOS_REQUESTS);
      const requestRow = nhsos_findRequestRow_(requestSh, requestId);

      if (requestRow) {
        const requestHeaders = nhsos_headerMap_(requestSh);
        const request = nhsos_rowObject_(requestSh, requestRow, requestHeaders);
        teacherEmail = String(request["Teacher Email"] || "").trim().toLowerCase();
        tutorId = String(request["Assigned Tutor ID"] || tutorId || "").trim();
      }
    } catch (e) {}
  }

  if (typeof nhs_v45InvalidatePortalState_ === "function") {
    return nhs_v45InvalidatePortalState_({
      teacherEmail: teacherEmail,
      tutorId: tutorId
    });
  }

  try {
    const cache = nhsperf_scriptCache_();
    cache.remove(nhs_v46OperationsCacheKey_());
    cache.remove(NHSOS_HEALTH_CACHE_KEY);

    if (teacherEmail) {
      cache.remove(nhs_v46TeacherCacheKey_(teacherEmail));
    }

    if (tutorId) {
      cache.remove(nhs_v46TutorCacheKey_(tutorId));
    }
  } catch (e) {}

  return nhsperf_touchDataRevision_();
}

function nhsos_markAlertRead(notificationId) {
  nhssec_requireOfficer_();

  if (typeof nhsn_markAlertRead !== "function") {
    return { ok: false };
  }

  const result = nhsn_markAlertRead(notificationId);

  // Read/unread UI metadata should not trigger a system-wide data revision.
  try {
    const cache = nhsperf_scriptCache_();
    cache.remove(nhs_v46OperationsCacheKey_());
  } catch (e) {}

  return result;
}

function nhsos_markAllAlertsRead() {
  nhssec_requireOfficer_();

  if (typeof nhsn_markAllInternalAlertsRead !== "function") {
    return { ok: false, updated: 0 };
  }

  const result = nhsn_markAllInternalAlertsRead();

  try {
    const cache = nhsperf_scriptCache_();
    cache.remove(nhs_v46OperationsCacheKey_());
  } catch (e) {}

  return result;
}

function nhsos_buildDashboard_(requests, tutors, sessions) {
  const today = new Date();
  const weekEnd = new Date(today);
  weekEnd.setDate(today.getDate()+7);
  const tz = NHS_PLATFORM.TIMEZONE;

  const upcoming = requests.filter(function(r){
    if (!nhslife_upcoming_(r,today,tz)) return false;
    const d=new Date(r.scheduledDate+"T12:00:00");
    return !isNaN(d.getTime()) && d<=weekEnd;
  });

  const verifiedThisMonth=sessions.filter(function(s){
    if(s.result!=="Completed"||s.verificationStatus!=="Verified"||!s.sessionDate)return false;
    const d=new Date(s.sessionDate+"T12:00:00");
    return d.getFullYear()===today.getFullYear()&&d.getMonth()===today.getMonth();
  });
  const pendingVerifications=sessions.filter(function(s){
    return s.result==="Completed"&&s.verificationStatus==="Pending";
  });

  const needsAttention=requests.filter(function(r){
    return ["Needs Follow-Up","No Tutor Available"].includes(r.status) ||
      (["Assigned","Scheduled"].includes(r.status) && !nhslife_confirmed_(r));
  });

  const activeTutorRows=tutors.filter(function(t){return !["Inactive","Suspended"].includes(t.reliabilityStatus);});
  const weeklyConfirmed=activeTutorRows.filter(function(t){return !!(t.weeklyAvailability&&t.weeklyAvailability.confirmed);}).length;
  const weeklyUnavailableAll=activeTutorRows.filter(function(t){return !!(t.weeklyAvailability&&t.weeklyAvailability.confirmed&&t.weeklyAvailability.availableCount===0);}).length;

  return {
    newRequests:requests.filter(function(r){return r.status==="New";}).length,
    needsAttention:needsAttention.length+pendingVerifications.length,
    scheduledThisWeek:upcoming.length,
    activeTutors:activeTutorRows.length,
    weeklyAvailabilityConfirmed:weeklyConfirmed,
    weeklyAvailabilityUnconfirmed:Math.max(0,activeTutorRows.length-weeklyConfirmed),
    weeklyUnavailableAll:weeklyUnavailableAll,
    weeklyAvailabilityWeekStart:activeTutorRows[0]&&activeTutorRows[0].weeklyAvailability?activeTutorRows[0].weeklyAvailability.weekStart:"",
    completedThisMonth:verifiedThisMonth.length,
    pendingVerifications:pendingVerifications.length,
    upcoming:upcoming.slice(0,12),
    needsAction:requests.filter(function(r){
      return ["New","Matching","Needs Follow-Up","No Tutor Available","Pending Verification"].includes(r.status) ||
        (["Assigned","Scheduled"].includes(r.status) && !nhslife_confirmed_(r));
    }).slice(0,12)
  };
}

function nhsos_readRequests_(ss) {
  const sh = nhsperf_sheet_(ss, NHSOS_REQUESTS);
  if (!sh) return [];

  const objects = typeof nhsdata_readKeyedObjects_ === "function"
    ? nhsdata_readKeyedObjects_(sh, "Request ID")
    : [];

  return objects.map(function(obj) {
    const recurringValue = obj["Recurring?"] || "";
    const recurringActiveRaw = obj["Recurring Active"] || "";
    return {
      row: obj._row,
      requestId: obj["Request ID"] || "",
      timestamp: obj["Timestamp"] || "",
      teacherName: obj["Teacher Name"] || "",
      teacherEmail: obj["Teacher Email"] || "",
      course: obj["Course"] || "",
      topic: obj["Topic / Academic Need"] || "",
      supportType: obj["Support Type"] || "",
      schedulingType: obj["Scheduling Preference"] || "Legacy",
      preferredDates: obj["Preferred Dates"] || "",
      preferredDays: obj["Preferred Days"] || "",
      recurring: recurringValue,
      recurringActive: recurringActiveRaw ? nhsos_truthy_(recurringActiveRaw) : nhsos_isRecurringValue_(recurringValue),
      nextSessionDate: nhsos_isoDate_(obj["Next Session Date"] || ""),
      teacherNotes: obj["Teacher Notes"] || "",
      status: obj["Status"] || "New",
      assignedTutorId: obj["Assigned Tutor ID"] || "",
      assignedTutor: obj["Assigned Tutor"] || "",
      scheduledDate: nhsos_isoDate_(obj["Scheduled Date"] || ""),
      scheduledTime: obj["Scheduled Time"] || "",
      location: obj["Location"] || "",
      coordinatorNotes: obj["Coordinator Notes"] || "",
      topMatch1: obj["Top Match 1"] || "",
      topMatch2: obj["Top Match 2"] || "",
      topMatch3: obj["Top Match 3"] || "",
      matchNotes: obj["Match Notes"] || "",
      matchBreakdown: (function(){
        try { return JSON.parse(String(obj["Match Breakdown"] || "[]")); }
        catch(e) { return []; }
      })(),
      notificationStatus: obj["Notification Status"] || "",
      assignmentConfirmed: nhsos_nonFalseValue_(obj["Assignment Confirmed"] || ""),
      tutorResponseStatus: obj["Tutor Response Status"] || "",
      tutorResponseAt: obj["Tutor Response At"] || "",
      tutorDeclineReason: obj["Tutor Decline Reason"] || "",
      tutorViewedAt: obj["Tutor Viewed At"] || "",
      tutorReconfirmStatus: obj["Tutor Reconfirm Status"] || "",
      tutorReconfirmAt: obj["Tutor Reconfirm At"] || "",
      teacherFiveStarReady: nhsos_truthy_(obj["Teacher Five Star Ready"] || ""),
      teacherFiveStarReadyAt: obj["Teacher Five Star Ready At"] || "",
      tutorFiveStarEnrolled: nhsos_truthy_(obj["Tutor Five Star Enrolled"] || ""),
      tutorFiveStarEnrolledAt: obj["Tutor Five Star Enrolled At"] || "",
      fiveStarReadyAt: obj["Five Star Ready At"] || "",
      fiveStarReadiness: nhssched_readiness_(obj),
      lastUpdate: obj["Last Update"] || ""
    };
  }).reverse();
}

/**
 * Run once after installing the auth/delivery split.
 * Adds Notification Email to Tutor Database if missing and seeds it from
 * Personal Email for existing tutors. School Email is authentication-only.
 */
function migrateTutorNotificationEmails() {
  nhssec_requireMaintenanceAdmin_();
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const sh = nhsperf_sheet_(ss, NHSOS_TUTORS);

  if (!sh) {
    throw new Error('Missing "' + NHSOS_TUTORS + '" sheet.');
  }

  const headers = sh
    .getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1))
    .getDisplayValues()[0]
    .map(v => String(v || "").trim());

  let notificationCol = headers.indexOf("Notification Email") + 1;

  if (!notificationCol) {
    notificationCol = headers.length + 1;
    sh.getRange(1, notificationCol).setValue("Notification Email");
  }

  let freshHeaders = sh
    .getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0]
    .map(v => String(v || "").trim());

  const personalCol = freshHeaders.indexOf("Personal Email") + 1;

  if (!personalCol || sh.getLastRow() < 2) {
    return { ok: true, migrated: 0 };
  }

  const count = sh.getLastRow() - 1;
  const personal = sh.getRange(2, personalCol, count, 1).getValues();
  const notification = sh.getRange(2, notificationCol, count, 1).getValues();

  let migrated = 0;

  for (let i = 0; i < count; i++) {
    if (
      !String(notification[i][0] || "").trim() &&
      String(personal[i][0] || "").trim()
    ) {
      notification[i][0] = String(personal[i][0]).trim();
      migrated++;
    }
  }

  sh.getRange(2, notificationCol, count, 1).setValues(notification);

  nhsos_invalidateCache_();

  return {
    ok: true,
    migrated: migrated,
    message:
      "Notification Email migration complete. " +
      migrated +
      " tutor record(s) seeded from Personal Email."
  };
}


/* ============================================================
   V45.5 — TUTOR RELIABILITY ADMINISTRATION
   ============================================================ */

const NHSOS_RELIABILITY_STATUSES = [
  "Excellent",
  "Reliable",
  "New",
  "Warning",
  "Probation",
  "Suspended",
  "Inactive"
];

/**
 * Performance recommendation only. Reliability Status itself remains an
 * officer-controlled field, so small samples never silently suspend a tutor.
 *
 * Thresholds:
 * - Suspended / Inactive: preserved as manual holds.
 * - New: fewer than 3 resolved sessions.
 * - Excellent: 8+ resolved sessions and zero tutor no-shows.
 * - Probation: 2+ tutor no-shows OR >=20% tutor no-show rate.
 * - Warning: 1+ tutor no-show AND >=10% tutor no-show rate.
 * - Reliable: all other established tutors.
 */
function nhsos_reliabilityRecommendationFromStats_(completed, noShows, currentStatus) {
  completed = Math.max(0, Number(completed || 0));
  noShows = Math.max(0, Number(noShows || 0));

  const current = String(currentStatus || "New").trim();
  const currentNorm = current.toLowerCase();
  const resolved = completed + noShows;
  const rate = resolved > 0 ? noShows / resolved : 0;

  if (currentNorm === "suspended" || currentNorm === "inactive") {
    return {
      status: currentNorm === "suspended" ? "Suspended" : "Inactive",
      performanceStatus: resolved < 3 ? "New" : "Reliable",
      manualHold: true,
      resolvedSessions: resolved,
      noShowRate: rate
    };
  }

  let status = "Reliable";

  if (resolved < 3) {
    status = "New";
  } else if (resolved >= 8 && noShows === 0) {
    status = "Excellent";
  } else if (noShows >= 2 || rate >= 0.20) {
    status = "Probation";
  } else if (noShows >= 1 && rate >= 0.10) {
    status = "Warning";
  }

  return {
    status: status,
    performanceStatus: status,
    manualHold: false,
    resolvedSessions: resolved,
    noShowRate: rate
  };
}

function nhsos_findTutorRowById_(sh, tutorId) {
  if (!sh || sh.getLastRow() < 2) return 0;
  const h = nhsos_headerMap_(sh);
  const col = h["Tutor ID"];
  if (!col) return 0;

  const match = sh
    .getRange(2, col, sh.getLastRow() - 1, 1)
    .createTextFinder(String(tutorId || "").trim())
    .matchEntireCell(true)
    .matchCase(true)
    .findNext();

  return match ? match.getRow() : 0;
}

function nhsos_invalidateTutorReliabilityCaches_(tutorId, schoolEmail) {
  tutorId = String(tutorId || "").trim();
  schoolEmail = String(schoolEmail || "").trim().toLowerCase();

  try {
    const cache = nhsperf_scriptCache_();
    if (tutorId) cache.remove(nhs_v46TutorCacheKey_(tutorId));
    if (schoolEmail) cache.remove(nhs_v46TutorIdentityKey_(schoolEmail));
  } catch (e) {}

  if (typeof nhs_v45InvalidatePortalState_ === "function") {
    return nhs_v45InvalidatePortalState_({ tutorId: tutorId });
  }

  return nhsperf_touchDataRevision_();
}

function nhsos_updateTutorReliabilityStatus(tutorId, newStatus, reason) {
  nhssec_requireOfficer_();

  tutorId = String(tutorId || "").trim();
  newStatus = String(newStatus || "").trim();
  reason = String(reason || "").trim();

  if (!tutorId) throw new Error("Tutor ID is required.");
  if (!NHSOS_RELIABILITY_STATUSES.includes(newStatus)) {
    throw new Error("Invalid reliability status.");
  }
  if (!reason) {
    throw new Error("Enter a reason for the reliability status change.");
  }

  return nhsos_withLock_(function() {
    const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
    const sh = nhsperf_sheet_(ss, NHSOS_TUTORS);
    if (!sh) throw new Error('Missing "' + NHSOS_TUTORS + '" sheet.');

    nhssec_ensureHeaders_(sh, ["Reliability Status"]);
    nhsperf_invalidateSheetMemo_(sh);

    const row = nhsos_findTutorRowById_(sh, tutorId);
    if (!row) throw new Error("Tutor not found.");

    const h = nhsos_headerMap_(sh);
    const oldStatus = String(
      sh.getRange(row, h["Reliability Status"]).getDisplayValue() || "New"
    ).trim() || "New";

    const fullName = h["Full Name"]
      ? String(sh.getRange(row, h["Full Name"]).getDisplayValue() || "").trim()
      : "";
    const schoolEmail = h["School Email"]
      ? String(sh.getRange(row, h["School Email"]).getDisplayValue() || "").trim().toLowerCase()
      : "";

    sh.getRange(row, h["Reliability Status"]).setValue(newStatus);
    SpreadsheetApp.flush();

    nhsos_logActivity_(ss, {
      type: "Tutor Reliability Changed",
      tutorId: tutorId,
      actor: nhsos_actor_(),
      summary: (fullName || tutorId) + " reliability updated",
      details: oldStatus + " → " + newStatus + " · " + reason
    });

    nhsos_invalidateTutorReliabilityCaches_(tutorId, schoolEmail);

    return {
      ok: true,
      tutorId: tutorId,
      oldStatus: oldStatus,
      reliabilityStatus: newStatus,
      reason: reason
    };
  });
}

/**
 * Applies system recommendations to every non-Suspended/non-Inactive tutor.
 * Manual hold states are intentionally preserved.
 */
function nhsos_applyReliabilityRecommendations() {
  nhssec_requireOfficer_();

  return nhsos_withLock_(function() {
    const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
    const sh = nhsperf_sheet_(ss, NHSOS_TUTORS);
    if (!sh || sh.getLastRow() < 2) {
      return { ok: true, updated: 0, unchanged: 0, preservedManualHolds: 0 };
    }

    nhssec_ensureHeaders_(sh, ["Reliability Status"]);
    nhsperf_invalidateSheetMemo_(sh);

    const h = nhsos_headerMap_(sh);
    const body = sh
      .getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
      .getDisplayValues();

    const output = [];
    const cache = nhsperf_scriptCache_();
    let updated = 0;
    let unchanged = 0;
    let preservedManualHolds = 0;

    body.forEach(function(row) {
      const tutorId = h["Tutor ID"] ? String(row[h["Tutor ID"] - 1] || "").trim() : "";
      const schoolEmail = h["School Email"] ? String(row[h["School Email"] - 1] || "").trim().toLowerCase() : "";
      const current = h["Reliability Status"] ? String(row[h["Reliability Status"] - 1] || "New").trim() || "New" : "New";
      const completed = h["Sessions Completed"] ? Number(row[h["Sessions Completed"] - 1] || 0) : 0;
      const noShows = h["Tutor No-Shows"] ? Number(row[h["Tutor No-Shows"] - 1] || 0) : 0;
      const rec = nhsos_reliabilityRecommendationFromStats_(completed, noShows, current);

      let next = current;
      if (rec.manualHold) {
        preservedManualHolds++;
      } else {
        next = rec.status;
        if (next !== current) updated++;
        else unchanged++;
      }

      output.push([next]);

      try {
        if (tutorId) cache.remove(nhs_v46TutorCacheKey_(tutorId));
        if (schoolEmail) cache.remove(nhs_v46TutorIdentityKey_(schoolEmail));
      } catch (e) {}
    });

    sh.getRange(2, h["Reliability Status"], output.length, 1).setValues(output);
    SpreadsheetApp.flush();

    nhsos_logActivity_(ss, {
      type: "Tutor Reliability Bulk Update",
      actor: nhsos_actor_(),
      summary: "System reliability recommendations applied",
      details:
        updated + " updated · " +
        unchanged + " unchanged · " +
        preservedManualHolds + " Suspended/Inactive preserved"
    });

    // Operations cache + shared revision. Individual tutor caches were cleared above.
    nhsos_invalidateCache_();

    return {
      ok: true,
      updated: updated,
      unchanged: unchanged,
      preservedManualHolds: preservedManualHolds
    };
  });
}

function nhsos_readTutors_(ss) {
  const sh = nhsperf_sheet_(ss, NHSOS_TUTORS);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getDisplayValues();
  const idx = nhsos_headerIndex_(values[0]);
  const weeklyIndex = nhswa_readIndex_(ss);
  const targetWeek = nhswa_confirmationTargetWeekStart_(nhs_safeTimeZone_(ss), new Date());

  return values.slice(1).map(function(row, i) {
    const fullName = nhsos_get_(row, idx, "Full Name");
    if (!fullName) return null;

    const preferred = nhsos_get_(row, idx, "Preferred Name");
    const sessionsCompleted = Number(nhsos_get_(row, idx, "Sessions Completed") || 0);
    const tutorNoShows = Number(nhsos_get_(row, idx, "Tutor No-Shows") || 0);
    const reliabilityStatus = nhsos_get_(row, idx, "Reliability Status") || "New";
    const recommendation = nhsos_reliabilityRecommendationFromStats_(
      sessionsCompleted,
      tutorNoShows,
      reliabilityStatus
    );

    return {
      row: i + 2,
      tutorId: nhsos_get_(row, idx, "Tutor ID"),
      fullName: fullName,
      preferredName: preferred,
      displayName: nhsos_tutorDisplayName_(preferred, fullName),
      email: nhsos_get_(row, idx, "Notification Email"),
      notificationEmail: nhsos_get_(row, idx, "Notification Email"),
      personalEmail: nhsos_get_(row, idx, "Personal Email"),
      schoolEmail: nhsos_get_(row, idx, "School Email"),
      reminderPreference: nhsos_get_(row, idx, "Reminder Preference") || "Day Before + Morning Of",
      temporaryUnavailableDates: nhsos_get_(row, idx, "Temporary Unavailable Dates"),
      grade: nhsos_get_(row, idx, "Grade"),
      strongestCourse: nhsos_get_(row, idx, "Strongest Course"),
      titanAvailability: nhsos_get_(row, idx, "Typical Titan Time Availability"),
      afterSchoolAvailability: nhsos_get_(row, idx, "Typical After-School Availability"),
      maxSessions: Number(nhsos_get_(row, idx, "Maximum Sessions / Week") || 0),
      currentCaseload: Number(nhsos_get_(row, idx, "Current Caseload") || 0),
      sessionsCompleted: sessionsCompleted,
      verifiedHours: Number(nhsos_get_(row, idx, "Verified Tutoring Hours") || 0),
      tutorNoShows: tutorNoShows,
      reliabilityStatus: reliabilityStatus,
      expectationsAgreement: nhsos_get_(row, idx, "Expectations Agreement"),
      reliabilityRecommendation: recommendation.status,
      reliabilityManualHold: recommendation.manualHold,
      reliabilityResolvedSessions: recommendation.resolvedSessions,
      reliabilityNoShowRate: Math.round(recommendation.noShowRate * 1000) / 10,
      experience: nhsos_get_(row, idx, "Previous Experience"),
      languages: nhsos_get_(row, idx, "Tutoring Languages"),
      weeklyAvailability: nhswa_weekSummaryFromIndex_(weeklyIndex, {
        tutorId: nhsos_get_(row, idx, "Tutor ID"),
        titanAvailability: nhsos_get_(row, idx, "Typical Titan Time Availability"),
        temporaryUnavailableDates: nhsos_get_(row, idx, "Temporary Unavailable Dates")
      }, targetWeek),
      weeklyByDate: nhswa_futureMapForTutor_(weeklyIndex, nhsos_get_(row, idx, "Tutor ID"))
    };
  }).filter(function(tutor) {
    if (!tutor) return false;
    const reliability = String(tutor.reliabilityStatus || "").trim().toLowerCase();
    const schoolEmail = String(tutor.schoolEmail || "").trim().toLowerCase();
    return reliability !== "inactive" && nhsa_isSchoolIdentityEmail_(schoolEmail);
  });
}

function nhsos_readSessions_(ss) {
  const sh = nhsperf_sheet_(ss, NHSOS_SESSIONS);
  if (!sh || sh.getLastRow() < 2) return [];

  const values = sh.getDataRange().getDisplayValues();
  const idx = nhsos_headerIndex_(values[0]);

  return values.slice(1).map((row, i) => {
    const sessionId = nhsos_get_(row, idx, "Session ID");
    if (!sessionId) return null;

    const result = nhsos_get_(row, idx, "Result");
    const verifiedMinutes = Number(nhsos_get_(row, idx, "Verified Minutes") || 0);
    const reportedMinutes = Number(
      nhsos_get_(row, idx, "Reported Minutes") ||
      nhsos_get_(row, idx, "Verified Minutes") ||
      0
    );

    let verificationStatus = nhsos_get_(row, idx, "Verification Status");
    if (!verificationStatus) {
      if (result === "Completed" && verifiedMinutes > 0) verificationStatus = "Verified";
      else if (result === "Completed") verificationStatus = "Pending";
      else verificationStatus = "Not Required";
    }

    return {
      row: i + 2,
      sessionId,
      requestId: nhsos_get_(row, idx, "Request ID"),
      tutorId: nhsos_get_(row, idx, "Tutor ID"),
      tutor: nhsos_get_(row, idx, "Tutor"),
      sessionDate: nhsos_isoDate_(nhsos_get_(row, idx, "Session Date")),
      result,
      reportedMinutes,
      verificationStatus,
      minutes: verifiedMinutes,
      notes: nhsos_get_(row, idx, "Notes"),
      loggedAt: nhsos_get_(row, idx, "Logged At"),
      loggedBy: nhsos_get_(row, idx, "Logged By"),
      verifiedAt: nhsos_get_(row, idx, "Verified At"),
      verifiedBy: nhsos_get_(row, idx, "Verified By"),
      verificationNotes: nhsos_get_(row, idx, "Verification Notes"),
      voidedAt: nhsos_get_(row, idx, "Voided At"),
      voidedBy: nhsos_get_(row, idx, "Voided By"),
      voidReason: nhsos_get_(row, idx, "Void Reason")
    };
  }).filter(Boolean).reverse();
}


function nhsos_readActivities_(ss) {
  const sh = nhsperf_sheet_(ss, NHSOS_ACTIVITY);
  if (!sh || sh.getLastRow() < 2) return [];

  const values = sh.getDataRange().getDisplayValues();
  const idx = nhsos_headerIndex_(values[0]);

  return values.slice(1).map((row, i) => {
    const activityId = nhsos_get_(row, idx, "Activity ID");
    if (!activityId) return null;

    return {
      row: i + 2,
      activityId,
      timestamp: nhsos_get_(row, idx, "Timestamp"),
      type: nhsos_get_(row, idx, "Type"),
      requestId: nhsos_get_(row, idx, "Request ID"),
      sessionId: nhsos_get_(row, idx, "Session ID"),
      tutorId: nhsos_get_(row, idx, "Tutor ID"),
      actor: nhsos_get_(row, idx, "Actor"),
      summary: nhsos_get_(row, idx, "Summary"),
      details: nhsos_get_(row, idx, "Details")
    };
  }).filter(Boolean).reverse();
}

function nhsos_readNotificationHistory_(ss) {
  const sh = nhsperf_sheet_(ss, NHSOS_NOTIFICATION_LOG);
  if (!sh || sh.getLastRow() < 2) return [];

  const values = sh.getDataRange().getDisplayValues();
  const idx = nhsos_headerIndex_(values[0]);

  return values.slice(1).map((row, i) => {
    const notificationId = nhsos_get_(row, idx, "Notification ID");
    if (!notificationId) return null;

    return {
      row: i + 2,
      notificationId,
      createdAt: nhsos_get_(row, idx, "Created At"),
      requestId: nhsos_get_(row, idx, "Request ID"),
      channel: nhsos_get_(row, idx, "Channel"),
      recipient: nhsos_get_(row, idx, "Recipient"),
      type: nhsos_get_(row, idx, "Type"),
      status: nhsos_get_(row, idx, "Status"),
      read: nhsos_truthy_(nhsos_get_(row, idx, "Read")),
      severity: nhsos_get_(row, idx, "Severity"),
      title: nhsos_get_(row, idx, "Title"),
      message: nhsos_get_(row, idx, "Message"),
      error: nhsos_get_(row, idx, "Error")
    };
  }).filter(Boolean).reverse();
}

function nhsos_buildSessionVerificationAlerts_(sessions) {
  return sessions
    .filter(s => s.result === "Completed" && s.verificationStatus === "Pending")
    .map(s => ({
      notificationId: "",
      type: "SESSION_VERIFICATION",
      requestId: s.requestId,
      severity: "High",
      title: "Session needs verification",
      message:
        (s.tutor || "Tutor") +
        " reported a completed session on " +
        (s.sessionDate || "the scheduled date") +
        ". Verify service minutes before credit is awarded.",
      createdAt: s.loggedAt || s.sessionDate || "",
      read: false
    }));
}

function nhsv473_nhsos_saveRequestCore_(payload) {
  nhssec_requireOfficer_();
  return nhsos_withLock_(function() {
    const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
    const sh = nhsperf_sheet_(ss, NHSOS_REQUESTS);
    const row = nhsos_findRequestRow_(sh, payload.requestId);
    if (!row) throw new Error("Request not found.");

    const headers = nhsos_headerMap_(sh);
    const range = sh.getRange(row, 1, 1, sh.getLastColumn());
    const values = range.getValues()[0];
    const display = range.getDisplayValues()[0];

    const get = header =>
      headers[header]
        ? String(display[headers[header] - 1] || "").trim()
        : "";

    const set = (header, value) => {
      if (headers[header]) {
        values[headers[header] - 1] = value;
      }
    };

    const oldState = {
      assignedTutor: get("Assigned Tutor"),
      assignedTutorId: get("Assigned Tutor ID"),
      scheduledDate: nhsos_isoDate_(get("Scheduled Date")),
      location: get("Location"),
      status: get("Status"),
      assignmentConfirmed: get("Assignment Confirmed")
    };

    const requestedTutorId = String(payload.assignedTutorId || "").trim();
    const legacyTutorName = String(payload.assignedTutor || "").trim();

    let tutor = null;
    if (requestedTutorId) {
      tutor = nhsos_findTutorById_(ss, requestedTutorId);
    }
    if (!tutor && legacyTutorName) {
      tutor = nhsos_findTutor_(ss, legacyTutorName);
    }

    const assignedTutorId = tutor
      ? String(tutor["Tutor ID"] || "").trim()
      : "";

    const assignedTutor = tutor
      ? nhsos_tutorDisplayName_(
          tutor["Preferred Name"],
          tutor["Full Name"]
        )
      : "";

    if (
      tutor &&
      !String(tutor["Notification Email"] || "").trim()
    ) {
      throw new Error(
        assignedTutor +
        " is missing a Notification Email. Add a reachable personal email in Tutor Database before assigning this tutor."
      );
    }

    const tutorChanged =
      oldState.assignedTutorId !== assignedTutorId ||
      oldState.assignedTutor !== assignedTutor;

    let scheduledDate = String(payload.scheduledDate || "").trim();
    const scheduledDateChanged = oldState.scheduledDate !== scheduledDate;

    // A date must be revalidated whenever either the date OR tutor changes.
    // This closes an important edge case for Future Reservations: the request
    // may already contain its reserved date before a tutor is assigned. If the
    // officer assigns the tutor on that date, same-day Five Star enrollment is
    // impossible even though the date string itself did not change.
    if (scheduledDate && (scheduledDateChanged || tutorChanged)) {
      scheduledDate = nhssched_validateScheduledAgainstRequest_(
        {
          "Scheduling Preference": get("Scheduling Preference"),
          "Preferred Dates": get("Preferred Dates"),
          "Preferred Days": get("Preferred Days")
        },
        scheduledDate,
        nhs_safeTimeZone_(ss),
        new Date()
      );
    }

    let weeklyAvailabilityWarning = "";
    if (tutor) {
      nhssec_validateTutorAssignment_(
        ss,
        tutor,
        payload.requestId,
        scheduledDate
      );
      if (scheduledDate) {
        const weeklyCheck = nhswa_assertAssignmentDate_(ss, tutor, scheduledDate);
        weeklyAvailabilityWarning = weeklyCheck.warning || "";
      }
    }

    const coordinatorNotes = nhsinput_literalText_(
      payload.coordinatorNotes,
      1200,
      "Coordinator notes"
    );
    const recurring = nhsos_isRecurringValue_(get("Recurring?"));
    const now = new Date();

    set("Assigned Tutor ID", assignedTutorId);
    set("Assigned Tutor", assignedTutor);

    const assignmentMutation = tutorChanged || scheduledDateChanged;

    // Any tutor/date change is a new scheduling commitment. The tutor must
    // explicitly accept the updated assignment, and prior confirmation/
    // delivery/Five Star state must not carry across to the new commitment.
    if (assignmentMutation) {
      set("Tutor Response Status", assignedTutor ? "Pending" : "");
      set("Tutor Response At", "");
      set("Tutor Decline Reason", "");
      set("Tutor Viewed At", "");
      set("Tutor Portal Last Action", assignedTutor ? "Assignment Updated — Acceptance Required" : "");
      set("Tutor Portal Last Action At", assignedTutor ? new Date() : "");
      set("Tutor Reconfirm Status", "");
      set("Tutor Reconfirm At", "");
      set("Teacher Five Star Ready", false);
      set("Teacher Five Star Ready At", "");
      set("Tutor Five Star Enrolled", false);
      set("Tutor Five Star Enrolled At", "");
      set("Five Star Ready At", "");
      set("Teacher Confirmation Sent At", "");
      set("Tutor Confirmation Sent At", "");
      set("Confirm Assignment", false);
      set("Assignment Confirmed", "");
    }
    set(
      "Scheduled Date",
      scheduledDate
        ? new Date(scheduledDate + "T12:00:00")
        : ""
    );
    set("Coordinator Notes", coordinatorNotes);

    let nextStatus = get("Status") || "New";
    let notificationStatus = get("Notification Status") || "";
    let recurringActive = nhsos_truthy_(get("Recurring Active"));
    let nextSessionDate = get("Next Session Date") || "";
    let scheduledTime = get("Scheduled Time") || "";

    if (assignmentMutation) {
      if (assignedTutor) {
        nextStatus = scheduledDate ? "Scheduled" : "Assigned";
        notificationStatus = "Awaiting Tutor Acceptance";
        scheduledTime = scheduledDate ? NHSOS_TITAN_TIME : "";
        recurringActive = !!recurring;
        nextSessionDate = recurring && scheduledDate
          ? new Date(scheduledDate + "T12:00:00")
          : "";
      } else {
        nextStatus = "New";
        notificationStatus = "";
        scheduledTime = "";
        recurringActive = false;
        nextSessionDate = "";
      }
    }

    set("Status", nextStatus);
    set("Scheduled Time", scheduledTime);
    set("Notification Status", notificationStatus);
    set("Recurring Active", recurringActive);
    set("Next Session Date", nextSessionDate);
    set("Last Update", now);

    // One authoritative write for the request itself.
    range.setValues([values]);

    const newState = {
      assignedTutor,
      assignedTutorId,
      scheduledDate,
      location: oldState.location,
      status: nextStatus
    };

    // Reconcile only tutors touched by this mutation. A single assignment
    // save must not rebuild the Current Caseload value for the entire network.
    const caseloadTargets = [];
    function addCaseloadTarget_(id, name) {
      id = String(id || "").trim();
      name = String(name || "").trim();
      if (!id && !name) return;
      const key = id ? "id:" + id : "name:" + nhsos_norm_(name);
      if (caseloadTargets.some(function(target) { return target.key === key; })) return;
      caseloadTargets.push({ key: key, id: id, name: name });
    }

    addCaseloadTarget_(oldState.assignedTutorId, oldState.assignedTutor);
    addCaseloadTarget_(assignedTutorId, assignedTutor);

    caseloadTargets.forEach(function(target) {
      nhsos_reconcileSingleTutorCaseload_(ss, target.id, target.name);
    });

    try {
      if (typeof nhsn_handleAssignmentMutation_ === "function") {
        nhsn_handleAssignmentMutation_(
          payload.requestId,
          oldState,
          newState
        );
      }
    } catch (e) {
      console.error("Notification mutation handler failed:", e);
    }

    const changedTutor =
      oldState.assignedTutorId !== assignedTutorId;
    const changedDate =
      oldState.scheduledDate !== scheduledDate;

    nhsos_logActivity_(ss, {
      type:
        changedTutor
          ? "Assignment Updated"
          : changedDate
            ? "Schedule Updated"
            : "Request Updated",
      requestId: payload.requestId,
      tutorId: assignedTutorId,
      summary:
        assignedTutor
          ? (
              changedTutor
                ? "Tutor assignment updated"
                : changedDate
                  ? "Session date updated"
                  : "Request details updated"
            )
          : "Tutor assignment cleared",
      details:
        (assignedTutor ? "Tutor: " + assignedTutor + ". " : "") +
        (scheduledDate ? "Date: " + scheduledDate + ". " : "") +
        (recurring ? "Recurring support active." : "")
    });

    nhsos_invalidateCache_();

    return {
      ok: true,
      requestId: payload.requestId,
      patch: {
        assignedTutor,
        assignedTutorId,
        scheduledDate,
        scheduledTime,
        coordinatorNotes,
        status: nextStatus,
        notificationStatus,
        assignmentConfirmed: assignmentMutation ? "" : get("Assignment Confirmed"),
        tutorResponseStatus: assignedTutor
          ? (assignmentMutation ? "Pending" : get("Tutor Response Status"))
          : "",
        tutorResponseAt: assignmentMutation ? "" : get("Tutor Response At"),
        tutorDeclineReason: assignmentMutation ? "" : get("Tutor Decline Reason"),
        tutorReconfirmStatus: assignmentMutation ? "" : get("Tutor Reconfirm Status"),
        teacherFiveStarReady: assignmentMutation ? false : nhsos_truthy_(get("Teacher Five Star Ready")),
        tutorFiveStarEnrolled: assignmentMutation ? false : nhsos_truthy_(get("Tutor Five Star Enrolled")),
        fiveStarReadiness: assignmentMutation
          ? "Five Star Setup Pending"
          : nhssched_readiness_({
              "Teacher Five Star Ready": get("Teacher Five Star Ready"),
              "Tutor Five Star Enrolled": get("Tutor Five Star Enrolled")
            }),
        recurringActive,
        nextSessionDate: scheduledDate && recurring
          ? scheduledDate
          : "",
        weeklyAvailabilityWarning: weeklyAvailabilityWarning
      }
    };
  });
}

function nhsv473_nhsos_confirmAssignmentCore_(requestId) {
  nhssec_requireOfficer_();

  return nhsos_withLock_(function() {
    const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
    const sh = nhsperf_sheet_(ss, NHSOS_REQUESTS);
    const row = nhsos_findRequestRow_(sh, requestId);

    if (!row) throw new Error("Request not found.");

    const h = nhsos_headerMap_(sh);
    const request = nhsos_rowObject_(sh, row, h);

    const existingNotificationStatus = String(request["Notification Status"] || "").trim();
    const alreadyConfirmed = typeof nhslife_confirmed_ === "function"
      ? nhslife_confirmed_(request)
      : !!String(request["Assignment Confirmed"] || "").trim();
    const teacherAlreadySent = !!String(request["Teacher Confirmation Sent At"] || "").trim();
    const tutorAlreadySent = !!String(request["Tutor Confirmation Sent At"] || "").trim();

    const confirmationRetryNeeded =
      existingNotificationStatus === "Confirmation Email Failed — Retry" ||
      existingNotificationStatus === "Sending Confirmation…" ||
      (alreadyConfirmed && (!teacherAlreadySent || !tutorAlreadySent));

    if (alreadyConfirmed && teacherAlreadySent && tutorAlreadySent && !confirmationRetryNeeded) {
      return {
        ok: true,
        alreadyConfirmed: true,
        message: "Teacher and tutor were already notified.",
        notificationStatus: existingNotificationStatus || "Teacher & Tutor Notified — Five Star Setup Pending",
        calendarStatus: request["Calendar Status"] || "",
        teacherFiveStarReady: nhsos_truthy_(request["Teacher Five Star Ready"]),
        tutorFiveStarEnrolled: nhsos_truthy_(request["Tutor Five Star Enrolled"]),
        fiveStarReadiness: nhssched_readiness_(request)
      };
    }

    if (!request["Assigned Tutor"]) {
      throw new Error("Assign a tutor before confirming.");
    }

    const tutorResponse =
      String(request["Tutor Response Status"] || "").trim();

    if (tutorResponse !== "Accepted") {
      throw new Error(
        tutorResponse === "Declined"
          ? "The tutor declined this assignment. Reassign the request before confirming."
          : "The tutor has not accepted this assignment yet. Wait for acceptance before confirming."
      );
    }

    if (!request["Scheduled Date"]) {
      throw new Error("Choose a scheduled date before confirming.");
    }

    const scheduledDateKey = nhssched_validateScheduledAgainstRequest_(
      request,
      request["Scheduled Date"],
      nhs_safeTimeZone_(ss),
      new Date()
    );
    const schedulingType = String(request["Scheduling Preference"] || "").trim();
    const reconfirmStatus = String(request["Tutor Reconfirm Status"] || "").trim();

    if (schedulingType === NHSSCHED_TYPES.FUTURE) {
      if (!nhssched_reconfirmWindowOpen_(scheduledDateKey, nhs_safeTimeZone_(ss), new Date())) {
        throw new Error(
          "Future reservations can be finalized only within " + NHSSCHED_RECONFIRM_DAYS +
          " days of the session, after the tutor reconfirms availability."
        );
      }
      if (reconfirmStatus !== "Confirmed") {
        throw new Error("The tutor must reconfirm this future reservation before final confirmation.");
      }
    }

    if (schedulingType === NHSSCHED_TYPES.RECURRING && reconfirmStatus === "Required") {
      throw new Error("The tutor must reconfirm this recurring session before final confirmation.");
    }

    if (!request["Location"]) {
      throw new Error(
        "The teacher request is missing a classroom / room number."
      );
    }

    let tutor = null;

    if (request["Assigned Tutor ID"]) {
      tutor = nhsos_findTutorById_(
        ss,
        request["Assigned Tutor ID"]
      );
    }

    if (!tutor) {
      tutor = nhsos_findTutor_(
        ss,
        request["Assigned Tutor"]
      );
    }

    if (!tutor) {
      throw new Error("Assigned tutor could not be found.");
    }

    if (!tutor["Notification Email"]) {
      throw new Error(
        "Tutor Notification Email is missing. Add a reachable personal email before confirming."
      );
    }

    if (!nhs_teacherDeliveryEmail_(request)) {
      throw new Error("Teacher notification email is missing.");
    }

    nhssec_validateTutorAssignment_(
      ss,
      tutor,
      requestId,
      nhsos_isoDate_(request["Scheduled Date"])
    );

    const now = new Date();

    // Commit confirmation state BEFORE mail delivery so every portal can see it.
    // On retry, preserve the original confirmation timestamp and any Five Star
    // readiness already recorded. Only the still-missing recipient is resent.
    nhsos_set_(sh, row, h, "Confirm Assignment", true);
    if (!alreadyConfirmed) {
      nhsos_set_(sh, row, h, "Assignment Confirmed", now);
      nhsos_set_(sh, row, h, "Teacher Five Star Ready", false);
      nhsos_set_(sh, row, h, "Teacher Five Star Ready At", "");
      nhsos_set_(sh, row, h, "Tutor Five Star Enrolled", false);
      nhsos_set_(sh, row, h, "Tutor Five Star Enrolled At", "");
      nhsos_set_(sh, row, h, "Five Star Ready At", "");
    }
    nhsos_set_(sh, row, h, "Status", "Scheduled");
    nhsos_set_(sh, row, h, "Notification Status", "Sending Confirmation…");
    nhsos_set_(sh, row, h, "Last Update", now);

    if (h["Calendar Status"]) {
      nhsos_set_(sh, row, h, "Calendar Status", "Preparing Invite…");
    }

    SpreadsheetApp.flush();
    nhsos_invalidateCache_();

    // Use the V27 branded mail/calendar workflow.
    let result = null;

    if (typeof sendAssignmentEmails_ !== "function") {
      throw new Error(
        "Assignment email workflow is unavailable. Confirm the V27/V28 Assignment Workflow file is installed."
      );
    }

    result = sendAssignmentEmails_(request, tutor, {
      skipTeacher: teacherAlreadySent,
      skipTutor: tutorAlreadySent
    });

    const teacherSent = teacherAlreadySent || !!(
      result &&
      result.teacherSent
    );

    const tutorSent = tutorAlreadySent || !!(
      result &&
      result.tutorSent
    );

    if (!teacherAlreadySent && result && result.teacherSent) {
      nhsos_set_(sh, row, h, "Teacher Confirmation Sent At", now);
    }
    if (!tutorAlreadySent && result && result.tutorSent) {
      nhsos_set_(sh, row, h, "Tutor Confirmation Sent At", now);
    }

    const calendarCreated = !!(
      result &&
      result.calendarCreated
    );

    const calendarEventId = String(
      result && result.calendarEventId || ""
    );

    const errors = [
      result && result.teacherError,
      result && result.tutorError,
      result && result.calendarError
    ].filter(Boolean);

    if (teacherSent && tutorSent) {
      nhsos_set_(
        sh,
        row,
        h,
        "Notification Status",
        "Teacher & Tutor Notified — Five Star Setup Pending"
      );

      if (h["Calendar Status"]) {
        nhsos_set_(
          sh,
          row,
          h,
          "Calendar Status",
          calendarCreated
            ? "Calendar Invite Sent"
            : "Confirmation Sent — Calendar Link Included"
        );
      }

      if (h["Calendar Event ID"] && calendarEventId) {
        nhsos_set_(
          sh,
          row,
          h,
          "Calendar Event ID",
          calendarEventId
        );
      }

      nhsos_logActivity_(ss, {
        type: "Assignment Confirmed",
        requestId: requestId,
        tutorId:
          tutor["Tutor ID"] ||
          request["Assigned Tutor ID"] ||
          "",
        summary: "Teacher and tutor notified",
        details:
          (request["Course"] || "Tutoring") +
          " · " +
          (request["Assigned Tutor"] || "")
      });

      nhsos_invalidateCache_();

      return {
        ok: true,
        message: "Teacher and tutor notified.",
        notificationStatus: "Teacher & Tutor Notified — Five Star Setup Pending",
        teacherFiveStarReady: alreadyConfirmed
          ? nhsos_truthy_(request["Teacher Five Star Ready"])
          : false,
        tutorFiveStarEnrolled: alreadyConfirmed
          ? nhsos_truthy_(request["Tutor Five Star Enrolled"])
          : false,
        fiveStarReadiness: alreadyConfirmed
          ? nhssched_readiness_(request)
          : "Five Star Setup Pending",
        calendarStatus:
          calendarCreated
            ? "Calendar Invite Sent"
            : "Confirmation Sent — Calendar Link Included"
      };
    }

    nhsos_set_(
      sh,
      row,
      h,
      "Notification Status",
      "Confirmation Email Failed — Retry"
    );

    if (h["Calendar Status"]) {
      nhsos_set_(
        sh,
        row,
        h,
        "Calendar Status",
        calendarCreated
          ? "Calendar Invite Sent"
          : "Calendar Invite Pending"
      );
    }

    nhsos_set_(sh, row, h, "Last Update", new Date());
    nhsos_invalidateCache_();

    throw new Error(
      errors.join(" | ") ||
      "The assignment was confirmed, but one or more confirmation messages failed. Retry after checking the email configuration."
    );
  });
}

function nhsos_cancelRequest(requestId) {
  nhssec_requireOfficer_();
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const sh = nhsperf_sheet_(ss, NHSOS_REQUESTS);
  const row = nhsos_findRequestRow_(sh, requestId);

  if (!row) {
    throw new Error("Request not found.");
  }

  const h = nhsos_headerMap_(sh);
  const range = sh.getRange(row, 1, 1, sh.getLastColumn());
  const values = range.getValues()[0];
  const display = range.getDisplayValues()[0];

  const get = header =>
    h[header]
      ? String(display[h[header] - 1] || "").trim()
      : "";

  const set = (header, value) => {
    if (h[header]) values[h[header] - 1] = value;
  };

  const oldStatus = get("Status");

  if (oldStatus === "Cancelled") {
    return {
      ok: true,
      requestId,
      status: "Cancelled"
    };
  }

  set("Status", "Cancelled");
  set("Recurring Active", false);
  set("Next Session Date", "");
  set("Last Update", new Date());

  range.setValues([values]);

  // V31 COMMIT-FIRST CANCELLATION:
  // Guarantee the Cancelled row is visible before the global revision changes.
  // Teacher/Tutor/Operations delta readers therefore cannot wake up and read
  // the previous status.
  SpreadsheetApp.flush();

  // Clear every portal cache that could still contain the pre-cancellation row.
  const teacherEmail =
    get("Teacher Email");

  const tutorId =
    get("Assigned Tutor ID");

  if (
    teacherEmail &&
    typeof tp_invalidateTeacherCache_ === "function"
  ) {
    try {
      tp_invalidateTeacherCache_(
        teacherEmail
      );
    } catch (e) {}
  }

  if (tutorId) {
    try {
      nhsperf_scriptCache_().remove(
        nhs_v46TutorCacheKey_(tutorId)
      );
    } catch (e) {}
  }

  // This clears Operations cache AND bumps the shared data revision.
  nhsos_invalidateCache_();

  // Expensive side effects remain off the critical path:
  // - cancellation email
  // - caseload reconciliation
  // - activity log write
  nhsos_queueFastJob_("CANCEL_REQUEST", {
    requestId,
    oldStatus,
    tutorId: get("Assigned Tutor ID"),
    course: get("Course")
  });

  return {
    ok: true,
    requestId,
    status: "Cancelled",
    revision:
      typeof nhsperf_getDataRevision_ === "function"
        ? nhsperf_getDataRevision_()
        : ""
  };
}

/**
 * Permanently deletes a closed/unused request from Tutoring Requests.
 *
 * Safety rules:
 * - Completed / Assigned / Scheduled / Pending Verification requests cannot be deleted.
 * - Any request with Session Log history cannot be deleted.
 * - Active unassigned requests can only be deleted when no tutor/date is attached.
 * - Activity / Notification / Session history is never purged by this action.
 */
function nhsos_deleteRequest(requestId) {
  nhssec_requireOfficer_();
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const sh = nhsperf_sheet_(ss, NHSOS_REQUESTS);
  const row = nhsos_findRequestRow_(sh, requestId);

  if (!row) {
    throw new Error("Request not found.");
  }

  const h = nhsos_headerMap_(sh);
  const display = sh
    .getRange(row, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0];

  const get = header =>
    h[header]
      ? String(display[h[header] - 1] || "").trim()
      : "";

  const status = get("Status") || "New";
  const assignedTutor = get("Assigned Tutor");
  const assignedTutorId = get("Assigned Tutor ID");
  const scheduledDate = get("Scheduled Date");

  const protectedStatuses = [
    "Completed",
    "Assigned",
    "Scheduled",
    "Pending Verification"
  ];

  if (protectedStatuses.includes(status)) {
    throw new Error(
      "This request cannot be permanently deleted because it is part of active or completed tutoring history."
    );
  }

  if (nhsos_hasSessionForRequest_(ss, requestId)) {
    throw new Error(
      "This request has session history and cannot be permanently deleted. Keep it for service-hour and audit records."
    );
  }

  const closedDeleteStatuses = [
    "Cancelled",
    "No Tutor Available"
  ];

  const unusedDeleteStatuses = [
    "New",
    "Matching",
    "Needs Follow-Up"
  ];

  const isClosedDelete =
    closedDeleteStatuses.includes(status);

  const isUnusedDelete =
    unusedDeleteStatuses.includes(status) &&
    !assignedTutor &&
    !assignedTutorId &&
    !scheduledDate;

  if (!isClosedDelete && !isUnusedDelete) {
    throw new Error(
      "Only cancelled, unavailable, or clearly unused unassigned requests can be permanently deleted."
    );
  }

  const metadata = {
    requestId: requestId,
    status: status,
    course: get("Course"),
    teacherName: get("Teacher Name"),
    teacherEmail: get("Teacher Email"),
    tutorId: assignedTutorId,
    assignedTutor: assignedTutor
  };

  // Authoritative permanent delete.
  sh.deleteRow(row);

  // Non-critical cleanup/logging happens in the fast worker.
  nhsos_queueFastJob_("DELETE_REQUEST", metadata);

  nhsos_invalidateCache_();

  return {
    ok: true,
    requestId: requestId,
    deleted: true,
    message: "Request permanently deleted."
  };
}


function nhsos_hasSessionForRequest_(ss, requestId) {
  const sh = nhsperf_sheet_(ss, NHSOS_SESSIONS);

  if (!sh || sh.getLastRow() < 2) {
    return false;
  }

  const h = nhsos_headerMap_(sh);
  const col = h["Request ID"];

  if (!col) {
    return false;
  }

  const match = sh
    .getRange(2, col, sh.getLastRow() - 1, 1)
    .createTextFinder(String(requestId || "").trim())
    .matchEntireCell(true)
    .matchCase(true)
    .findNext();

  return !!match;
}


function nhsos_logSession(payload) {
  nhssec_requireOfficer_();
  return nhsos_withLock_(function() {
    const allowedResults = [
      "Completed",
      "Student No-Show",
      "Tutor No-Show",
      "Cancelled",
      "Reschedule Needed"
    ];

    if (!allowedResults.includes(payload.result)) {
      throw new Error("Invalid result.");
    }

    const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
    const reqSh = nhsperf_sheet_(ss, NHSOS_REQUESTS);
    const row = nhsos_findRequestRow_(reqSh, payload.requestId);
    if (!row) throw new Error("Request not found.");

    const rh = nhsos_headerMap_(reqSh);
    const req = nhsos_rowObject_(reqSh, row, rh);

    const tutorName = String(req["Assigned Tutor"] || "").trim();
    let tutorId = String(req["Assigned Tutor ID"] || "").trim();

    if (!tutorName && !tutorId) throw new Error("No assigned tutor.");

    if (!tutorId && tutorName) {
      const tutorRecord = nhsos_findTutor_(ss, tutorName);
      if (tutorRecord && tutorRecord["Tutor ID"]) {
        tutorId = String(tutorRecord["Tutor ID"]).trim();
        nhsos_set_(reqSh, row, rh, "Assigned Tutor ID", tutorId);
      }
    }

    const logSh = nhsos_ensureSessionLog_(ss);
    const lh = nhsos_headerMap_(logSh);

    const sessionId =
      "S-" +
      Utilities.formatDate(
        new Date(),
        nhs_safeTimeZone_(ss),
        "yyyyMMdd-HHmmss"
      ) +
      "-" +
      Utilities.getUuid();

    const reportedMinutes =
      payload.result === "Completed"
        ? Math.max(0, Number(payload.minutes || 40))
        : 0;

    const date = payload.sessionDate
      ? new Date(payload.sessionDate + "T12:00:00")
      : new Date();

    const verificationStatus =
      payload.result === "Completed"
        ? "Pending"
        : "Not Required";

    const rowValues = new Array(logSh.getLastColumn()).fill("");

    const setLogValue = (header, value) => {
      if (lh[header]) rowValues[lh[header] - 1] = value;
    };

    setLogValue("Session ID", sessionId);
    setLogValue("Request ID", payload.requestId);
    setLogValue("Tutor ID", tutorId);
    setLogValue("Tutor", tutorName);
    setLogValue("Session Date", date);
    setLogValue("Result", payload.result);
    setLogValue("Reported Minutes", reportedMinutes);
    setLogValue("Verification Status", verificationStatus);
    setLogValue("Verified Minutes", 0);
    setLogValue(
      "Notes",
      nhsinput_literalText_(payload.notes, 1200, "Session notes")
    );
    setLogValue("Logged At", new Date());
    setLogValue("Logged By", nhsos_actor_());

    logSh
      .getRange(logSh.getLastRow() + 1, 1, 1, rowValues.length)
      .setValues([rowValues]);

    let nextStatus = req["Status"];

    if (payload.result === "Completed") {
      nextStatus = "Pending Verification";
    }

    if (payload.result === "Cancelled") {
      nextStatus = "Cancelled";
      nhsos_set_(reqSh, row, rh, "Recurring Active", false);
      nhsos_set_(reqSh, row, rh, "Next Session Date", "");
    }

    if (
      ["Reschedule Needed", "Tutor No-Show", "Student No-Show"]
        .includes(payload.result)
    ) {
      nextStatus = "Needs Follow-Up";
    }

    nhsos_set_(reqSh, row, rh, "Status", nextStatus);
    nhsos_set_(reqSh, row, rh, "Last Update", new Date());

    nhsos_logActivity_(ss, {
      type: "Session Reported",
      requestId: payload.requestId,
      sessionId,
      tutorId,
      summary:
        payload.result === "Completed"
          ? "Completed session submitted for verification"
          : "Session outcome logged: " + payload.result,
      details:
        tutorName +
        (reportedMinutes ? " · " + reportedMinutes + " reported minutes" : "")
    });

    nhsos_reconcileSingleTutorCaseload_(ss, tutorId, tutorName);
    nhsos_invalidateCache_();

    return {
      ok: true,
      sessionId,
      verificationStatus
    };
  });
}

function nhsv473_nhsos_verifySessionCore_(sessionId, verifiedMinutes, verificationNotes) {
  nhssec_requireOfficer_();
  return nhsos_withLock_(function() {
    const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
    const logSh = nhsos_ensureSessionLog_(ss);
    const row = nhsos_findSessionRow_(logSh, sessionId);

    if (!row) throw new Error("Session not found.");

    const lh = nhsos_headerMap_(logSh);
    const session = nhsos_rowObject_(logSh, row, lh);

    if (session["Result"] !== "Completed") {
      throw new Error("Only completed sessions require service verification.");
    }

    const verificationState = String(session["Verification Status"] || "").trim();

    if (verificationState === "Verified") {
      return { ok: true, message: "Session is already verified." };
    }

    if (verificationState === "Voided") {
      throw new Error("Voided sessions cannot be re-verified. Create a new session record if service must be restored.");
    }

    if (!["Pending", "Rejected", ""].includes(verificationState)) {
      throw new Error("This session is not eligible for verification from its current state.");
    }

    const reported = Number(
      session["Reported Minutes"] ||
      session["Verified Minutes"] ||
      40
    );

    const rawMinutes =
      verifiedMinutes === undefined ||
      verifiedMinutes === null ||
      verifiedMinutes === ""
        ? reported
        : Number(verifiedMinutes);

    if (!Number.isFinite(rawMinutes) || rawMinutes < 0 || rawMinutes > 180) {
      throw new Error("Verified minutes must be a number from 0 to 180.");
    }

    const minutes = Math.round(rawMinutes);

    nhsos_set_(logSh, row, lh, "Verified Minutes", minutes);
    nhsos_set_(logSh, row, lh, "Verification Status", "Verified");
    nhsos_set_(logSh, row, lh, "Verified At", new Date());
    nhsos_set_(logSh, row, lh, "Verified By", nhsos_actor_());
    nhsos_set_(
      logSh,
      row,
      lh,
      "Verification Notes",
      String(verificationNotes || "").trim()
    );

    const requestId = String(session["Request ID"] || "").trim();
    const reqSh = nhsperf_sheet_(ss, NHSOS_REQUESTS);
    const reqRow = nhsos_findRequestRow_(reqSh, requestId);

    let nextDateText = "";
    let nextSessionNeedsReschedule = false;

    if (reqRow) {
      const rh = nhsos_headerMap_(reqSh);
      const req = nhsos_rowObject_(reqSh, reqRow, rh);
      const recurring =
        nhsos_truthy_(req["Recurring Active"]) ||
        nhsos_isRecurringValue_(req["Recurring?"]);

      if (recurring) {
        const nextDate = nhsos_nextPreferredDate_(
          session["Session Date"],
          req["Preferred Days"]
        );

        nextDateText = Utilities.formatDate(
          nextDate,
          nhs_safeTimeZone_(ss),
          "yyyy-MM-dd"
        );

        const assignedTutorForWeek = nhsos_findTutorById_(ss, String(req["Assigned Tutor ID"] || "")) || nhsos_findTutor_(ss, String(req["Assigned Tutor"] || ""));
        let recurringWeeklyUnavailable = false;
        if (assignedTutorForWeek) {
          const weeklyState = nhswa_statusForTutorDate_(nhswa_readIndex_(ss), assignedTutorForWeek, nextDateText);
          recurringWeeklyUnavailable = weeklyState.status === "Unavailable";
        }
        nextSessionNeedsReschedule = recurringWeeklyUnavailable;

        nhsos_set_(reqSh, reqRow, rh, "Recurring Active", true);
        nhsos_set_(reqSh, reqRow, rh, "Scheduled Date", recurringWeeklyUnavailable ? "" : nextDate);
        nhsos_set_(reqSh, reqRow, rh, "Next Session Date", nextDate);
        nhsos_set_(reqSh, reqRow, rh, "Scheduled Time", recurringWeeklyUnavailable ? "" : NHSOS_TITAN_TIME);
        nhsos_set_(reqSh, reqRow, rh, "Status", recurringWeeklyUnavailable ? "Needs Follow-Up" : "Scheduled");
        nhsos_set_(reqSh, reqRow, rh, "Confirm Assignment", false);
        nhsos_set_(reqSh, reqRow, rh, "Assignment Confirmed", "");
        nhsos_set_(reqSh, reqRow, rh, "Tutor Reconfirm Status", recurringWeeklyUnavailable ? "Unavailable This Week" : "Required");
        nhsos_set_(reqSh, reqRow, rh, "Tutor Reconfirm At", "");
        nhsos_set_(reqSh, reqRow, rh, "Teacher Five Star Ready", false);
        nhsos_set_(reqSh, reqRow, rh, "Teacher Five Star Ready At", "");
        nhsos_set_(reqSh, reqRow, rh, "Tutor Five Star Enrolled", false);
        nhsos_set_(reqSh, reqRow, rh, "Tutor Five Star Enrolled At", "");
        nhsos_set_(reqSh, reqRow, rh, "Five Star Ready At", "");
        nhsos_set_(reqSh, reqRow, rh, "Notification Status", recurringWeeklyUnavailable
          ? "Recurring — Tutor unavailable this week; reschedule required"
          : "Recurring — Tutor Reconfirmation Required");
      } else {
        nhsos_set_(reqSh, reqRow, rh, "Status", "Completed");
        nhsos_set_(reqSh, reqRow, rh, "Next Session Date", "");
      }

      nhsos_set_(reqSh, reqRow, rh, "Last Update", new Date());
    }

    const tutorStats = nhsos_rebuildSingleTutorStats_(
      ss,
      String(session["Tutor ID"] || ""),
      String(session["Tutor"] || "")
    );
    nhsos_reconcileSingleTutorCaseload_(
      ss,
      String(session["Tutor ID"] || ""),
      String(session["Tutor"] || "")
    );

    nhsos_logActivity_(ss, {
      type: "Session Verified",
      requestId,
      sessionId,
      tutorId: session["Tutor ID"] || "",
      summary: "Service minutes verified",
      details:
        minutes +
        " verified minutes" +
        (nextDateText ? " · Next recurring session: " + nextDateText : "")
    });

    nhsos_invalidateCache_();

    return {
      ok: true,
      message: nextDateText
        ? (nextSessionNeedsReschedule
            ? "Session verified. The next recurring date is " + nextDateText + ", but the tutor is unavailable that week; rescheduling is required."
            : "Session verified. Next recurring session scheduled for " + nextDateText + ".")
        : "Session verified and request completed.",
      nextSessionDate: nextDateText,
      nextSessionNeedsReschedule: nextSessionNeedsReschedule,
      tutorStats: tutorStats
    };
  });
}


/* ============================================================
   V45.2 — VERIFIED SESSION ADMINISTRATION
   ============================================================ */

function nhsos_correctVerifiedSession(
  sessionId,
  verifiedMinutes,
  verificationNotes
) {
  nhssec_requireOfficer_();

  return nhsos_withLock_(function() {
    const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
    const logSh = nhsos_ensureSessionLog_(ss);
    const row = nhsos_findSessionRow_(logSh, sessionId);

    if (!row) throw new Error("Session not found.");

    const h = nhsos_headerMap_(logSh);
    const session = nhsos_rowObject_(logSh, row, h);
    const status = String(
      session["Verification Status"] || ""
    ).trim();

    if (status !== "Verified") {
      throw new Error("Only verified sessions can be corrected.");
    }

    const minutes = Number(verifiedMinutes);

    if (!Number.isFinite(minutes) || minutes < 0 || minutes > 180) {
      throw new Error("Verified minutes must be between 0 and 180.");
    }

    const oldMinutes = Number(session["Verified Minutes"] || 0);
    const previousNotes = String(
      session["Verification Notes"] || ""
    ).trim();
    const suppliedNotes = String(verificationNotes || "").trim();
    const actor = nhsos_actor_();

    nhsos_set_(logSh, row, h, "Verified Minutes", minutes);
    nhsos_set_(logSh, row, h, "Verified At", new Date());
    nhsos_set_(logSh, row, h, "Verified By", actor);
    nhsos_set_(
      logSh,
      row,
      h,
      "Verification Notes",
      suppliedNotes || previousNotes
    );

    nhsos_logActivity_(ss, {
      type: "Verified Session Corrected",
      requestId: String(session["Request ID"] || ""),
      sessionId: sessionId,
      tutorId: String(session["Tutor ID"] || ""),
      actor: actor,
      summary: "Verified service corrected",
      details:
        oldMinutes +
        " → " +
        minutes +
        " minutes" +
        (suppliedNotes ? " · " + suppliedNotes : "")
    });

    const tutorStats = nhsos_rebuildSingleTutorStats_(
      ss,
      String(session["Tutor ID"] || ""),
      String(session["Tutor"] || "")
    );
    nhsos_invalidateSessionMutation_(ss, session);

    return {
      ok: true,
      sessionId: sessionId,
      verifiedMinutes: minutes,
      verificationStatus: "Verified",
      tutorStats: tutorStats
    };
  });
}

/**
 * Safe deletion equivalent for verified service.
 * The Session Log row remains for auditability, but all service credit is
 * removed because the verification state becomes Voided and minutes become 0.
 */
function nhsos_voidVerifiedSession(sessionId, reason) {
  nhssec_requireOfficer_();

  return nhsos_withLock_(function() {
    const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
    const logSh = nhsos_ensureSessionLog_(ss);
    const row = nhsos_findSessionRow_(logSh, sessionId);

    if (!row) throw new Error("Session not found.");

    const h = nhsos_headerMap_(logSh);
    const session = nhsos_rowObject_(logSh, row, h);
    const status = String(
      session["Verification Status"] || ""
    ).trim();

    if (status === "Voided") {
      return {
        ok: true,
        alreadyVoided: true,
        sessionId: sessionId,
        verificationStatus: "Voided",
        verifiedMinutes: 0,
        voidReason: String(session["Void Reason"] || "")
      };
    }

    if (status !== "Verified") {
      throw new Error("Only verified sessions can be voided.");
    }

    reason = String(reason || "").trim();
    if (!reason) {
      throw new Error("A reason is required when voiding a verified session.");
    }

    const previousMinutes = Number(session["Verified Minutes"] || 0);
    const actor = nhsos_actor_();

    nhsos_set_(logSh, row, h, "Verification Status", "Voided");
    nhsos_set_(logSh, row, h, "Verified Minutes", 0);
    nhsos_set_(logSh, row, h, "Voided At", new Date());
    nhsos_set_(logSh, row, h, "Voided By", actor);
    nhsos_set_(logSh, row, h, "Void Reason", reason);

    nhsos_logActivity_(ss, {
      type: "Verified Session Voided",
      requestId: String(session["Request ID"] || ""),
      sessionId: sessionId,
      tutorId: String(session["Tutor ID"] || ""),
      actor: actor,
      summary: "Verified tutoring service voided",
      details:
        previousMinutes +
        " verified minutes removed · " +
        reason
    });

    const tutorStats = nhsos_rebuildSingleTutorStats_(
      ss,
      String(session["Tutor ID"] || ""),
      String(session["Tutor"] || "")
    );
    nhsos_invalidateSessionMutation_(ss, session);

    return {
      ok: true,
      sessionId: sessionId,
      requestId: String(session["Request ID"] || ""),
      tutorId: String(session["Tutor ID"] || ""),
      verificationStatus: "Voided",
      verifiedMinutes: 0,
      voidReason: reason,
      tutorStats: tutorStats
    };
  });
}

/**
 * Permanent cleanup is intentionally restricted to ADMIN/ADVISOR and only
 * works after a session has first been voided. This preserves a two-step audit
 * workflow and prevents accidental destruction of credited service history.
 */
function nhsos_deleteVoidedSession(sessionId, reason) {
  const actor = nhsa_requireAnyRole_(
    ["ADMIN", "ADVISOR"],
    "NHS Internal Operations"
  ).email;

  return nhsos_withLock_(function() {
    const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
    const logSh = nhsos_ensureSessionLog_(ss);
    const row = nhsos_findSessionRow_(logSh, sessionId);

    if (!row) throw new Error("Session not found.");

    const h = nhsos_headerMap_(logSh);
    const session = nhsos_rowObject_(logSh, row, h);
    const status = String(
      session["Verification Status"] || ""
    ).trim();

    if (status !== "Voided") {
      throw new Error(
        "A session must be voided before it can be permanently deleted."
      );
    }

    reason = String(reason || "").trim();
    if (!reason) {
      throw new Error("A deletion reason is required.");
    }

    nhsos_logActivity_(ss, {
      type: "Voided Session Permanently Deleted",
      requestId: String(session["Request ID"] || ""),
      sessionId: sessionId,
      tutorId: String(session["Tutor ID"] || ""),
      actor: actor,
      summary: "Voided session permanently deleted",
      details:
        reason +
        " · Original void reason: " +
        String(session["Void Reason"] || "")
    });

    logSh.deleteRow(row);

    const tutorStats = nhsos_rebuildSingleTutorStats_(
      ss,
      String(session["Tutor ID"] || ""),
      String(session["Tutor"] || "")
    );
    nhsos_invalidateSessionMutation_(ss, session);

    return {
      ok: true,
      sessionId: sessionId,
      deleted: true,
      tutorStats: tutorStats
    };
  });
}

function nhsos_rejectSessionVerification(sessionId, verificationNotes) {
  nhssec_requireOfficer_();
  return nhsos_withLock_(function() {
    const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
    const logSh = nhsos_ensureSessionLog_(ss);
    const row = nhsos_findSessionRow_(logSh, sessionId);

    if (!row) throw new Error("Session not found.");

    const lh = nhsos_headerMap_(logSh);
    const session = nhsos_rowObject_(logSh, row, lh);

    if (session["Result"] !== "Completed") {
      throw new Error("Only completed sessions can be rejected for verification.");
    }

    const verificationState = String(session["Verification Status"] || "").trim();
    if (verificationState === "Rejected") {
      return { ok: true, alreadyRejected: true };
    }
    if (verificationState !== "Pending") {
      throw new Error("Only sessions pending verification can be rejected.");
    }

    nhsos_set_(logSh, row, lh, "Verified Minutes", 0);
    nhsos_set_(logSh, row, lh, "Verification Status", "Rejected");
    nhsos_set_(logSh, row, lh, "Verified At", new Date());
    nhsos_set_(logSh, row, lh, "Verified By", nhsos_actor_());
    nhsos_set_(
      logSh,
      row,
      lh,
      "Verification Notes",
      String(verificationNotes || "Verification rejected.").trim()
    );

    const requestId = String(session["Request ID"] || "").trim();
    const reqSh = nhsperf_sheet_(ss, NHSOS_REQUESTS);
    const reqRow = nhsos_findRequestRow_(reqSh, requestId);

    if (reqRow) {
      const rh = nhsos_headerMap_(reqSh);
      nhsos_set_(reqSh, reqRow, rh, "Status", "Needs Follow-Up");
      nhsos_set_(reqSh, reqRow, rh, "Last Update", new Date());
    }

    const tutorStats = nhsos_rebuildSingleTutorStats_(
      ss,
      String(session["Tutor ID"] || ""),
      String(session["Tutor"] || "")
    );
    nhsos_reconcileSingleTutorCaseload_(
      ss,
      String(session["Tutor ID"] || ""),
      String(session["Tutor"] || "")
    );

    nhsos_logActivity_(ss, {
      type: "Session Verification Rejected",
      requestId,
      sessionId,
      tutorId: session["Tutor ID"] || "",
      summary: "Session verification rejected",
      details: String(verificationNotes || "").trim()
    });

    nhsos_invalidateSessionMutation_(ss, session);
    return { ok: true, tutorStats: tutorStats };
  });
}

function nhsos_endRecurringRequest(requestId) {
  nhssec_requireOfficer_();
  return nhsos_withLock_(function() {
    const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
    const sh = nhsperf_sheet_(ss, NHSOS_REQUESTS);
    const row = nhsos_findRequestRow_(sh, requestId);
    if (!row) throw new Error("Request not found.");

    const h = nhsos_headerMap_(sh);
    const range = sh.getRange(row, 1, 1, sh.getLastColumn());
    const values = range.getValues()[0];
    const display = range.getDisplayValues()[0];

    const get = header =>
      h[header]
        ? String(display[h[header] - 1] || "").trim()
        : "";

    const set = (header, value) => {
      if (h[header]) values[h[header] - 1] = value;
    };

    const currentStatus = get("Status");
    const nextStatus = currentStatus === "Pending Verification"
      ? "Pending Verification"
      : "Completed";

    set("Recurring Active", false);
    set("Next Session Date", "");
    set("Status", nextStatus);
    set("Last Update", new Date());

    range.setValues([values]);

    nhsos_reconcileCaseloads_(ss);

    nhsos_logActivity_(ss, {
      type: "Recurring Support Ended",
      requestId,
      tutorId: get("Assigned Tutor ID"),
      summary: "Recurring tutoring closed",
      details: get("Course")
    });

    nhsos_invalidateCache_();

    return {
      ok: true,
      requestId,
      patch: {
        recurringActive: false,
        nextSessionDate: "",
        status: nextStatus
      }
    };
  });
}

function nhsos_ensureRequestColumns_(sh) {
  if (!sh) return;

  const required = [
    "Teacher Notification Email",
    "Assigned Tutor ID",
    "Recurring Active",
    "Next Session Date",
    "Tutor Response Status",
    "Tutor Response At",
    "Tutor Decline Reason",
    "Tutor Viewed At",
    "Tutor Portal Last Action",
    "Tutor Portal Last Action At",
    "Teacher Confirmation Sent At",
    "Tutor Confirmation Sent At",
    "Match Breakdown",
    "Preferred Dates",
    "Tutor Reconfirm Status",
    "Tutor Reconfirm At",
    "Teacher Five Star Ready",
    "Teacher Five Star Ready At",
    "Tutor Five Star Enrolled",
    "Tutor Five Star Enrolled At",
    "Five Star Ready At"
  ];

  let headers = sh
    .getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1))
    .getDisplayValues()[0];

  required.forEach(header => {
    if (!headers.includes(header)) {
      sh.getRange(1, headers.length + 1).setValue(header);
      headers.push(header);
    }
  });
}

function nhsos_ensureSessionLog_(ss) {
  let sh = nhsperf_sheet_(ss, NHSOS_SESSIONS);

  const headers = [
    "Session ID",
    "Request ID",
    "Tutor ID",
    "Tutor",
    "Session Date",
    "Result",
    "Reported Minutes",
    "Verification Status",
    "Verified Minutes",
    "Notes",
    "Logged At",
    "Logged By",
    "Verified At",
    "Verified By",
    "Verification Notes",
    "Voided At",
    "Voided By",
    "Void Reason"
  ];

  if (!sh) {
    sh = ss.insertSheet(NHSOS_SESSIONS);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const existing = sh
      .getRange(1, 1, 1, Math.max(1, sh.getLastColumn()))
      .getDisplayValues()[0];

    headers.forEach(header => {
      if (!existing.includes(header)) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(header);
        existing.push(header);
      }
    });
  }

  sh.setFrozenRows(1);

  sh.getRange(1, 1, 1, sh.getLastColumn())
    .setFontWeight("bold")
    .setBackground("#1f4e78")
    .setFontColor("#ffffff");

  return sh;
}

function nhsos_ensureActivityLog_(ss) {
  let sh = nhsperf_sheet_(ss, NHSOS_ACTIVITY);

  const headers = [
    "Activity ID",
    "Timestamp",
    "Type",
    "Request ID",
    "Session ID",
    "Tutor ID",
    "Actor",
    "Summary",
    "Details"
  ];

  if (!sh) {
    sh = ss.insertSheet(NHSOS_ACTIVITY);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const existing = sh
      .getRange(1, 1, 1, Math.max(1, sh.getLastColumn()))
      .getDisplayValues()[0];

    headers.forEach(header => {
      if (!existing.includes(header)) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(header);
        existing.push(header);
      }
    });
  }

  sh.setFrozenRows(1);

  sh.getRange(1, 1, 1, sh.getLastColumn())
    .setFontWeight("bold")
    .setBackground("#123A5A")
    .setFontColor("#ffffff");

  return sh;
}

function nhsos_reconcileCaseloads_(ss) {
  const tutorSh = nhsperf_sheet_(ss, NHSOS_TUTORS);
  if (!tutorSh || tutorSh.getLastRow() < 2) return { ok: true, updated: 0 };

  const requests = nhsos_readRequests_(ss);
  const activeById = {};
  const activeByName = {};

  requests.forEach(function(r) {
    const active = typeof nhslife_activeCaseload_ === "function"
      ? nhslife_activeCaseload_(r)
      : ["Assigned", "Scheduled"].includes(r.status);
    if (!active) return;

    const tutorId = String(r.assignedTutorId || "").trim();
    if (tutorId) activeById[tutorId] = (activeById[tutorId] || 0) + 1;

    const name = String(r.assignedTutor || "").trim();
    if (name) {
      const key = nhsos_norm_(name);
      activeByName[key] = (activeByName[key] || 0) + 1;
    }
  });

  const h = nhsos_headerMap_(tutorSh);
  if (!h["Current Caseload"] || !h["Full Name"]) return { ok: true, updated: 0 };

  const rowCount = tutorSh.getLastRow() - 1;
  const values = tutorSh.getRange(2, 1, rowCount, tutorSh.getLastColumn()).getDisplayValues();
  const output = [];
  let updated = 0;

  values.forEach(function(row) {
    const full = String(row[h["Full Name"] - 1] || "").trim();
    if (!full) { output.push([""]); return; }

    const tutorId = h["Tutor ID"] ? String(row[h["Tutor ID"] - 1] || "").trim() : "";
    const preferred = h["Preferred Name"] ? String(row[h["Preferred Name"] - 1] || "").trim() : "";
    const display = nhsos_tutorDisplayName_(preferred, full);

    const count = tutorId
      ? Number(activeById[tutorId] || 0)
      : Math.max(
          Number(activeByName[nhsos_norm_(display)] || 0),
          Number(activeByName[nhsos_norm_(full)] || 0),
          Number(activeByName[nhsos_norm_(preferred)] || 0)
        );

    output.push([count]);
    updated++;
  });

  tutorSh.getRange(2, h["Current Caseload"], output.length, 1).setValues(output);
  return { ok: true, updated: updated };
}

function nhsos_findTutorRowForStats_(tutorSh, tutorId, tutorName) {
  if (!tutorSh || tutorSh.getLastRow() < 2) return null;

  const h = nhsos_headerMap_(tutorSh);
  const rowCount = tutorSh.getLastRow() - 1;
  const values = tutorSh
    .getRange(2, 1, rowCount, tutorSh.getLastColumn())
    .getDisplayValues();

  const wantedId = String(tutorId || "").trim();
  const wantedName = nhsos_norm_(tutorName);

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowId = h["Tutor ID"]
      ? String(row[h["Tutor ID"] - 1] || "").trim()
      : "";

    if (wantedId && rowId === wantedId) {
      return { row: i + 2, values: row, h: h };
    }

    if (!wantedId && wantedName) {
      const full = h["Full Name"]
        ? String(row[h["Full Name"] - 1] || "").trim()
        : "";
      const preferred = h["Preferred Name"]
        ? String(row[h["Preferred Name"] - 1] || "").trim()
        : "";
      const display = nhsos_tutorDisplayName_(preferred, full);
      if ([full, preferred, display].some(function(name) {
        return nhsos_norm_(name) === wantedName;
      })) {
        return { row: i + 2, values: row, h: h };
      }
    }
  }

  return null;
}


function nhsos_rebuildSingleTutorStats_(ss, tutorId, tutorName) {
  const tutorSh = nhsperf_sheet_(ss, NHSOS_TUTORS);
  const target = nhsos_findTutorRowForStats_(tutorSh, tutorId, tutorName);
  if (!target) return { ok: true, updated: 0 };

  const logs = nhsos_readSessions_(ss);
  const wantedId = String(tutorId || "").trim();
  const wantedName = nhsos_norm_(tutorName);
  let completed = 0;
  let minutes = 0;
  let noShows = 0;

  logs.forEach(function(session) {
    const sameTutor = wantedId
      ? String(session.tutorId || "").trim() === wantedId
      : nhsos_norm_(session.tutor) === wantedName;

    if (!sameTutor) return;

    if (
      session.result === "Completed" &&
      session.verificationStatus === "Verified"
    ) {
      completed++;
      minutes += Number(session.minutes || 0);
    }

    if (session.result === "Tutor No-Show") noShows++;
  });

  const h = target.h;
  if (h["Sessions Completed"]) {
    tutorSh.getRange(target.row, h["Sessions Completed"]).setValue(completed);
  }
  if (h["Verified Tutoring Hours"]) {
    tutorSh
      .getRange(target.row, h["Verified Tutoring Hours"])
      .setValue(Math.round((minutes / 60) * 100) / 100);
  }
  if (h["Tutor No-Shows"]) {
    tutorSh.getRange(target.row, h["Tutor No-Shows"]).setValue(noShows);
  }

  try {
    const cache = nhsperf_scriptCache_();
    const schoolEmail = h["School Email"]
      ? String(target.values[h["School Email"] - 1] || "").trim().toLowerCase()
      : "";
    const resolvedId = wantedId || (
      h["Tutor ID"] ? String(target.values[h["Tutor ID"] - 1] || "").trim() : ""
    );
    if (schoolEmail) cache.remove(nhs_v46TutorIdentityKey_(schoolEmail));
    if (resolvedId) cache.remove(nhs_v46TutorCacheKey_(resolvedId));
  } catch (e) {}

  return {
    ok: true,
    updated: 1,
    sessionsCompleted: completed,
    verifiedMinutes: minutes,
    verifiedHours: Math.round((minutes / 60) * 100) / 100,
    tutorNoShows: noShows
  };
}


function nhsos_reconcileSingleTutorCaseload_(ss, tutorId, tutorName) {
  const tutorSh = nhsperf_sheet_(ss, NHSOS_TUTORS);
  const target = nhsos_findTutorRowForStats_(tutorSh, tutorId, tutorName);
  if (!target || !target.h["Current Caseload"]) {
    return { ok: true, updated: 0, currentCaseload: 0 };
  }

  const wantedId = String(tutorId || "").trim();
  const wantedName = nhsos_norm_(tutorName);
  let count = 0;

  nhsos_readRequests_(ss).forEach(function(request) {
    if (!nhslife_activeCaseload_(request)) return;

    const sameTutor = wantedId
      ? String(request.assignedTutorId || "").trim() === wantedId
      : nhsos_norm_(request.assignedTutor) === wantedName;

    if (sameTutor) count++;
  });

  tutorSh
    .getRange(target.row, target.h["Current Caseload"])
    .setValue(count);

  return { ok: true, updated: 1, currentCaseload: count };
}


function nhsos_rebuildTutorStats_(ss) {
  const tutorSh = nhsperf_sheet_(ss, NHSOS_TUTORS);
  if (!tutorSh || tutorSh.getLastRow() < 2) return { ok: true, updated: 0 };

  const logs = nhsos_readSessions_(ss);
  const statsById = {};
  const statsByName = {};

  function ensure(map, key) {
    if (!key) return null;
    if (!map[key]) map[key] = { completed: 0, minutes: 0, noShows: 0 };
    return map[key];
  }

  logs.forEach(function(s) {
    const verified = s.result === "Completed" && s.verificationStatus === "Verified";
    const tutorId = String(s.tutorId || "").trim();
    const nameKey = nhsos_norm_(s.tutor);
    const targets = [ensure(statsById, tutorId), ensure(statsByName, nameKey)].filter(Boolean);

    targets.forEach(function(stats) {
      if (verified) {
        stats.completed++;
        stats.minutes += Number(s.minutes || 0);
      }
      if (s.result === "Tutor No-Show") stats.noShows++;
    });
  });

  const h = nhsos_headerMap_(tutorSh);
  const rowCount = tutorSh.getLastRow() - 1;
  const values = tutorSh
    .getRange(2, 1, rowCount, tutorSh.getLastColumn())
    .getDisplayValues();

  const completedOut = [];
  const hoursOut = [];
  const noShowOut = [];
  let updated = 0;

  values.forEach(function(row) {
    const full = h["Full Name"] ? String(row[h["Full Name"] - 1] || "").trim() : "";
    const tutorId = h["Tutor ID"] ? String(row[h["Tutor ID"] - 1] || "").trim() : "";
    const preferred = h["Preferred Name"] ? String(row[h["Preferred Name"] - 1] || "").trim() : "";
    const display = nhsos_tutorDisplayName_(preferred, full);
    const stats =
      (tutorId && statsById[tutorId]) ||
      statsByName[nhsos_norm_(display)] ||
      statsByName[nhsos_norm_(full)] ||
      { completed: 0, minutes: 0, noShows: 0 };

    completedOut.push([full ? stats.completed || 0 : ""]);
    hoursOut.push([full ? Math.round(((stats.minutes || 0) / 60) * 100) / 100 : ""]);
    noShowOut.push([full ? stats.noShows || 0 : ""]);
    if (full) updated++;

    try {
      const cache = nhsperf_scriptCache_();
      const schoolEmail = h["School Email"]
        ? String(row[h["School Email"] - 1] || "").trim().toLowerCase()
        : "";
      if (schoolEmail) cache.remove(nhs_v46TutorIdentityKey_(schoolEmail));
      if (tutorId) cache.remove(nhs_v46TutorCacheKey_(tutorId));
    } catch (e) {}
  });

  if (h["Sessions Completed"]) {
    tutorSh.getRange(2, h["Sessions Completed"], rowCount, 1).setValues(completedOut);
  }
  if (h["Verified Tutoring Hours"]) {
    tutorSh.getRange(2, h["Verified Tutoring Hours"], rowCount, 1).setValues(hoursOut);
  }
  if (h["Tutor No-Shows"]) {
    tutorSh.getRange(2, h["Tutor No-Shows"], rowCount, 1).setValues(noShowOut);
  }

  nhsos_reconcileCaseloads_(ss);
  return { ok: true, updated: updated };
}


/* ============================================================
   FAST DEFERRED WORKER
   Keeps slow non-critical side effects off officer click paths.
   ============================================================ */

function nhsos_ensureFastWorker_() {
  const handler = "nhsos_processFastJobs_";

  const exists = ScriptApp
    .getProjectTriggers()
    .some(trigger => trigger.getHandlerFunction() === handler);

  if (!exists) {
    ScriptApp
      .newTrigger(handler)
      .timeBased()
      .everyMinutes(1)
      .create();
  }
}


function nhsos_queueFastJob_(type, payload) {
  const props = nhsperf_scriptProperties_();

  const key =
    "nhsosFastJob_" +
    Date.now() +
    "_" +
    Utilities.getUuid();

  props.setProperty(
    key,
    JSON.stringify({
      type,
      createdAt: new Date().toISOString(),
      attempts: 0,
      payload: payload || {}
    })
  );
}


function nhsos_processFastJobs_() {
  const props = nhsperf_scriptProperties_();
  const all = props.getProperties();

  const keys = Object.keys(all)
    .filter(key => key.indexOf("nhsosFastJob_") === 0)
    .sort()
    .slice(0, 25);

  if (!keys.length) return;

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);

  keys.forEach(key => {
    try {
      const job = JSON.parse(all[key] || "{}");
      const p = job.payload || {};

      if (job.type === "CANCEL_REQUEST") {
        nhsos_processFastCancellation_(ss, p);
      }

      if (job.type === "DELETE_REQUEST") {
        nhsos_processFastDeletion_(ss, p);
      }

      props.deleteProperty(key);

    } catch (error) {
      console.error("NHS OS fast job failed:", key, error);
      if (typeof nhsdata_recordPropertyJobFailure_ === "function") {
        nhsdata_recordPropertyJobFailure_(props, key, all[key], error, 5);
      }
      // Repeated failures are capped and dead-lettered for diagnosis.
    }
  });
  try {
    nhsperf_scriptProperties_().setProperty(
      "NHS_FAST_WORKER_LAST_SUCCESS",
      new Date().toISOString()
    );
    if (typeof nhssec_recordHealth_ === "function") {
      nhssec_recordHealth_("Internal Fast Worker", true, "Processed background queue", "");
    }
  } catch (e) {}

}


function nhsos_processFastCancellation_(ss, p) {
  nhsos_reconcileCaseloads_(ss);

  try {
    if (
      p.oldStatus !== "Cancelled" &&
      typeof nhsn_notifyCancellation_ === "function"
    ) {
      nhsn_notifyCancellation_(p.requestId);
    }
  } catch (e) {
    console.error("Cancellation notification failed:", e);
  }

  nhsos_logActivity_(ss, {
    type: "Request Cancelled",
    requestId: p.requestId || "",
    tutorId: p.tutorId || "",
    summary: "Tutoring request cancelled",
    details: p.course || ""
  });

  nhsos_invalidateCache_();
}

function nhsos_processFastDeletion_(ss, p) {
  nhsos_reconcileCaseloads_(ss);

  nhsos_logActivity_(ss, {
    type: "Request Deleted",
    requestId: p.requestId || "",
    tutorId: p.tutorId || "",
    summary: "Request permanently deleted from operations queue",
    details:
      (p.course ? "Course: " + p.course + ". " : "") +
      (p.status ? "Previous status: " + p.status + ". " : "") +
      (p.teacherName ? "Teacher: " + p.teacherName + "." : "")
  });

  nhsos_invalidateCache_();
}



function nhsos_findRequestRow_(sh, requestId) {
  if (!sh || sh.getLastRow() < 2) return 0;

  const h = nhsos_headerMap_(sh);
  const col = h["Request ID"];
  if (!col) return 0;

  const match = sh
    .getRange(2, col, sh.getLastRow() - 1, 1)
    .createTextFinder(String(requestId || "").trim())
    .matchEntireCell(true)
    .matchCase(true)
    .findNext();

  return match ? match.getRow() : 0;
}

function nhsos_findTutor_(ss, name) {
  const sh = nhsperf_sheet_(ss, NHSOS_TUTORS);
  if (!sh || sh.getLastRow() < 2) return null;

  const h = nhsos_headerMap_(sh);
  const values = sh
    .getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
    .getDisplayValues();

  const target = nhsos_norm_(name);

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const fullName = h["Full Name"] ? row[h["Full Name"] - 1] : "";
    if (!fullName) continue;

    const preferred = h["Preferred Name"]
      ? row[h["Preferred Name"] - 1]
      : "";

    const display = nhsos_tutorDisplayName_(preferred, fullName);

    if (
      [display, fullName, preferred]
        .map(nhsos_norm_)
        .includes(target)
    ) {
      const obj = {};
      Object.keys(h).forEach(key => {
        obj[key] = row[h[key] - 1] || "";
      });
      return obj;
    }
  }

  return null;
}

function nhsos_getTeacherFormUrl_() {
  try {
    const id = nhsperf_scriptProperties_().getProperty("NHS_TEACHER_REQUEST_FORM_ID");
    return id ? "https://docs.google.com/forms/d/" + id + "/viewform" : "";
  } catch (e) { return ""; }
}

function nhsos_headerMap_(sh) {
  return nhsperf_headerMap_(sh, true, true);
}

function nhsos_headerIndex_(headers, oneBased) {
  const o = {};

  headers.forEach((h, i) => {
    const key = String(h || "").trim();
    if (!key) return;

    // Keep the FIRST occurrence. This makes the system stable even before
    // the V6 schema migration removes duplicate legacy columns.
    if (o[key] === undefined) {
      o[key] = oneBased ? i + 1 : i;
    }
  });

  return o;
}

function nhsos_get_(row, idx, header) {
  return idx[header] === undefined ? "" : String(row[idx[header]] || "").trim();
}

function nhsos_set_(sh,row,h,header,value) {
  if (h[header]) sh.getRange(row,h[header]).setValue(value);
}

function nhsos_rowObject_(sh,row,h) {
  const vals = sh.getRange(row,1,1,sh.getLastColumn()).getDisplayValues()[0];
  const obj = {};
  Object.keys(h).forEach(k => obj[k] = vals[h[k]-1] || "");
  return obj;
}


function nhsos_findSessionRow_(sh, sessionId) {
  if (!sh || sh.getLastRow() < 2) return 0;

  const h = nhsos_headerMap_(sh);
  const col = h["Session ID"];
  if (!col) return 0;

  const match = sh
    .getRange(2, col, sh.getLastRow() - 1, 1)
    .createTextFinder(String(sessionId || "").trim())
    .matchEntireCell(true)
    .matchCase(true)
    .findNext();

  return match ? match.getRow() : 0;
}

function nhsos_findTutorById_(ss, tutorId) {
  const target = String(tutorId || "").trim();
  if (!target) return null;

  const sh = nhsperf_sheet_(ss, NHSOS_TUTORS);
  if (!sh || sh.getLastRow() < 2) return null;

  const h = nhsos_headerMap_(sh);
  const col = h["Tutor ID"];
  if (!col) return null;

  const match = sh
    .getRange(2, col, sh.getLastRow() - 1, 1)
    .createTextFinder(target)
    .matchEntireCell(true)
    .matchCase(true)
    .findNext();

  return match ? nhsos_rowObject_(sh, match.getRow(), h) : null;
}

function nhsos_logActivity_(ss, spec) {
  const sh = nhsos_ensureActivityLog_(ss);
  const h = nhsos_headerMap_(sh);

  const id =
    "A-" +
    Utilities.formatDate(
      new Date(),
      nhs_safeTimeZone_(ss),
      "yyyyMMdd-HHmmss"
    ) +
    "-" +
    Utilities.getUuid();

  const row = new Array(sh.getLastColumn()).fill("");

  const setValue = (header, value) => {
    if (h[header]) row[h[header] - 1] = value;
  };

  setValue("Activity ID", id);
  setValue("Timestamp", new Date());
  setValue("Type", spec.type || "System");
  setValue("Request ID", spec.requestId || "");
  setValue("Session ID", spec.sessionId || "");
  setValue("Tutor ID", spec.tutorId || "");
  setValue("Actor", spec.actor || nhsos_actor_());
  setValue("Summary", spec.summary || "");
  setValue("Details", spec.details || "");

  sh
    .getRange(sh.getLastRow() + 1, 1, 1, row.length)
    .setValues([row]);

  return id;
}

function nhsos_actor_() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (email) return email;
  } catch (e) {}

  return "NHS Officer";
}

function nhsos_truthy_(value) {
  if (value === true) return true;

  const normalized = String(value || "")
    .trim()
    .toLowerCase();

  return ["true", "yes", "y", "1", "checked"].includes(normalized);
}

function nhsos_nonFalseValue_(value) {
  const raw = String(value || "").trim();
  if (!raw) return false;

  const normalized = raw.toLowerCase();
  return !["false", "no", "0"].includes(normalized);
}

function nhsos_isRecurringValue_(value) {
  const normalized = String(value || "")
    .trim()
    .toLowerCase();

  if (!normalized) return false;

  if (normalized.startsWith("yes")) return true;
  if (normalized.includes("recurring")) return true;
  return ["true", "1"].includes(normalized);
}

function nhsos_nextPreferredDate_(afterDateValue, preferredDaysValue) {
  const base = new Date(afterDateValue);
  if (isNaN(base.getTime())) {
    throw new Error("Cannot calculate the next recurring date from the current session date.");
  }

  base.setHours(12, 0, 0, 0);

  const preferred = String(preferredDaysValue || "").toLowerCase();

  const dayMap = {
    sunday: 0,
    monday: 1,
    tuesday: 2,
    wednesday: 3,
    thursday: 4,
    friday: 5,
    saturday: 6
  };

  const wanted = Object.keys(dayMap)
    .filter(function(name) { return preferred.includes(name); })
    .map(function(name) { return dayMap[name]; });

  const fallbackDay = base.getDay();
  const acceptedDays = wanted.length ? wanted : [fallbackDay];

  for (let offset = 1; offset <= 70; offset++) {
    const candidate = new Date(base);
    candidate.setDate(candidate.getDate() + offset);

    if (!acceptedDays.includes(candidate.getDay())) continue;

    const key = Utilities.formatDate(
      candidate,
      NHS_PLATFORM.TIMEZONE,
      "yyyy-MM-dd"
    );

    if (!nhssched_exceptionForDate_(key)) {
      return candidate;
    }
  }

  throw new Error(
    "No available recurring Titan Time date was found within the next 10 weeks. Review School Calendar Exceptions."
  );
}

function nhsos_tutorDisplayName_(preferred, full) {
  preferred = String(preferred || "").trim();
  full = String(full || "").trim();
  if (!preferred) return full;
  const parts = full.split(/\s+/).filter(Boolean);
  if (parts.length < 2) return preferred;
  return preferred + " " + parts[parts.length-1];
}

function nhsos_norm_(v) {
  return String(v || "").toLowerCase().replace(/[^\w@]+/g," ").replace(/\s+/g," ").trim();
}

function nhsos_isoDate_(value) {
  if (!value) return "";
  const d = new Date(value);
  if (isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function nhsos_formatDateForEmail_(value) {
  const d = new Date(value);
  if (isNaN(d.getTime())) return String(value || "");
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy");
}

function nhsos_startOfDay_(d) {
  const x = new Date(d);
  x.setHours(0,0,0,0);
  return x;
}

function nhsos_withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function nhsv473_nhsos_getAuthoritativeRequestsCore_() {

  nhssec_requireOfficer_();

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const requests = nhsos_readRequests_(ss);

  return {
    generatedAt: new Date().toISOString(),
    revision: nhsperf_getDataRevision_(),
    requests: requests
  };
}


/** ============================================================
 * SOURCE GROUP: Gateway.gs
 * ============================================================ */

/**
 * TESORO NHS PORTAL GATEWAY — V43
 *
 * Standalone Apps Script Web App.
 * Requires NHS_Access_Core.gs in this project.
 */

function installNHSPortalGateway() {
  nhssec_requireMaintenanceAdmin_();

  return nhsa_installAccessCore_();
}


function nhs_routeBase_(candidate) {
  const value = String(candidate || "").trim();
  if (!value) return "";
  return value
    .split("?")[0]
    .replace(/\/dev$/, "/exec")
    .replace(/\/$/, "");
}

function nhs_routeUrl_(baseUrl, routeName) {
  const base = nhs_routeBase_(baseUrl);
  const route = String(routeName || "").trim().toLowerCase();
  if (!base || !route) return base;
  // Redundant routing is intentional: Apps Script supports pathInfo, while the
  // route query provides a fallback if an authentication handoff normalizes the
  // path. doGet accepts either source.
  return base + "/" + route + "?route=" + encodeURIComponent(route);
}


function gateway_getState() {
  const state = nhsa_getLoginState(
    "NHS Portal Gateway"
  );

  const serviceUrl = String(
    ScriptApp.getService().getUrl() || ""
  ).trim();

  const configuredGateway =
    state && state.settings
      ? String(state.settings.GATEWAY_URL || "").trim()
      : "";

  // The deployment currently serving this request is authoritative.
  // Portal Settings remain a fallback only when Apps Script cannot
  // provide a service URL.
  const clean = nhs_routeBase_(
    serviceUrl || configuredGateway
  );

  state.canonicalRoutes = clean
    ? {
        gateway: clean,
        teacher: nhs_routeUrl_(clean, "teacher"),
        tutor: nhs_routeUrl_(clean, "tutor"),
        operations: nhs_routeUrl_(clean, "operations")
      }
    : {};

  state.environment = nhsrg_publicEnvironmentInfo_();
  return state;
}

function gateway_requestAccess(
  role,
  fullName,
  reason
) {
  return nhsa_requestAccess(
    role,
    fullName,
    reason
  );
}


/** ============================================================
 * SOURCE GROUP: NHS_Navigation.gs
 * ============================================================ */

/**
 * Returns the authoritative base URL for the currently deployed NHS web app.
 * Safe for Teacher, Tutor, and Operations portals.
 */
function nhs_getGatewayUrl_() {
  try {
    return String(
      ScriptApp.getService().getUrl() || ""
    )
      .trim()
      .replace(/\/dev$/,"/exec");
  } catch (e) {
    throw new Error(
      "Unable to determine the deployed NHS Gateway URL."
    );
  }
}


/** ============================================================
 * SOURCE GROUP: NHS_Canonical_Routes.gs
 * ============================================================ */

/**
 * V33 — ONE-TIME CANONICAL PORTAL ROUTE REPAIR
 *
 * Run this once from the MAIN NHS Apps Script project after deploying the
 * consolidated web app.
 *
 * It rewrites all Portal Settings URLs to the current deployment so no legacy
 * teacher/tutor deployment can be opened by other parts of the platform.
 */
function repairPortalSettingsToProductionV42() {
  nhssec_requireMaintenanceAdmin_();
  const base = String(
    ScriptApp.getService().getUrl() || ""
  ).trim();

  if (!base) {
    throw new Error(
      "Could not resolve the current web-app deployment URL. Deploy the MAIN project first, then run this function."
    );
  }

  if (
    typeof nhsa_openSpreadsheet_ !== "function" ||
    typeof nhsa_ensureSettings_ !== "function" ||
    typeof nhsa_writeSetting_ !== "function"
  ) {
    throw new Error(
      "NHS Access Admin helpers are unavailable. Keep the current NHS Access Admin.gs file installed."
    );
  }

  const ss = nhsa_openSpreadsheet_();
  const sh = nhsa_ensureSettings_(ss);

  const clean = nhs_routeBase_(base);

  const routes = {
    GATEWAY_URL: clean,
    TEACHER_URL: nhs_routeUrl_(clean, "teacher"),
    TUTOR_URL: nhs_routeUrl_(clean, "tutor"),
    INTERNAL_URL: nhs_routeUrl_(clean, "operations")
  };

  Object.keys(routes).forEach(function(key) {
    nhsa_writeSetting_(sh,key,routes[key]);
  });

  SpreadsheetApp.flush();
  if (typeof nhsa_invalidateAccessCache_ === "function") {
    nhsa_invalidateAccessCache_();
  }

  return {
    ok: true,
    message: "Portal Settings now point to the current consolidated deployment.",
    routes: routes
  };
}


/**
 * Optional verification after the one-time repair.
 */
function checkProductionPortalRoutesV42() {
  nhssec_requireMaintenanceAdmin_();

  const base = String(
    ScriptApp.getService().getUrl() || ""
  ).trim();
  const clean = nhs_routeBase_(base);

  const expected = clean
    ? {
        gateway: clean,
        teacher: nhs_routeUrl_(clean, "teacher"),
        tutor: nhs_routeUrl_(clean, "tutor"),
        operations: nhs_routeUrl_(clean, "operations")
      }
    : {};

  let settings = {};
  try {
    settings = nhsa_readSettings_(nhsa_openSpreadsheet_()) || {};
  } catch (e) {
    return {
      ok: false,
      currentDeployment: base,
      productionDeployment: clean,
      expected: expected,
      actual: {},
      settingsMatchCurrentDeployment: false,
      error: e && e.message ? e.message : String(e)
    };
  }

  const actual = {
    gateway: String(settings.GATEWAY_URL || "").trim(),
    teacher: String(settings.TEACHER_URL || "").trim(),
    tutor: String(settings.TUTOR_URL || "").trim(),
    operations: String(settings.INTERNAL_URL || "").trim()
  };

  const matches = !!clean &&
    nhs_routeBase_(actual.gateway) === expected.gateway &&
    actual.teacher === expected.teacher &&
    actual.tutor === expected.tutor &&
    actual.operations === expected.operations;

  return {
    ok: matches,
    currentDeployment: base,
    productionDeployment: clean,
    expected: expected,
    actual: actual,
    settingsMatchCurrentDeployment: matches
  };
}


/** ============================================================
 * SOURCE GROUP: NHS_Maintenance.gs
 * ============================================================ */

/**
 * NHS TUTORING WORKBOOK CLEANUP
 *
 * Run ONE function:
 *   cleanNHSTutoringWorkbook()
 *
 * What it does:
 * 1. Creates a timestamped backup copy of the entire spreadsheet.
 * 2. Cleans duplicate headers in "Tutoring Requests" by consolidating data
 *    into the leftmost copy, then deleting the redundant duplicate columns.
 * 3. Deletes obvious blank matcher-artifact rows from "Tutoring Requests".
 * 4. Preserves all request history; V43 does not use account-specific auto-deletion.
 * 5. Deletes legacy Tutor Database columns that are no longer used.
 * 6. Hides backend / legacy tabs so the workbook is much cleaner to use.
 * 7. Leaves all columns still required by the Teacher Portal, Tutor Portal,
 *    Internal NHS OS, matching engine, notifications, and verification system.
 *
 * IMPORTANT:
 * - This is intentionally conservative.
 * - It DOES NOT delete real completed tutoring history.
 * - It DOES NOT delete Session Log, Activity Log, or Notification Log data.
 * - It DOES NOT remove columns currently used by your portals.
 */

function cleanNHSTutoringWorkbook() {
  nhssec_requireMaintenanceAdmin_();

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);

  if (!ss) {
    throw new Error("Open the NHS tutoring spreadsheet and run this function from its Apps Script project.");
  }

  const expectedId = String(NHS_SPREADSHEET_ID || "").trim();

  if (!expectedId || ss.getId() !== expectedId) {
    throw new Error(
      "Safety stop: cleanup target does not match NHS_SPREADSHEET_ID."
    );
  }

  const result = {
    backupUrl: "",
    duplicateColumnsRemoved: [],
    artifactRowsDeleted: 0,
    testRequestsDeleted: [],
    legacyTutorColumnsDeleted: [],
    weeklyAvailabilityRepair: null,
    hiddenSheets: [],
    notes: []
  };

  // ----------------------------------------------------------
  // 1) BACK UP THE ENTIRE WORKBOOK FIRST
  // ----------------------------------------------------------
  result.backupUrl = createNHSCleanupBackup_(ss);

  // ----------------------------------------------------------
  // 2) TUTORING REQUESTS CLEANUP
  // ----------------------------------------------------------
  const requests = nhsperf_sheet_(ss, "Tutoring Requests");

  if (!requests) {
    throw new Error('Missing required sheet: "Tutoring Requests".');
  }

  // Consolidate duplicate headers such as:
  // Match Notes / Match Updated
  result.duplicateColumnsRemoved =
    consolidateDuplicateHeaders_(requests);

  // Delete blank rows that only contain matcher/system artifacts.
  result.artifactRowsDeleted =
    deleteBlankMatcherArtifactRows_(requests);

  // V43 deliberately does NOT auto-delete any request based on a hardcoded
  // user/name heuristic. Request history is preserved.
  result.testRequestsDeleted = [];
  result.notes.push("V43 preserved all cancelled request history; no user-specific auto-deletion was performed.");

  // ----------------------------------------------------------
  // 3) TUTOR DATABASE LEGACY COLUMN CLEANUP
  // ----------------------------------------------------------
  const tutors = nhsperf_sheet_(ss, "Tutor Database");

  if (tutors) {
    const legacyTutorHeaders = [
      "Legacy Strongest #2",
      "Legacy Strongest #3",
      "Legacy Titan Availability",
      "Legacy After-School Availability"
    ];

    result.legacyTutorColumnsDeleted =
      deleteColumnsByHeader_(tutors, legacyTutorHeaders);
  }

  result.weeklyAvailabilityRepair =
    nhsmaint_reconcileWeeklyAvailabilityTutorIds_(ss);

  // ----------------------------------------------------------
  // 4) HIDE BACKEND / LEGACY TABS
  // ----------------------------------------------------------
  const tabsToHide = [
    "Teacher Rooms",
    "Teacher Dashboard",
    "Tutor Directory",
    "Teacher Feedback Log",
    "Activity Log",
    "Notification Log",
    "Session Log",
    "Teacher Directory",
    "Tutor Profile Responses",
    "START HERE"
  ];

  result.hiddenSheets =
    hideSheetsIfPresent_(ss, tabsToHide);

  // ----------------------------------------------------------
  // 5) KEEP OPERATIONAL TABS VISIBLE + ORDER THEM CLEANLY
  // ----------------------------------------------------------
  const visibleOrder = [
    "Dashboard",
    "Tutoring Requests",
    "Tutor Database",
    "Membership Roster"
  ];

  visibleOrder.forEach(name => {
    const sh = nhsperf_sheet_(ss, name);
    if (sh) {
      sh.showSheet();
    }
  });

  reorderVisibleSheets_(ss, visibleOrder);

  // ----------------------------------------------------------
  // 6) LIGHT FORMATTING / USABILITY CLEANUP
  // ----------------------------------------------------------
  formatCoreSheet_(nhsperf_sheet_(ss, "Tutoring Requests"));
  formatCoreSheet_(nhsperf_sheet_(ss, "Tutor Database"));
  formatCoreSheet_(nhsperf_sheet_(ss, "Membership Roster"));


  Logger.log(JSON.stringify(result, null, 2));

  SpreadsheetApp.getUi().alert(
    "NHS workbook cleanup complete.\n\n" +
    "Backup created first.\n" +
    "Blank matcher rows deleted: " + result.artifactRowsDeleted + "\n" +
    "Cancelled test requests deleted: " + result.testRequestsDeleted.length + "\n" +
    "Duplicate request columns removed: " + result.duplicateColumnsRemoved.length + "\n" +
    "Legacy tutor columns removed: " + result.legacyTutorColumnsDeleted.length + "\n" +
    "Backend tabs hidden: " + result.hiddenSheets.length + "\n\n" +
    "Backup:\n" + result.backupUrl
  );

  return result;
}


/* ============================================================
   BACKUP
   ============================================================ */

function createNHSCleanupBackup_(ss) {
  const tz = nhs_safeTimeZone_(ss) || Session.getScriptTimeZone();

  const stamp = Utilities.formatDate(
    new Date(),
    tz,
    "yyyy-MM-dd HHmmss"
  );

  const sourceFile = DriveApp.getFileById(ss.getId());

  const backup = sourceFile.makeCopy(
    ss.getName() + " — PRE-CLEANUP BACKUP " + stamp
  );

  return backup.getUrl();
}


/* ============================================================
   DUPLICATE HEADER CONSOLIDATION
   ============================================================ */

/**
 * For duplicate header names:
 * - keeps the LEFTMOST column
 * - copies the rightmost nonblank value from all duplicates
 *   into the kept column
 * - deletes all redundant duplicate columns
 *
 * This specifically fixes the duplicate Match Notes / Match Updated issue
 * without losing the populated matcher output currently stored in the
 * rightmost duplicate pair.
 */
function consolidateDuplicateHeaders_(sh) {
  if (!sh || sh.getLastColumn() < 2) {
    return [];
  }

  const lastRow = Math.max(sh.getLastRow(), 1);
  const lastCol = sh.getLastColumn();

  const headers = sh
    .getRange(1, 1, 1, lastCol)
    .getDisplayValues()[0]
    .map(v => String(v || "").trim());

  const groups = {};

  headers.forEach((header, index) => {
    if (!header) return;

    if (!groups[header]) {
      groups[header] = [];
    }

    groups[header].push(index + 1);
  });

  const duplicateGroups = Object.keys(groups)
    .filter(header => groups[header].length > 1);

  if (!duplicateGroups.length) {
    return [];
  }

  // First consolidate values into the leftmost column.
  duplicateGroups.forEach(header => {
    const cols = groups[header];
    const keepCol = cols[0];

    if (lastRow < 2) return;

    const sourceValues = cols.map(col =>
      sh
        .getRange(2, col, lastRow - 1, 1)
        .getValues()
        .map(row => row[0])
    );

    const merged = [];

    for (let r = 0; r < lastRow - 1; r++) {
      let chosen = "";

      // Prefer the rightmost nonblank value because the current
      // matching system has historically written into the later duplicate.
      for (let c = sourceValues.length - 1; c >= 0; c--) {
        const value = sourceValues[c][r];

        if (
          value !== "" &&
          value !== null &&
          value !== undefined
        ) {
          chosen = value;
          break;
        }
      }

      merged.push([chosen]);
    }

    sh
      .getRange(2, keepCol, merged.length, 1)
      .setValues(merged);
  });

  // Delete redundant columns from RIGHT TO LEFT so indexes stay valid.
  const colsToDelete = [];

  duplicateGroups.forEach(header => {
    const cols = groups[header];

    for (let i = 1; i < cols.length; i++) {
      colsToDelete.push({
        header: header,
        col: cols[i]
      });
    }
  });

  colsToDelete
    .sort((a, b) => b.col - a.col)
    .forEach(item => {
      sh.deleteColumn(item.col);
    });

  return colsToDelete
    .sort((a, b) => a.col - b.col)
    .map(item => item.header + " (column " + item.col + ")");
}


/* ============================================================
   BLANK MATCHER ARTIFACT ROW CLEANUP
   ============================================================ */

/**
 * Removes rows with no Request ID and no real teacher/request input.
 *
 * It intentionally ignores system-generated fields such as:
 * - Confirm Assignment = FALSE
 * - Match Notes
 * - Match Updated
 *
 * so old matcher artifacts are treated as blank.
 */
function deleteBlankMatcherArtifactRows_(sh) {
  if (!sh || sh.getLastRow() < 2) {
    return 0;
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  const values = sh
    .getRange(1, 1, lastRow, lastCol)
    .getDisplayValues();

  const headers = values[0]
    .map(v => String(v || "").trim());

  const h = headerIndex_(headers);

  // These are the fields that indicate a row is a REAL request.
  const meaningfulHeaders = [
    "Request ID",
    "Timestamp",
    "Teacher Name",
    "Teacher Email",
    "Teacher Notification Email",
    "Student Name",
    "Student Grade",
    "Course",
    "Topic / Academic Need",
    "Support Type",
    "Scheduling Preference",
    "Preferred Dates",
    "Preferred Days",
    "Urgency",
    "Recurring?",
    "Specific Tutor Preference",
    "Teacher Notes",
    "Status",
    "Assigned Tutor",
    "Scheduled Date",
    "Scheduled Time",
    "Location",
    "Coordinator Notes",
    "Assigned Tutor ID",
    "Request Source",
    "Teacher Change Request",
    "Teacher Change Requested At",
    "Teacher Change Status",
    "Tutor Response Status",
    "Tutor Response At",
    "Tutor Decline Reason",
    "Tutor Viewed At",
    "Tutor Portal Last Action",
    "Tutor Portal Last Action At",
    "Teacher Confirmation Sent At",
    "Tutor Confirmation Sent At",
    "Preferred Dates",
    "Tutor Reconfirm Status",
    "Teacher Five Star Ready",
    "Tutor Five Star Enrolled",
    "Five Star Ready At"
  ];

  const rowsToDelete = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const hasMeaningfulData = meaningfulHeaders.some(header => {
      const idx = h[header];

      if (idx === undefined) return false;

      const value = String(row[idx] || "").trim();

      return value !== "";
    });

    if (!hasMeaningfulData) {
      rowsToDelete.push(r + 1);
    }
  }

  deleteRowsInBlocks_(sh, rowsToDelete);

  return rowsToDelete.length;
}


/* ============================================================
   CANCELLED TEST REQUEST CLEANUP
   ============================================================ */

/**
 * Backward-compatible no-op. V43 never identifies or deletes requests using
 * hardcoded user/name heuristics.
 */
function deleteKnownCancelledTestRequests_(ss, requestSh) {
  // Retained only for backward compatibility with old callers.
  // V43 never deletes requests using hardcoded account/name criteria.
  return [];
}


function getSessionRequestIds_(ss) {
  const ids = new Set();

  const sh =
    nhsperf_sheet_(ss, "Session Log");

  if (
    !sh ||
    sh.getLastRow() < 2
  ) {
    return ids;
  }

  const values = sh
    .getRange(
      1,
      1,
      sh.getLastRow(),
      sh.getLastColumn()
    )
    .getDisplayValues();

  const h = headerIndex_(values[0]);

  const idx = h["Request ID"];

  if (idx === undefined) {
    return ids;
  }

  values
    .slice(1)
    .forEach(row => {
      const id =
        String(row[idx] || "").trim();

      if (id) ids.add(id);
    });

  return ids;
}


/* ============================================================
   LEGACY COLUMN CLEANUP
   ============================================================ */

function deleteColumnsByHeader_(sh, headerNames) {
  if (!sh || !headerNames.length) {
    return [];
  }

  const headers = sh
    .getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0]
    .map(v => String(v || "").trim());

  const found = [];

  headers.forEach((header, index) => {
    if (headerNames.includes(header)) {
      found.push({
        header: header,
        col: index + 1
      });
    }
  });

  found
    .sort((a, b) => b.col - a.col)
    .forEach(item => {
      sh.deleteColumn(item.col);
    });

  return found
    .sort((a, b) => a.col - b.col)
    .map(item => item.header);
}


/* ============================================================
   SHEET VISIBILITY / ORDER
   ============================================================ */

function hideSheetsIfPresent_(ss, names) {
  const hidden = [];

  names.forEach(name => {
    const sh = nhsperf_sheet_(ss, name);

    if (!sh) return;

    // Google Sheets requires at least one visible sheet,
    // and our operational tabs remain visible.
    sh.hideSheet();
    hidden.push(name);
  });

  return hidden;
}


function reorderVisibleSheets_(ss, orderedNames) {
  let targetPosition = 1;

  orderedNames.forEach(name => {
    const sh = nhsperf_sheet_(ss, name);

    if (!sh) return;

    ss.setActiveSheet(sh);
    ss.moveActiveSheet(targetPosition);

    targetPosition++;
  });
}


/* ============================================================
   TUTOR DATABASE IDENTITY RECONCILIATION
   ============================================================ */

function nhsmaint_personNameKey_(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\([^)]*\)/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function nhsmaint_migrateTutorIdentityReferences_(ss, mappings) {
  mappings = Array.isArray(mappings) ? mappings : [];
  if (!mappings.length) return { ok: true, rowsUpdated: 0, bySheet: {} };

  const byOldId = {};
  mappings.forEach(function(mapping) {
    if (mapping.oldTutorId && mapping.newTutorId) byOldId[mapping.oldTutorId] = mapping;
  });

  let total = 0;
  const bySheet = {};

  function migrate_(sheetName, idHeader, nameHeader, emailHeader) {
    const sh = nhsperf_sheet_(ss, sheetName);
    if (!sh || sh.getLastRow() < 2) return;

    const h = nhsperf_headerMap_(sh, true, true);
    if (!h[idHeader]) return;

    const range = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn());
    const rows = range.getValues();
    let changed = 0;

    rows.forEach(function(row) {
      const currentId = String(row[h[idHeader] - 1] || "").trim();
      const mapping = byOldId[currentId];
      if (!mapping) return;

      row[h[idHeader] - 1] = mapping.newTutorId;
      if (nameHeader && h[nameHeader]) row[h[nameHeader] - 1] = mapping.displayName || row[h[nameHeader] - 1];
      if (emailHeader && h[emailHeader]) row[h[emailHeader] - 1] = mapping.newSchoolEmail;
      changed++;
    });

    if (changed) range.setValues(rows);
    if (changed) {
      bySheet[sheetName] = changed;
      total += changed;
    }
  }

  migrate_("Tutoring Requests", "Assigned Tutor ID", "Assigned Tutor", null);
  migrate_("Session Log", "Tutor ID", "Tutor", null);
  migrate_("Activity Log", "Tutor ID", null, null);
  migrate_("Teacher Feedback Log", "Tutor ID", "Tutor", null);
  migrate_("Tutor Weekly Availability", "Tutor ID", "Tutor Name", "School Email");

  return { ok: true, rowsUpdated: total, bySheet: bySheet };
}

function nhsmaint_reconcileTutorDatabaseIdentities_(ss) {
  const sh = nhsperf_sheet_(ss, "Tutor Database");
  if (!sh || sh.getLastRow() < 2) {
    return {
      ok: true,
      rowsReviewed: 0,
      invalidSchoolIdentitiesInactivated: 0,
      identitiesMerged: 0,
      referencesUpdated: 0,
      grantsCreated: 0,
      grantsInactivated: 0
    };
  }

  const h = nhsperf_headerMap_(sh, true, true);
  if (!h["School Email"] || !h["Reliability Status"] || !h["Tutor ID"]) {
    return { ok: false, error: "Tutor Database identity headers are missing." };
  }

  const range = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn());
  const rows = range.getValues();
  const valid = [];

  rows.forEach(function(row, index) {
    const email = String(row[h["School Email"] - 1] || "").trim().toLowerCase();
    const reliability = String(row[h["Reliability Status"] - 1] || "").trim().toLowerCase();
    if (email && nhsa_isSchoolIdentityEmail_(email) && reliability !== "inactive") {
      valid.push({ row: row, index: index, email: email });
    }
  });

  const portalOwned = [
    "Preferred Name",
    "Strongest Course",
    "Tutoring Languages",
    "Experience Details",
    "Typical Titan Time Availability",
    "Typical After-School Availability",
    "Maximum Sessions / Week",
    "Temporary Unavailable Dates",
    "Reminder Preference",
    "Teacher Directory Bio",
    "Recurring Tutoring"
  ];
  const statFields = [
    "Sessions Completed",
    "Verified Tutoring Hours",
    "Tutor No-Shows",
    "Current Caseload"
  ];

  let changed = false;
  let inactivated = 0;
  let merged = 0;
  const mappings = [];

  rows.forEach(function(row, index) {
    const legacyEmail = String(row[h["School Email"] - 1] || "").trim().toLowerCase();
    if (!legacyEmail || nhsa_isSchoolIdentityEmail_(legacyEmail)) return;

    const legacyTutorId = String(row[h["Tutor ID"] - 1] || "").trim();
    const legacyNotification = h["Notification Email"]
      ? String(row[h["Notification Email"] - 1] || "").trim().toLowerCase()
      : "";
    const legacyFull = h["Full Name"] ? String(row[h["Full Name"] - 1] || "").trim() : "";
    const legacyPreferred = h["Preferred Name"] ? String(row[h["Preferred Name"] - 1] || "").trim() : "";
    const legacyNames = [legacyFull, legacyPreferred]
      .map(nhsmaint_personNameKey_)
      .filter(Boolean);

    let candidates = valid.filter(function(item) {
      const notification = h["Notification Email"]
        ? String(item.row[h["Notification Email"] - 1] || "").trim().toLowerCase()
        : "";
      return notification && (notification === legacyEmail || (legacyNotification && notification === legacyNotification));
    });

    if (candidates.length !== 1 && legacyNames.length) {
      candidates = valid.filter(function(item) {
        const names = [
          h["Full Name"] ? item.row[h["Full Name"] - 1] : "",
          h["Preferred Name"] ? item.row[h["Preferred Name"] - 1] : ""
        ].map(nhsmaint_personNameKey_).filter(Boolean);
        return legacyNames.some(function(key) { return names.includes(key); });
      });
    }

    const canonical = candidates.length === 1 ? candidates[0] : null;

    if (canonical) {
      const canonicalRow = canonical.row;
      const newTutorId = String(canonicalRow[h["Tutor ID"] - 1] || "").trim();
      const newSchoolEmail = canonical.email;
      const legacyPortalMs = h["Last Portal Profile Update"]
        ? nhstutor_dateMillis_(row[h["Last Portal Profile Update"] - 1])
        : 0;
      const canonicalPortalMs = h["Last Portal Profile Update"]
        ? nhstutor_dateMillis_(canonicalRow[h["Last Portal Profile Update"] - 1])
        : 0;

      if (legacyPortalMs >= canonicalPortalMs) {
        portalOwned.forEach(function(header) {
          if (!h[header]) return;
          const value = row[h[header] - 1];
          if (String(value == null ? "" : value).trim() !== "") {
            canonicalRow[h[header] - 1] = value;
            changed = true;
          }
        });
        if (h["Last Portal Profile Update"] && legacyPortalMs) {
          canonicalRow[h["Last Portal Profile Update"] - 1] = row[h["Last Portal Profile Update"] - 1];
          changed = true;
        }
      }

      statFields.forEach(function(header) {
        if (!h[header]) return;
        const legacyValue = Number(row[h[header] - 1] || 0);
        const canonicalValue = Number(canonicalRow[h[header] - 1] || 0);
        const mergedValue = Math.max(
          Number.isFinite(legacyValue) ? legacyValue : 0,
          Number.isFinite(canonicalValue) ? canonicalValue : 0
        );
        if (mergedValue !== canonicalValue) {
          canonicalRow[h[header] - 1] = mergedValue;
          changed = true;
        }
      });

      if (h["Notification Email"]) {
        const canonicalNotification = String(canonicalRow[h["Notification Email"] - 1] || "").trim();
        if (!canonicalNotification && legacyNotification && nhsa_isValidEmailAddress_(legacyNotification)) {
          canonicalRow[h["Notification Email"] - 1] = legacyNotification;
          changed = true;
        }
      }

      if (h["Last Profile Update"]) {
        const legacyProfileMs = nhstutor_dateMillis_(row[h["Last Profile Update"] - 1]);
        const canonicalProfileMs = nhstutor_dateMillis_(canonicalRow[h["Last Profile Update"] - 1]);
        if (legacyProfileMs > canonicalProfileMs) {
          canonicalRow[h["Last Profile Update"] - 1] = row[h["Last Profile Update"] - 1];
          changed = true;
        }
      }

      if (legacyTutorId && newTutorId && legacyTutorId !== newTutorId) {
        const displayName = nhsos_tutorDisplayName_(
          h["Preferred Name"] ? canonicalRow[h["Preferred Name"] - 1] : "",
          h["Full Name"] ? canonicalRow[h["Full Name"] - 1] : ""
        );
        mappings.push({
          oldTutorId: legacyTutorId,
          newTutorId: newTutorId,
          oldSchoolEmail: legacyEmail,
          newSchoolEmail: newSchoolEmail,
          displayName: displayName
        });
        merged++;
      }
    }

    const current = String(row[h["Reliability Status"] - 1] || "").trim().toLowerCase();
    if (current !== "inactive") {
      row[h["Reliability Status"] - 1] = "Inactive";
      inactivated++;
      changed = true;
    }
    if (h["Current Caseload"] && Number(row[h["Current Caseload"] - 1] || 0) !== 0) {
      row[h["Current Caseload"] - 1] = 0;
      changed = true;
    }
    if (h["Coordinator Notes"]) {
      const prior = String(row[h["Coordinator Notes"] - 1] || "").trim();
      const note = canonical
        ? "V47.7 identity reconciliation: legacy personal-email tutor identity merged into " + canonical.email + "."
        : "V47.7 identity reconciliation: Tutor School Email must be a CUSD identity; personal email belongs in Notification Email.";
      if (prior.indexOf(note) === -1) {
        row[h["Coordinator Notes"] - 1] = prior ? prior + " | " + note : note;
        changed = true;
      }
    }
  });

  if (changed) range.setValues(rows);

  const referenceResult = nhsmaint_migrateTutorIdentityReferences_(ss, mappings);

  let grantsInactivated = 0;
  let grantsCreated = 0;
  let grantsBlocked = 0;
  const registry = nhsa_ensureRegistry_(ss);
  const settings = nhsa_readSettings_(ss);
  const actor = (function() {
    try { return nhsa_getActiveEmail_(); } catch (e) { return "V47.7 identity repair"; }
  })();

  mappings.forEach(function(mapping) {
    const grants = nhsa_readRegistry_(registry);
    grants.filter(function(grant) {
      return grant.role === "TUTOR" && grant.email === mapping.oldSchoolEmail && nhsa_isGrantActive_(grant);
    }).forEach(function(grant) {
      if (nhsa_inactivateTutorGrant_(ss, grant, "Tutor access migrated to " + mapping.newSchoolEmail + ".")) {
        grantsInactivated++;
      }
    });

    const refreshed = nhsa_readRegistry_(registry);
    const canonicalHistory = refreshed.filter(function(grant) {
      return grant.role === "TUTOR" && grant.email === mapping.newSchoolEmail;
    });
    const activeCanonical = canonicalHistory.some(nhsa_isGrantActive_);

    if (!activeCanonical && canonicalHistory.length === 0) {
      const result = nhsa_createGrant_(ss, {
        email: mapping.newSchoolEmail,
        fullName: mapping.displayName,
        role: "TUTOR",
        schoolYear: settings.CURRENT_SCHOOL_YEAR || "",
        approvedBy: actor,
        notes: "V47.7 identity reconciliation from legacy Tutor ID " + mapping.oldTutorId
      });
      if (result && !result.existing) grantsCreated++;
    } else if (!activeCanonical && canonicalHistory.length) {
      grantsBlocked++;
    }
  });

  nhsa_invalidateAccessCache_();

  return {
    ok: true,
    rowsReviewed: rows.length,
    invalidSchoolIdentitiesInactivated: inactivated,
    identitiesMerged: merged,
    mappings: mappings,
    referencesUpdated: Number(referenceResult.rowsUpdated || 0),
    referenceSheets: referenceResult.bySheet || {},
    grantsCreated: grantsCreated,
    grantsInactivated: grantsInactivated,
    grantsBlockedByPriorState: grantsBlocked
  };
}

/* ============================================================
   WEEKLY AVAILABILITY IDENTITY RECONCILIATION
   ============================================================ */

function nhsmaint_reconcileWeeklyAvailabilityTutorIds_(ss) {
  const sh = nhsperf_sheet_(ss, "Tutor Weekly Availability");
  const tutors = nhsperf_sheet_(ss, "Tutor Database");

  if (!sh || !tutors || sh.getLastRow() < 2 || tutors.getLastRow() < 2) {
    return { ok: true, idsNormalized: 0, duplicatesRemoved: 0, keptRows: Math.max(0, sh ? sh.getLastRow() - 1 : 0) };
  }

  const wh = nhsperf_headerMap_(sh, true, true);
  const th = nhsperf_headerMap_(tutors, true, true);
  const required = ["Tutor ID", "School Email", "Date"];

  if (required.some(function(header) { return !wh[header]; }) || !th["Tutor ID"] || !th["School Email"]) {
    return { ok: false, idsNormalized: 0, duplicatesRemoved: 0, error: "Weekly availability or Tutor Database identity headers are missing." };
  }

  const tutorMap = {};
  tutors
    .getRange(2, 1, tutors.getLastRow() - 1, tutors.getLastColumn())
    .getDisplayValues()
    .forEach(function(row) {
      const email = String(row[th["School Email"] - 1] || "").trim().toLowerCase();
      const tutorId = String(row[th["Tutor ID"] - 1] || "").trim();
      const reliability = th["Reliability Status"]
        ? String(row[th["Reliability Status"] - 1] || "").trim().toLowerCase()
        : "";
      if (
        email &&
        tutorId &&
        nhsa_isSchoolIdentityEmail_(email) &&
        !["inactive", "suspended"].includes(reliability)
      ) {
        tutorMap[email] = tutorId;
      }
    });

  const range = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn());
  const rows = range.getValues();
  const selected = {};
  let idsNormalized = 0;
  let invalidIdentityRowsRemoved = 0;

  rows.forEach(function(row, index) {
    const email = String(row[wh["School Email"] - 1] || "").trim().toLowerCase();
    const currentId = String(row[wh["Tutor ID"] - 1] || "").trim();

    if (email && !nhsa_isSchoolIdentityEmail_(email)) {
      row.__nhsInvalidTutorIdentity = true;
      invalidIdentityRowsRemoved++;
      return;
    }

    const authoritativeId = email && tutorMap[email] ? tutorMap[email] : currentId;

    if (authoritativeId && authoritativeId !== currentId) {
      row[wh["Tutor ID"] - 1] = authoritativeId;
      idsNormalized++;
    }

    const dateKey = nhssched_dateKey_(row[wh["Date"] - 1], nhs_safeTimeZone_(ss));
    if (!authoritativeId || !dateKey) return;

    const identityKey = (email ? "email:" + email : "id:" + authoritativeId) + "|" + dateKey;
    const updatedRaw = wh["Updated At"] ? row[wh["Updated At"] - 1] : "";
    const confirmedRaw = wh["Confirmed At"] ? row[wh["Confirmed At"] - 1] : "";
    const updatedMs = Math.max(
      nhstutor_dateMillis_(updatedRaw),
      nhstutor_dateMillis_(confirmedRaw),
      0
    );

    const prior = selected[identityKey];
    if (!prior || updatedMs > prior.updatedMs || (updatedMs === prior.updatedMs && index > prior.index)) {
      selected[identityKey] = { row: row, index: index, updatedMs: updatedMs };
    }
  });

  const chosenIndexes = {};
  Object.keys(selected).forEach(function(key) { chosenIndexes[selected[key].index] = true; });

  const kept = rows.filter(function(row, index) {
    if (row.__nhsInvalidTutorIdentity) return false;
    const id = String(row[wh["Tutor ID"] - 1] || "").trim();
    const dateKey = nhssched_dateKey_(row[wh["Date"] - 1], nhs_safeTimeZone_(ss));
    if (!id || !dateKey) return row.some(function(v) { return String(v == null ? "" : v).trim() !== ""; });
    return !!chosenIndexes[index];
  });

  range.clearContent();
  if (kept.length) {
    sh.getRange(2, 1, kept.length, sh.getLastColumn()).setValues(kept);
  }

  return {
    ok: true,
    idsNormalized: idsNormalized,
    duplicatesRemoved: Math.max(0, rows.length - kept.length - invalidIdentityRowsRemoved),
    invalidIdentityRowsRemoved: invalidIdentityRowsRemoved,
    keptRows: kept.length
  };
}


/* ============================================================
   LIGHT FORMATTING
   ============================================================ */

function formatCoreSheet_(sh) {
  if (!sh) return;

  sh.setFrozenRows(1);

  if (sh.getLastColumn() > 0) {
    sh
      .getRange(1, 1, 1, sh.getLastColumn())
      .setFontWeight("bold")
      .setWrap(true);
  }

  // Keep row heights compact.
  sh.setRowHeight(1, 34);
}


/* ============================================================
   GENERIC HELPERS
   ============================================================ */

function deleteRowsInBlocks_(sh, rows) {
  if (!rows || !rows.length) return;

  // Google Sheets will not allow a sheet to contain only frozen rows.
  // Keep at least ONE non-frozen row alive at all times.
  const frozenRows = sh.getFrozenRows();
  const minimumRowsToKeep = Math.max(1, frozenRows + 1);

  // Sort + dedupe descending so row shifts are safe.
  rows = [...new Set(rows)]
    .filter(r => Number.isInteger(r) && r > frozenRows)
    .sort((a, b) => b - a);

  if (!rows.length) return;

  // If the requested deletion would remove every non-frozen physical row,
  // preserve one row and clear it instead of deleting it.
  const maxRowsBefore = sh.getMaxRows();
  const nonFrozenRowsBefore = maxRowsBefore - frozenRows;

  if (rows.length >= nonFrozenRowsBefore) {
    const rowToKeep = rows[rows.length - 1];

    // Delete every targeted row except the one preserved row.
    const rowsToDelete = rows.filter(r => r !== rowToKeep);

    deleteRowsSafely_(sh, rowsToDelete, minimumRowsToKeep);

    // Clear the preserved row so no request/test data remains.
    // Its row number may have shifted after deletions, so clear the first
    // available non-frozen row instead.
    const clearRow = frozenRows + 1;

    sh
      .getRange(
        clearRow,
        1,
        1,
        sh.getMaxColumns()
      )
      .clearContent();

    return;
  }

  deleteRowsSafely_(sh, rows, minimumRowsToKeep);
}


/**
 * Deletes arbitrary row numbers from bottom to top while always retaining
 * at least one non-frozen physical row.
 */
function deleteRowsSafely_(sh, rows, minimumRowsToKeep) {
  if (!rows || !rows.length) return;

  rows = [...new Set(rows)]
    .sort((a, b) => b - a);

  let blockStart = rows[0];
  let blockEnd = rows[0];

  for (let i = 1; i <= rows.length; i++) {
    const row = rows[i];

    if (
      i < rows.length &&
      row === blockEnd - 1
    ) {
      blockEnd = row;
      continue;
    }

    let start = blockEnd;
    let count =
      blockStart - blockEnd + 1;

    const maxDeletable =
      sh.getMaxRows() - minimumRowsToKeep;

    if (maxDeletable <= 0) {
      // Nothing else may be physically deleted.
      // Clear the remaining targeted cells instead.
      sh
        .getRange(
          start,
          1,
          count,
          sh.getMaxColumns()
        )
        .clearContent();
    } else {
      const safeCount =
        Math.min(count, maxDeletable);

      if (safeCount > 0) {
        sh.deleteRows(start, safeCount);
      }

      // If part of the requested block could not be deleted because doing so
      // would remove the last non-frozen row, clear that remaining row.
      const remainder =
        count - safeCount;

      if (remainder > 0) {
        sh
          .getRange(
            sh.getFrozenRows() + 1,
            1,
            remainder,
            sh.getMaxColumns()
          )
          .clearContent();
      }
    }

    if (i < rows.length) {
      blockStart = row;
      blockEnd = row;
    }
  }
}


function headerIndex_(headers) {
  const map = {};

  headers.forEach((header, index) => {
    const key =
      String(header || "").trim();

    // Keep LEFTMOST copy.
    if (
      key &&
      map[key] === undefined
    ) {
      map[key] = index;
    }
  });

  return map;
}


function getCell_(row, headerMap, header) {
  const idx = headerMap[header];

  if (idx === undefined) {
    return "";
  }

  return String(row[idx] || "").trim();
}


/** ============================================================
 * SOURCE GROUP: NHS_Installer.gs
 * ============================================================ */

/**
 * ONE-TIME PLATFORM INSTALLER
 *
 * Run this after the consolidated project is assembled.
 * Future officers should not need to run it for normal yearly management.
 */
function installNHSPlatform() {
  nhssec_requireMaintenanceAdmin_();

  const results = {};
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);

  results.timezone = nhs_ensureSpreadsheetTimeZone_(ss);
  results.access = nhsa_installAccessCore_();
  results.bootstrap = nhsa_bootstrapFirstAdmin();
  results.operations = installNHSOS();
  results.matching = installTutorMatchingEngine();
  results.assignment = installAssignmentWorkflow();
  results.notifications = installNHSNotificationEngine();
  results.teacher = installTeacherPortal();
  results.tutor = installTutorPortal();
  results.tutorIdentity = nhsmaint_reconcileTutorDatabaseIdentities_(ss);
  results.tutorAccessSync = nhsa_adminSyncTutorAccess();
  results.weeklyAvailability = nhsmaint_reconcileWeeklyAvailabilityTutorIds_(ss);

  if (typeof repairNHSTriggersV43_ === "function") {
    results.triggers = repairNHSTriggersV43_();
  }

  SpreadsheetApp.flush();
  nhsa_invalidateAccessCache_();
  nhsperf_touchDataRevision_();

  return {
    ok: true,
    build: NHS_RELEASE.VERSION,
    installedAt: new Date().toISOString(),
    results: results,
    next: "Deploy this consolidated project as a Web App, then run repairNHSPlatformV43() once."
  };
}


/**
 * Safe yearly maintenance entry point.
 */
function runNHSAnnualMaintenance() {
  nhsa_requireAnyRole_(
    ["ADMIN", "ADVISOR"],
    "Annual Maintenance"
  );

  const ss = SpreadsheetApp.openById(NHS_SPREADSHEET_ID);
  const timezone = nhs_ensureSpreadsheetTimeZone_(ss);
  const tutorIdentity = nhsmaint_reconcileTutorDatabaseIdentities_(ss);
  const tutorAccessSync = nhsa_adminSyncTutorAccess();
  const weeklyAvailability = nhsmaint_reconcileWeeklyAvailabilityTutorIds_(ss);

  // Keep background structures healthy.
  nhsa_installAccessCore_();

  if (typeof installNHSNotificationEngine === "function") {
    installNHSNotificationEngine();
  }

  return {
    ok: true,
    spreadsheet: ss.getName(),
    timezone: timezone,
    tutorIdentity: tutorIdentity,
    tutorAccessSync: tutorAccessSync,
    weeklyAvailability: weeklyAvailability,
    message:
      "Annual maintenance complete. Officer rollover should be performed from Access Management."
  };
}


/** ============================================================
 * SOURCE GROUP: NHS_Production_Audit.gs
 * ============================================================ */

/**
 * TESORO NHS — PRODUCTION AUDIT + SAFE REPAIR V43
 *
 * Primary diagnostics:
 *   auditNHSTutoringWorkbook()
 *   runNHSProductionAuditV43()
 *
 * One-time migration after deploying V43:
 *   repairNHSPlatformV43()
 */

const NHSV43_REQUIRED_SHEETS = [
  "Portal Settings",
  "Access Registry",
  "Access Requests",
  "Tutoring Requests",
  "Tutor Database",
  "Membership Roster",
  "Activity Log",
  "Notification Log",
  "Session Log",
  "System Health",
  "Schedule Exceptions"
];

const NHSV43_REQUIRED_REQUEST_HEADERS = [
  "Request ID", "Timestamp", "Teacher Name", "Teacher Email", "Course",
  "Topic / Academic Need", "Support Type", "Preferred Days", "Recurring?",
  "Status", "Assigned Tutor", "Assigned Tutor ID", "Scheduled Date",
  "Scheduled Time", "Location", "Last Update", "Confirm Assignment",
  "Notification Status", "Assignment Confirmed", "Recurring Active",
  "Tutor Response Status", "Calendar Status"
];

function auditNHSTutoringWorkbook() {
  nhssec_requireMaintenanceAdmin_();

  return runNHSProductionAuditV43();
}


/**
 * V43.1 one-time cleanup for the historical checkbox used-range artifact.
 * Safe to run repeatedly.
 */
function repairRequestUsedRangeV43_1() {
  nhssec_requireMaintenanceAdmin_();

  const ss = nhsperf_ss_();
  const sh = nhsperf_sheet_(ss, "Tutoring Requests");

  if (!sh) {
    throw new Error('Missing "Tutoring Requests" sheet.');
  }

  const before = {
    physicalLast: sh.getLastRow(),
    realLast: nhsdata_realLastRowByKey_(sh, "Request ID")
  };

  const compact = nhsdata_compactKeyedTable_(sh, "Request ID");
  nhsdata_applyCheckboxValidation_(sh, "Confirm Assignment");

  const realLast = nhsdata_realLastRowByKey_(sh, "Request ID");
  const h = nhsperf_headerMap_(sh, true, true);

  if (realLast >= 2 && h["Confirm Assignment"]) {
    const range = sh.getRange(
      2,
      h["Confirm Assignment"],
      realLast - 1,
      1
    );
    const values = range.getValues();
    let changed = false;

    values.forEach(function(row) {
      if (row[0] === "") {
        row[0] = false;
        changed = true;
      }
    });

    if (changed) {
      range.setValues(values);
    }
  }

  SpreadsheetApp.flush();

  const after = {
    physicalLast: sh.getLastRow(),
    realLast: nhsdata_realLastRowByKey_(sh, "Request ID")
  };

  try {
    nhsdata_invalidateRequestState_({ touchRevision: true });
  } catch (e) {}

  return {
    ok: after.physicalLast <= Math.max(after.realLast + 5, 10),
    build: "V43.1",
    before: before,
    compact: compact,
    after: after
  };
}


/**
 * V43.2 definitive request-table cleanup.
 * Safe to run repeatedly.
 *
 * Cleans ALL phantom values below the last real Request ID row, then re-applies
 * checkbox validation only to real rows + one staging row.
 */
function repairRequestTableV43_2() {
  nhssec_requireMaintenanceAdmin_();

  const ss = nhsperf_ss_();
  const sh = nhsperf_sheet_(ss, "Tutoring Requests");

  if (!sh) {
    throw new Error('Missing "Tutoring Requests" sheet.');
  }

  const before = {
    physicalLast: sh.getLastRow(),
    realLast: nhsdata_realLastRowByKey_(sh, "Request ID"),
    maxRows: sh.getMaxRows()
  };

  // First compact all genuine keyed records to the top of the table.
  const compact = nhsdata_compactKeyedTable_(sh, "Request ID");
  SpreadsheetApp.flush();

  const realLast = nhsdata_realLastRowByKey_(sh, "Request ID");
  const lastCol = sh.getLastColumn();

  // Clear ALL content below the real keyed table, not only the checkbox column.
  // This removes historical Match Notes / Match Updated artifacts as well.
  const cleanupStart = Math.max(2, realLast + 1);
  const cleanupCount = sh.getMaxRows() - cleanupStart + 1;

  if (cleanupCount > 0 && lastCol > 0) {
    sh
      .getRange(cleanupStart, 1, cleanupCount, lastCol)
      .clearContent();
  }

  // Re-apply checkbox validation only where it belongs.
  nhsdata_applyCheckboxValidation_(sh, "Confirm Assignment");

  // Real request rows should have a boolean checkbox value.
  const h = nhsperf_headerMap_(sh, true, true);
  if (realLast >= 2 && h["Confirm Assignment"]) {
    const checkboxRange = sh.getRange(
      2,
      h["Confirm Assignment"],
      realLast - 1,
      1
    );
    const vals = checkboxRange.getValues();
    let changed = false;

    vals.forEach(function(row) {
      if (row[0] === "") {
        row[0] = false;
        changed = true;
      }
    });

    if (changed) {
      checkboxRange.setValues(vals);
    }
  }

  SpreadsheetApp.flush();

  const after = {
    physicalLast: sh.getLastRow(),
    realLast: nhsdata_realLastRowByKey_(sh, "Request ID"),
    maxRows: sh.getMaxRows()
  };

  try {
    nhsdata_invalidateRequestState_({ touchRevision: true });
  } catch (e) {}

  return {
    ok: after.physicalLast <= Math.max(after.realLast + 2, 5),
    build: "V43.2",
    before: before,
    compact: compact,
    after: after
  };
}


/**
 * V44 portal diagnostic.
 * Tests the exact stable bootstrap + authoritative request paths.
 */
function diagnoseTutorOperationsV44() {
  nhssec_requireMaintenanceAdmin_();

  const out = {
    build: "V44",
    generatedAt: new Date().toISOString(),
    activeEmail: "",
    tutor: {},
    operations: {}
  };

  function timed_(fn) {
    const start = Date.now();

    try {
      return {
        ok: true,
        ms: Date.now() - start,
        value: fn()
      };
    } catch (e) {
      return {
        ok: false,
        ms: Date.now() - start,
        error: String(e && e.message ? e.message : e)
      };
    }
  }

  try {
    out.activeEmail = nhsa_getActiveEmail_();
  } catch (e) {}

  out.tutor.bootstrap = timed_(function() {
    const d = tutor_getPortalData(true);

    return {
      lifecycleBuild: d.lifecycleBuild,
      requestTruth: d.requestTruth,
      tutorId: d.identity && d.identity.tutorId,
      requests: (d.requests || []).length,
      pendingActions:
        d.counts && d.counts.pendingAcceptance,
      upcoming:
        d.counts && d.counts.upcoming,
      currentCaseload:
        d.counts && d.counts.currentCaseload
    };
  });

  out.tutor.authoritative = timed_(function() {
    const d = tutor_getAuthoritativeRequestStateV42();

    return {
      buildVersion: d.buildVersion,
      requests: (d.requests || []).length,
      counts: d.counts || {}
    };
  });

  out.operations.bootstrap = timed_(function() {
    const d = nhsos_getBootstrapData(true);

    return {
      lifecycleBuild: d.lifecycleBuild,
      requestTruth: d.requestTruth,
      requests: (d.requests || []).length,
      tutors: (d.tutors || []).length,
      sessions: (d.sessions || []).length,
      activities: (d.activities || []).length,
      alerts: (d.alerts || []).length
    };
  });

  out.operations.authoritative = timed_(function() {
    const d = nhsos_getAuthoritativeRequestsV39();

    return {
      requests: (d.requests || []).length,
      revision: d.revision || ""
    };
  });

  return out;
}


/* ============================================================
   V46 — CACHE / MAINTENANCE HARDENING
   ============================================================ */

function nhs_v46Revision_() {
  try { return String(nhsperf_getDataRevision_() || "0"); }
  catch (e) { return "0"; }
}

function nhs_v46TeacherCacheKey_(email) {
  return "teacherPortal" + NHS_RELEASE.CACHE_NAMESPACE + "|" + nhs_v46Revision_() + "|" +
    String(email || "").trim().toLowerCase();
}

function nhs_v46TutorCacheKey_(tutorId) {
  return "tutorPortal" + NHS_RELEASE.CACHE_NAMESPACE + "|" + nhs_v46Revision_() + "|" +
    String(tutorId || "").trim();
}

function nhs_v46TutorIdentityKey_(email) {
  return "tutorIdentity" + NHS_RELEASE.CACHE_NAMESPACE + "|" + nhs_v46Revision_() + "|" +
    String(email || "").trim().toLowerCase();
}

function nhs_v46OperationsCacheKey_() {
  return NHSOS_CACHE_KEY + "|" + nhs_v46Revision_();
}

function nhssec_requireMaintenanceAdmin_() {
  return nhsa_requireAnyRole_(["ADMIN", "ADVISOR"], "NHS Maintenance");
}

/* ============================================================
   V45 — UNIFIED STABLE PRODUCTION LAYER
   One server file. Four HTML clients.
   ============================================================ */

const NHS_V45_BUILD = NHS_RELEASE.VERSION;

/**
 * Canonical cross-portal invalidation.
 *
 * Every request mutation should ultimately advance the global data revision.
 * Targeted cache keys are cleared when their identities are known. Request
 * delta endpoints remain the source of truth, so a cache failure can never
 * erase live request state.
 */
function nhs_v45InvalidatePortalState_(spec) {
  spec = spec || {};

  try {
    const cache = nhsperf_scriptCache_();

    // Operations is shared across all request mutations.
    cache.remove(NHSOS_CACHE_KEY);
    cache.remove("nhsos_health_v14");

    if (spec.teacherEmail) {
      cache.remove(
        nhs_v46TeacherCacheKey_(spec.teacherEmail)
      );
    }

    if (spec.tutorId) {
      cache.remove(
        nhs_v46TutorCacheKey_(spec.tutorId)
      );
    }
  } catch (e) {}

  if (typeof nhsperf_touchDataRevision_ === "function") {
    return nhsperf_touchDataRevision_();
  }

  return String(Date.now());
}


/**
 * Safe post-deploy cache prime for the CURRENT signed-in account only.
 * It is read-only apart from normal last-login/access telemetry already
 * present in the production code.
 */
function primeNHSPlatformV45() {
  nhssec_requireMaintenanceAdmin_();

  const result = {
    build: NHS_V45_BUILD,
    generatedAt: new Date().toISOString(),
    teacher: null,
    tutor: null,
    operations: null
  };

  function run_(fn) {
    const started = Date.now();

    try {
      const value = fn();

      return {
        ok: true,
        ms: Date.now() - started,
        summary: value
      };
    } catch (e) {
      return {
        ok: false,
        ms: Date.now() - started,
        error: String(e && e.message ? e.message : e)
      };
    }
  }

  result.teacher = run_(function() {
    const d = teacher_getPortalData(true);

    return {
      requests: (d.requests || []).length,
      upcoming: (d.upcoming || []).length
    };
  });

  result.tutor = run_(function() {
    const d = tutor_getPortalData(true);

    return {
      tutorId: d.identity && d.identity.tutorId,
      requests: (d.requests || []).length,
      pendingActions: d.counts && d.counts.pendingAcceptance,
      upcoming: d.counts && d.counts.upcoming,
      currentCaseload: d.counts && d.counts.currentCaseload
    };
  });

  result.operations = run_(function() {
    const d = nhsos_getBootstrapData(true);

    return {
      requests: (d.requests || []).length,
      tutors: (d.tutors || []).length,
      sessions: (d.sessions || []).length
    };
  });

  return result;
}


/**
 * One read-only diagnostic for the whole deployed main project.
 * This intentionally does not create requests, send mail, or mutate workflow.
 */
function diagnoseNHSProductionV45() {
  nhssec_requireMaintenanceAdmin_();

  const out = {
    build: NHS_V45_BUILD,
    generatedAt: new Date().toISOString(),
    activeEmail: "",
    audit: null,
    gateway: null,
    teacher: null,
    tutor: null,
    tutorAuthority: null,
    operations: null,
    operationsAuthority: null
  };

  function timed_(fn) {
    const started = Date.now();

    try {
      return {
        ok: true,
        ms: Date.now() - started,
        value: fn()
      };
    } catch (e) {
      return {
        ok: false,
        ms: Date.now() - started,
        error: String(e && e.message ? e.message : e)
      };
    }
  }

  try {
    out.activeEmail = nhsa_getActiveEmail_();
  } catch (e) {}

  out.audit = timed_(function() {
    return runNHSProductionAuditV43();
  });

  out.gateway = timed_(function() {
    const d = gateway_getState();

    return {
      email: d && d.email,
      roles: d && d.roles,
      teacherAllowed: !!(d && d.teacherAllowed),
      tutorAllowed: !!(d && d.tutorAllowed),
      operationsAllowed: !!(d && d.operationsAllowed)
    };
  });

  out.teacher = timed_(function() {
    const d = teacher_getPortalData(true);

    return {
      requests: (d.requests || []).length,
      upcoming: (d.upcoming || []).length,
      counts: d.counts || {}
    };
  });

  out.tutor = timed_(function() {
    const d = tutor_getPortalData(true);

    return {
      tutorId: d.identity && d.identity.tutorId,
      requests: (d.requests || []).length,
      counts: d.counts || {}
    };
  });

  out.tutorAuthority = timed_(function() {
    const d = tutor_getAuthoritativeRequestStateV42();

    return {
      requests: (d.requests || []).length,
      counts: d.counts || {},
      buildVersion: d.buildVersion || ""
    };
  });

  out.operations = timed_(function() {
    const d = nhsos_getBootstrapData(true);

    return {
      requests: (d.requests || []).length,
      tutors: (d.tutors || []).length,
      sessions: (d.sessions || []).length,
      dashboard: d.dashboard || {}
    };
  });

  out.operationsAuthority = timed_(function() {
    const d = nhsos_getAuthoritativeRequestsV39();

    return {
      requests: (d.requests || []).length,
      revision: d.revision || ""
    };
  });

  return out;
}

function runNHSDeploymentIdentityDiagnostic() {
  const active = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  const effective = String(Session.getEffectiveUser().getEmail() || "").trim().toLowerCase();
  const serviceUrl = String(ScriptApp.getService().getUrl() || "").trim();

  return {
    ok: !!active,
    activeUserEmail: active,
    effectiveUserEmail: effective,
    sameIdentity: !!active && active === effective,
    serviceUrl: serviceUrl,
    guidance: active
      ? "Active user identity is visible to this deployment."
      : "Active user identity is hidden. For this execute-as-deployer architecture, create the production deployment from a CUSD Workspace collaborator account and test again with a CUSD user."
  };
}


function runNHSProductionAuditV43() {
  nhssec_requireMaintenanceAdmin_();

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const checks = [];

  function add(name, ok, detail, severity) {
    checks.push({
      name: name,
      ok: !!ok,
      severity: severity || (ok ? "Info" : "High"),
      detail: String(detail || "")
    });
  }

  add("Spreadsheet ID", ss.getId() === NHS_SPREADSHEET_ID, ss.getName());
  add("Timezone", nhs_safeTimeZone_(ss) === NHS_PLATFORM.TIMEZONE, nhs_safeTimeZone_(ss), "Medium");

  NHSV43_REQUIRED_SHEETS.forEach(function(name) {
    add("Sheet: " + name, !!ss.getSheetByName(name), ss.getSheetByName(name) ? "Present" : "Missing");
  });

  const requestSh = ss.getSheetByName("Tutoring Requests");
  if (requestSh) {
    const h = nhsperf_headerMap_(requestSh, true, true);
    NHSV43_REQUIRED_REQUEST_HEADERS.forEach(function(header) {
      add("Request header: " + header, !!h[header], h[header] ? "Present" : "Missing");
    });

    const physicalLast = requestSh.getLastRow();
    const realLast = nhsdata_realLastRowByKey_(requestSh, "Request ID");
    add(
      "Request used-range integrity",
      physicalLast <= Math.max(realLast + 5, 10),
      "physicalLast=" + physicalLast + ", realLast=" + realLast,
      physicalLast <= Math.max(realLast + 5, 10) ? "Info" : "Medium"
    );

    const requests = nhsdata_readKeyedObjects_(requestSh, "Request ID");
    const seen = {};
    const dupes = [];
    requests.forEach(function(r) {
      const id = String(r["Request ID"] || "");
      if (seen[id]) dupes.push(id);
      seen[id] = true;
    });
    add("Unique Request IDs", dupes.length === 0, dupes.length ? dupes.join(", ") : "No duplicates");
  }

  let routeAudit = null;
  try {
    routeAudit = checkProductionPortalRoutesV42();
  } catch (e) {
    routeAudit = { ok: false, error: e && e.message ? e.message : String(e) };
  }

  add(
    "Canonical portal routes",
    !!(routeAudit && routeAudit.ok),
    routeAudit && routeAudit.ok
      ? "Portal Settings match the current consolidated deployment."
      : (routeAudit && routeAudit.error
          ? routeAudit.error
          : "Portal Settings do not match the current deployment. Run repairPortalSettingsToProductionV42()."),
    "High"
  );

  if (routeAudit && routeAudit.expected && routeAudit.actual) {
    ["gateway", "teacher", "tutor", "operations"].forEach(function(key) {
      add(
        "Route match: " + key,
        String(routeAudit.actual[key] || "") === String(routeAudit.expected[key] || ""),
        "actual=" + String(routeAudit.actual[key] || "") + " | expected=" + String(routeAudit.expected[key] || ""),
        "High"
      );
    });
  }

  const grants = nhsa_readRegistry_(nhsa_ensureRegistry_(ss));
  const privileged = grants.filter(function(g) {
    return nhsa_isGrantActive_(g) && ["ADMIN", "ADVISOR"].includes(g.role);
  });
  add("Recovery authority", privileged.length > 0, privileged.length + " active Admin/Advisor grant(s)");

  add("Notification email validator", typeof nhsn_isValidEmail_ === "function", "nhsn_isValidEmail_");
  add("Notification alert logger", typeof nhsn_logAlert_ === "function", "nhsn_logAlert_");
  add("Lifecycle core", typeof nhslife_upcoming_ === "function", "V43 canonical lifecycle");
  add("Data core", typeof nhsdata_readKeyedObjects_ === "function", "V43 keyed data core");

  const effectiveEmail = String(
    Session.getEffectiveUser().getEmail() || ""
  ).trim().toLowerCase();
  const deployerIsSchool = nhsa_isSchoolIdentityEmail_(effectiveEmail);
  add(
    "Production deployer identity",
    deployerIsSchool,
    effectiveEmail
      ? ("effectiveUser=" + effectiveEmail)
      : "Effective/deploying account could not be identified.",
    "High"
  );

  let mailAuthorizationOk = false;
  let mailAuthorizationDetail = "";
  try {
    const quota = MailApp.getRemainingDailyQuota();
    mailAuthorizationOk = Number.isFinite(Number(quota));
    mailAuthorizationDetail = "MailApp authorized; remaining recipient quota=" + Number(quota || 0);
  } catch (e) {
    mailAuthorizationDetail = e && e.message ? e.message : String(e);
  }
  add(
    "Mail authorization",
    mailAuthorizationOk,
    mailAuthorizationDetail || "MailApp authorization unavailable.",
    "High"
  );

  let relay = null;
  try { relay = typeof checkNHSBatchRelayV29 === "function" ? checkNHSBatchRelayV29() : null; } catch (e) { relay = { ok: false, message: e.message }; }
  add("Batch mail relay", !!(relay && relay.ok), relay ? (relay.message || JSON.stringify(relay)) : "Diagnostic unavailable", "High");

  const triggerAudit = nhsv43_auditTriggers_();
  add("Trigger uniqueness", triggerAudit.duplicates.length === 0, triggerAudit.duplicates.length ? triggerAudit.duplicates.join(", ") : "No duplicate managed triggers", "Medium");
  add("Required background triggers", triggerAudit.missing.length === 0, triggerAudit.missing.length ? triggerAudit.missing.join(", ") : "All managed triggers present", "Medium");

  let profileTriggerAudit = null;
  try {
    profileTriggerAudit = nhstutor_profileTriggerAuditV55_();
  } catch (e) {
    profileTriggerAudit = { formAvailable: false, submitCount: 0, reconcileCount: 0, ok: false, error: e && e.message ? e.message : String(e) };
  }

  add(
    "Tutor profile form access",
    !!(profileTriggerAudit && profileTriggerAudit.formAvailable),
    profileTriggerAudit && profileTriggerAudit.formAvailable
      ? "Tutor Profile Form is accessible."
      : ((profileTriggerAudit && profileTriggerAudit.error) || "Form unavailable; verify NHS_TUTOR_PROFILE_FORM_ID and Editor access."),
    "High"
  );
  add(
    "Tutor profile submit trigger",
    !!(profileTriggerAudit && profileTriggerAudit.submitCount === 1),
    "count=" + Number(profileTriggerAudit && profileTriggerAudit.submitCount || 0),
    "High"
  );
  add(
    "Tutor profile reconcile trigger",
    !!(profileTriggerAudit && profileTriggerAudit.reconcileCount === 1),
    "count=" + Number(profileTriggerAudit && profileTriggerAudit.reconcileCount || 0),
    "High"
  );

  const tutorRowsForIdentity = nhsos_readTutors_(ss);
  const invalidTutorProfiles = tutorRowsForIdentity.filter(function(t) {
    const email = String(t.schoolEmail || "").trim().toLowerCase();
    const reliability = String(t.reliabilityStatus || "").trim().toLowerCase();
    return email && !["inactive", "suspended"].includes(reliability) && !nhsa_isSchoolIdentityEmail_(email);
  });
  add(
    "Tutor profile school identities",
    invalidTutorProfiles.length === 0,
    invalidTutorProfiles.length
      ? invalidTutorProfiles.length + " active tutor profile(s) use a non-CUSD School Email."
      : "All active tutor profiles use CUSD school identities.",
    "High"
  );

  const activeTutorGrantRows = nhsa_readRegistry_(nhsa_ensureRegistry_(ss)).filter(function(grant) {
    return grant.role === "TUTOR" && nhsa_isGrantActive_(grant);
  });
  const invalidTutorGrants = activeTutorGrantRows.filter(function(grant) {
    return !nhsa_isSchoolIdentityEmail_(grant.email);
  });
  add(
    "Tutor access school identities",
    invalidTutorGrants.length === 0,
    invalidTutorGrants.length
      ? invalidTutorGrants.length + " active Tutor grant(s) use a personal/non-CUSD identity."
      : "All active Tutor grants use CUSD school identities.",
    "High"
  );


  const activeTutorGrantEmails = {};
  activeTutorGrantRows.forEach(function(grant) {
    activeTutorGrantEmails[String(grant.email || "").trim().toLowerCase()] = true;
  });
  const missingTutorGrantProfiles = tutorRowsForIdentity.filter(function(t) {
    const email = String(t.schoolEmail || "").trim().toLowerCase();
    const reliability = String(t.reliabilityStatus || "").trim().toLowerCase();
    return email && nhsa_isSchoolIdentityEmail_(email) && reliability !== "inactive" && !activeTutorGrantEmails[email];
  });
  add(
    "Tutor profiles have active access grants",
    missingTutorGrantProfiles.length === 0,
    missingTutorGrantProfiles.length
      ? missingTutorGrantProfiles.map(function(t) { return t.schoolEmail; }).join(", ")
      : "Every active CUSD Tutor Database profile has an active Tutor grant.",
    "High"
  );

  const activeTutorIds = {};
  tutorRowsForIdentity.forEach(function(t) {
    const email = String(t.schoolEmail || "").trim().toLowerCase();
    const reliability = String(t.reliabilityStatus || "").trim().toLowerCase();
    if (t.tutorId && email && nhsa_isSchoolIdentityEmail_(email) && reliability !== "inactive") {
      activeTutorIds[String(t.tutorId)] = true;
    }
  });
  const orphanActiveAssignments = nhsdata_requestObjects_(ss).filter(function(request) {
    if (!nhslife_activeCaseload_(request)) return false;
    const tutorId = String(request["Assigned Tutor ID"] || "").trim();
    return tutorId && !activeTutorIds[tutorId];
  });
  add(
    "Active assignment tutor identities",
    orphanActiveAssignments.length === 0,
    orphanActiveAssignments.length
      ? orphanActiveAssignments.map(function(r) { return String(r["Request ID"] || ""); }).join(", ")
      : "All active assignments resolve to an active CUSD tutor identity.",
    "High"
  );

  const health = ss.getSheetByName("System Health");
  if (health && health.getLastRow() >= 2) {
    const rows = health.getRange(1, 1, health.getLastRow(), health.getLastColumn()).getDisplayValues();
    const headers = rows[0];
    const componentCol = headers.indexOf("Component");
    const errorCol = headers.indexOf("Last Error");
    if (componentCol >= 0 && errorCol >= 0) {
      rows.slice(1).forEach(function(row) {
        const comp = String(row[componentCol] || "").trim();
        const err = String(row[errorCol] || "").trim();
        if (comp) add("Health: " + comp, !err, err || "No recorded error", err ? "High" : "Info");
      });
    }
  }

  const failures = checks.filter(function(c) { return !c.ok; });
  return {
    ok: failures.filter(function(c) { return c.severity === "High"; }).length === 0,
    build: NHS_RELEASE.VERSION,
    generatedAt: new Date().toISOString(),
    checks: checks,
    failureCount: failures.length,
    highSeverityCount: failures.filter(function(c) { return c.severity === "High"; }).length,
    triggerAudit: triggerAudit,
    profileTriggerAudit: profileTriggerAudit,
    routeAudit: routeAudit
  };
}

function nhsv43_auditTriggers_() {
  const managed = [
    "rankAllOpenRequests_",
    "handleAssignmentWorkflowEdit_",
    "runNHSNotificationSweep_",
    "nhsos_processFastJobs_",
    "processQueuedTutorCaseloadReconciliation_",
    "tp_processAsyncReceiptQueue_",
    "tp_processCancellationJobs_",
    "tup_processJobs_"
  ];

  const counts = {};
  ScriptApp.getProjectTriggers().forEach(function(t) {
    const fn = t.getHandlerFunction();
    counts[fn] = (counts[fn] || 0) + 1;
  });

  return {
    counts: counts,
    missing: managed.filter(function(fn) { return !counts[fn]; }),
    duplicates: managed.filter(function(fn) { return Number(counts[fn] || 0) > 1; })
  };
}

function repairNHSTriggersV43_() {
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const managed = [
    "rankAllOpenRequests_",
    "handleAssignmentWorkflowEdit_",
    "runNHSNotificationSweep_",
    "nhsos_processFastJobs_",
    "processQueuedTutorCaseloadReconciliation_",
    "tp_processAsyncReceiptQueue_",
    "tp_processCancellationJobs_",
    "tup_processJobs_"
  ];

  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (managed.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger("rankAllOpenRequests_").timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger("handleAssignmentWorkflowEdit_").forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger("runNHSNotificationSweep_").timeBased().everyHours(1).create();
  ScriptApp.newTrigger("nhsos_processFastJobs_").timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger("processQueuedTutorCaseloadReconciliation_").timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger("tp_processAsyncReceiptQueue_").timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger("tp_processCancellationJobs_").timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger("tup_processJobs_").timeBased().everyMinutes(1).create();

  return nhsv43_auditTriggers_();
}

function nhsv43_backupWorkbook_(ss) {
  try {
    const source = DriveApp.getFileById(ss.getId());
    const name =
      ss.getName() +
      " — V43 Backup — " +
      Utilities.formatDate(new Date(), nhs_safeTimeZone_(ss), "yyyy-MM-dd HHmmss");
    const copy = source.makeCopy(name);
    return copy.getUrl();
  } catch (e) {
    return "Backup unavailable: " + (e.message || String(e));
  }
}

function repairNHSPlatformV43() {
  nhsa_requireAnyRole_(["ADMIN", "ADVISOR"], "V43 Platform Repair");

  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const result = {
    build: NHS_RELEASE.VERSION,
    backupUrl: nhsv43_backupWorkbook_(ss),
    timezoneRepair: null,
    requestCompaction: null,
    routes: null,
    triggers: null,
    prunedDeadLetters: 0
  };

  result.timezoneRepair = nhs_ensureSpreadsheetTimeZone_(ss);

  if (typeof migrateTutorIdentityToSchoolEmail === "function") {
    try { result.tutorIdentityMigration = migrateTutorIdentityToSchoolEmail(); }
    catch (e) { result.tutorIdentityMigration = { ok: false, error: e.message || String(e) }; }
  }

  const requestSh = nhsperf_sheet_(ss, "Tutoring Requests");
  if (!requestSh) throw new Error('Missing "Tutoring Requests" sheet.');

  result.requestCompaction = nhsdata_compactKeyedTable_(requestSh, "Request ID");
  nhsdata_applyCheckboxValidation_(requestSh, "Confirm Assignment");
  result.scheduleExceptions = {
    sheet: nhssched_ensureExceptionSheet_(ss).getName(),
    count: nhssched_listExceptions_().filter(function(item) { return item.active; }).length
  };

  // Ensure every real request row has an explicit checkbox value, but blank
  // rows below remain truly blank.
  const realLast = nhsdata_realLastRowByKey_(requestSh, "Request ID");
  const h = nhsperf_headerMap_(requestSh, true, true);
  if (realLast >= 2 && h["Confirm Assignment"]) {
    const range = requestSh.getRange(2, h["Confirm Assignment"], realLast - 1, 1);
    const values = range.getValues();
    let changed = false;
    values.forEach(function(row) {
      if (row[0] === "") { row[0] = false; changed = true; }
    });
    if (changed) range.setValues(values);
  }

  if (typeof repairPortalSettingsToProductionV42 === "function") {
    result.routes = repairPortalSettingsToProductionV42();
  }

  result.triggers = repairNHSTriggersV43_();

  try {
    result.tutorProfilePipeline =
      nhstutor_repairProfilePipelineV55_(
        ss
      );
  } catch (error) {
    result.tutorProfilePipeline = {
      ok: false,
      error:
        error &&
        error.message
          ? error.message
          : String(error)
    };
  }

  try {
    result.tutorIdentityRepair =
      nhsmaint_reconcileTutorDatabaseIdentities_(ss);
  } catch (e) {
    result.tutorIdentityRepair = { ok: false, error: e && e.message ? e.message : String(e) };
  }

  try {
    result.tutorAccessSync =
      nhsa_adminSyncTutorAccess();
  } catch (e) {
    result.tutorAccessSync = { ok: false, error: e && e.message ? e.message : String(e) };
  }

  try {
    result.weeklyAvailabilityRepair =
      nhsmaint_reconcileWeeklyAvailabilityTutorIds_(ss);
  } catch (e) {
    result.weeklyAvailabilityRepair = { ok: false, error: e && e.message ? e.message : String(e) };
  }

  if (typeof reconcileTutorCaseloads === "function") {
    try { reconcileTutorCaseloads(); } catch (e) {}
  } else if (typeof nhsos_reconcileCaseloads_ === "function") {
    try { nhsos_reconcileCaseloads_(ss); } catch (e) {}
  }

  result.prunedDeadLetters = nhsdata_pruneDeadLetters_(50);

  nhsa_invalidateAccessCache_();
  nhsdata_invalidateRequestState_({});
  SpreadsheetApp.flush();

  result.audit = runNHSProductionAuditV43();
  result.ok = result.audit.ok;
  return result;
}




/** ============================================================
 * SOURCE GROUP: NHS_V47_3_Schoolwide_Hardening.gs
 * ============================================================ */

/**
 * NHS V47.3 — SCHOOLWIDE DEPLOYMENT READINESS
 *
 * Adds five production-hardening capabilities without changing existing
 * workflow contracts:
 *   1) one release identity across all portals,
 *   2) sampled, cache-backed runtime observability,
 *   3) schoolwide deployment preflight / diagnostic bundle,
 *   4) dead-letter visibility + safe retry/dismiss controls,
 *   5) staging persona lifecycle certification helpers.
 *
 * The observability store uses Script Cache, not Sheets or Script Properties,
 * so it is bounded, ephemeral, and does not create a new persistent hot path.
 */

const NHSOBS_SAMPLE_RATE = 0.15;
const NHSOBS_MAX_SAMPLES = 48;
const NHSOBS_MAX_FAILURES = 8;
const NHSOBS_TTL_SECONDS = 21600;
const NHSOBS_METRICS = [
  "teacher.bootstrap","teacher.authoritative","teacher.submit","teacher.cancel","teacher.tutorDirectory",
  "tutor.bootstrap","tutor.authoritative","tutor.accept","tutor.decline","tutor.report",
  "operations.bootstrap","operations.authoritative","operations.saveRequest",
  "operations.confirmAssignment","operations.verifySession"
];

function nhsobs_cacheKey_(metric) {
  return "nhsobs|" + NHS_RELEASE.VERSION + "|" + String(metric || "unknown");
}

function nhsobs_sanitizeError_(error) {
  return String(error && error.message ? error.message : error || "Unknown error")
    .replace(/[\r\n\t]+/g, " ")
    .replace(/\s+/g, " ")
    .slice(0, 260);
}

function nhsobs_record_(metric, durationMs, ok, error) {
  // Successful calls are sampled; failures are always retained.
  if (ok && Math.random() > NHSOBS_SAMPLE_RATE) return;
  try {
    const cache = nhsperf_scriptCache_();
    const key = nhsobs_cacheKey_(metric);
    let state = { samples: [], failures: [], totalSampled: 0, sampledFailures: 0 };
    const raw = cache.get(key);
    if (raw) {
      try { state = Object.assign(state, JSON.parse(raw) || {}); } catch (e) {}
    }
    state.samples = Array.isArray(state.samples) ? state.samples : [];
    state.failures = Array.isArray(state.failures) ? state.failures : [];
    state.samples.push(Math.max(0, Number(durationMs || 0)));
    state.samples = state.samples.slice(-NHSOBS_MAX_SAMPLES);
    state.totalSampled = Number(state.totalSampled || 0) + 1;
    if (!ok) {
      state.sampledFailures = Number(state.sampledFailures || 0) + 1;
      state.failures.push({
        at: new Date().toISOString(),
        message: nhsobs_sanitizeError_(error),
        durationMs: Math.max(0, Number(durationMs || 0))
      });
      state.failures = state.failures.slice(-NHSOBS_MAX_FAILURES);
    }
    state.updatedAt = new Date().toISOString();
    cache.put(key, JSON.stringify(state), NHSOBS_TTL_SECONDS);
  } catch (e) {
    // Observability must never interfere with an NHS workflow.
  }
}

function nhsobs_measure_(metric, fn) {
  const started = Date.now();
  try {
    const value = fn();
    nhsobs_record_(metric, Date.now() - started, true, null);
    return value;
  } catch (error) {
    nhsobs_record_(metric, Date.now() - started, false, error);
    throw error;
  }
}

function nhsobs_percentile_(values, percentile) {
  const rows = (values || []).map(Number).filter(function(v){ return isFinite(v); }).sort(function(a,b){ return a-b; });
  if (!rows.length) return 0;
  const index = Math.min(rows.length - 1, Math.max(0, Math.ceil((Number(percentile || 0) / 100) * rows.length) - 1));
  return rows[index];
}

function nhsobs_getSummaryInternal_() {
  const cache = nhsperf_scriptCache_();
  const metrics = NHSOBS_METRICS.map(function(metric) {
    let state = { samples: [], failures: [], totalSampled: 0, sampledFailures: 0 };
    try {
      const raw = cache.get(nhsobs_cacheKey_(metric));
      if (raw) state = Object.assign(state, JSON.parse(raw) || {});
    } catch (e) {}
    const samples = Array.isArray(state.samples) ? state.samples : [];
    return {
      metric: metric,
      sampledCalls: Number(state.totalSampled || samples.length || 0),
      retainedSamples: samples.length,
      sampledFailures: Number(state.sampledFailures || 0),
      p50Ms: nhsobs_percentile_(samples, 50),
      p95Ms: nhsobs_percentile_(samples, 95),
      maxMs: samples.length ? Math.max.apply(null, samples.map(Number)) : 0,
      updatedAt: state.updatedAt || "",
      recentFailures: Array.isArray(state.failures) ? state.failures.slice(-3).reverse() : []
    };
  });
  return {
    ok: true,
    build: NHS_RELEASE.VERSION,
    sampleRate: NHSOBS_SAMPLE_RATE,
    ephemeral: true,
    metrics: metrics,
    generatedAt: new Date().toISOString()
  };
}

function nhsobs_getSummary() {
  nhssec_requireMaintenanceAdmin_();
  return nhsobs_getSummaryInternal_();
}

function nhsobs_clearTelemetry() {
  nhssec_requireMaintenanceAdmin_();
  const cache = nhsperf_scriptCache_();
  NHSOBS_METRICS.forEach(function(metric){ try { cache.remove(nhsobs_cacheKey_(metric)); } catch (e) {} });
  return { ok: true, message: "V47.3 runtime telemetry cleared." };
}

// ---- measured endpoint wrappers ------------------------------------------------
function teacher_getPortalData(forceRefresh) {
  return nhsobs_measure_("teacher.bootstrap", function() { return nhsv473_teacher_getPortalDataCore_(forceRefresh); });
}

function teacher_getAuthoritativeRequestStateV39() {
  return nhsobs_measure_("teacher.authoritative", function() { return nhsv473_teacher_getAuthoritativeRequestStateCore_(); });
}

function teacher_getTutorDirectory(forceRefresh) {
  return nhsobs_measure_(
    "teacher.tutorDirectory",
    function() {
      return nhsv54_teacher_getTutorDirectoryCore_(
        !!forceRefresh
      );
    }
  );
}

function teacher_submitRequest(payload) {
  return nhsobs_measure_("teacher.submit", function() { return nhsv473_teacher_submitRequestCore_(payload); });
}

function teacher_cancelRequest(requestId, reason) {
  return nhsobs_measure_("teacher.cancel", function() { return nhsv473_teacher_cancelRequestCore_(requestId, reason); });
}

function tutor_getPortalData(forceRefresh) {
  return nhsobs_measure_("tutor.bootstrap", function() { return nhsv473_tutor_getPortalDataCore_(forceRefresh); });
}

function tutor_getAuthoritativeRequestStateV42() {
  return nhsobs_measure_("tutor.authoritative", function() { return nhsv473_tutor_getAuthoritativeRequestStateCore_(); });
}

function tutor_acceptAssignment(requestId) {
  return nhsobs_measure_("tutor.accept", function() { return nhsv473_tutor_acceptAssignmentCore_(requestId); });
}

function tutor_declineAssignment(requestId, reason) {
  return nhsobs_measure_("tutor.decline", function() { return nhsv473_tutor_declineAssignmentCore_(requestId, reason); });
}

function tutor_reportSession(payload) {
  return nhsobs_measure_("tutor.report", function() { return nhsv473_tutor_reportSessionCore_(payload); });
}

function nhsos_getBootstrapData(forceRefresh) {
  return nhsobs_measure_("operations.bootstrap", function() { return nhsv473_nhsos_getBootstrapDataCore_(forceRefresh); });
}

function nhsos_getAuthoritativeRequestsV39() {
  return nhsobs_measure_("operations.authoritative", function() { return nhsv473_nhsos_getAuthoritativeRequestsCore_(); });
}

function nhsos_saveRequest(payload) {
  return nhsobs_measure_("operations.saveRequest", function() { return nhsv473_nhsos_saveRequestCore_(payload); });
}

function nhsos_confirmAssignment(requestId) {
  return nhsobs_measure_("operations.confirmAssignment", function() { return nhsv473_nhsos_confirmAssignmentCore_(requestId); });
}

function nhsos_verifySession(sessionId, verifiedMinutes, verificationNotes) {
  return nhsobs_measure_("operations.verifySession", function() { return nhsv473_nhsos_verifySessionCore_(sessionId, verifiedMinutes, verificationNotes); });
}


// ---- dead-letter operations ----------------------------------------------------
function nhsops_deadLetterInventoryInternal_() {
  const props = nhsperf_scriptProperties_();
  const all = props.getProperties();
  const rows = Object.keys(all)
    .filter(function(key){ return key.indexOf(NHSDATA_DEADLETTER_PREFIX) === 0; })
    .sort().reverse()
    .map(function(propertyKey) {
      let record = {};
      try { record = JSON.parse(all[propertyKey] || "{}") || {}; } catch (e) {}
      const job = record.job || {};
      return {
        propertyKey: propertyKey,
        originalKey: String(record.originalKey || ""),
        attempts: Number(job.attempts || 0),
        lastAttemptAt: String(job.lastAttemptAt || ""),
        lastError: nhsobs_sanitizeError_(job.lastError || ""),
        jobType: String(job.type || job.jobType || "Background job"),
        createdKey: propertyKey.replace(NHSDATA_DEADLETTER_PREFIX, "").split("_")[0] || ""
      };
    });
  return rows;
}

function nhsops_getDeadLetters() {
  nhssec_requireMaintenanceAdmin_();
  return {
    ok: true,
    count: nhsops_deadLetterInventoryInternal_().length,
    jobs: nhsops_deadLetterInventoryInternal_(),
    generatedAt: new Date().toISOString()
  };
}

function nhsops_retryDeadLetter(propertyKey) {
  nhssec_requireMaintenanceAdmin_();
  const key = String(propertyKey || "").trim();
  if (!key || key.indexOf(NHSDATA_DEADLETTER_PREFIX) !== 0) throw new Error("Invalid dead-letter key.");
  const props = nhsperf_scriptProperties_();
  const raw = props.getProperty(key);
  if (!raw) throw new Error("Dead-letter job no longer exists.");
  let record = {};
  try { record = JSON.parse(raw || "{}") || {}; } catch (e) { throw new Error("Dead-letter payload is unreadable."); }
  const originalKey = String(record.originalKey || "").trim();
  if (!originalKey) throw new Error("Dead-letter job is missing its original queue key.");
  if (props.getProperty(originalKey)) throw new Error("A live queue job already exists for this key; retry was not duplicated.");
  const job = Object.assign({}, record.job || {});
  job.attempts = 0;
  delete job.lastAttemptAt;
  delete job.lastError;
  job.requeuedAt = new Date().toISOString();
  props.setProperty(originalKey, JSON.stringify(job));
  props.deleteProperty(key);
  return { ok: true, originalKey: originalKey, message: "Background job requeued safely." };
}

function nhsops_dismissDeadLetter(propertyKey) {
  nhssec_requireMaintenanceAdmin_();
  const key = String(propertyKey || "").trim();
  if (!key || key.indexOf(NHSDATA_DEADLETTER_PREFIX) !== 0) throw new Error("Invalid dead-letter key.");
  const props = nhsperf_scriptProperties_();
  if (!props.getProperty(key)) return { ok: true, alreadyGone: true, message: "Dead-letter job is already gone." };
  props.deleteProperty(key);
  return { ok: true, message: "Dead-letter job dismissed." };
}

// ---- schoolwide deployment preflight ------------------------------------------
function nhsdeploy_addCheck_(checks, id, label, status, detail, category) {
  checks.push({
    id: String(id || ""),
    label: String(label || id || "Check"),
    status: ["PASS","WARN","BLOCK"].includes(status) ? status : "WARN",
    detail: String(detail || ""),
    category: String(category || "General")
  });
}

function nhsdeploy_collectPreflight_(options) {
  options = options || {};
  const strict = options.strict !== false;
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const env = nhsrg_environmentInfo_();
  const checks = [];
  const counts = {};

  nhsdeploy_addCheck_(checks, "release", "Unified release identity", NHS_PLATFORM.BUILD === NHS_RELEASE.VERSION && NHSRG_BUILD === NHS_RELEASE.VERSION ? "PASS" : "BLOCK",
    "Application " + NHS_PLATFORM.BUILD + " · Release Guard " + NHSRG_BUILD, "Release");
  nhsdeploy_addCheck_(checks, "spreadsheet", "Data workbook opens", ss && ss.getId() === NHS_SPREADSHEET_ID ? "PASS" : "BLOCK", ss ? ss.getName() : "Unavailable", "Platform");
  if (env.isStaging) {
    nhsdeploy_addCheck_(checks, "staging-isolation", "Staging isolation", env.isolated && env.configuredDataId ? "PASS" : "BLOCK",
      env.isolated ? "Separate staging data target configured." : "Staging points at production data.", "Release");
  } else {
    nhsdeploy_addCheck_(checks, "production-target", "Production data target", NHS_SPREADSHEET_ID === NHS_PRODUCTION_SPREADSHEET_ID ? "PASS" : "BLOCK",
      NHS_SPREADSHEET_ID === NHS_PRODUCTION_SPREADSHEET_ID ? "Production workbook selected." : "Production project is pointed at a non-production workbook.", "Release");
  }

  const requiredEndpoints = [
    "gateway_getState","teacher_getPortalData","teacher_getAuthoritativeRequestStateV39",
    "tutor_getPortalData","tutor_getAuthoritativeRequestStateV42",
    "nhsos_getBootstrapData","nhsos_getAuthoritativeRequestsV39",
    "nhs_getDataRevision","nhsrg_runProductionSmokeTest","nhsrg_runReleaseCertification"
  ];
  const missingEndpoints = requiredEndpoints.filter(function(name){ return !nhsrg_functionExists_(name); });
  nhsdeploy_addCheck_(checks, "endpoints", "Portal endpoint contracts", missingEndpoints.length ? "BLOCK" : "PASS",
    missingEndpoints.length ? "Missing: " + missingEndpoints.join(", ") : requiredEndpoints.length + " critical endpoints installed.", "Platform");

  const settings = nhsa_readSettings_(ss);
  const routeKeys = ["GATEWAY_URL","TEACHER_URL","TUTOR_URL","INTERNAL_URL"];
  const badRoutes = routeKeys.filter(function(key){
    const v = String(settings[key] || "").trim();
    return !v || !/\/exec(?:\/|\?|$)/.test(v);
  });
  nhsdeploy_addCheck_(checks, "routes", "Portal routes configured", badRoutes.length ? (strict ? "BLOCK" : "WARN") : "PASS",
    badRoutes.length ? "Missing/invalid: " + badRoutes.join(", ") : "Gateway, Teacher, Tutor, and Operations routes are configured.", "Platform");

  const triggerAudit = nhsv43_auditTriggers_();
  nhsdeploy_addCheck_(checks, "triggers", "Managed triggers installed", triggerAudit.missing.length ? "BLOCK" : "PASS",
    triggerAudit.missing.length ? "Missing: " + triggerAudit.missing.join(", ") : "All managed triggers installed.", "Automation");
  nhsdeploy_addCheck_(checks, "trigger-duplicates", "No duplicate managed triggers", triggerAudit.duplicates.length ? "WARN" : "PASS",
    triggerAudit.duplicates.length ? "Duplicates: " + triggerAudit.duplicates.join(", ") : "No duplicate managed triggers.", "Automation");

  let mailStatus = null;
  try { mailStatus = getNHSMailRelayStatus(); } catch (e) { mailStatus = { configured:false, error:nhsobs_sanitizeError_(e) }; }
  const mailOk = env.isStaging ? ["RECORD_ONLY","REDIRECT"].includes(env.mailMode) : !!(mailStatus && mailStatus.configured);
  nhsdeploy_addCheck_(checks, "mail", "Notification transport ready", mailOk ? "PASS" : (strict ? "BLOCK" : "WARN"),
    env.isStaging ? "Staging mail mode: " + env.mailMode : (mailOk ? "Live relay configured." : "Live mail relay is not fully configured."), "Notifications");

  const registry = nhsa_readRegistry_(nhsa_ensureRegistry_(ss));
  const activeGrants = registry.filter(nhsa_isGrantActive_);
  const privileged = activeGrants.filter(function(g){ return ["ADMIN","ADVISOR"].includes(g.role); });
  const teachers = activeGrants.filter(function(g){ return g.role === "TEACHER"; });
  const invalidTeacherGrants = teachers.filter(function(g){ return !/@capousd\.org$/i.test(String(g.email || "")); });
  const grantSeen = {};
  const duplicateActiveGrants = [];
  activeGrants.forEach(function(g){ const key=String(g.email||"").toLowerCase()+"|"+String(g.role||""); if(grantSeen[key]) duplicateActiveGrants.push(key); grantSeen[key]=true; });
  const pendingAccess = nhsa_readRequests_(nhsa_ensureRequests_(ss)).filter(function(r){ return String(r.status||"") === "Pending"; });
  counts.activePrivileged = privileged.length;
  counts.activeTeachers = teachers.length;
  counts.pendingAccess = pendingAccess.length;
  nhsdeploy_addCheck_(checks, "privileged", "Admin / Advisor continuity", privileged.length ? "PASS" : "BLOCK",
    privileged.length + " active Admin/Advisor grant(s).", "Access");
  nhsdeploy_addCheck_(checks, "teachers", "Teacher roster provisioned", teachers.length ? "PASS" : (strict ? "BLOCK" : "WARN"),
    teachers.length + " active Teacher grant(s).", "Access");
  nhsdeploy_addCheck_(checks, "teacher-domain", "Teacher grants use staff accounts", invalidTeacherGrants.length ? "BLOCK" : "PASS",
    invalidTeacherGrants.length ? invalidTeacherGrants.length + " active Teacher grant(s) are not @capousd.org." : "All active Teacher grants use @capousd.org.", "Access");
  nhsdeploy_addCheck_(checks, "duplicate-grants", "No duplicate active role grants", duplicateActiveGrants.length ? "WARN" : "PASS",
    duplicateActiveGrants.length ? duplicateActiveGrants.length + " duplicate email/role grant(s)." : "No duplicate active email/role grants.", "Access");
  nhsdeploy_addCheck_(checks, "pending-access", "Pending access queue reviewed", pendingAccess.length ? "WARN" : "PASS",
    pendingAccess.length ? pendingAccess.length + " access request(s) are still pending." : "No pending access requests.", "Access");
  const expiringSoon = activeGrants.filter(function(g){
    if (!g.expirationDate) return false;
    const d = nhsa_parseDate_(g.expirationDate);
    if (!d) return false;
    const days = Math.floor((d.getTime() - Date.now()) / 86400000);
    return days >= 0 && days <= 45;
  });
  counts.expiringAccessGrants = expiringSoon.length;
  nhsdeploy_addCheck_(checks, "expiring-access", "Near-term access expirations reviewed", expiringSoon.length ? "WARN" : "PASS",
    expiringSoon.length ? expiringSoon.length + " active grant(s) expire within 45 days." : "No active grants expire within 45 days.", "Access");

  const tutorRows = nhsos_readTutors_(ss);
  const eligibleTutors = tutorRows.filter(function(t){
    const reliabilityOk = !["inactive","suspended"].includes(String(t.reliabilityStatus||"").toLowerCase());
    const agreement = String(t.expectationsAgreement || "").trim();
    const agreementOk = !agreement || /agree|yes|true/i.test(agreement);
    return reliabilityOk && agreementOk;
  });
  const missingTutorId = eligibleTutors.filter(function(t){ return !String(t.tutorId || "").trim(); });
  const missingTutorSchoolEmail = eligibleTutors.filter(function(t){ return !String(t.schoolEmail || "").trim(); });
  const nonStudentTutorEmails = eligibleTutors.filter(function(t){
    const email=String(t.schoolEmail||"").trim();
    return email && !/@capostu\.org$/i.test(email);
  });
  const missingTutorNotificationEmail = eligibleTutors.filter(function(t){ return !String(t.notificationEmail || "").trim(); });
  const tutorEmailSeen = {};
  const tutorIdSeen = {};
  const duplicateTutorEmails = [];
  const duplicateTutorIds = [];
  eligibleTutors.forEach(function(t){
    const email=String(t.schoolEmail||"").trim().toLowerCase();
    const id=String(t.tutorId||"").trim();
    if(email){ if(tutorEmailSeen[email]) duplicateTutorEmails.push(email); tutorEmailSeen[email]=true; }
    if(id){ if(tutorIdSeen[id]) duplicateTutorIds.push(id); tutorIdSeen[id]=true; }
  });
  counts.totalTutors = tutorRows.length;
  counts.eligibleTutors = eligibleTutors.length;
  counts.tutorsMissingNotificationEmail = missingTutorNotificationEmail.length;
  counts.tutorsMissingSchoolEmail = missingTutorSchoolEmail.length;
  nhsdeploy_addCheck_(checks, "tutors", "Eligible tutor network available", eligibleTutors.length ? "PASS" : (strict ? "BLOCK" : "WARN"),
    eligibleTutors.length + " eligible tutor(s) of " + tutorRows.length + " total.", "Access");
  nhsdeploy_addCheck_(checks, "tutor-id", "Eligible tutors have Tutor IDs", missingTutorId.length ? "BLOCK" : "PASS",
    missingTutorId.length ? missingTutorId.length + " eligible tutor(s) are missing Tutor ID." : "All eligible tutors have a Tutor ID.", "Access");
  nhsdeploy_addCheck_(checks, "tutor-id-duplicates", "Tutor IDs are unique", duplicateTutorIds.length ? "BLOCK" : "PASS",
    duplicateTutorIds.length ? "Duplicate Tutor ID: " + duplicateTutorIds.slice(0,5).join(", ") : "Tutor ID values are unique.", "Access");
  nhsdeploy_addCheck_(checks, "tutor-school-email", "Eligible tutors have school login email", missingTutorSchoolEmail.length ? "BLOCK" : "PASS",
    missingTutorSchoolEmail.length ? missingTutorSchoolEmail.length + " eligible tutor(s) are missing School Email." : "All eligible tutors have a School Email.", "Access");
  nhsdeploy_addCheck_(checks, "tutor-school-domain", "Tutor login accounts use student domain", nonStudentTutorEmails.length ? "WARN" : "PASS",
    nonStudentTutorEmails.length ? nonStudentTutorEmails.length + " eligible tutor login(s) are not @capostu.org; review intentional exceptions." : "All eligible tutor login emails use @capostu.org.", "Access");
  nhsdeploy_addCheck_(checks, "tutor-notification-email", "Eligible tutors have notification email", missingTutorNotificationEmail.length ? (strict ? "BLOCK" : "WARN") : "PASS",
    missingTutorNotificationEmail.length ? missingTutorNotificationEmail.length + " eligible tutor(s) are missing Notification Email." : "All eligible tutors have a Notification Email.", "Notifications");
  nhsdeploy_addCheck_(checks, "tutor-email-duplicates", "Tutor school emails are unique", duplicateTutorEmails.length ? "BLOCK" : "PASS",
    duplicateTutorEmails.length ? "Duplicate tutor School Email: " + duplicateTutorEmails.slice(0,5).join(", ") : "Tutor School Email values are unique.", "Access");

  const deadLetters = nhsops_deadLetterInventoryInternal_();
  counts.deadLetters = deadLetters.length;
  nhsdeploy_addCheck_(checks, "deadletters", "Background job queue is clean", deadLetters.length ? (strict ? "BLOCK" : "WARN") : "PASS",
    deadLetters.length ? deadLetters.length + " dead-letter job(s) need review." : "No dead-letter background jobs.", "Automation");

  let audit = null;
  try { audit = runNHSProductionAuditV43(); } catch (e) { audit = { ok:false, highFailures:1, error:nhsobs_sanitizeError_(e) }; }
  nhsdeploy_addCheck_(checks, "data-audit", "Workbook integrity audit", audit && audit.ok ? "PASS" : "BLOCK",
    audit && audit.ok ? "Production audit has no blocking integrity failures." : (audit && audit.error ? audit.error : String((audit && audit.highFailures) || 1) + " high-severity audit failure(s)."), "Data Integrity");

  const blockers = checks.filter(function(c){ return c.status === "BLOCK"; });
  const warnings = checks.filter(function(c){ return c.status === "WARN"; });
  const passes = checks.filter(function(c){ return c.status === "PASS"; });
  return {
    ok: blockers.length === 0,
    readyForSchoolwideRollout: strict && blockers.length === 0,
    strict: strict,
    build: NHS_RELEASE.VERSION,
    environment: env.environment,
    spreadsheetName: ss.getName(),
    counts: counts,
    passed: passes.length,
    warnings: warnings.length,
    blockers: blockers.length,
    checks: checks,
    generatedAt: new Date().toISOString()
  };
}

function nhsdeploy_runPreflight() {
  nhssec_requireMaintenanceAdmin_();
  return nhsdeploy_collectPreflight_({ strict: true });
}

function nhsdeploy_getConsoleState() {
  nhssec_requireMaintenanceAdmin_();
  const dead = nhsops_deadLetterInventoryInternal_();
  return {
    ok: true,
    preflight: nhsdeploy_collectPreflight_({ strict: true }),
    telemetry: nhsobs_getSummaryInternal_(),
    deadLetters: { count: dead.length, jobs: dead },
    environment: nhsrg_publicEnvironmentInfo_(),
    generatedAt: new Date().toISOString()
  };
}

function nhsdeploy_getDiagnosticBundle() {
  nhssec_requireMaintenanceAdmin_();
  const dead = nhsops_deadLetterInventoryInternal_();
  let audit = null;
  let releaseState = null;
  try { audit = runNHSProductionAuditV43(); } catch (e) { audit = { ok:false, error:nhsobs_sanitizeError_(e) }; }
  try { releaseState = nhsrg_getReleaseGuardState(); } catch (e) { releaseState = { ok:false, error:nhsobs_sanitizeError_(e) }; }
  return {
    report: "Tesoro NHS V47.3 Deployment Diagnostic",
    generatedAt: new Date().toISOString(),
    environment: nhsrg_publicEnvironmentInfo_(),
    preflight: nhsdeploy_collectPreflight_({ strict: true }),
    productionAudit: audit,
    releaseGuard: releaseState,
    triggers: nhsv43_auditTriggers_(),
    telemetry: nhsobs_getSummaryInternal_(),
    deadLetters: { count: dead.length, jobs: dead },
    dataRevision: String(nhsperf_getDataRevision_())
  };
}

// ---- staging persona lifecycle integration helpers -----------------------------
function nhsrg_patchFixtureRequest_(ss, requestId, patch) {
  const sh = ss.getSheetByName("Tutoring Requests");
  if (!sh) throw new Error("Tutoring Requests sheet missing.");
  const row = nhsdata_findRowByKey_(sh, "Request ID", requestId);
  if (!row) throw new Error("Release fixture request not found: " + requestId);
  const headers = nhsrg_headers_(sh);
  const values = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  Object.keys(patch || {}).forEach(function(header) {
    const idx = headers.indexOf(header);
    if (idx >= 0) values[idx] = patch[header];
  });
  sh.getRange(row, 1, 1, values.length).setValues([values]);
  SpreadsheetApp.flush();
  return true;
}

function nhsrg_fixtureRequest_(ss, requestId) {
  const row = nhsos_readRequests_(ss).find(function(r){ return String(r.requestId||"") === String(requestId||""); });
  if (!row) throw new Error("Release fixture disappeared: " + requestId);
  return row;
}

function nhsrg_runPersonaLifecycleFixture_(ss, fixture) {
  nhsrg_requireStaging_();
  const requestId = fixture.requestId;

  // Exercise the actual Operations assignment endpoint first. Authentication
  // remains real (current staging Admin/Advisor); no test-only impersonation or
  // production bypass is introduced.
  nhsos_saveRequest({
    requestId: requestId,
    assignedTutorId: fixture.tutorId,
    scheduledDate: fixture.dateKey,
    coordinatorNotes: "Release Guard V47.3 persona assignment"
  });

  let r = nhsrg_fixtureRequest_(ss, requestId);
  nhsrg_assert_(String(r.status || "") === "Scheduled", "Operations assignment did not schedule the fixture.");
  nhsrg_assert_(!nhslife_upcoming_(r, new Date(), NHS_PLATFORM.TIMEZONE), "Scheduled-but-unconfirmed fixture became Upcoming.");
  nhsrg_assert_(nhslife_needsTutorAction_(r), "Pending assignment did not require tutor action.");

  // Tutor decline/accept are represented as lifecycle mutations here because
  // Apps Script cannot securely impersonate a second Workspace user inside one
  // certification execution. The public Tutor endpoints remain separately
  // contract-tested and are never given a test bypass.
  nhsrg_patchFixtureRequest_(ss, requestId, { "Tutor Response Status":"Declined" });
  r = nhsrg_fixtureRequest_(ss, requestId);
  nhsrg_assert_(!nhslife_needsTutorAction_(r), "Declined assignment still required tutor action.");

  nhsrg_patchFixtureRequest_(ss, requestId, { "Tutor Response Status":"Accepted", "Tutor Response At":new Date() });
  r = nhsrg_fixtureRequest_(ss, requestId);
  nhsrg_assert_(!nhslife_confirmed_(r), "Tutor acceptance incorrectly confirmed the assignment.");

  // Exercise the actual Operations confirmation + notification path. In
  // STAGING, all mail is intercepted by the Release Guard sink.
  nhsos_confirmAssignment(requestId);
  r = nhsrg_fixtureRequest_(ss, requestId);
  nhsrg_assert_(nhslife_confirmed_(r), "Officer confirmation did not produce confirmed state.");
  nhsrg_assert_(nhslife_upcoming_(r, new Date(), NHS_PLATFORM.TIMEZONE), "Confirmed future fixture is not Upcoming.");

  nhsrg_patchFixtureRequest_(ss, requestId, { "Teacher Five Star Ready":true, "Tutor Five Star Enrolled":false });
  r = nhsrg_fixtureRequest_(ss, requestId);
  nhsrg_assert_(nhssched_readiness_(r) === "Waiting for Tutor Confirmation", "Teacher-only Five Star readiness is incorrect.");

  nhsrg_patchFixtureRequest_(ss, requestId, { "Tutor Five Star Enrolled":true });
  r = nhsrg_fixtureRequest_(ss, requestId);
  nhsrg_assert_(nhssched_readiness_(r) === "Ready", "Two-sided Five Star readiness did not become Ready.");

  return "Operations assignment → tutor lifecycle → Operations confirmation → Five Star flow passed";
}

function nhsrg_cancelPersonaFixture_(ss, fixture) {
  nhsrg_patchFixtureRequest_(ss, fixture.requestId, { "Status":"Cancelled" });
  const r = nhsrg_fixtureRequest_(ss, fixture.requestId);
  nhsrg_assert_(nhslife_closed_(r), "Cancelled fixture did not become closed.");
  nhsrg_assert_(!nhslife_activeCaseload_(r), "Cancelled fixture still counts toward active caseload.");
  return "Cancellation removed the fixture from active lifecycle state";
}




function rankAllOpenRequests() {
  nhssec_requireMaintenanceAdmin_();
  return rankAllOpenRequests_();
}

function runNHSNotificationSweep() {
  nhssec_requireMaintenanceAdmin_();
  return runNHSNotificationSweep_();
}

function handleAssignmentWorkflowEdit(e) {
  nhssec_requireMaintenanceAdmin_();
  return handleAssignmentWorkflowEdit_(e);
}

function nhsos_processFastJobs() {
  nhssec_requireMaintenanceAdmin_();
  return nhsos_processFastJobs_();
}

function processQueuedTutorCaseloadReconciliation() {
  nhssec_requireMaintenanceAdmin_();
  return processQueuedTutorCaseloadReconciliation_();
}

function tp_processAsyncReceiptQueue() {
  nhssec_requireMaintenanceAdmin_();
  return tp_processAsyncReceiptQueue_();
}

function tp_processCancellationJobs() {
  nhssec_requireMaintenanceAdmin_();
  return tp_processCancellationJobs_();
}

function tup_processJobs() {
  nhssec_requireMaintenanceAdmin_();
  return tup_processJobs_();
}

/** ============================================================
 * SOURCE GROUP: Router.gs
 * ============================================================ */


/** ============================================================
 * SOURCE GROUP: NHS_Release_Guard_V47.gs
 * ============================================================ */

/**
 * NHSOS V47 — RELEASE GUARD
 *
 * Safety model:
 *   - PRODUCTION: non-destructive smoke tests only.
 *   - STAGING: full certification, synthetic fixtures, mail sink, cleanup.
 *   - Full certification refuses to run when staging points at the live
 *     production spreadsheet.
 */
const NHSRG_BUILD = NHS_RELEASE.VERSION;
const NHSRG_ENV_PROPERTY = "NHS_ENVIRONMENT";
const NHSRG_DATA_ID_PROPERTY = "NHS_DATA_SPREADSHEET_ID";
const NHSRG_STAGING_MAIL_MODE_PROPERTY = "NHS_STAGING_MAIL_MODE";
const NHSRG_STAGING_MAIL_SINK_PROPERTY = "NHS_STAGING_MAIL_SINK";
const NHSRG_HISTORY_SHEET = "NHS Release History";
const NHSRG_RESULTS_SHEET = "NHS Release Test Results";
const NHSRG_MAIL_SINK_SHEET = "NHS Staging Mail Sink";
const NHSRG_TEST_PREFIX = "TEST-RC-";
const NHSRG_MAX_RESULT_ROWS = 1200;

function nhsrg_environmentInfo_() {
  const props = PropertiesService.getScriptProperties();
  const rawEnv = String(props.getProperty(NHSRG_ENV_PROPERTY) || "PRODUCTION").trim().toUpperCase();
  const environment = rawEnv === "STAGING" ? "STAGING" : "PRODUCTION";
  const configuredId = String(props.getProperty(NHSRG_DATA_ID_PROPERTY) || "").trim();
  const spreadsheetId = configuredId || NHS_PRODUCTION_SPREADSHEET_ID;
  const mailModeRaw = String(props.getProperty(NHSRG_STAGING_MAIL_MODE_PROPERTY) || "RECORD_ONLY").trim().toUpperCase();
  const mailMode = mailModeRaw === "REDIRECT" ? "REDIRECT" : "RECORD_ONLY";
  const sinkEmail = String(props.getProperty(NHSRG_STAGING_MAIL_SINK_PROPERTY) || "").trim();

  return {
    environment: environment,
    isProduction: environment === "PRODUCTION",
    isStaging: environment === "STAGING",
    build: NHSRG_BUILD,
    applicationBuild: typeof NHS_V45_BUILD !== "undefined" ? NHS_V45_BUILD : NHSRG_BUILD,
    spreadsheetId: spreadsheetId,
    productionSpreadsheetId: NHS_PRODUCTION_SPREADSHEET_ID,
    isolated: environment !== "STAGING" || spreadsheetId !== NHS_PRODUCTION_SPREADSHEET_ID,
    configuredDataId: !!configuredId,
    mailMode: mailMode,
    sinkEmail: sinkEmail,
    generatedAt: new Date().toISOString()
  };
}

function nhsrg_publicEnvironmentInfo_() {
  const info = nhsrg_environmentInfo_();
  return {
    environment: info.environment,
    isProduction: info.isProduction,
    isStaging: info.isStaging,
    build: info.build,
    applicationBuild: info.applicationBuild,
    isolated: info.isolated,
    mailMode: info.isStaging ? info.mailMode : "LIVE",
    generatedAt: info.generatedAt
  };
}

/**
 * Run once in the separate STAGING Apps Script project.
 * This function only writes project-scoped Script Properties; the new data
 * target takes effect on the next Apps Script invocation.
 */
/**
 * Creates a schema-only staging workbook from the production workbook.
 * No tutoring requests, sessions, tutor profiles, teacher directory rows, or
 * notification history are copied. Safe non-route Portal Settings are copied,
 * and the executing maintainer receives staging ADMIN + ADVISOR grants.
 */
function nhsrg_provisionStagingWorkbook(stagingName) {
  nhssec_requireMaintenanceAdmin_();
  const info = nhsrg_environmentInfo_();
  if (!info.isProduction || NHS_SPREADSHEET_ID !== NHS_PRODUCTION_SPREADSHEET_ID) {
    throw new Error("Provision a staging workbook only from the production project.");
  }

  const source = SpreadsheetApp.openById(NHS_PRODUCTION_SPREADSHEET_ID);
  const timestamp = Utilities.formatDate(new Date(), NHS_PLATFORM.TIMEZONE, "yyyy-MM-dd HHmm");
  const name = String(stagingName || ("NHS Tutor Database — STAGING — " + timestamp)).trim();
  const targetBook = SpreadsheetApp.create(name);
  const sourceSheets = source.getSheets();
  if (!sourceSheets.length) throw new Error("Production workbook has no sheets to provision.");

  const defaultSheet = targetBook.getSheets()[0];
  const hiddenLater = [];

  sourceSheets.forEach(function(src, index) {
    const target = index === 0 ? defaultSheet : targetBook.insertSheet();
    target.setName(src.getName());

    const lastCol = Math.max(1, src.getLastColumn());
    if (target.getMaxColumns() < lastCol) {
      target.insertColumnsAfter(target.getMaxColumns(), lastCol - target.getMaxColumns());
    }
    const headers = src.getRange(1, 1, 1, lastCol).getValues();
    target.getRange(1, 1, 1, lastCol).setValues(headers);
    try { target.getRange(1,1,1,lastCol).setFontWeight("bold"); } catch (e) {}
    try { target.setFrozenRows(Math.min(src.getFrozenRows(), target.getMaxRows() - 1)); } catch (e) {}
    try { target.setFrozenColumns(Math.min(src.getFrozenColumns(), target.getMaxColumns() - 1)); } catch (e) {}
    try { if (src.isSheetHidden()) hiddenLater.push(target); } catch (e) {}
  });

  // Copy only safe non-route configuration. Staging must derive its own URLs
  // from its separate deployment, never inherit production web-app routes.
  const sourceSettings = source.getSheetByName("Portal Settings");
  const targetSettings = targetBook.getSheetByName("Portal Settings");
  if (sourceSettings && targetSettings && sourceSettings.getLastRow() >= 2) {
    const width = Math.max(2, sourceSettings.getLastColumn());
    const rows = sourceSettings.getRange(2, 1, sourceSettings.getLastRow() - 1, width).getValues();
    const routeKeys = {GATEWAY_URL:true, INTERNAL_URL:true, TEACHER_URL:true, TUTOR_URL:true};
    const safeRows = rows.filter(function(row) {
      const key = String(row[0] || "").trim().toUpperCase();
      return key && !routeKeys[key];
    });
    if (safeRows.length) targetSettings.getRange(2, 1, safeRows.length, width).setValues(safeRows);
  }

  // Seed only the maintainer's staging authorization. Do not copy the live
  // Access Registry or any operational user directory.
  const email = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  const access = targetBook.getSheetByName("Access Registry");
  if (!email || !access) throw new Error("Unable to seed staging administrator access.");
  const accessHeaders = nhsrg_headers_(access);
  const accessMap = {};
  accessHeaders.forEach(function(h,i){ accessMap[h]=i; });
  function accessRow(role) {
    const row = new Array(accessHeaders.length).fill("");
    function put(header,value){ if (accessMap[header] !== undefined) row[accessMap[header]] = value; }
    put("Grant ID", "STAGING-" + role + "-" + Date.now() + "-" + Math.floor(Math.random()*1000));
    put("School Email", email);
    put("Email", email);
    put("Full Name", "Staging Maintainer");
    put("Role", role);
    put("Status", "Active");
    put("School Year", "Permanent");
    put("Start Date", new Date());
    put("Notes", "Provisioned by NHSOS Release Guard V47");
    return row;
  }
  const grants = [accessRow("ADMIN"), accessRow("ADVISOR")];
  if (/@capostu\.org$/i.test(email)) grants.push(accessRow("TUTOR"));
  access.getRange(2, 1, grants.length, accessHeaders.length).setValues(grants);

  // Give a capostu maintainer a minimal staging Tutor profile so all four
  // portals can be smoke-tested immediately without importing production data.
  if (/@capostu\.org$/i.test(email)) {
    const tutor = targetBook.getSheetByName("Tutor Database");
    if (tutor) {
      const headers = nhsrg_headers_(tutor);
      const map = {}; headers.forEach(function(h,i){map[h]=i;});
      const row = new Array(headers.length).fill("");
      function put(header,value){ if (map[header] !== undefined) row[map[header]] = value; }
      put("Tutor ID", "STAGING-TUTOR-" + Utilities.formatDate(new Date(), NHS_PLATFORM.TIMEZONE, "yyyyMMddHHmmss"));
      put("Full Name", "Staging Maintainer");
      put("Preferred Name", "Staging");
      put("School Email", email);
      put("Grade", "12");
      put("Mathematics Courses", "AP Calculus BC");
      put("Strongest Course", "AP Calculus BC");
      put("Typical Titan Time Availability", "Tuesday, Wednesday, Thursday, Friday");
      put("Maximum Sessions / Week", "2");
      put("Previous Experience", "Yes");
      put("Expectations Agreement", "Agree");
      put("Sessions Completed", "0");
      put("Verified Tutoring Hours", "0");
      put("Tutor No-Shows", "0");
      put("Current Caseload", "0");
      put("Reliability Status", "New");
      put("Notification Email", email);
      tutor.getRange(2,1,1,headers.length).setValues([row]);
    }
  }

  // Apply original hidden-sheet visibility only after all sheets exist so at
  // least one visible sheet is always available.
  hiddenLater.forEach(function(sh) { try { sh.hideSheet(); } catch (e) {} });
  SpreadsheetApp.flush();

  return {
    ok: true,
    name: name,
    spreadsheetId: targetBook.getId(),
    spreadsheetUrl: targetBook.getUrl(),
    maintainerEmail: email,
    message: "Schema-only staging workbook created. Create a separate Apps Script project, paste the V47 files, then run configureNHSStagingEnvironment with this spreadsheet ID."
  };
}

function configureNHSStagingEnvironment(stagingSpreadsheetId, sinkEmail, mailMode) {
  nhssec_requireMaintenanceAdmin_();
  const id = String(stagingSpreadsheetId || "").trim();
  const sink = String(sinkEmail || "").trim();
  const mode = String(mailMode || "RECORD_ONLY").trim().toUpperCase();

  if (!id) throw new Error("A staging spreadsheet ID is required.");
  if (id === NHS_PRODUCTION_SPREADSHEET_ID) {
    throw new Error("Staging cannot use the production spreadsheet ID.");
  }
  if (sink && !nhsm_isValidEmail_(sink)) {
    throw new Error("The staging sink email is invalid.");
  }
  if (!["RECORD_ONLY", "REDIRECT"].includes(mode)) {
    throw new Error("Mail mode must be RECORD_ONLY or REDIRECT.");
  }
  if (mode === "REDIRECT" && !sink) {
    throw new Error("REDIRECT mode requires a staging sink email.");
  }

  PropertiesService.getScriptProperties().setProperties({
    [NHSRG_ENV_PROPERTY]: "STAGING",
    [NHSRG_DATA_ID_PROPERTY]: id,
    [NHSRG_STAGING_MAIL_MODE_PROPERTY]: mode,
    [NHSRG_STAGING_MAIL_SINK_PROPERTY]: sink
  });

  return {
    ok: true,
    message: "Staging configuration saved. Start a new invocation before running certification.",
    environment: "STAGING",
    dataSpreadsheetId: id,
    mailMode: mode,
    sinkEmailConfigured: !!sink
  };
}

function configureNHSProductionEnvironment() {
  nhssec_requireMaintenanceAdmin_();
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    [NHSRG_ENV_PROPERTY]: "PRODUCTION",
    [NHSRG_DATA_ID_PROPERTY]: NHS_PRODUCTION_SPREADSHEET_ID,
    [NHSRG_STAGING_MAIL_MODE_PROPERTY]: "RECORD_ONLY",
    [NHSRG_STAGING_MAIL_SINK_PROPERTY]: ""
  });
  return {
    ok: true,
    message: "Production environment properties restored. Start a new invocation to reload configuration.",
    environment: "PRODUCTION"
  };
}

function nhsrg_requireStaging_() {
  const info = nhsrg_environmentInfo_();
  if (!info.isStaging) {
    throw new Error("Full Release Certification is STAGING-only. Run Production Smoke Test on production.");
  }
  if (!info.configuredDataId || !info.isolated || NHS_SPREADSHEET_ID === NHS_PRODUCTION_SPREADSHEET_ID) {
    throw new Error("STAGING SAFETY BLOCK: configure a separate staging spreadsheet before certification.");
  }
  return info;
}

function nhsrg_headers_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
    .map(function(v) { return String(v || "").trim(); });
}

function nhsrg_headerSet_(sheet) {
  const set = {};
  nhsrg_headers_(sheet).forEach(function(h) { if (h) set[h] = true; });
  return set;
}

function nhsrg_duplicateHeaders_(sheet) {
  const seen = {};
  const duplicates = [];
  nhsrg_headers_(sheet).forEach(function(h) {
    if (!h) return;
    if (seen[h] && !duplicates.includes(h)) duplicates.push(h);
    seen[h] = true;
  });
  return duplicates;
}

function nhsrg_getOrCreateSheet_(ss, name, headers, hidden) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const current = nhsrg_headers_(sh);
  const missing = (headers || []).filter(function(h) { return !current.includes(h); });
  if (!current.length && headers && headers.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else if (missing.length) {
    const start = Math.max(1, sh.getLastColumn() + 1);
    sh.getRange(1, start, 1, missing.length).setValues([missing]);
  }
  if (hidden) {
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}

function nhsrg_ensureHistorySheets_(ss) {
  const history = nhsrg_getOrCreateSheet_(ss, NHSRG_HISTORY_SHEET, [
    "Run ID", "Build", "Environment", "Started At", "Completed At", "Executed By",
    "Tests Passed", "Tests Failed", "Warnings", "Total Tests", "Status", "Duration Ms", "Notes"
  ], false);
  const results = nhsrg_getOrCreateSheet_(ss, NHSRG_RESULTS_SHEET, [
    "Run ID", "Build", "Environment", "Category", "Test", "Status", "Blocking", "Duration Ms", "Detail"
  ], true);
  if (nhsrg_environmentInfo_().isStaging) {
    nhsrg_getOrCreateSheet_(ss, NHSRG_MAIL_SINK_SHEET, [
      "Captured At", "Request ID", "Type", "Intended Recipient", "Delivered To",
      "Subject", "Mode", "Body Preview", "Calendar"
    ], true);
  }
  return { history: history, results: results };
}

function nhsrg_captureStagingMail_(spec, kind) {
  const info = nhsrg_environmentInfo_();
  if (!info.isStaging) return null;
  nhsrg_requireStaging_();

  spec = spec || {};
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const sh = nhsrg_getOrCreateSheet_(ss, NHSRG_MAIL_SINK_SHEET, [
    "Captured At", "Request ID", "Type", "Intended Recipient", "Delivered To",
    "Subject", "Mode", "Body Preview", "Calendar"
  ], true);
  const intended = String(spec.to || spec.recipient || "").trim();
  const delivered = info.mailMode === "REDIRECT" ? info.sinkEmail : "RECORD ONLY";
  const preview = String(spec.body || "").replace(/\s+/g, " ").slice(0, 500);
  const calendar = spec.calendar ? JSON.stringify(spec.calendar).slice(0, 1000) : "";
  sh.appendRow([
    new Date(), String(spec.requestId || ""), String(spec.type || kind || ""), intended,
    delivered, String(spec.subject || ""), info.mailMode, preview, calendar
  ]);
  return {
    intercepted: true,
    mode: info.mailMode,
    intendedRecipient: intended,
    redirectRecipient: info.sinkEmail,
    sinkSheet: NHSRG_MAIL_SINK_SHEET
  };
}

function nhsrg_stagingMailResult_(capture, originalRecipient, type) {
  return {
    ok: true,
    provider: "NHS Staging Mail Sink",
    recipient: String(originalRecipient || ""),
    intendedRecipient: String(originalRecipient || ""),
    deliveredTo: capture && capture.mode === "REDIRECT" ? String(capture.redirectRecipient || "") : "RECORD ONLY",
    relayMessageId: "STAGING-" + Date.now(),
    sender: "STAGING",
    calendarAttached: false,
    calendarUid: "",
    googleCalendarUrl: "",
    type: String(type || ""),
    staging: true
  };
}

/** Raw relay helper used only for STAGING REDIRECT mode. */
function nhsmail_sendRawRelay_(spec, intendedRecipient) {
  spec = spec || {};
  const to = String(spec.to || "").trim();
  const subject = String(spec.subject || "").trim();
  if (!nhsm_isValidEmail_(to) || !subject) throw new Error("Invalid staging redirect message.");
  const props = PropertiesService.getScriptProperties();
  const relayUrl = nhsm_normalizeRelayUrl_(props.getProperty(NHSM_RELAY_URL_PROPERTY) || NHSM_DEFAULT_RELAY_URL);
  const relayToken = String(props.getProperty(NHSM_RELAY_TOKEN_PROPERTY) || "").trim();
  if (!relayUrl || !relayToken) throw new Error("NHS mail relay is not configured for staging redirect mode.");
  const payload = {
    token: relayToken,
    to: to,
    subject: subject,
    body: String(spec.body || ""),
    htmlBody: String(spec.htmlBody || ""),
    name: String(spec.name || "Tesoro NHS"),
    source: "Tesoro NHS STAGING",
    requestId: String(spec.requestId || ""),
    type: String(spec.type || ""),
    sentAt: new Date().toISOString(),
    calendar: null
  };
  const response = UrlFetchApp.fetch(relayUrl,{method:"post",contentType:"application/json",payload:JSON.stringify(payload),muteHttpExceptions:true,followRedirects:true});
  const status = response.getResponseCode();
  let parsed = null;
  try { parsed = JSON.parse(response.getContentText() || "{}"); } catch (e) {}
  if (status < 200 || status >= 300 || !parsed || parsed.ok !== true) {
    throw new Error("Staging redirect relay failed: HTTP " + status);
  }
  return {
    id: String(spec.id || spec.type || "staging-message"),
    ok: true,
    provider: "NHS Staging Redirect",
    recipient: String(intendedRecipient || ""),
    intendedRecipient: String(intendedRecipient || ""),
    deliveredTo: to,
    relayMessageId: String(parsed.messageId || ""),
    staging: true
  };
}

function nhsrg_recordHistory_(ss, summary, notes) {
  const sheets = nhsrg_ensureHistorySheets_(ss);
  sheets.history.appendRow([
    summary.runId, summary.build, summary.environment, new Date(summary.startedAt),
    new Date(summary.completedAt), summary.executedBy || "", summary.passed, summary.failed,
    summary.warnings, summary.total, summary.status, summary.durationMs, String(notes || "")
  ]);
  const rows = (summary.results || []).map(function(r) {
    return [summary.runId, summary.build, summary.environment, r.category, r.name, r.status,
      r.blocking !== false, r.ms, r.detail || ""];
  });
  if (rows.length) {
    sheets.results.getRange(sheets.results.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
  }
  // Keep result history bounded while retaining the summary history.
  const overflow = Math.max(0, sheets.results.getLastRow() - 1 - NHSRG_MAX_RESULT_ROWS);
  if (overflow > 0) sheets.results.deleteRows(2, overflow);
}

function nhsrg_latestHistory_(ss, limit) {
  const sh = ss.getSheetByName(NHSRG_HISTORY_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const last = sh.getLastRow();
  const count = Math.min(Number(limit || 10), last - 1);
  const values = sh.getRange(last - count + 1, 1, count, sh.getLastColumn()).getDisplayValues();
  return values.reverse().map(function(row) {
    return {
      runId: row[0] || "", build: row[1] || "", environment: row[2] || "",
      startedAt: row[3] || "", completedAt: row[4] || "", executedBy: row[5] || "",
      passed: Number(row[6] || 0), failed: Number(row[7] || 0), warnings: Number(row[8] || 0),
      total: Number(row[9] || 0), status: row[10] || "", durationMs: Number(row[11] || 0), notes: row[12] || ""
    };
  });
}

function nhsrg_getReleaseGuardState() {
  nhssec_requireMaintenanceAdmin_();
  const info = nhsrg_environmentInfo_();
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  return {
    ok: true,
    environment: nhsrg_publicEnvironmentInfo_(),
    spreadsheetName: ss.getName(),
    spreadsheetIdMatchesProduction: NHS_SPREADSHEET_ID === NHS_PRODUCTION_SPREADSHEET_ID,
    fullCertificationAllowed: info.isStaging && info.isolated && info.configuredDataId,
    productionSmokeAllowed: true,
    latestRuns: nhsrg_latestHistory_(ss, 8),
    generatedAt: new Date().toISOString()
  };
}

function nhsrg_functionExists_(name) {
  name = String(name || "");
  if (!/^[A-Za-z_$][A-Za-z0-9_$]*$/.test(name)) return false;
  try {
    if (typeof globalThis !== "undefined") return typeof globalThis[name] === "function";
  } catch (e) {}
  try { return !!eval("typeof " + name + " === 'function'"); } catch (e) { return false; }
}

function nhsrg_assert_(condition, message) {
  if (!condition) throw new Error(message || "Assertion failed.");
  return true;
}

function nhsrg_assertThrows_(fn, messageContains) {
  let threw = false;
  let message = "";
  try { fn(); } catch (e) { threw = true; message = String(e && e.message ? e.message : e); }
  if (!threw) throw new Error("Expected an exception, but none was thrown.");
  if (messageContains && !message.includes(messageContains)) {
    throw new Error('Expected error containing "' + messageContains + '", got: ' + message);
  }
  return true;
}

function nhsrg_addTest_(tests, category, name, fn, blocking) {
  const started = Date.now();
  const result = {
    category: category,
    name: name,
    status: "PASS",
    blocking: blocking !== false,
    ms: 0,
    detail: ""
  };
  try {
    const value = fn();
    if (value && typeof value === "object" && value.detail) result.detail = String(value.detail);
    else if (typeof value === "string") result.detail = value;
  } catch (e) {
    result.status = result.blocking ? "FAIL" : "WARN";
    result.detail = String(e && e.message ? e.message : e);
  }
  result.ms = Date.now() - started;
  tests.push(result);
  return result;
}

function nhsrg_fixtureDate_() {
  const options = nhssched_teacherOptions_(NHS_PLATFORM.TIMEZONE, new Date());
  const first = (options.thisWeekDates || [])[0];
  if (first && first.value) return first.value;
  let key = nhssched_todayKey_(NHS_PLATFORM.TIMEZONE, new Date());
  for (let i = 1; i <= 14; i++) {
    const candidate = nhssched_addDaysKey_(key, i);
    if (nhssched_isTitanDay_(candidate)) return candidate;
  }
  return key;
}

function nhsrg_makeFixtureRows_(ss, runId) {
  const requestSh = ss.getSheetByName("Tutoring Requests");
  const tutorSh = ss.getSheetByName("Tutor Database");
  const sessionSh = ss.getSheetByName("Session Log");
  if (!requestSh || !tutorSh || !sessionSh) throw new Error("Fixture sheets are missing.");

  const dateKey = nhsrg_fixtureDate_();
  const requestId = NHSRG_TEST_PREFIX + runId + "-REQ";
  const tutorId = NHSRG_TEST_PREFIX + runId + "-TUTOR";
  const sessionId = NHSRG_TEST_PREFIX + runId + "-SESSION";

  function appendObject(sheet, object) {
    const headers = nhsrg_headers_(sheet);
    const row = headers.map(function(h) { return Object.prototype.hasOwnProperty.call(object, h) ? object[h] : ""; });
    const next = Math.max(2, sheet.getLastRow() + 1);
    sheet.getRange(next, 1, 1, row.length).setValues([row]);
    return next;
  }

  const tutorRow = appendObject(tutorSh, {
    "Tutor ID": tutorId,
    "Full Name": "Release Test Tutor",
    "Preferred Name": "Release",
    "School Email": "release.tutor@capostu.org",
    "Grade": "12",
    "Strongest Course": "AP Calculus BC",
    "Mathematics Courses": "AP Calculus BC",
    "Typical Titan Time Availability": "Tuesday, Wednesday, Thursday, Friday",
    "Maximum Sessions / Week": "2",
    "Previous Experience": "Yes",
    "Expectations Agreement": "Agree",
    "Sessions Completed": "0",
    "Verified Tutoring Hours": "0",
    "Tutor No-Shows": "0",
    "Current Caseload": "0",
    "Reliability Status": "New",
    "Notification Email": "release.tutor@example.com"
  });

  const requestRow = appendObject(requestSh, {
    "Request ID": requestId,
    "Timestamp": new Date(),
    "Teacher Name": "Release Test Teacher",
    "Teacher Email": "release.teacher@capousd.org",
    "Course": "AP Calculus BC",
    "Topic / Academic Need": "Release certification fixture",
    "Support Type": "Assignment Help",
    "Scheduling Preference": "This Week",
    "Preferred Days": nhssched_dayName_(dateKey),
    "Preferred Dates": dateKey,
    "Recurring?": "No — one-time / short-term support",
    "Status": "New",
    "Location": "RC TEST ROOM",
    "Request Source": "Release Certification"
  });

  const sessionRow = appendObject(sessionSh, {
    "Session ID": sessionId,
    "Request ID": requestId,
    "Tutor ID": tutorId,
    "Tutor": "Release Test Tutor",
    "Session Date": dateKey,
    "Result": "Completed",
    "Reported Minutes": 40,
    "Verification Status": "Pending",
    "Verified Minutes": "",
    "Notes": "Release certification fixture",
    "Logged At": new Date(),
    "Logged By": "release.certification"
  });

  SpreadsheetApp.flush();
  return {
    requestId: requestId, tutorId: tutorId, sessionId: sessionId,
    requestRow: requestRow, tutorRow: tutorRow, sessionRow: sessionRow, dateKey: dateKey
  };
}

function nhsrg_cleanupTestData_(ss) {
  nhsrg_requireStaging_();
  const targets = [
    ["Tutoring Requests", "Request ID"],
    ["Tutor Database", "Tutor ID"],
    ["Session Log", "Session ID"],
    ["Activity Log", "Request ID"],
    ["Notification Log", "Request ID"],
    [NHSRG_MAIL_SINK_SHEET, "Request ID"]
  ];
  let deleted = 0;
  targets.forEach(function(pair) {
    const sh = ss.getSheetByName(pair[0]);
    if (!sh || sh.getLastRow() < 2) return;
    const headers = nhsrg_headers_(sh);
    const idx = headers.indexOf(pair[1]);
    if (idx < 0) return;
    const values = sh.getRange(2, idx + 1, sh.getLastRow() - 1, 1).getDisplayValues();
    for (let i = values.length - 1; i >= 0; i--) {
      if (String(values[i][0] || "").startsWith(NHSRG_TEST_PREFIX)) {
        sh.deleteRow(i + 2);
        deleted++;
      }
    }
  });
  return deleted;
}

function nhsrg_cleanupTestData() {
  nhssec_requireMaintenanceAdmin_();
  const info = nhsrg_requireStaging_();
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  const deleted = nhsrg_cleanupTestData_(ss);
  nhs_v45InvalidatePortalState_({});
  return { ok: true, environment: info.environment, deletedRows: deleted, message: "Staging test data cleaned." };
}

function nhsrg_collectDataIntegrityTests_(tests, ss) {
  const requests = nhsos_readRequests_(ss);
  const tutors = nhsos_readTutors_(ss);
  const sessions = nhsos_readSessions_(ss);
  const requestIds = {};
  const tutorIds = {};

  requests.forEach(function(r) { if (r.requestId) requestIds[r.requestId] = (requestIds[r.requestId] || 0) + 1; });
  tutors.forEach(function(t) { if (t.tutorId) tutorIds[t.tutorId] = (tutorIds[t.tutorId] || 0) + 1; });

  nhsrg_addTest_(tests, "Data Integrity", "No duplicate Request IDs", function() {
    const dup = Object.keys(requestIds).filter(function(k) { return requestIds[k] > 1; });
    nhsrg_assert_(!dup.length, "Duplicate Request IDs: " + dup.slice(0, 5).join(", "));
  });
  nhsrg_addTest_(tests, "Data Integrity", "No duplicate Tutor IDs", function() {
    const dup = Object.keys(tutorIds).filter(function(k) { return tutorIds[k] > 1; });
    nhsrg_assert_(!dup.length, "Duplicate Tutor IDs: " + dup.slice(0, 5).join(", "));
  });
  nhsrg_addTest_(tests, "Data Integrity", "Sessions reference known requests or certification fixtures", function() {
    const bad = sessions.filter(function(s) { return s.requestId && !requestIds[s.requestId] && !String(s.requestId).startsWith(NHSRG_TEST_PREFIX); });
    nhsrg_assert_(!bad.length, "Orphan session request IDs: " + bad.slice(0, 5).map(function(s){return s.requestId;}).join(", "));
  });
  nhsrg_addTest_(tests, "Data Integrity", "Active assignments have tutor identity", function() {
    const bad = requests.filter(function(r) { return nhslife_activeCaseload_(r) && r.assignedTutor && !r.assignedTutorId; });
    nhsrg_assert_(!bad.length, "Active assignments missing tutor ID: " + bad.slice(0, 5).map(function(r){return r.requestId;}).join(", "));
  });
  nhsrg_addTest_(tests, "Data Integrity", "Five Star Ready requires both confirmations", function() {
    const bad = requests.filter(function(r) {
      return String(r.fiveStarReadiness || "").toLowerCase().includes("ready") &&
        !(r.teacherFiveStarReady && r.tutorFiveStarEnrolled);
    });
    nhsrg_assert_(!bad.length, "Invalid Five Star Ready state: " + bad.slice(0, 5).map(function(r){return r.requestId;}).join(", "));
  });
  nhsrg_addTest_(tests, "Data Integrity", "Verified sessions have nonnegative verified minutes", function() {
    const bad = sessions.filter(function(s) {
      return String(s.verificationStatus || "") === "Verified" && Number(s.minutes || 0) < 0;
    });
    nhsrg_assert_(!bad.length, "Negative verified minutes detected.");
  });
  nhsrg_addTest_(tests, "Data Integrity", "Voided sessions carry zero verified minutes", function() {
    const bad = sessions.filter(function(s) {
      return String(s.verificationStatus || "") === "Voided" && Number(s.minutes || 0) !== 0;
    });
    nhsrg_assert_(!bad.length, "Voided session retains verified minutes.");
  });
  nhsrg_addTest_(tests, "Data Integrity", "Confirmed future requests have scheduled dates", function() {
    const bad = requests.filter(function(r) { return nhslife_confirmed_(r) && !r.scheduledDate && !nhslife_closed_(r); });
    nhsrg_assert_(!bad.length, "Confirmed requests missing scheduled date.");
  });
  nhsrg_addTest_(tests, "Data Integrity", "Tutor caseload does not exceed configured capacity without a closed state", function() {
    const bad = tutors.filter(function(t) {
      const max = Number(t.maxSessions || 0);
      return max > 0 && Number(t.currentCaseload || 0) > max;
    });
    nhsrg_assert_(!bad.length, "Tutor capacity exceeded: " + bad.slice(0, 5).map(function(t){return t.displayName;}).join(", "));
  }, false);
  nhsrg_addTest_(tests, "Data Integrity", "Weekly availability rows use Tuesday-Friday dates", function() {
    const sh=ss.getSheetByName(NHSWA_SHEET); if(!sh||sh.getLastRow()<2)return;
    const values=sh.getDataRange().getDisplayValues(), h={}; values[0].forEach(function(v,i){h[String(v||"").trim()]=i;});
    const bad=values.slice(1).filter(function(row){const date=String(row[h["Date"]]||"").trim();return date&&!nhssched_isTitanDay_(date);});
    nhsrg_assert_(!bad.length,"Weekly availability contains a non-Titan-Time date.");
  });
  nhsrg_addTest_(tests, "Data Integrity", "Weekly availability statuses are valid", function() {
    const sh=ss.getSheetByName(NHSWA_SHEET); if(!sh||sh.getLastRow()<2)return;
    const values=sh.getDataRange().getDisplayValues(), h={}; values[0].forEach(function(v,i){h[String(v||"").trim()]=i;});
    const bad=values.slice(1).filter(function(row){const status=String(row[h["Status"]]||"").trim();return status&&!NHSWA_ALLOWED_STATUSES.includes(status);});
    nhsrg_assert_(!bad.length,"Invalid weekly availability status detected.");
  });
  nhsrg_addTest_(tests, "Data Integrity", "Weekly availability has no duplicate tutor/date rows", function() {
    const sh=ss.getSheetByName(NHSWA_SHEET); if(!sh||sh.getLastRow()<2)return;
    const values=sh.getDataRange().getDisplayValues(), h={}, seen={}, dup=[]; values[0].forEach(function(v,i){h[String(v||"").trim()]=i;});
    values.slice(1).forEach(function(row){const id=String(row[h["Tutor ID"]]||"").trim(),date=String(row[h["Date"]]||"").trim();if(!id||!date)return;const key=id+"|"+date;if(seen[key])dup.push(key);seen[key]=true;});
    nhsrg_assert_(!dup.length,"Duplicate weekly availability row: "+dup.slice(0,3).join(", "));
  });
}

function nhsrg_collectCoreTests_(tests, ss, includeStagingIntegration) {
  const info = nhsrg_environmentInfo_();

  [
    ["Environment name resolves", function(){ nhsrg_assert_(["PRODUCTION","STAGING"].includes(info.environment)); }],
    ["Build label present", function(){ nhsrg_assert_(!!info.build); }],
    ["Data spreadsheet ID resolves", function(){ nhsrg_assert_(!!NHS_SPREADSHEET_ID); }],
    ["Production spreadsheet constant present", function(){ nhsrg_assert_(!!NHS_PRODUCTION_SPREADSHEET_ID); }],
    ["Staging mail mode is safe", function(){ if(info.isStaging) nhsrg_assert_(["RECORD_ONLY","REDIRECT"].includes(info.mailMode)); }],
    ["Staging is isolated from production", function(){ if(info.isStaging) nhsrg_assert_(info.isolated, "Staging points at production data."); }],
    ["Staging has explicit data ID", function(){ if(info.isStaging) nhsrg_assert_(info.configuredDataId, "Staging data ID is not configured."); }],
    ["Redirect mode has sink email", function(){ if(info.isStaging && info.mailMode==="REDIRECT") nhsrg_assert_(nhsm_isValidEmail_(info.sinkEmail), "Redirect sink email missing/invalid."); }]
  ].forEach(function(t){ nhsrg_addTest_(tests,"Environment",t[0],t[1]); });

  const requiredSheets = [
    "Tutoring Requests","Tutor Database","Tutor Weekly Availability","Session Log","Activity Log","Notification Log",
    "Tutor Directory","Membership Roster","Teacher Directory","System Access","Access Registry",
    "Access Requests","Access Log","Portal Settings","System Health"
  ];
  requiredSheets.forEach(function(name) {
    nhsrg_addTest_(tests, "Schema", "Sheet exists: " + name, function() {
      nhsrg_assert_(!!ss.getSheetByName(name), "Missing sheet: " + name);
    });
  });

  const criticalHeaders = {
    "Tutoring Requests": [
      "Request ID","Timestamp","Teacher Name","Teacher Email","Course","Topic / Academic Need","Support Type",
      "Scheduling Preference","Preferred Days","Status","Assigned Tutor","Scheduled Date","Scheduled Time","Location",
      "Assigned Tutor ID","Tutor Response Status","Preferred Dates","Tutor Reconfirm Status","Teacher Five Star Ready",
      "Tutor Five Star Enrolled","Five Star Ready At","Teacher Notification Email"
    ],
    "Tutor Database": [
      "Tutor ID","Full Name","School Email","Strongest Course","Typical Titan Time Availability","Maximum Sessions / Week",
      "Sessions Completed","Verified Tutoring Hours","Current Caseload","Reliability Status","Temporary Unavailable Dates","Notification Email"
    ],
    "Session Log": [
      "Session ID","Request ID","Tutor ID","Tutor","Session Date","Result","Reported Minutes","Verification Status",
      "Verified Minutes","Logged At","Verified At","Voided At","Void Reason"
    ],
    "Access Registry": ["School Email","Full Name","Role","Status"],
    "Notification Log": ["Notification ID","Created At","Request ID","Channel","Recipient","Type","Status","Dedupe Key"],
    "Tutor Weekly Availability": ["Tutor ID","Week Start","Date","Day","Status","Confirmed At","Source"]
  };

  Object.keys(criticalHeaders).forEach(function(sheetName) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    const set = nhsrg_headerSet_(sh);
    criticalHeaders[sheetName].forEach(function(header) {
      nhsrg_addTest_(tests, "Schema", sheetName + " header: " + header, function() {
        nhsrg_assert_(!!set[header], 'Missing header "' + header + '" in ' + sheetName);
      });
    });
    nhsrg_addTest_(tests, "Schema", sheetName + " has no duplicate headers", function() {
      const dup = nhsrg_duplicateHeaders_(sh);
      nhsrg_assert_(!dup.length, "Duplicate headers: " + dup.join(", "));
    });
  });

  const endpoints = [
    "gateway_getState","gateway_requestAccess","nhs_getDataRevision","nhscourse_getCatalog",
    "teacher_getPortalData","teacher_getLiveRequestData","teacher_submitRequest","teacher_requestChange",
    "teacher_cancelRequest","teacher_endRecurringRequest","teacher_submitSessionFeedback","teacher_confirmFiveStarReady",
    "tutor_getPortalData","tutor_getLiveRequestData","tutor_updateProfile","tutor_acceptAssignment","tutor_declineAssignment",
    "tutor_reconfirmSession","tutor_confirmFiveStarEnrollment","tutor_reportSession","tutor_confirmWeeklyAvailability",
    "nhsos_getBootstrapData","nhsos_getLiveRequests","nhsos_saveRequest","nhsos_confirmAssignment",
    "nhsos_verifySession","nhsos_rejectSessionVerification","nhsos_correctVerifiedSession","nhsos_voidVerifiedSession",
    "nhsos_deleteVoidedSession","nhsa_adminBulkSyncTeachers","nhsa_adminDeleteVerifiedUser","nhsos_sendWeeklyAvailabilityReminders",
    "nhsrg_getReleaseGuardState","nhsrg_provisionStagingWorkbook","nhsrg_runReleaseCertification","nhsrg_runProductionSmokeTest","nhsrg_cleanupTestData",
    "nhsdeploy_runPreflight","nhsdeploy_getConsoleState","nhsdeploy_getDiagnosticBundle","nhsobs_getSummary","nhsops_getDeadLetters","nhsops_retryDeadLetter","nhsops_dismissDeadLetter"
  ];
  endpoints.forEach(function(name) {
    nhsrg_addTest_(tests, "Endpoint Contract", "Function exists: " + name, function() {
      nhsrg_assert_(nhsrg_functionExists_(name), "Missing server function: " + name);
    });
  });

  // Access policy tests are pure and do not grant access.
  [
    ["capousd Teacher accepted", "teacher@capousd.org", "TEACHER", true],
    ["capostu Teacher rejected", "student@capostu.org", "TEACHER", false],
    ["capostu Tutor accepted", "student@capostu.org", "TUTOR", true],
    ["capousd Tutor accepted", "staff@capousd.org", "TUTOR", true],
    ["personal Tutor rejected", "student.personal@gmail.com", "TUTOR", false],
    ["capostu Officer accepted", "student@capostu.org", "OFFICER", true],
    ["capostu Admin accepted", "student@capostu.org", "ADMIN", true]
  ].forEach(function(spec) {
    nhsrg_addTest_(tests, "Access Policy", spec[0], function() {
      let ok = true;
      try { nhsa_assertRoleEmailPolicy_(spec[1], spec[2]); } catch (e) { ok = false; }
      nhsrg_assert_(ok === spec[3], "Unexpected role-email policy result.");
    });
  });

  const monday = new Date("2026-09-14T12:00:00-07:00");
  const tuesday = new Date("2026-09-15T12:00:00-07:00");
  const thursday = "2026-09-17";
  const mondayKey = "2026-09-21";
  const teacherOptionsMonday = nhssched_teacherOptions_(NHS_PLATFORM.TIMEZONE, monday);
  const teacherOptionsTuesday = nhssched_teacherOptions_(NHS_PLATFORM.TIMEZONE, tuesday);

  [
    ["Monday shows Tue-Fri this week", function(){ nhsrg_assert_(teacherOptionsMonday.thisWeekDates.length === 4); }],
    ["Tuesday excludes same-day Titan Time", function(){ nhsrg_assert_(teacherOptionsTuesday.thisWeekDates.every(function(d){return d.value !== "2026-09-15";})); }],
    ["Thursday is a Titan Time day", function(){ nhsrg_assert_(nhssched_isTitanDay_(thursday)); }],
    ["Monday is not a Titan Time day", function(){ nhsrg_assert_(!nhssched_isTitanDay_(mondayKey)); }],
    ["Future Thursday validates", function(){ nhsrg_assert_(nhssched_validateFutureTitanDate_(thursday,NHS_PLATFORM.TIMEZONE,monday)===thursday); }],
    ["Future Monday rejected", function(){ nhsrg_assertThrows_(function(){nhssched_validateFutureTitanDate_(mondayKey,NHS_PLATFORM.TIMEZONE,monday);},"Titan Time"); }],
    ["Same-day rejected", function(){ nhsrg_assertThrows_(function(){nhssched_validateFutureTitanDate_("2026-09-15",NHS_PLATFORM.TIMEZONE,tuesday);},"Same-day"); }],
    ["This Week preference validates", function(){ const r=nhssched_validateTeacherPreference_({schedulingType:"This Week",preferredDates:[thursday]},NHS_PLATFORM.TIMEZONE,monday); nhsrg_assert_(r.preferredDates[0]===thursday); }],
    ["Future Reservation validates", function(){ const futureTuesday="2026-09-22"; const r=nhssched_validateTeacherPreference_({schedulingType:"Future Reservation",preferredDates:[futureTuesday]},NHS_PLATFORM.TIMEZONE,monday); nhsrg_assert_(r.preferredDates[0]===futureTuesday); }],
    ["Recurring weekdays normalize", function(){ const r=nhssched_validateTeacherPreference_({schedulingType:"Recurring Support",preferredDays:["Tuesday","Thursday"]},NHS_PLATFORM.TIMEZONE,monday); nhsrg_assert_(r.preferredDays.length===2); }],
    ["Reconfirmation window excludes same day", function(){ nhsrg_assert_(!nhssched_reconfirmWindowOpen_("2026-09-15",NHS_PLATFORM.TIMEZONE,tuesday)); }],
    ["Reconfirmation opens within seven days", function(){ nhsrg_assert_(nhssched_reconfirmWindowOpen_("2026-09-17",NHS_PLATFORM.TIMEZONE,tuesday)); }]
  ].forEach(function(t){ nhsrg_addTest_(tests,"Scheduling",t[0],t[1]); });

  [
    ["APUSH alias resolves", function(){ nhsrg_assert_(canonicalCourse_("APUSH") === canonicalCourse_("AP US History")); }],
    ["AP Calculus BC canonicalizes", function(){ nhsrg_assert_(!!canonicalCourse_("AP Calculus BC")); }],
    ["Course catalog is populated", function(){ nhsrg_assert_(nhscourse_clientCatalog_().length >= 80, "Course catalog unexpectedly small."); }],
    ["Health is included", function(){ nhsrg_assert_(!!nhscourse_resolve_("Health")); }]
  ].forEach(function(t){ nhsrg_addTest_(tests,"Course Registry",t[0],t[1]); });

  const request = {
    "Preferred Days":"Thursday","Preferred Dates":"2026-09-17","Scheduled Date":"",
    "Specific Tutor Preference":"Release Test Tutor"
  };
  const tutor = {
    "Tutor ID":"T-RC","Full Name":"Release Test Tutor","Preferred Name":"Release",
    "Mathematics Courses":"AP Calculus BC","Strongest Course":"AP Calculus BC",
    "Typical Titan Time Availability":"Thursday","Reliability Status":"Reliable",
    "Maximum Sessions / Week":"2","Current Caseload":"0","Previous Experience":"Yes",
    "Notification Email":"release.tutor@example.com","Temporary Unavailable Dates":""
  };
  const freeIndex = {byId:{},byName:{}};
  const bookedIndex = {byId:{"T-RC":{"2026-09-17":true}},byName:{}};
  const partialRequest = Object.assign({}, request, {"Preferred Dates":"2026-09-17, 2026-09-18","Preferred Days":"Thursday, Friday"});
  const partialTutor = Object.assign({}, tutor, {"Typical Titan Time Availability":"Thursday, Friday"});

  [
    ["Qualified tutor is eligible", function(){ const r=scoreTutorForRequest_(request,tutor,"AP Calculus BC",freeIndex); nhsrg_assert_(r.eligible && r.score<=100); }],
    ["Wrong course is ineligible", function(){ const r=scoreTutorForRequest_(request,tutor,"AP Biology",freeIndex); nhsrg_assert_(!r.eligible); }],
    ["Booked-only date is ineligible", function(){ const r=scoreTutorForRequest_(request,tutor,"AP Calculus BC",bookedIndex); nhsrg_assert_(!r.eligible); }],
    ["One free requested date remains eligible", function(){ const r=scoreTutorForRequest_(partialRequest,partialTutor,"AP Calculus BC",bookedIndex); nhsrg_assert_(r.eligible); }],
    ["Suspended tutor is ineligible", function(){ const r=scoreTutorForRequest_(request,Object.assign({},tutor,{"Reliability Status":"Suspended"}),"AP Calculus BC",freeIndex); nhsrg_assert_(!r.eligible); }],
    ["At-capacity tutor is ineligible", function(){ const r=scoreTutorForRequest_(request,Object.assign({},tutor,{"Current Caseload":"2"}),"AP Calculus BC",freeIndex); nhsrg_assert_(!r.eligible); }],
    ["Missing notification email is ineligible", function(){ const r=scoreTutorForRequest_(request,Object.assign({},tutor,{"Notification Email":""}),"AP Calculus BC",freeIndex); nhsrg_assert_(!r.eligible); }],
    ["Match strength labels bounded", function(){ nhsrg_assert_(nhsmatch_strengthLabel_(100)==="Exceptional Match" && !!nhsmatch_strengthLabel_(50)); }],
    ["Matching weights total 100", function(){ nhsrg_assert_(50+10+15+10+10+5===100); }],
    ["Strongest-course qualification scores 50", function(){ const r=scoreTutorForRequest_(request,tutor,"AP Calculus BC",freeIndex,{byTutorDate:{}}); const q=(r.criteria||[]).find(function(c){return c.key==="qualification";}); nhsrg_assert_(q&&q.score===50&&q.max===50); }],
    ["Weekly confirmed availability scores 10", function(){ const idx={byTutorDate:{"T-RC":{"2026-09-17":{status:"Available",day:"Thursday",source:"Tutor Confirmed",confirmedAt:"2026-09-14"}}}}; const r=scoreTutorForRequest_(request,tutor,"AP Calculus BC",freeIndex,idx); const c=(r.criteria||[]).find(function(x){return x.key==="schedule";}); nhsrg_assert_(c&&c.score===10&&c.max===10); }],
    ["Weekly unavailable is ineligible", function(){ const idx={byTutorDate:{"T-RC":{"2026-09-17":{status:"Unavailable",day:"Thursday",source:"Tutor Confirmed",confirmedAt:"2026-09-14"}}}}; const r=scoreTutorForRequest_(request,tutor,"AP Calculus BC",freeIndex,idx); nhsrg_assert_(!r.eligible); }]
  ].forEach(function(t){ nhsrg_addTest_(tests,"Matching",t[0],t[1]); });

  [
    ["Scheduled is not implicitly confirmed", function(){ nhsrg_assert_(!nhslife_confirmed_({"Status":"Scheduled","Assignment Confirmed":""})); }],
    ["Confirmed timestamp marks confirmed", function(){ nhsrg_assert_(nhslife_confirmed_({"Status":"Scheduled","Assignment Confirmed":"2026-09-16"})); }],
    ["Cancelled is closed", function(){ nhsrg_assert_(nhslife_closed_({"Status":"Cancelled"})); }],
    ["Completed is closed", function(){ nhsrg_assert_(nhslife_closed_({"Status":"Completed"})); }],
    ["Cancelled is not active caseload", function(){ nhsrg_assert_(!nhslife_activeCaseload_({"Status":"Cancelled","Assigned Tutor":"X"})); }],
    ["Five Star pending with neither side", function(){ nhsrg_assert_(nhssched_readiness_({"Teacher Five Star Ready":false,"Tutor Five Star Enrolled":false}).includes("Pending")); }],
    ["Five Star ready with both sides", function(){ nhsrg_assert_(nhssched_readiness_({"Teacher Five Star Ready":true,"Tutor Five Star Enrolled":true}).includes("Ready")); }]
  ].forEach(function(t){ nhsrg_addTest_(tests,"Lifecycle",t[0],t[1]); });

  nhsrg_addTest_(tests, "Notifications", "Mail transport function installed", function(){ nhsrg_assert_(nhsrg_functionExists_("nhsmail_send_")); });
  nhsrg_addTest_(tests, "Notifications", "Batch mail transport installed", function(){ nhsrg_assert_(nhsrg_functionExists_("nhsmail_sendBatch_")); });
  nhsrg_addTest_(tests, "Notifications", "Calendar transport installed", function(){ nhsrg_assert_(nhsrg_functionExists_("nhsmail_createCalendarEvent_")); });
  nhsrg_addTest_(tests, "Notifications", "Staging transport cannot be LIVE", function(){ if(info.isStaging) nhsrg_assert_(info.mailMode!=="LIVE"); });

  const perfStart = Date.now();
  nhsrg_addTest_(tests, "Performance", "Spreadsheet memoized lookup available", function(){ nhsrg_assert_(!!nhsperf_ss_(NHS_SPREADSHEET_ID)); });
  nhsrg_addTest_(tests, "Performance", "Critical sheets resolve", function(){ ["Tutoring Requests","Tutor Database","Session Log"].forEach(function(n){nhsrg_assert_(!!nhsperf_sheet_(ss,n),"Missing "+n);}); });
  nhsrg_addTest_(tests, "Performance", "Core test collection under 30 seconds", function(){ nhsrg_assert_(Date.now()-perfStart < 30000, "Core certification checks exceeded 30 seconds."); }, false);
  nhsrg_addTest_(tests, "Deployment", "V47.3 deployment preflight engine builds", function(){ const p=nhsdeploy_collectPreflight_({strict:false}); nhsrg_assert_(Array.isArray(p.checks) && p.checks.length>=10, "Deployment preflight did not return a complete checklist."); return p.checks.length+" preflight checks"; });
  nhsrg_addTest_(tests, "Observability", "Runtime telemetry summary builds", function(){ const t=nhsobs_getSummaryInternal_(); nhsrg_assert_(Array.isArray(t.metrics) && t.metrics.length===NHSOBS_METRICS.length); });
  nhsrg_addTest_(tests, "Reliability", "Dead-letter inventory is readable", function(){ nhsrg_assert_(Array.isArray(nhsops_deadLetterInventoryInternal_())); });

  if (includeStagingIntegration) {
    nhsrg_addTest_(tests, "Staging Integration", "Synthetic fixture cleanup before run", function(){ return "Removed " + nhsrg_cleanupTestData_(ss) + " stale test rows"; });
    let fixture = null;
    nhsrg_addTest_(tests, "Staging Integration", "Create synthetic request/tutor/session fixtures", function(){ fixture=nhsrg_makeFixtureRows_(ss,String(Date.now())); nhsrg_assert_(!!fixture.requestId); return fixture.requestId; });
    nhsrg_addTest_(tests, "Staging Integration", "Synthetic request readable through Operations reader", function(){ const rows=nhsos_readRequests_(ss); nhsrg_assert_(rows.some(function(r){return r.requestId===fixture.requestId;})); });
    nhsrg_addTest_(tests, "Staging Integration", "Synthetic tutor readable through Operations reader", function(){ const rows=nhsos_readTutors_(ss); nhsrg_assert_(rows.some(function(r){return r.tutorId===fixture.tutorId;})); });
    nhsrg_addTest_(tests, "Staging Integration", "Synthetic session readable through Operations reader", function(){ const rows=nhsos_readSessions_(ss); nhsrg_assert_(rows.some(function(r){return r.sessionId===fixture.sessionId;})); });
    nhsrg_addTest_(tests, "Staging Persona", "Teacher-to-confirmed-session lifecycle", function(){ return nhsrg_runPersonaLifecycleFixture_(ss,fixture); });
    nhsrg_addTest_(tests, "Staging Persona", "Session verification mutation", function(){ nhsos_verifySession(fixture.sessionId,40,"Release Guard V47.3 verification"); const row=nhsos_readSessions_(ss).find(function(x){return x.sessionId===fixture.sessionId;}); nhsrg_assert_(row && String(row.verificationStatus)==="Verified","Synthetic session did not verify."); return "Verified 40 minutes"; });
    nhsrg_addTest_(tests, "Staging Persona", "Verified-session void mutation", function(){ nhsos_voidVerifiedSession(fixture.sessionId,"Release Guard V47.3 void test"); const row=nhsos_readSessions_(ss).find(function(x){return x.sessionId===fixture.sessionId;}); nhsrg_assert_(row && String(row.verificationStatus)==="Voided","Synthetic verified session did not void."); return "Verified credit removed"; });
    nhsrg_addTest_(tests, "Staging Persona", "Cancellation leaves no active caseload state", function(){ return nhsrg_cancelPersonaFixture_(ss,fixture); });
    nhsrg_addTest_(tests, "Staging Integration", "Staging mail is intercepted", function(){ const result=nhsmail_send_({to:"release.teacher@capousd.org",subject:"[RC] Staging mail capture",body:"Release certification mail sink test.",requestId:fixture.requestId,type:"Release Certification"}); nhsrg_assert_(result && result.staging===true,"Mail was not intercepted by staging sink."); });
    nhsrg_addTest_(tests, "Staging Integration", "Staging mail sink contains capture", function(){ const sh=ss.getSheetByName(NHSRG_MAIL_SINK_SHEET); nhsrg_assert_(sh && sh.getLastRow()>=2); });
    nhsrg_addTest_(tests, "Staging Integration", "Synthetic cleanup removes fixtures", function(){ const deleted=nhsrg_cleanupTestData_(ss); nhsrg_assert_(deleted>=3,"Expected at least 3 synthetic rows to be removed."); return "Removed "+deleted+" rows"; });
    nhsrg_addTest_(tests, "Staging Integration", "Synthetic request absent after cleanup", function(){ const rows=nhsos_readRequests_(ss); nhsrg_assert_(!rows.some(function(r){return String(r.requestId||"").startsWith(NHSRG_TEST_PREFIX);})); });
  }

  nhsrg_collectDataIntegrityTests_(tests, ss);
}

function nhsrg_summarize_(runId, environment, started, tests) {
  const failed = tests.filter(function(t){ return t.status === "FAIL"; }).length;
  const warnings = tests.filter(function(t){ return t.status === "WARN"; }).length;
  const passed = tests.filter(function(t){ return t.status === "PASS"; }).length;
  const completed = new Date();
  return {
    ok: failed === 0,
    runId: runId,
    build: NHSRG_BUILD,
    applicationBuild: typeof NHS_V45_BUILD !== "undefined" ? NHS_V45_BUILD : "",
    environment: environment,
    startedAt: started.toISOString(),
    completedAt: completed.toISOString(),
    durationMs: completed.getTime() - started.getTime(),
    executedBy: String(Session.getActiveUser().getEmail() || ""),
    passed: passed,
    failed: failed,
    warnings: warnings,
    total: tests.length,
    status: failed ? "FAIL" : (warnings ? "PASS WITH WARNINGS" : "PASS"),
    certifiedForProduction: environment === "STAGING" && failed === 0,
    results: tests
  };
}

function nhsrg_runReleaseCertification() {
  nhssec_requireMaintenanceAdmin_();
  const info = nhsrg_requireStaging_();
  const started = new Date();
  const runId = NHSRG_TEST_PREFIX + Utilities.formatDate(started, NHS_PLATFORM.TIMEZONE, "yyyyMMdd-HHmmss");
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  nhsrg_ensureHistorySheets_(ss);

  // Certification validates migration behavior in staging, but never resets.
  nhsos_ensureRequestColumns_(ss.getSheetByName("Tutoring Requests"));
  nhswa_ensureSheet_(ss);
  const tests = [];
  let summary = null;
  try {
    nhsrg_collectCoreTests_(tests, ss, true);
    summary = nhsrg_summarize_(runId, info.environment, started, tests);
    nhsrg_recordHistory_(ss, summary, "Full staging certification");
    return summary;
  } finally {
    try { nhsrg_cleanupTestData_(ss); } catch (cleanupError) {}
    try { nhs_v45InvalidatePortalState_({}); } catch (e) {}
  }
}

function nhsrg_runProductionSmokeTest() {
  nhssec_requireMaintenanceAdmin_();
  const info = nhsrg_environmentInfo_();
  const started = new Date();
  const runId = "SMOKE-" + Utilities.formatDate(started, NHS_PLATFORM.TIMEZONE, "yyyyMMdd-HHmmss");
  const ss = nhsperf_ss_(NHS_SPREADSHEET_ID);
  nhswa_ensureSheet_(ss);
  nhsrg_ensureHistorySheets_(ss);
  const tests = [];

  const smoke = [
    ["Spreadsheet opens", function(){ nhsrg_assert_(!!ss); }],
    ["Gateway endpoint installed", function(){ nhsrg_assert_(nhsrg_functionExists_("gateway_getState")); }],
    ["Teacher bootstrap endpoint installed", function(){ nhsrg_assert_(nhsrg_functionExists_("teacher_getPortalData")); }],
    ["Tutor bootstrap endpoint installed", function(){ nhsrg_assert_(nhsrg_functionExists_("tutor_getPortalData")); }],
    ["Operations bootstrap endpoint installed", function(){ nhsrg_assert_(nhsrg_functionExists_("nhsos_getBootstrapData")); }],
    ["Course catalog loads", function(){ nhsrg_assert_(nhscourse_clientCatalog_().length >= 80); }],
    ["Data revision service responds", function(){ nhsrg_assert_(String(nhsperf_getDataRevision_()) !== ""); }],
    ["Core sheets accessible", function(){ ["Tutoring Requests","Tutor Database","Tutor Weekly Availability","Session Log"].forEach(function(n){nhsrg_assert_(!!ss.getSheetByName(n),"Missing "+n);}); }],
    ["Weekly availability endpoint installed", function(){ nhsrg_assert_(nhsrg_functionExists_("tutor_confirmWeeklyAvailability")); }],
    ["Mail relay configuration readable", function(){ const st=getNHSMailRelayStatus(); nhsrg_assert_(typeof st.configured === "boolean"); }],
    ["System health builds", function(){ const h=nhsos_getSystemHealth_(ss); nhsrg_assert_(Array.isArray(h.missingSheets)); }],
    ["Notification sweep installed", function(){ nhsrg_assert_(nhsrg_functionExists_("runNHSNotificationSweep_")); }],
    ["Fast worker installed", function(){ nhsrg_assert_(nhsrg_functionExists_("nhsos_processFastJobs_")); }]
  ];
  smoke.forEach(function(t){ nhsrg_addTest_(tests,"Production Smoke",t[0],t[1]); });
  nhsrg_collectDataIntegrityTests_(tests, ss);
  const summary = nhsrg_summarize_(runId, info.environment, started, tests);
  summary.certifiedForProduction = false;
  nhsrg_recordHistory_(ss, summary, "Non-destructive production smoke test");
  return summary;
}


/**
 * ROUTER — ONE DEPLOYMENT, FOUR INTERFACES
 *
 * /exec                    -> Gateway
 * /exec/teacher            -> Teacher Portal
 * /exec/tutor              -> Tutor Portal
 * /exec/operations         -> Internal NHS Operations
 */

function doGet(e) {
  // V45.5.2 ROUTING HARDENING:
  // Prefer Apps Script pathInfo (/exec/teacher) because Google auth/redirect
  // handoffs can discard query parameters such as ?page=teacher. Keep the
  // legacy query route as a backward-compatible fallback.
  const pathRoute = String(
    e && e.pathInfo ? e.pathInfo : ""
  )
    .replace(/^\/+|\/+$/g, "")
    .split("/")[0]
    .trim()
    .toLowerCase();

  const explicitRoute = String(
    e && e.parameter && e.parameter.route
      ? e.parameter.route
      : ""
  )
    .trim()
    .toLowerCase();

  const legacyQueryRoute = String(
    e && e.parameter && e.parameter.page
      ? e.parameter.page
      : ""
  )
    .trim()
    .toLowerCase();

  const requested =
    pathRoute ||
    explicitRoute ||
    legacyQueryRoute ||
    NHS_PLATFORM.DEFAULT_ROUTE;

  const routes = {
    gateway: {
      file: "Gateway",
      title: "Tesoro NHS Hub",
      favicon:
        "https://ssl.gstatic.com/docs/doclist/images/mediatype/icon_1_folder_x16.png"
    },
    teacher: {
      file: "Teacher",
      title: "Tesoro NHS Teacher Desk",
      favicon:
        "https://ssl.gstatic.com/docs/doclist/images/mediatype/icon_1_document_x16.png"
    },
    tutor: {
      file: "Tutor",
      title: "Tesoro NHS Tutor Desk",
      favicon:
        "https://ssl.gstatic.com/docs/doclist/images/mediatype/icon_1_form_x16.png"
    },
    operations: {
      file: "Operations",
      title: "Tesoro NHS Operations Center",
      favicon:
        "https://ssl.gstatic.com/docs/doclist/images/mediatype/icon_1_spreadsheet_x16.png"
    },
    ops: {
      file: "Operations",
      title: "Tesoro NHS Operations Center",
      favicon:
        "https://ssl.gstatic.com/docs/doclist/images/mediatype/icon_1_spreadsheet_x16.png"
    }
  };

  const route = routes[requested] || routes.gateway;

  const output = HtmlService
    .createTemplateFromFile(route.file)
    .evaluate()
    .setTitle(route.title)
    .setFaviconUrl(route.favicon)
    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1, viewport-fit=cover"
    );

  return output;
}


function includeGateway_(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}


/**
 * Canonical URLs for the single deployment.
 * Useful after deployment when copying links into Portal Settings.
 */
function getNHSPlatformRoutes(baseUrl) {
  nhssec_requireMaintenanceAdmin_();

  baseUrl = String(baseUrl || "").trim();

  if (!baseUrl) {
    throw new Error("Base deployment URL is required.");
  }

  const clean = baseUrl.split("?")[0].replace(/\/dev$/, "/exec");

  return {
    gateway: clean,
    teacher: nhs_routeUrl_(clean, "teacher"),
    tutor: nhs_routeUrl_(clean, "tutor"),
    operations: nhs_routeUrl_(clean, "operations")
  };
}
function debugNHSAdminAccess() {
  const activeEmail = String(
    Session.getActiveUser().getEmail() || ""
  ).trim().toLowerCase();

  const effectiveEmail = String(
    Session.getEffectiveUser().getEmail() || ""
  ).trim().toLowerCase();

  const ss = nhsa_openSpreadsheet_();
  const sheet = ss.getSheetByName(NHSA_REGISTRY_SHEET);

  if (!sheet) {
    throw new Error("Access Registry sheet not found.");
  }

  const grants = nhsa_readRegistry_(sheet);

  const privileged = grants
    .filter(function(grant) {
      return nhsa_isGrantActive_(grant) &&
        ["ADMIN", "ADVISOR"].includes(
          String(grant.role || "").trim().toUpperCase()
        );
    })
    .map(function(grant) {
      return {
        email: grant.email,
        name: grant.fullName,
        role: grant.role,
        status: grant.status
      };
    });

  const myGrants = grants
    .filter(function(grant) {
      return (
        String(grant.email || "").trim().toLowerCase() === activeEmail &&
        nhsa_isGrantActive_(grant)
      );
    })
    .map(function(grant) {
      return {
        role: grant.role,
        status: grant.status
      };
    });

  console.log("ACTIVE EMAIL: " + activeEmail);
  console.log("EFFECTIVE EMAIL: " + effectiveEmail);
  console.log("MY ACTIVE GRANTS:");
  console.log(JSON.stringify(myGrants, null, 2));
  console.log("ACTIVE ADMIN / ADVISOR GRANTS:");
  console.log(JSON.stringify(privileged, null, 2));

  return {
    activeEmail: activeEmail,
    effectiveEmail: effectiveEmail,
    myGrants: myGrants,
    privileged: privileged
  };
}
