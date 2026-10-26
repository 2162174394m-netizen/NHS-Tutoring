/**
 * TESORO NHS — CURRENT ACCESS BULK CHECK V47.7
 * READ-ONLY
 *
 * Lists every account with an ACTIVE NHS Access Registry grant, combines
 * multiple grants by email, shows portal authorization, and reports the
 * current Apps Script identity context. This script does not mutate data.
 */

function runNHSCurrentAccessCheck() {
  nhssec_requireMaintenanceAdmin_();
  const report = getNHSCurrentAccessNow_();

  console.log("============================================================");
  console.log("TESORO NHS — CURRENT ACCESS V47.7");
  console.log("Generated: " + report.generatedAt);
  console.log("Current authorized users: " + report.count);

  if (report.deploymentIdentity) {
    console.log("------------------------------------------------------------");
    console.log("IDENTITY CONTEXT");
    console.log("Active user: " + (report.deploymentIdentity.activeUserEmail || "(hidden)"));
    console.log("Effective user: " + (report.deploymentIdentity.effectiveUserEmail || "(hidden)"));
    console.log("Service URL: " + (report.deploymentIdentity.serviceUrl || "(not deployed)"));
    console.log("Identity visible: " + (report.deploymentIdentity.ok ? "YES" : "NO"));
  }

  console.log("============================================================");

  if (!report.users.length) {
    console.log("No currently active Access Registry grants were found.");
    return report;
  }

  report.users.forEach(function(user, index) {
    const portals = [];
    if (user.access.gateway) portals.push("Gateway");
    if (user.access.teacher) portals.push("Teacher");
    if (user.access.tutor) portals.push("Tutor");
    if (user.access.operations) portals.push("Operations");

    console.log(
      (index + 1) + ". " +
      (user.fullName || "(No name)") +
      " <" + user.email + ">" +
      " | Roles: " + user.roles.join(", ") +
      " | Access: " + portals.join(", ") +
      (user.expirationDates.length
        ? " | Expires: " + user.expirationDates.join(", ")
        : "")
    );
  });

  console.log("============================================================");
  console.log("COUNTS BY ROLE");
  Object.keys(report.countsByRole).sort().forEach(function(role) {
    console.log(role + ": " + report.countsByRole[role]);
  });

  console.log("============================================================");
  console.log("COUNTS BY PORTAL");
  console.log("Gateway: " + report.countsByPortal.gateway);
  console.log("Teacher: " + report.countsByPortal.teacher);
  console.log("Tutor: " + report.countsByPortal.tutor);
  console.log("Operations: " + report.countsByPortal.operations);

  if (report.issues.length) {
    console.log("============================================================");
    console.log("ACCESS WARNINGS");
    report.issues.forEach(function(issue) {
      console.log("- " + issue);
    });
  }

  console.log("============================================================");
  return report;
}

function getNHSCurrentAccessNow_() {
  const ss = nhsbulk_openAccessSpreadsheet_();
  const registryName =
    typeof NHSA_REGISTRY_SHEET !== "undefined"
      ? NHSA_REGISTRY_SHEET
      : "Access Registry";

  const sheet = ss.getSheetByName(registryName);
  if (!sheet) {
    throw new Error('Access Registry sheet not found: "' + registryName + '"');
  }

  if (typeof nhsa_readRegistry_ !== "function") {
    throw new Error(
      "Missing nhsa_readRegistry_(). Add this checker to the main NHS Apps Script project containing the access-control backend."
    );
  }

  const grants = nhsa_readRegistry_(sheet) || [];
  const now = new Date();
  const activeGrants = grants.filter(function(grant) {
    if (typeof nhsa_isGrantActive_ === "function") {
      return !!nhsa_isGrantActive_(grant);
    }
    return nhsbulk_isGrantActiveFallback_(grant, now);
  });

  const usersByEmail = {};
  const issues = [];

  activeGrants.forEach(function(grant) {
    const email = String(grant.email || "").trim().toLowerCase();
    if (!email) return;

    if (!usersByEmail[email]) {
      usersByEmail[email] = {
        email: email,
        fullName: String(grant.fullName || "").trim(),
        roles: [],
        grantIds: [],
        startDates: [],
        expirationDates: []
      };
    }

    const user = usersByEmail[email];
    if (!user.fullName && grant.fullName) {
      user.fullName = String(grant.fullName).trim();
    }

    const role = String(grant.role || "").trim().toUpperCase();
    if (role && user.roles.indexOf(role) === -1) user.roles.push(role);

    const grantId = String(grant.grantId || "").trim();
    if (grantId && user.grantIds.indexOf(grantId) === -1) user.grantIds.push(grantId);

    const startDate = nhsbulk_dateLabel_(grant.startDate);
    if (startDate && user.startDates.indexOf(startDate) === -1) user.startDates.push(startDate);

    const expirationDate = nhsbulk_dateLabel_(grant.expirationDate);
    if (expirationDate && user.expirationDates.indexOf(expirationDate) === -1) {
      user.expirationDates.push(expirationDate);
    }
  });

  const users = Object.keys(usersByEmail).map(function(email) {
    const user = usersByEmail[email];
    user.roles.sort();

    const roleSet = {};
    user.roles.forEach(function(role) { roleSet[role] = true; });

    user.access = {
      gateway: user.roles.length > 0,
      teacher: !!(roleSet.TEACHER || roleSet.ADVISOR || roleSet.ADMIN),
      tutor: !!roleSet.TUTOR,
      operations: !!(roleSet.OFFICER || roleSet.ADVISOR || roleSet.ADMIN)
    };

    if (
      roleSet.TUTOR &&
      typeof nhsa_isSchoolIdentityEmail_ === "function" &&
      !nhsa_isSchoolIdentityEmail_(email)
    ) {
      issues.push("Active TUTOR grant uses a non-CUSD identity: " + email);
    }

    return user;
  });

  users.sort(function(a, b) {
    const nameA = (a.fullName || a.email).toLowerCase();
    const nameB = (b.fullName || b.email).toLowerCase();
    return nameA.localeCompare(nameB) || a.email.localeCompare(b.email);
  });

  const countsByRole = {};
  const countsByPortal = { gateway: 0, teacher: 0, tutor: 0, operations: 0 };

  users.forEach(function(user) {
    user.roles.forEach(function(role) {
      countsByRole[role] = (countsByRole[role] || 0) + 1;
    });
    Object.keys(countsByPortal).forEach(function(portal) {
      if (user.access[portal]) countsByPortal[portal]++;
    });
  });

  let deploymentIdentity = null;
  try {
    if (typeof runNHSDeploymentIdentityDiagnostic === "function") {
      deploymentIdentity = runNHSDeploymentIdentityDiagnostic();
      if (!deploymentIdentity.ok) {
        issues.push(
          "The current runtime cannot see the active user's Google identity. " +
          "Use the CUSD-domain production deployment described in the V47.7 deployment checklist."
        );
      }
    }
  } catch (e) {
    issues.push("Deployment identity diagnostic failed: " + (e && e.message ? e.message : String(e)));
  }

  return {
    generatedAt: new Date().toISOString(),
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    count: users.length,
    users: users,
    countsByRole: countsByRole,
    countsByPortal: countsByPortal,
    deploymentIdentity: deploymentIdentity,
    issues: issues,
    note:
      "This report shows NHS application-level authorization from Access Registry. " +
      "Google Apps Script deployment access is a separate security layer."
  };
}

function getNHSCurrentAccessSimple_() {
  return getNHSCurrentAccessNow_().users.map(function(user) {
    return {
      name: user.fullName,
      email: user.email,
      roles: user.roles.join(", "),
      gateway: user.access.gateway,
      teacher: user.access.teacher,
      tutor: user.access.tutor,
      operations: user.access.operations
    };
  });
}

function getNHSCurrentAccessNow() {
  nhssec_requireMaintenanceAdmin_();
  return getNHSCurrentAccessNow_();
}

function getNHSCurrentAccessSimple() {
  nhssec_requireMaintenanceAdmin_();
  return getNHSCurrentAccessSimple_();
}

function nhsbulk_openAccessSpreadsheet_() {
  if (typeof nhsa_openSpreadsheet_ === "function") return nhsa_openSpreadsheet_();

  let id = "";
  if (typeof NHS_ACCESS_SPREADSHEET_ID !== "undefined") {
    id = String(NHS_ACCESS_SPREADSHEET_ID || "").trim();
  }
  if (!id && typeof NHS_SPREADSHEET_ID !== "undefined") {
    id = String(NHS_SPREADSHEET_ID || "").trim();
  }
  if (!id) {
    throw new Error(
      "Could not determine the NHS access spreadsheet. Expected nhsa_openSpreadsheet_(), NHS_ACCESS_SPREADSHEET_ID, or NHS_SPREADSHEET_ID."
    );
  }
  return SpreadsheetApp.openById(id);
}

function nhsbulk_isGrantActiveFallback_(grant, now) {
  const status = String(grant.status || "").trim().toLowerCase();
  if (status !== "active") return false;

  const start = nhsbulk_asDate_(grant.startDate);
  const expiration = nhsbulk_asDate_(grant.expirationDate);
  if (start && start.getTime() > now.getTime()) return false;

  if (expiration) {
    const end = new Date(expiration);
    end.setHours(23, 59, 59, 999);
    if (end.getTime() < now.getTime()) return false;
  }
  return true;
}

function nhsbulk_asDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return isNaN(value.getTime()) ? null : value;
  }
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function nhsbulk_dateLabel_(value) {
  const date = nhsbulk_asDate_(value);
  if (!date) return "";

  let tz = "America/Los_Angeles";
  try {
    tz = String(Session.getScriptTimeZone() || tz).trim() || tz;
  } catch (e) {}

  return Utilities.formatDate(date, tz, "yyyy-MM-dd");
}
