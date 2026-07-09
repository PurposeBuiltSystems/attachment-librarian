/* Offline unit tests for attachment lineage grouping. Run: npm test */
"use strict";
var L = require("../src/librarian.js");

var failures = 0;
function check(label, actual, expected) {
  if (actual !== expected) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) + "\n  actual:   " + JSON.stringify(actual));
  }
}

function att(name, date, from, subject) {
  return { name: name, size: 1000, date: date, fromName: from || "Ann Lee", fromAddress: "ann@x.com", subject: subject || "FY27 planning", messageId: "m", webLink: "" };
}

// Versions of the same doc group together, newest first
var groups = L.groupAttachments([
  att("Budget_v1.xlsx", "2026-06-01T00:00:00Z"),
  att("Budget_v2.xlsx", "2026-06-15T00:00:00Z"),
  att("Budget FINAL.xlsx", "2026-07-01T00:00:00Z"),
  att("Budget (1).xlsx", "2026-06-20T00:00:00Z"),
  att("Budget.pdf", "2026-06-10T00:00:00Z"),           // different ext = different lineage
  att("Site Photos 2026-06-05.zip", "2026-06-05T00:00:00Z"),
]);

check("three lineages", groups.length, 3);
check("xlsx lineage has 4 versions", groups[0].versions.length, 4);
check("newest first in group", groups[0].versions[0].name, "Budget FINAL.xlsx");
check("group display name = newest filename", groups[0].displayName, "Budget FINAL.xlsx");
check("groups sorted by latest desc", groups[0].latest, "2026-07-01T00:00:00Z");
check("pdf separate", groups.some(function (g) { return g.ext === "pdf" && g.versions.length === 1; }), true);

// lineage keys
check("v-token stripped", L.lineageKey("Report_v3.docx"), "report|docx");
check("final stripped", L.lineageKey("Report FINAL draft.docx"), "report|docx");
check("date stripped", L.lineageKey("Report 2026-07-09.docx"), "report|docx");
check("copy counter stripped", L.lineageKey("Report (2).docx"), "report|docx");
check("all-version name survives", L.lineageKey("v2.docx"), "v2|docx");
check("different docs differ", L.lineageKey("Budget.xlsx") === L.lineageKey("Roster.xlsx"), false);

// filter: by name and by sender
var f1 = L.filterGroups(groups, "budget");
check("filter by name", f1.length, 2);
var f2 = L.filterGroups(groups, "ann lee");
check("filter by sender", f2.length, 3);
var f3 = L.filterGroups(groups, "zzz");
check("filter no match", f3.length, 0);

// size formatting
check("size KB", L.fmtSize(2048), "2 KB");
check("size MB", L.fmtSize(3 * 1048576), "3.0 MB");

if (failures) {
  console.error("\n" + failures + " librarian test(s) FAILED");
  process.exit(1);
}
console.log("All librarian tests passed.");
