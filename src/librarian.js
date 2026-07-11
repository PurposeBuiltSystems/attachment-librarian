/*
 * Attachment Librarian — file-lineage grouping (pure logic, no Office/Graph).
 *
 * Takes a flat list of attachments seen across the mailbox and groups the
 * versions of the "same" document together — "Budget_v2.xlsx",
 * "Budget final.xlsx", and "Budget (1).xlsx" are one lineage — newest first,
 * so "which one is the latest?" has an answer.
 *
 * Deterministic name normalization, no AI. Works in the browser (global
 * `Librarian`) and in Node (module.exports) for offline tests.
 */
(function (root) {
  "use strict";

  // Tokens that mark a version/copy rather than a different document.
  // Applied AFTER separators (._-) are normalized to spaces, so \b works.
  var VERSION_TOKENS = [
    /\(\d+\)/g,                                     // (1) (2) copy counters
    /\bv(er(sion)?)? ?\d+[a-z]?\b/gi,               // v2, ver3, version 4, v2b
    /\brev(ised|ision)? ?\d*\b/gi,                  // rev, rev2, revised
    /\b(final|draft|updated?|copy|edit(ed)?|clean|redline)\b/gi,
    /\b\d{4} \d{1,2} \d{1,2}\b/g,                   // 2026-07-09 (post-normalize)
    /\b\d{1,2} \d{1,2} (\d{4}|\d{2})\b/g,           // 7-9-2026, 07.09.26
  ];

  function splitName(filename) {
    var name = String(filename || "");
    var dot = name.lastIndexOf(".");
    if (dot <= 0) { return { base: name, ext: "" }; }
    return { base: name.slice(0, dot), ext: name.slice(dot + 1).toLowerCase() };
  }

  /** "Budget_v2 FINAL (3).xlsx" -> "budget" (+ ext kept separately). */
  function lineageKey(filename) {
    var parts = splitName(filename);
    var base = parts.base.toLowerCase().replace(/[._\-]+/g, " "); // separators first, so \b works
    VERSION_TOKENS.forEach(function (re) { base = base.replace(re, " "); });
    base = base.replace(/\s+/g, " ").trim();
    if (!base) { base = parts.base.toLowerCase(); } // name was ALL version tokens
    return base + "|" + parts.ext;
  }

  /**
   * @param atts [{name, size, contentType, date (ISO), fromName, fromAddress,
   *               subject, messageId, webLink}]
   * @returns groups sorted by latest date desc:
   *   [{key, displayName, ext, versions: [attachment, ...newest first], latest}]
   */
  function groupAttachments(atts) {
    var byKey = {};
    (atts || []).forEach(function (a) {
      if (!a || !a.name) { return; }
      var key = lineageKey(a.name);
      (byKey[key] = byKey[key] || []).push(a);
    });
    var groups = Object.keys(byKey).map(function (key) {
      var versions = byKey[key].sort(function (x, y) {
        return (Date.parse(y.date || 0) || 0) - (Date.parse(x.date || 0) || 0);
      });
      // Same exact name + size (+ file-modified date when present) = the same
      // file re-sent, not a new version. Newest copy stays the "version";
      // older identical copies get flagged as re-sends.
      var seen = {};
      var distinct = 0;
      versions.forEach(function (v) {
        var sig = v.name.toLowerCase() + "|" + (v.size || 0) + "|" + (v.fileModified || "");
        if (seen[sig]) {
          v.resend = true;
        } else {
          seen[sig] = true;
          v.resend = false;
          distinct++;
        }
      });
      return {
        key: key,
        displayName: versions[0].name,   // most recent original filename
        ext: splitName(versions[0].name).ext,
        versions: versions,
        distinctVersions: distinct,
        latest: versions[0].date,
      };
    });
    groups.sort(function (a, b) {
      return (Date.parse(b.latest || 0) || 0) - (Date.parse(a.latest || 0) || 0);
    });
    return groups;
  }

  /** Case-insensitive filter across filename, sender, and subject. */
  function filterGroups(groups, query) {
    var q = String(query || "").toLowerCase().trim();
    if (!q) { return groups; }
    return groups.filter(function (g) {
      if (g.displayName.toLowerCase().indexOf(q) !== -1) { return true; }
      return g.versions.some(function (v) {
        return (v.fromName || "").toLowerCase().indexOf(q) !== -1 ||
               (v.fromAddress || "").toLowerCase().indexOf(q) !== -1 ||
               (v.subject || "").toLowerCase().indexOf(q) !== -1;
      });
    });
  }

  function fmtSize(bytes) {
    var n = Number(bytes) || 0;
    if (n >= 1048576) { return (n / 1048576).toFixed(1) + " MB"; }
    if (n >= 1024) { return Math.round(n / 1024) + " KB"; }
    return n + " B";
  }

  var api = {
    groupAttachments: groupAttachments,
    filterGroups: filterGroups,
    lineageKey: lineageKey,
    fmtSize: fmtSize,
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.Librarian = api; }
})(typeof self !== "undefined" ? self : this);
