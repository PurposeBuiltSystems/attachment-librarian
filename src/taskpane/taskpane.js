/*
 * Attachment Librarian — task pane UI wiring.
 *
 * Index: pull attachment metadata via GraphData, group into document lineages
 * via the pure Librarian module, render with live search. Per version:
 * "Open email" opens the message that carried it (webLink fallback).
 */
/* global Office, GraphData, Librarian, document */
(function () {
  "use strict";

  var groups = [];

  Office.onReady(function () {
    on("scan", "click", scan);
    on("search", "input", function () { render(); });
    if (window.innerWidth < 480) { byId("options").removeAttribute("open"); }
  });

  function byId(id) { return document.getElementById(id); }

  /**
   * Outlook caches the pane HTML but the ?v= query string makes it fetch
   * JavaScript fresh, so a returning user can run today's JS against
   * yesterday's page. Binding through this helper means a missing element
   * costs one feature instead of throwing and leaving every later button
   * unbound — a whole dead pane.
   */
  function on(id, ev, fn) {
    var el = byId(id);
    if (el) { el.addEventListener(ev, fn); }
    return el;
  }

  function setStatus(kind, text) {
    var el = byId("status");
    if (!text) { el.hidden = true; return; }
    el.hidden = false;
    el.className = "status " + kind;
    el.textContent = text;
  }

  async function scan() {
    var daysBack = Math.max(7, Math.min(365, parseInt(byId("daysBack").value, 10) || 90));
    byId("scan").disabled = true;
    byId("results").hidden = true;
    try {
      setStatus("work", "Indexing attachments from the last " + daysBack + " days…");
      var token = await GraphData.getToken();
      var atts = await GraphData.attachmentsInWindow(token, daysBack, function (n) {
        setStatus("work", "Reading " + n + " messages with attachments…");
      });
      groups = Librarian.groupAttachments(atts);
      if (!groups.length) {
        setStatus("info", "No file attachments found in the last " + daysBack + " days.");
        return;
      }
      render();
      byId("results").hidden = false;
      setStatus("info", atts.length + " attachments in " + groups.length + " documents. Newest first — search to narrow.");
    } catch (e) {
      var msg = (e && e.message) || String(e);
      if (/REPLACE_WITH_ENTRA_CLIENT_ID/.test(GraphData._config.clientId)) {
        msg = "Set your Entra client ID in src/graph.js before running. (" + msg + ")";
      }
      setStatus("error", "Indexing failed: " + msg);
    } finally {
      byId("scan").disabled = false;
    }
  }

  function fmtDate(iso) {
    var d = new Date(iso);
    return isNaN(d) ? "" : d.toLocaleDateString();
  }

  function render() {
    var host = byId("groups");
    host.innerHTML = "";
    var shown = Librarian.filterGroups(groups, byId("search").value);
    if (!shown.length) {
      var none = document.createElement("div");
      none.className = "muted";
      none.textContent = "No matches.";
      host.appendChild(none);
      return;
    }
    shown.slice(0, 200).forEach(function (g) {
      var det = document.createElement("details");
      det.className = "item";
      var sum = document.createElement("summary");
      sum.className = "top";
      var name = document.createElement("span");
      name.className = "subject";
      name.textContent = g.displayName;
      var meta = document.createElement("span");
      meta.className = "age";
      var counts = "";
      if (g.versions.length > 1) {
        counts = g.distinctVersions === g.versions.length
          ? g.versions.length + " versions · "
          : g.distinctVersions + " version" + (g.distinctVersions > 1 ? "s" : "") +
            " · " + g.versions.length + " copies · ";
      }
      meta.textContent = counts + fmtDate(g.latest);
      sum.appendChild(name);
      sum.appendChild(meta);
      det.appendChild(sum);

      g.versions.forEach(function (v) {
        var row = document.createElement("div");
        row.className = "who";
        var line = v.name + " — " + Librarian.fmtSize(v.size) + " — " +
          (v.fromName || "?") + " — " + fmtDate(v.date);
        // Show the file's own modified date when it meaningfully differs from
        // the email date (email doesn't always preserve it, but when it does
        // it's the real "change date").
        if (v.fileModified) {
          var diff = Math.abs((Date.parse(v.fileModified) || 0) - (Date.parse(v.date) || 0));
          if (diff > 864e5) { line += " — file modified " + fmtDate(v.fileModified); }
        }
        if (v.resend) { line += " — same file (re-sent)"; }
        row.textContent = line + "  ";
        var open = document.createElement("button");
        open.textContent = "Open email";
        open.addEventListener("click", function () { openMessage(v); });
        row.appendChild(open);
        det.appendChild(row);
      });
      host.appendChild(det);
    });
    if (shown.length > 200) {
      var more = document.createElement("div");
      more.className = "muted";
      more.textContent = "Showing the newest 200 documents — search to narrow further.";
      host.appendChild(more);
    }
  }

  function openMessage(v) {
    // Preferred: open in this Outlook client. Fallback: web link.
    try {
      var ewsId = Office.context.mailbox.convertToEwsId(
        v.messageId,
        Office.MailboxEnums.RestVersion.v2_0
      );
      Office.context.mailbox.displayMessageForm(ewsId);
      return;
    } catch (e) { /* fall through */ }
    if (v.webLink) {
      try {
        if (Office.context.ui && Office.context.ui.openBrowserWindow) {
          Office.context.ui.openBrowserWindow(v.webLink);
        } else {
          window.open(v.webLink, "_blank");
        }
      } catch (e2) {
        setStatus("error", "Could not open the message.");
      }
    }
  }
})();
