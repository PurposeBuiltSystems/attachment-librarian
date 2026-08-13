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

  /**
   * Account row. Certification policy 1100.5.7.1 requires a visible way out
   * wherever an add-in signs a user in. Every element access is guarded:
   * Outlook desktop caches the pane HTML while ?v= fetches fresh JS, so this
   * code can run against a page that predates these controls, and an
   * unguarded dereference here would throw inside Office.onReady and take
   * the whole pane down as "Add-in Error".
   */
  function authSet(id, k, v) { var e = document.getElementById(id); if (e) { e[k] = v; } }

  async function renderAuthState() {
    var who = null;
    try { who = await GraphData.currentAccount(); } catch (e) { who = null; }
    authSet("authWho", "textContent", who ? ("Signed in as " + who) : "Not signed in");
    authSet("signOut", "hidden", !who);
    authSet("signIn", "hidden", !!who);
  }

  async function doSignIn() {
    authSet("signIn", "disabled", true);
    try { await GraphData.getToken(); }
    catch (e) { authSet("authWho", "textContent", "Sign-in failed: " + ((e && e.message) || e)); }
    finally { authSet("signIn", "disabled", false); renderAuthState(); }
  }

  async function doSignOut() {
    authSet("signOut", "disabled", true);
    try {
      await GraphData.signOut();
      authSet("authWho", "textContent", "Signed out \u2014 this add-in's saved tokens are cleared. " +
        "Your Outlook session is separate and is not affected; no add-in can end it.");
    } catch (e) {
      authSet("authWho", "textContent", "Sign-out failed: " + ((e && e.message) || e));
    } finally {
      authSet("signOut", "disabled", false);
      setTimeout(renderAuthState, 2500);
    }
  }


  Office.onReady(function () {
    // Certification 1100.5.7.1 - sign-out must be reachable.
    var _si = document.getElementById("signIn");
    if (_si) { _si.addEventListener("click", doSignIn); }
    var _so = document.getElementById("signOut");
    if (_so) { _so.addEventListener("click", doSignOut); }
    renderAuthState();
    on("scan", "click", scan);
    on("search", "input", function () { render(); });
    if (window.innerWidth < 480) { rmAttrIf("options", "open"); }
  });

  function byId(id) { return document.getElementById(id); }

  /**
   * Guarded element access. Outlook desktop caches the pane HTML far harder
   * than the web client while ?v= still fetches today's JavaScript, so startup
   * routinely runs new code against an old page. One unguarded
   * `byId(x).value` there throws inside Office.onReady, and Outlook reports
   * that as "Add-in Error" - the whole pane, not one field. This is the exact
   * cause of certification finding 1120.3.7.8 on a sibling add-in.
   */
  function setVal(id, v) { var el = byId(id); if (el) { el.value = v; } }
  function setProp(id, k, v) { var el = byId(id); if (el) { el[k] = v; } }
  function setAttrIf(id, n, v) { var el = byId(id); if (el) { el.setAttribute(n, v); } }
  function rmAttrIf(id, n) { var el = byId(id); if (el) { el.removeAttribute(n); } }
  function isChecked(id) { var el = byId(id); return !!(el && el.checked); }

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

  /**
   * Open the original message.
   *
   * displayMessageForm is right on desktop, but on Outlook on the web and new
   * Outlook it can no-op WITHOUT throwing - so a try/catch around it never
   * fires, the fallback never runs, and the click does nothing and says
   * nothing. (Exactly that was reported against the sibling Waiting On
   * add-in.) So choose by host rather than by exception, keep each route as
   * the other's fallback, and never return silently.
   */
  function openMessage(v) {
    var host = "";
    try { host = (Office.context.mailbox.diagnostics || {}).hostName || ""; } catch (e) { host = ""; }
    var webFirst = /web|newoutlook/i.test(host);

    function viaForm() {
      try {
        var ewsId = Office.context.mailbox.convertToEwsId(
          v.messageId,
          Office.MailboxEnums.RestVersion.v2_0
        );
        Office.context.mailbox.displayMessageForm(ewsId);
        return true;
      } catch (e) { return false; }
    }
    function viaLink() {
      if (!v.webLink) { return false; }
      try {
        if (Office.context.ui && Office.context.ui.openBrowserWindow) {
          Office.context.ui.openBrowserWindow(v.webLink);
        } else {
          window.open(v.webLink, "_blank");
        }
        return true;
      } catch (e) { return false; }
    }

    var ok = webFirst ? (viaLink() || viaForm()) : (viaForm() || viaLink());
    if (!ok) {
      setStatus("error", "This client wouldn't open the message. Search your mail for: " +
        (v.subject || v.name || "that message"));
    }
  }

})();
