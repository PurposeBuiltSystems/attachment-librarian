/*
 * Attachment Librarian — Microsoft Graph data layer.
 *
 * AUTH: Nested App Authentication (NAA) via MSAL — no backend; identical
 * pattern to the other PurposeBuilt add-ins. Only ever reads the SIGNED-IN
 * user's own mailbox (single delegated scope: Mail.Read).
 *
 * Exposes a global `GraphData` object.
 */
/* global msal */
(function (root) {
  "use strict";

  var CLIENT_ID = "c52a0066-bb9c-4711-8107-0dd9ec78c900"; // "Attachment Librarian" Entra app (purposebuilt.systems tenant)
  var GRAPH = "https://graph.microsoft.com/v1.0";
  var SCOPES = ["Mail.Read"];

  var pcaPromise = null;

  function getPca() {
    if (!pcaPromise) {
      pcaPromise = msal.createNestablePublicClientApplication({
        auth: {
          clientId: CLIENT_ID,
          authority: "https://login.microsoftonline.com/common",
        },
      });
    }
    return pcaPromise;
  }

  /**
   * Sign-in must never hang the pane. An un-timed await on the popup flow
   * leaves a button disabled with nothing visible happening — which reads
   * to the user as "the button does nothing" and gives them nothing to act
   * on. Fail loudly instead, naming the two things that actually fix it.
   */
  function withTimeout(promise, ms, message) {
    var timer;
    return Promise.race([
      promise.then(function (v) { clearTimeout(timer); return v; },
                   function (e) { clearTimeout(timer); throw e; }),
      new Promise(function (_, reject) {
        timer = setTimeout(function () { reject(new Error(message)); }, ms);
      }),
    ]);
  }

  async function getToken() {
    var pca = await withTimeout(getPca(), 20000,
      "Sign-in didn't start. Fully quit Outlook (Cmd+Q) and reopen, then try again.");
    try {
      return (await withTimeout(pca.acquireTokenSilent({ scopes: SCOPES }), 20000, "silent timeout")).accessToken;
    } catch (e) {
      var interactive = await withTimeout(
        pca.acquireTokenPopup({ scopes: SCOPES }), 120000,
        "Sign-in didn't finish. A Microsoft sign-in window may have opened behind Outlook — " +
        "check for it (or Mission Control), finish signing in, and click again. If no window " +
        "appeared at all, fully quit Outlook (Cmd+Q), reopen, and retry.");
      return interactive.accessToken;
    }
  }

  /** Page through a Graph collection following @odata.nextLink. */
  async function graphAll(token, path, cap) {
    var items = [];
    var url = GRAPH + path;
    var guard = 0;
    var maxPages = cap || 30;
    while (url && guard++ < maxPages) {
      var res = await fetch(url, { headers: { Authorization: "Bearer " + token } });
      if (!res.ok) { throw new Error("Graph GET " + url + " -> " + res.status); }
      var page = await res.json();
      items = items.concat(page.value || []);
      url = page["@odata.nextLink"] || null;
    }
    return items;
  }

  /**
   * All real file attachments across the mailbox window, flattened to the
   * shape the Librarian grouper expects. Inline images are skipped.
   * $select inside $expand keeps contentBytes out of the payload.
   */
  async function attachmentsInWindow(token, daysBack, onProgress) {
    var since = new Date(Date.now() - daysBack * 864e5).toISOString();
    var msgs = await graphAll(
      token,
      "/me/messages" +
        "?$select=id,subject,from,receivedDateTime,webLink" +
        "&$filter=hasAttachments eq true and receivedDateTime ge " + since +
        "&$expand=attachments($select=id,name,size,contentType,isInline,lastModifiedDateTime)" +
        "&$top=50"
    );
    if (onProgress) { onProgress(msgs.length); }
    var out = [];
    msgs.forEach(function (m) {
      var from = (m.from && m.from.emailAddress) || {};
      (m.attachments || []).forEach(function (a) {
        if (a["@odata.type"] && a["@odata.type"] !== "#microsoft.graph.fileAttachment") { return; }
        if (a.isInline) { return; }
        if (!a.name) { return; }
        out.push({
          name: a.name,
          size: a.size,
          contentType: a.contentType,
          date: m.receivedDateTime,
          fileModified: a.lastModifiedDateTime || null,
          fromName: from.name || from.address || "",
          fromAddress: from.address || "",
          subject: m.subject || "",
          messageId: m.id,
          webLink: m.webLink || "",
        });
      });
    });
    return out;
  }

  root.GraphData = {
    getToken: getToken,
    attachmentsInWindow: attachmentsInWindow,
    _config: { clientId: CLIENT_ID },
  };
})(typeof self !== "undefined" ? self : this);
