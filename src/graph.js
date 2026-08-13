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

  // --- add-in sign-out state -------------------------------------------
  //
  // Certification rejected the naive version on a sibling add-in: "after
  // clicking sign-out there is no response or not signed out." The reason is
  // structural. Under nested app authentication Outlook owns the session and
  // getAllAccounts() reports the HOST's account, not a cache this add-in
  // controls - so clearing MSAL's cache changes nothing visible, the next
  // silent acquisition succeeds anyway, and the pane redraws as signed in.
  //
  // A sign-out this add-in cannot deliver should not be offered. What it CAN
  // deliver is refusing to act until the user authenticates again: while
  // signed out it reports itself signed out and will not use a silent token,
  // so the next action raises a real prompt. Outlook's own session is
  // untouched, and the pane says so.
  var SIGNED_OUT_KEY = "addinSignedOut";
  var signedOut = false;
  try { signedOut = Office.context.roamingSettings.get(SIGNED_OUT_KEY) === true; } catch (e) { signedOut = false; }

  function setSignedOut(v) {
    signedOut = !!v;
    try {
      Office.context.roamingSettings.set(SIGNED_OUT_KEY, signedOut);
      Office.context.roamingSettings.saveAsync(function () {});
    } catch (e) { /* in-memory is still correct for this session */ }
  }

  /** Signed-in account, or null. Reports null while signed out, by design. */

  /**
   * Microsoft Graph throttles, and this add-in makes bursts of calls - a
   * records bundle or a bulk post is dozens to hundreds. An unretried 429
   * aborts the whole run part-way, which is the worst possible failure for
   * work that is half-written. One respectful retry honouring Retry-After
   * absorbs the overwhelming majority of throttling without hammering the
   * service; anything past that is a real outage and should surface.
   */
  async function fetchRetry(url, opts) {
    var res = await fetch(url, opts);
    if (res.status === 429 || res.status === 503) {
      var wait = Number(res.headers.get("Retry-After") || 3) * 1000;
      await new Promise(function (r) { setTimeout(r, Math.min(wait, 15000)); });
      res = await fetch(url, opts);
    }
    return res;
  }

  async function currentAccount() {
    if (signedOut) { return null; }
    try {
      var pca = await getPca();
      var accts = (pca.getAllAccounts && pca.getAllAccounts()) || [];
      return accts.length ? (accts[0].username || accts[0].name || "signed in") : null;
    } catch (e) { return null; }
  }

  /**
   * Sign out of the add-in. The state flips SYNCHRONOUSLY before any awaiting
   * so the pane can respond instantly - awaiting a broker handshake first is
   * the "no response" half of the finding. Cache clearing is best-effort on
   * top; the enforced state is what makes this real.
   */
  function signOut() {
    setSignedOut(true);
    var pending = pcaPromise;      // only clear what exists; never start a handshake here
    pcaPromise = null;
    if (!pending) { return Promise.resolve(true); }
    return Promise.resolve(pending).then(function (pca) {
      var accts = (pca && pca.getAllAccounts && pca.getAllAccounts()) || [];
      var chain = Promise.resolve();
      accts.forEach(function (a) {
        chain = chain.then(function () {
          if (pca.clearCache) { return pca.clearCache({ account: a }); }
          if (pca.logoutPopup) { return pca.logoutPopup({ account: a }); }
        }).catch(function () { /* best effort; the enforced state stands */ });
      });
      return chain.then(function () { return true; });
    }).catch(function () { return true; });
  }

  /** True while the user has signed the add-in out. */
  function isSignedOut() { return signedOut; }




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
      // Signed out means signed out: skip silent so the user must re-authenticate.
      if (signedOut) { throw new Error("signed out of the add-in"); }
      return (await withTimeout(pca.acquireTokenSilent({ scopes: SCOPES }), 20000, "silent timeout")).accessToken;
    } catch (e) {
      var interactive = await withTimeout(
        pca.acquireTokenPopup({ scopes: SCOPES }), 120000,
        "Sign-in didn't finish. A Microsoft sign-in window may have opened behind Outlook — " +
        "check for it (or Mission Control), finish signing in, and click again. If no window " +
        "appeared at all, fully quit Outlook (Cmd+Q), reopen, and retry.");
      setSignedOut(false);   // a real interactive sign-in ends the signed-out state
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
      var res = await fetchRetry(url, { headers: { Authorization: "Bearer " + token } });
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
    signOut: signOut,
    currentAccount: currentAccount,
    isSignedOut: isSignedOut,
    getToken: getToken,
    attachmentsInWindow: attachmentsInWindow,
    _config: { clientId: CLIENT_ID },
  };
})(typeof self !== "undefined" ? self : this);
