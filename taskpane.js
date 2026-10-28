/* SMKPB Mail Merge task pane. Logic lives in merge.js; this file wires it to Outlook and Microsoft Graph. */
(function () {
  "use strict";

  const MM = window.MailMerge;
  const CONFIG = window.MAIL_MERGE_CONFIG || {};
  const GRAPH_SEND_URL = "https://graph.microsoft.com/v1.0/me/sendMail";
  const $ = (id) => document.getElementById(id);

  const state = {
    table: { headers: [], rows: [] },
    built: { emails: [], errors: [], warnings: [] },
    previewIndex: 0,
    attachment: null,
    lastField: null,
    sending: false,
    stopRequested: false,
    lastResults: [],
  };

  class UserError extends Error {}

  // ─── STORAGE ──────────────────────────────────────────────────────────────
  // Small settings (client ID, signature) roam with the mailbox; drafts stay in this browser.

  function roaming() {
    try { return window.Office && Office.context && Office.context.roamingSettings; } catch { return null; }
  }
  const local = {
    get(key) { try { return JSON.parse(localStorage.getItem(`mm.${key}`)); } catch { return null; } },
    set(key, value) { try { localStorage.setItem(`mm.${key}`, JSON.stringify(value)); } catch { /* storage unavailable */ } },
    remove(key) { try { localStorage.removeItem(`mm.${key}`); } catch { /* ignore */ } },
  };
  let roamingSaveTimer = null;
  const settings = {
    get(key) {
      const rs = roaming();
      if (rs) {
        const v = rs.get(key);
        if (v !== undefined && v !== null) return v;
      }
      return local.get(key);
    },
    set(key, value) {
      local.set(key, value);
      const rs = roaming();
      if (!rs) return;
      rs.set(key, value);
      clearTimeout(roamingSaveTimer);
      roamingSaveTimer = setTimeout(() => { try { rs.saveAsync(() => {}); } catch { /* ignore */ } }, 800);
    },
  };

  function myAddress() {
    try { return Office.context.mailbox.userProfile.emailAddress || ""; } catch { return ""; }
  }

  // ─── SIGN-IN (nested app authentication) ──────────────────────────────────

  let pca = null;
  let pcaClientId = null;
  let signedIn = false;

  function clientId() {
    return String(settings.get("clientId") || CONFIG.clientId || "").trim();
  }

  function naaSupported() {
    try { return Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"); } catch { return false; }
  }

  function authority() {
    // Signing in against the mailbox's own domain works whether the app is single- or multi-tenant.
    const domain = myAddress().split("@")[1];
    return `https://login.microsoftonline.com/${domain || "organizations"}`;
  }

  async function getToken({ interactive }) {
    const id = clientId();
    if (!id) throw new UserError("Add your Application (client) ID under Settings at the bottom first.");
    if (!naaSupported()) {
      throw new UserError("This version of Outlook can't sign in add-ins. Use Outlook on the web (outlook.office.com) or the new Outlook.");
    }
    if (!pca || pcaClientId !== id) {
      pca = await msal.createNestablePublicClientApplication({
        auth: { clientId: id, authority: authority() },
        cache: { cacheLocation: "localStorage" },
      });
      pcaClientId = id;
    }
    const request = { scopes: ["Mail.Send"] };
    let result;
    try {
      result = await pca.acquireTokenSilent(request);
    } catch (err) {
      if (!interactive) throw err;
      result = await pca.acquireTokenPopup(request);
    }
    signedIn = true;
    renderAccount(result && result.account && result.account.username);
    return result.accessToken;
  }

  function renderAccount(username) {
    const who = username || myAddress();
    $("account").textContent = signedIn ? `Sending as ${who}` : (who ? `${who} · not signed in` : "Not signed in");
  }

  function friendlyAuthError(err) {
    if (err instanceof UserError) return err.message;
    const code = (err && (err.errorCode || err.code)) || "";
    const msg = (err && err.message) || String(err);
    if (/user_cancelled|popup_window_error|cancel/i.test(code + msg)) return "Sign-in was cancelled.";
    if (/AADSTS700016/.test(msg)) return "Microsoft doesn't recognize that Application (client) ID. Check it under Settings.";
    if (/AADSTS50011|redirect/i.test(msg)) return "The app registration is missing this add-in's redirect URI. See step 3 of the setup guide.";
    if (/AADSTS65001|consent/i.test(msg)) return "Permission to send mail hasn't been granted yet. See step 3 of the setup guide.";
    return `Sign-in failed: ${msg}`;
  }

  // ─── GRAPH ────────────────────────────────────────────────────────────────

  async function graphSend(email, token) {
    const res = await fetch(GRAPH_SEND_URL, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(MM.graphMessage(email, state.attachment)),
    });
    if (res.ok) return { status: res.status };
    let detail = "";
    try { const j = await res.json(); detail = (j.error && j.error.message) || ""; } catch { /* not JSON */ }
    const retryAfter = parseInt(res.headers.get("Retry-After"), 10);
    return {
      status: res.status,
      retryAfter: Number.isFinite(retryAfter) ? retryAfter : undefined,
      error: friendlyGraphError(res.status, detail),
    };
  }

  function friendlyGraphError(status, detail) {
    if (status === 401) return "Sign-in expired. Use Retry failed to sign in again.";
    if (status === 403) return "Microsoft refused to send (check the Mail.Send permission).";
    if (status === 413) return "Email too large. Use a smaller attachment.";
    if (status === 429) return "Microsoft is rate-limiting this mailbox. Try again in a few minutes.";
    return detail || `Error ${status}`;
  }

  // ─── TABLE ────────────────────────────────────────────────────────────────

  function fillSelect(select, headers, selected, allowNone) {
    select.innerHTML = "";
    if (allowNone) select.add(new Option("None", ""));
    headers.forEach((h) => select.add(new Option(h, h)));
    select.value = selected && headers.includes(selected) ? selected : (allowNone ? "" : (headers[0] || ""));
  }

  function onTableChanged() {
    const text = $("table-input").value;
    local.set("table", text);
    const table = MM.parseTable(text);
    const prev = { email: $("email-column").value, cc: $("cc-column").value, bcc: $("bcc-column").value };
    state.table = table;
    state.previewIndex = 0;

    const hasTable = table.headers.length > 0;
    $("table-summary").hidden = !hasTable;
    if (hasTable) {
      const emailGuess = table.headers.includes(prev.email) ? prev.email : MM.findEmailColumn(table.headers);
      fillSelect($("email-column"), table.headers, emailGuess, true);
      fillSelect($("cc-column"), table.headers,
        table.headers.includes(prev.cc) ? prev.cc : MM.findColumn(table.headers, /^cc$/i), true);
      fillSelect($("bcc-column"), table.headers,
        table.headers.includes(prev.bcc) ? prev.bcc : MM.findColumn(table.headers, /^bcc$/i), true);
      const n = table.rows.length;
      $("table-count").textContent = `${n} ${n === 1 ? "person" : "people"} · ${table.headers.length} columns`;
    }
    renderChips();
    refresh();
  }

  function renderChips() {
    const wrap = $("field-chips");
    wrap.querySelectorAll(".chip").forEach((c) => c.remove());
    const headers = state.table.headers;
    wrap.hidden = headers.length === 0;
    headers.forEach((h) => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "chip";
      chip.textContent = h;
      chip.title = `Insert {{${h}}}`;
      chip.addEventListener("mousedown", (e) => e.preventDefault()); // keep the cursor where it is
      chip.addEventListener("click", () => insertField(h));
      wrap.appendChild(chip);
    });
  }

  function insertField(name) {
    const target = state.lastField || $("body");
    const token = `{{${name}}}`;
    const start = target.selectionStart ?? target.value.length;
    const end = target.selectionEnd ?? target.value.length;
    target.value = target.value.slice(0, start) + token + target.value.slice(end);
    target.focus();
    target.selectionStart = target.selectionEnd = start + token.length;
    target.dispatchEvent(new Event("input", { bubbles: true }));
  }

  async function readTextFile(file) {
    const buf = await file.arrayBuffer();
    try {
      return new TextDecoder("utf-8", { fatal: true }).decode(buf);
    } catch {
      // Excel on Windows often saves CSV as Windows-1252.
      return new TextDecoder("windows-1252").decode(buf);
    }
  }

  // ─── ATTACHMENT ───────────────────────────────────────────────────────────

  function readAsBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result).split(",")[1] || "");
      reader.onerror = () => reject(reader.error);
      reader.readAsDataURL(file);
    });
  }

  async function onAttachmentChosen(file) {
    const max = CONFIG.maxAttachmentBytes || 2.5 * 1024 * 1024;
    if (!file) return;
    if (file.size > max) {
      state.attachment = null;
      $("attachment-status").textContent = `${file.name} is ${(file.size / 1048576).toFixed(1)} MB. The limit is ${(max / 1048576).toFixed(1)} MB.`;
      $("remove-attachment").hidden = true;
      return;
    }
    state.attachment = {
      name: file.name,
      contentType: file.type || "application/octet-stream",
      contentBytes: await readAsBase64(file),
      size: file.size,
    };
    $("attachment-status").textContent = `${file.name} (${Math.max(1, Math.round(file.size / 1024))} KB) goes to everyone.`;
    $("remove-attachment").hidden = false;
  }

  // ─── BUILD + PREVIEW ──────────────────────────────────────────────────────

  let refreshTimer = null;
  function scheduleRefresh() {
    clearTimeout(refreshTimer);
    refreshTimer = setTimeout(refresh, 150);
  }

  function currentContext() {
    return {
      headers: state.table.headers,
      rows: state.table.rows,
      emailColumn: $("email-column").value || null,
      ccColumn: $("cc-column").value || null,
      bccColumn: $("bcc-column").value || null,
      subject: $("subject").value,
      body: $("body").value,
      signature: $("signature").value,
    };
  }

  function sendable() {
    return state.built.emails.filter((e) => e.sendable);
  }

  function refresh() {
    state.built = MM.buildEmails(currentContext());
    const { emails, errors, warnings } = state.built;

    const box = $("messages");
    box.innerHTML = "";
    const started = state.table.rows.length || $("subject").value || $("body").value;
    if (started) {
      errors.forEach((m) => box.appendChild(messageEl("error", m)));
      warnings.forEach((m) => box.appendChild(messageEl("warn", m)));
    } else {
      box.appendChild(messageEl("info", "Paste a table and write your message to see each person's email here."));
    }

    if (state.previewIndex >= emails.length) state.previewIndex = Math.max(0, emails.length - 1);
    renderPreview();
    renderSendButtons();
    renderHistory();
  }

  function messageEl(kind, text) {
    const div = document.createElement("div");
    div.className = `msg ${kind}`;
    div.textContent = text;
    return div;
  }

  function renderPreview() {
    const emails = state.built.emails;
    const show = emails.length > 0 && ($("subject").value || $("body").value);
    $("preview").hidden = !show;
    if (!show) return;
    const e = emails[state.previewIndex];
    $("pager-label").textContent = `Person ${state.previewIndex + 1} of ${emails.length} · row ${e.index + 2}`;
    $("prev").disabled = state.previewIndex === 0;
    $("next").disabled = state.previewIndex >= emails.length - 1;
    $("pv-to").textContent = e.to.join("; ") || "(none)";
    $("pv-cc-line").hidden = !e.cc.length;
    $("pv-cc").textContent = e.cc.join("; ");
    $("pv-bcc-line").hidden = !e.bcc.length;
    $("pv-bcc").textContent = e.bcc.join("; ");
    $("pv-subject").textContent = e.subject || "(no subject)";

    const issues = [];
    if (e.problems.length) issues.push(`Skipped: ${e.problems.join("; ")}.`);
    if (e.blanks.length) issues.push(`Blank: ${e.blanks.join(", ")}.`);
    $("pv-issue").hidden = !issues.length;
    $("pv-issue").textContent = issues.join(" ");
    $("pv-body").srcdoc = e.html;
  }

  function renderSendButtons() {
    const list = sendable();
    const history = loadHistory();
    const remaining = list.filter((e) => !history.has(MM.recipientKey(e))).length;
    const blocked = state.built.errors.length > 0 || list.length === 0;
    const btn = $("send-all");
    if (!list.length) btn.textContent = "Send";
    else if (!remaining) btn.textContent = "Everyone has been sent this";
    else btn.textContent = `Send to ${remaining} ${remaining === 1 ? "person" : "people"}`;
    btn.disabled = blocked || remaining === 0 || state.sending;
    $("send-test").disabled = blocked || state.sending;
  }

  // ─── SENT HISTORY (so a re-run never double-sends) ────────────────────────

  function historyKey() {
    return `sent.${MM.messageKey($("subject").value, $("body").value)}`;
  }
  function loadHistory() {
    return new Set(local.get(historyKey()) || []);
  }
  function saveHistory(set) {
    local.set(historyKey(), [...set]);
  }
  function renderHistory() {
    const history = loadHistory();
    const already = sendable().filter((e) => history.has(MM.recipientKey(e))).length;
    const p = $("sent-history");
    p.hidden = already === 0 || state.sending;
    if (already) {
      p.innerHTML = "";
      p.append(`${already} ${already === 1 ? "person has" : "people have"} already been sent this message from this computer and will be skipped. `);
      const a = document.createElement("button");
      a.type = "button";
      a.className = "btn btn-link";
      a.textContent = "Send to them again";
      a.addEventListener("click", () => { local.remove(historyKey()); renderHistory(); renderSendButtons(); setStatus(""); });
      p.appendChild(a);
    }
  }

  // ─── SENDING ──────────────────────────────────────────────────────────────

  function setStatus(text, kind) {
    const s = $("status");
    s.textContent = text;
    s.className = `status${kind ? ` ${kind}` : ""}`;
  }

  async function onSendTest() {
    const me = myAddress();
    const e = state.built.emails[state.previewIndex];
    if (!me || !e) return;
    setStatus("Sending a test to you…");
    try {
      const token = await getToken({ interactive: true });
      const test = { ...e, to: [me], cc: [], bcc: [], subject: `[Test] ${e.subject}` };
      const result = await graphSend(test, token);
      if (result.status >= 200 && result.status < 300) setStatus(`Test sent to ${me} (person ${state.previewIndex + 1}'s version).`, "ok");
      else setStatus(`Test failed: ${result.error}`, "error");
    } catch (err) {
      setStatus(friendlyAuthError(err), "error");
    }
  }

  function onSendAllClicked() {
    const list = sendable();
    const history = loadHistory();
    const toSend = list.filter((e) => !history.has(MM.recipientKey(e))).length;
    if (!toSend) {
      setStatus("Everyone on this list has already been sent this message.", "error");
      return;
    }
    const minutes = Math.ceil((toSend * (CONFIG.delayMs || 2400)) / 60000);
    const from = myAddress() ? ` from ${myAddress()}` : "";
    $("confirm-text").textContent =
      `Send ${toSend} ${toSend === 1 ? "email" : "emails"}${from}? ` +
      `${minutes > 1 ? `This takes about ${minutes} minutes; keep this panel open. ` : ""}Sent emails can't be recalled.`;
    $("confirm").hidden = false;
    $("send-actions").hidden = true;
  }

  async function runSend(list) {
    $("confirm").hidden = true;
    let token;
    try {
      token = await getToken({ interactive: true });
    } catch (err) {
      $("send-actions").hidden = false;
      setStatus(friendlyAuthError(err), "error");
      return;
    }

    state.sending = true;
    state.stopRequested = false;
    setLocked(true);
    $("send-actions").hidden = true;
    $("progress").hidden = false;
    $("results").hidden = true;
    $("stop").disabled = false;
    setStatus("");
    renderSendButtons();
    renderHistory();

    const history = loadHistory();
    const keyAtStart = historyKey();
    let sent = 0;
    let failed = 0;

    const results = await MM.sendAll(list, {
      delayMs: CONFIG.delayMs || 2400,
      alreadySent: history,
      shouldStop: () => state.stopRequested,
      send: async (email) => {
        try {
          token = await getToken({ interactive: false });
        } catch {
          return { status: 401, error: "Sign-in expired. Use Retry failed to sign in again." };
        }
        return graphSend(email, token);
      },
      onProgress: (r, done, total) => {
        if (r.status === "sent") { sent++; local.set(keyAtStart, [...history]); }
        if (r.status === "failed") failed++;
        $("progress-fill").style.width = `${Math.round((done / total) * 100)}%`;
        $("progress-text").textContent = `${done} of ${total} · ${sent} sent${failed ? ` · ${failed} failed` : ""}`;
      },
    });

    state.sending = false;
    setLocked(false);
    $("progress").hidden = true;
    $("send-actions").hidden = false;
    showResults(results);
    renderSendButtons();
    renderHistory();
  }

  function showResults(results) {
    state.lastResults = results;
    const count = (s) => results.filter((r) => r.status === s).length;
    const sent = count("sent");
    const failed = count("failed");
    const skipped = count("skipped");
    const stopped = count("stopped");
    const parts = [`${sent} sent`];
    if (failed) parts.push(`${failed} failed`);
    if (skipped) parts.push(`${skipped} skipped (already sent)`);
    if (stopped) parts.push(`${stopped} not sent (stopped)`);
    $("results-summary").textContent = parts.join(" · ");

    const ul = $("results-list");
    ul.innerHTML = "";
    // Problems first, so they're visible without scrolling.
    const order = { failed: 0, stopped: 1, skipped: 2, sent: 3 };
    [...results].sort((a, b) => order[a.status] - order[b.status] || a.email.index - b.email.index).forEach((r) => {
      const li = document.createElement("li");
      const who = document.createElement("span");
      who.className = "who";
      who.textContent = `Row ${r.email.index + 2} · ${r.email.to.join("; ")}`;
      who.title = r.detail;
      const st = document.createElement("span");
      st.className = r.status;
      st.textContent = r.status === "failed" ? r.detail : r.status[0].toUpperCase() + r.status.slice(1);
      li.append(who, st);
      ul.appendChild(li);
    });
    $("retry-failed").hidden = !(failed || stopped);
    $("results").hidden = false;
    setStatus(failed ? "Some emails didn't send. You can retry them." : (sent ? "Done." : ""), failed ? "error" : (sent ? "ok" : ""));
  }

  function setLocked(locked) {
    ["table-input", "csv-file", "email-column", "cc-column", "bcc-column", "subject", "body", "signature", "attachment-file", "clear-table"]
      .forEach((id) => { $(id).disabled = locked; });
  }

  async function copyText(text) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      const ta = document.createElement("textarea");
      ta.value = text;
      ta.style.position = "fixed";
      ta.style.opacity = "0";
      document.body.appendChild(ta);
      ta.select();
      let ok = false;
      try { ok = document.execCommand("copy"); } catch { ok = false; }
      ta.remove();
      return ok;
    }
  }

  // ─── WIRING ───────────────────────────────────────────────────────────────

  function init() {
    // Restore drafts and settings.
    $("table-input").value = local.get("table") || "";
    $("subject").value = local.get("subject") || "";
    $("body").value = local.get("body") || "";
    $("signature").value = settings.get("signature") || "";
    $("client-id").value = clientId();
    if (!clientId()) $("settings").open = true;
    renderAccount();

    $("table-input").addEventListener("input", () => { clearTimeout(refreshTimer); refreshTimer = setTimeout(onTableChanged, 200); });
    $("csv-file").addEventListener("change", async (ev) => {
      const file = ev.target.files[0];
      if (!file) return;
      $("table-input").value = await readTextFile(file);
      ev.target.value = "";
      onTableChanged();
    });
    $("clear-table").addEventListener("click", () => { $("table-input").value = ""; onTableChanged(); });
    ["email-column", "cc-column", "bcc-column"].forEach((id) => $(id).addEventListener("change", refresh));

    ["subject", "body", "signature"].forEach((id) => {
      const el = $(id);
      el.addEventListener("focus", () => { state.lastField = el; });
      el.addEventListener("input", () => {
        if (id === "signature") settings.set("signature", el.value);
        else local.set(id, el.value);
        scheduleRefresh();
      });
    });

    $("attachment-file").addEventListener("change", (ev) => onAttachmentChosen(ev.target.files[0]));
    $("remove-attachment").addEventListener("click", () => {
      state.attachment = null;
      $("attachment-file").value = "";
      $("remove-attachment").hidden = true;
      $("attachment-status").textContent = "Optional. The same file goes to everyone (up to 2.5 MB).";
    });

    $("prev").addEventListener("click", () => { state.previewIndex = Math.max(0, state.previewIndex - 1); renderPreview(); });
    $("next").addEventListener("click", () => {
      state.previewIndex = Math.min(state.built.emails.length - 1, state.previewIndex + 1);
      renderPreview();
    });

    $("send-test").addEventListener("click", onSendTest);
    $("send-all").addEventListener("click", onSendAllClicked);
    $("confirm-no").addEventListener("click", () => { $("confirm").hidden = true; $("send-actions").hidden = false; });
    $("confirm-yes").addEventListener("click", () => runSend(sendable()));
    $("stop").addEventListener("click", () => {
      state.stopRequested = true;
      $("stop").disabled = true;
      $("progress-text").textContent += " · stopping…";
    });
    $("retry-failed").addEventListener("click", () => {
      const again = state.lastResults.filter((r) => r.status === "failed" || r.status === "stopped").map((r) => r.email);
      if (again.length) runSend(again);
    });
    $("copy-results").addEventListener("click", async () => {
      const ok = await copyText(MM.resultsToTsv(state.lastResults));
      setStatus(ok ? "Results copied. Paste them into Excel." : "Couldn't copy. Select the list and copy it by hand.", ok ? "ok" : "error");
    });

    $("save-settings").addEventListener("click", () => {
      settings.set("clientId", $("client-id").value.trim());
      pca = null;
      signedIn = false;
      renderAccount();
      setStatus("Settings saved.", "ok");
    });
    $("sign-in").addEventListener("click", async () => {
      const typed = $("client-id").value.trim();
      if (typed && typed !== clientId()) settings.set("clientId", typed);
      try {
        await getToken({ interactive: true });
        setStatus("Signed in. You're ready to send.", "ok");
      } catch (err) {
        setStatus(friendlyAuthError(err), "error");
      }
    });

    onTableChanged();
  }

  if (window.Office && Office.onReady) Office.onReady(() => init());
  else document.addEventListener("DOMContentLoaded", init);
})();
