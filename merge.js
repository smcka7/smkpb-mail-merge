/*
 * SMKPB Mail Merge: pure logic (no Office or network code), so it can be tested on its own.
 * Loaded in the task pane as window.MailMerge, and in Node via require().
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.MailMerge = api;
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // ─── TABLE PARSING ────────────────────────────────────────────────────────

  /** Pick the delimiter from the header line: tab (pasted from Excel), then ; or , (CSV). */
  function detectDelimiter(text) {
    const firstLine = (text.split(/\r?\n/).find((l) => l.trim() !== "") || "");
    if (firstLine.includes("\t")) return "\t";
    const semis = (firstLine.match(/;/g) || []).length;
    const commas = (firstLine.match(/,/g) || []).length;
    return semis > commas ? ";" : ",";
  }

  /** RFC 4180-style parser: quoted fields, doubled quotes, delimiters and newlines inside quotes. */
  function parseRows(text, delimiter) {
    const rows = [];
    let row = [];
    let field = "";
    let inQuotes = false;
    let i = 0;
    const n = text.length;
    while (i < n) {
      const c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
          inQuotes = false; i++; continue;
        }
        field += c; i++; continue;
      }
      if (c === '"' && field === "") { inQuotes = true; i++; continue; }
      if (c === delimiter) { row.push(field); field = ""; i++; continue; }
      if (c === "\r") { i++; continue; }
      if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; i++; continue; }
      field += c; i++;
    }
    if (field !== "" || row.length) { row.push(field); rows.push(row); }
    return rows;
  }

  /**
   * Parse a pasted or uploaded table. The first non-empty row is the header.
   * Returns { headers: [..], rows: [{header: value}], delimiter }.
   */
  function parseTable(text) {
    const clean = String(text || "").replace(/^﻿/, "");
    if (!clean.trim()) return { headers: [], rows: [], delimiter: "," };
    const delimiter = detectDelimiter(clean);
    const raw = parseRows(clean, delimiter).filter((r) => r.some((v) => v.trim() !== ""));
    if (!raw.length) return { headers: [], rows: [], delimiter };

    const seen = {};
    const headers = raw[0].map((h, idx) => {
      let name = h.trim() || `Column ${idx + 1}`;
      const key = name.toLowerCase();
      if (seen[key]) { seen[key] += 1; name = `${name} (${seen[key]})`; } else seen[key] = 1;
      return name;
    });

    const rows = raw.slice(1).map((r) => {
      const obj = {};
      headers.forEach((h, idx) => { obj[h] = (r[idx] ?? "").trim(); });
      return obj;
    });
    return { headers, rows, delimiter };
  }

  /** Best guess at the column holding email addresses. */
  function findEmailColumn(headers) {
    const exact = headers.find((h) => /^(e-?mail|e-?mail address|email addr|to|recipient)$/i.test(h.trim()));
    if (exact) return exact;
    return headers.find((h) => /e-?mail/i.test(h)) || null;
  }

  /** Optional CC / BCC columns, matched by header name. */
  function findColumn(headers, pattern) {
    return headers.find((h) => pattern.test(h.trim())) || null;
  }

  // ─── TEMPLATES ────────────────────────────────────────────────────────────

  const PLACEHOLDER = /\{\{\s*([^{}|]+?)\s*(?:\|([^{}]*))?\}\}/g;

  /** Field names used in a template, e.g. "Hi {{First Name|there}}" -> ["First Name"]. */
  function placeholders(template) {
    const names = [];
    String(template || "").replace(PLACEHOLDER, (_, name) => { names.push(name.trim()); return ""; });
    return names;
  }

  /** Placeholders that have a "|fallback" so a blank value is fine. */
  function placeholdersWithDefault(template) {
    const names = new Set();
    String(template || "").replace(PLACEHOLDER, (_, name, def) => {
      if (def !== undefined) names.add(name.trim().toLowerCase());
      return "";
    });
    return names;
  }

  /** Look up a field case- and space-insensitively. */
  function lookup(row, name) {
    if (Object.prototype.hasOwnProperty.call(row, name)) return row[name];
    const want = name.toLowerCase().replace(/\s+/g, " ");
    for (const key of Object.keys(row)) {
      if (key.toLowerCase().replace(/\s+/g, " ") === want) return row[key];
    }
    return undefined;
  }

  function hasColumn(headers, name) {
    const want = name.toLowerCase().replace(/\s+/g, " ");
    return headers.some((h) => h.toLowerCase().replace(/\s+/g, " ") === want);
  }

  /** Fill {{Field}} and {{Field|fallback}} with plain-text values from a row. */
  function fill(template, row) {
    return String(template || "").replace(PLACEHOLDER, (_, name, def) => {
      const value = lookup(row, name.trim());
      if (value !== undefined && String(value).trim() !== "") return String(value);
      return def !== undefined ? def : "";
    });
  }

  // ─── HTML ─────────────────────────────────────────────────────────────────

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  /** Plain text -> safe HTML: escaped, links clickable, line breaks kept. */
  function textToHtml(text) {
    const escaped = escapeHtml(text);
    const linked = escaped.replace(/\b(https?:\/\/[^\s<]+[^\s<.,;:!?)\]'"])/g, '<a href="$1">$1</a>');
    return linked.replace(/\r?\n/g, "<br>");
  }

  const BODY_STYLE = "font-family: Aptos, Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; color: #000000;";

  function buildHtml(bodyText, signatureText) {
    let html = `<div style="${BODY_STYLE}">${textToHtml(bodyText)}`;
    if (signatureText && signatureText.trim()) html += `<br><br>${textToHtml(signatureText)}`;
    html += "</div>";
    return `<html><head><meta charset="utf-8"></head><body>${html}</body></html>`;
  }

  // ─── ADDRESSES ────────────────────────────────────────────────────────────

  const EMAIL_RE = /^[^\s@<>(),;:"]+@[^\s@<>(),;:"]+\.[^\s@<>(),;:"]{2,}$/;

  /** "a@x.com; b@y.com" or "Jane <a@x.com>" -> ["a@x.com", "b@y.com"]. */
  function splitAddresses(value) {
    return String(value || "")
      .split(/[;,]/)
      .map((part) => {
        const angle = part.match(/<([^>]+)>/);
        return (angle ? angle[1] : part).trim();
      })
      .filter(Boolean);
  }

  function isValidEmail(addr) {
    return EMAIL_RE.test(addr);
  }

  // ─── BUILD + VALIDATE ─────────────────────────────────────────────────────

  /**
   * Turn the table + template into one email per row.
   * ctx: { headers, rows, emailColumn, ccColumn, bccColumn, subject, body, signature }
   * Returns { emails: [{index, row, to, cc, bcc, subject, html, text, problems}], errors, warnings }.
   */
  function buildEmails(ctx) {
    const { headers = [], rows = [], emailColumn, ccColumn, bccColumn } = ctx;
    const subject = ctx.subject || "";
    const body = ctx.body || "";
    const signature = ctx.signature || "";
    const errors = [];
    const warnings = [];

    if (!rows.length) errors.push("Add a table with a header row and at least one person.");
    if (rows.length && !emailColumn) errors.push("Choose which column holds the email addresses.");
    if (!subject.trim()) errors.push("Add a subject.");
    if (!body.trim()) errors.push("Write the email body.");

    const allText = `${subject}\n${body}\n${signature}`;
    const used = placeholders(allText);
    const unknown = [...new Set(used.filter((name) => !hasColumn(headers, name)))];
    if (rows.length && unknown.length) {
      errors.push(`No column named ${unknown.map((u) => `"${u}"`).join(", ")}. Check the spelling against your table's header row.`);
    }
    const withDefault = placeholdersWithDefault(allText);
    const required = [...new Set(used.filter((n) => hasColumn(headers, n) && !withDefault.has(n.toLowerCase())))];

    const seen = new Map();
    const emails = rows.map((row, index) => {
      const problems = [];
      const to = emailColumn ? splitAddresses(lookup(row, emailColumn)) : [];
      const cc = ccColumn ? splitAddresses(lookup(row, ccColumn)) : [];
      const bcc = bccColumn ? splitAddresses(lookup(row, bccColumn)) : [];

      if (!to.length) problems.push("no email address");
      for (const addr of [...to, ...cc, ...bcc]) {
        if (!isValidEmail(addr)) problems.push(`"${addr}" doesn't look like an email address`);
      }
      const blanks = required.filter((name) => String(lookup(row, name) ?? "").trim() === "");
      const key = to.map((a) => a.toLowerCase()).sort().join(";");
      if (key && seen.has(key)) problems.push(`same address as row ${seen.get(key) + 2}`);
      else if (key) seen.set(key, index);

      const filledSubject = fill(subject, row).replace(/[\r\n]+/g, " ").trim();
      const filledBody = fill(body, row);
      const filledSig = fill(signature, row);
      return {
        index,
        row,
        to,
        cc,
        bcc,
        subject: filledSubject,
        text: filledBody,
        html: buildHtml(filledBody, filledSig),
        blanks,
        problems,
        // Rows with a bad or missing address, or a repeat of an earlier row, are left out.
        sendable: problems.length === 0,
      };
    });

    const excluded = emails.filter((e) => !e.sendable).length;
    if (excluded) warnings.push(`${excluded} ${excluded === 1 ? "row" : "rows"} will be skipped (missing, invalid or repeated address).`);
    const blankRows = emails.filter((e) => e.sendable && e.blanks.length).length;
    if (blankRows) {
      warnings.push(`${blankRows} ${blankRows === 1 ? "email has" : "emails have"} a blank field. Add a fallback like {{First Name|there}} or fill in the table.`);
    }
    return { emails, errors, warnings };
  }

  /** Microsoft Graph sendMail payload for one email. */
  function graphMessage(email, attachment) {
    const toRecipients = email.to.map((address) => ({ emailAddress: { address } }));
    const message = {
      subject: email.subject,
      body: { contentType: "HTML", content: email.html },
      toRecipients,
    };
    if (email.cc.length) message.ccRecipients = email.cc.map((address) => ({ emailAddress: { address } }));
    if (email.bcc.length) message.bccRecipients = email.bcc.map((address) => ({ emailAddress: { address } }));
    if (attachment) {
      message.attachments = [{
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: attachment.name,
        contentType: attachment.contentType || "application/octet-stream",
        contentBytes: attachment.contentBytes,
      }];
    }
    return { message, saveToSentItems: true };
  }

  // ─── SENDING ──────────────────────────────────────────────────────────────

  /** Stable short hash, used to remember who already got a given message. */
  function hash(str) {
    let h = 0x811c9dc5;
    for (let i = 0; i < str.length; i++) {
      h ^= str.charCodeAt(i);
      h = Math.imul(h, 0x01000193);
    }
    return (h >>> 0).toString(36);
  }

  function messageKey(subject, body) {
    return hash(`${subject}\u0000${body}`);
  }

  function recipientKey(email) {
    return email.to.map((a) => a.toLowerCase()).sort().join(";");
  }

  /**
   * Send emails one at a time, paced to stay under Exchange Online's 30-messages-per-minute limit.
   * opts: {
   *   send: async (email) => ({ status, retryAfter }) — status is the HTTP code,
   *   delayMs: pause between sends, sleep: async (ms), shouldStop: () => bool,
   *   onProgress: (result, done, total), alreadySent: Set of recipientKey
   * }
   * Retries 429 / 503 / 504 up to maxRetries, honoring Retry-After.
   */
  async function sendAll(emails, opts) {
    const {
      send,
      delayMs = 2400,
      sleep = (ms) => new Promise((r) => setTimeout(r, ms)),
      shouldStop = () => false,
      onProgress = () => {},
      alreadySent = new Set(),
      maxRetries = 4,
    } = opts;
    const results = [];
    let first = true;
    for (let i = 0; i < emails.length; i++) {
      const email = emails[i];
      if (shouldStop()) {
        for (const rest of emails.slice(i)) {
          const r = { email: rest, status: "stopped", detail: "Not sent (stopped)" };
          results.push(r);
          onProgress(r, results.length, emails.length);
        }
        break;
      }
      const key = recipientKey(email);
      if (alreadySent.has(key)) {
        const r = { email, status: "skipped", detail: "Already sent this message" };
        results.push(r);
        onProgress(r, results.length, emails.length);
        continue;
      }
      if (!first) await sleep(delayMs);
      first = false;

      let attempt = 0;
      let outcome;
      for (;;) {
        try {
          outcome = await send(email);
        } catch (err) {
          outcome = { status: 0, error: err && err.message ? err.message : String(err) };
        }
        const retryable = outcome.status === 429 || outcome.status === 503 || outcome.status === 504 || outcome.status === 0;
        if (!retryable || attempt >= maxRetries) break;
        attempt++;
        const wait = outcome.retryAfter ? outcome.retryAfter * 1000 : Math.min(60000, 2000 * 2 ** attempt);
        await sleep(wait);
      }
      const ok = outcome.status >= 200 && outcome.status < 300;
      const r = ok
        ? { email, status: "sent", detail: "Sent" }
        : { email, status: "failed", detail: outcome.error || `Error ${outcome.status}` };
      if (ok) alreadySent.add(key);
      results.push(r);
      onProgress(r, results.length, emails.length);
    }
    return results;
  }

  /** Tab-separated summary that pastes cleanly into Excel. */
  function resultsToTsv(results) {
    const lines = [["Row", "To", "Subject", "Result", "Detail"].join("\t")];
    for (const r of results) {
      lines.push([
        r.email.index + 2,
        r.email.to.join("; "),
        r.email.subject,
        r.status,
        r.detail,
      ].map((v) => String(v).replace(/[\t\r\n]+/g, " ")).join("\t"));
    }
    return lines.join("\n");
  }

  return {
    detectDelimiter,
    parseRows,
    parseTable,
    findEmailColumn,
    findColumn,
    placeholders,
    fill,
    escapeHtml,
    textToHtml,
    buildHtml,
    splitAddresses,
    isValidEmail,
    buildEmails,
    graphMessage,
    messageKey,
    recipientKey,
    sendAll,
    resultsToTsv,
  };
});
