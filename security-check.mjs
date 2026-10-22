#!/usr/bin/env node
// My Mail Notes – security self-check. Created by Vaibhav Girdhar.
// Proves, from the source itself, what the add-in can and cannot do. No dependencies.
// Run:  node security-check.mjs [folder]      (default: this script's folder)
// Exit code 0 = all checks passed. Runs automatically on every push (GitHub Actions).

import { readFileSync, readdirSync, statSync } from "node:fs";
import { join, extname, dirname, relative } from "node:path";
import { fileURLToPath } from "node:url";

const root = process.argv[2] || dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(root, f), "utf8");
const stripComments = (js) => js.replace(/\/\*[\s\S]*?\*\//g, "").replace(/(^|[^:\\])\/\/.*$/gm, "$1");

const results = [];
const check = (ok, title, detail) => results.push({ ok, title, detail: detail || "" });

const js = stripComments(read("taskpane.js"));
const html = read("taskpane.html");
const manifests = readdirSync(root).filter((f) => /^manifest.*\.xml$/.test(f));

/* 1. No network access from the add-in's own code */
const network = [
  ["fetch(", /\bfetch\s*\(/], ["XMLHttpRequest", /XMLHttpRequest/], ["WebSocket", /WebSocket/],
  ["EventSource", /EventSource/], ["navigator.sendBeacon", /sendBeacon/], ["dynamic import()", /\bimport\s*\(/],
  ["importScripts", /importScripts/], ["window.open", /window\.open\s*\(/], ["new Image (tracking pixel)", /new\s+Image\s*\(/],
  ["external URL in code", /https?:\/\/(?!www\.w3\.org\/2000\/svg)/],
];
network.forEach(([name, re]) => check(!re.test(js), `No ${name} in taskpane.js`, "The add-in cannot send data anywhere."));

/* 2. No access to mailbox tokens, EWS, or other mailbox-wide APIs */
const tokens = [
  ["getCallbackTokenAsync", /getCallbackTokenAsync/], ["getUserIdentityTokenAsync", /getUserIdentityTokenAsync/],
  ["getAccessToken / SSO", /getAccessToken|auth\.getAccessToken/], ["makeEwsRequestAsync (EWS)", /makeEwsRequestAsync/],
  ["REST URL", /restUrl|ewsUrl/], ["MSAL / sign-in library", /msal|PublicClientApplication/i],
];
tokens.forEach(([name, re]) => check(!re.test(js), `Never requests ${name}`, "Without a token it cannot reach any other email, folder or person."));

/* 3. Never reads the email body or attachments, never sends/forwards/deletes/changes the email */
const itemOps = [
  ["read the email body", /\bbody\s*\.\s*(getAsync|getTypeAsync)/], ["change the email body", /\bbody\s*\.\s*(setAsync|setSelectedDataAsync|prependAsync|appendOnSendAsync)/],
  ["read attachments", /getAttachment(Content)?Async|\.attachments\b/], ["add attachments", /addFileAttachment|addItemAttachment/],
  ["send email", /\.sendAsync\s*\(/], ["reply/forward", /displayReply(All)?Form|displayReplyAllForm|forward/i],
  ["create new emails", /displayNewMessageForm/], ["move or delete items", /\bmoveAsync|\bdeleteAsync|\bremoveAttachmentAsync|\.move\(/],
  ["change recipients", /\b(to|cc|bcc|requiredAttendees|optionalAttendees)\s*\.\s*(setAsync|addAsync)/],
  ["change internet headers", /internetHeaders\s*\.\s*setAsync/], ["read other items by ID", /loadItemByIdAsync|getItemIdAsync|convertToRestId|convertToEwsId/],
  ["create mail rules / forwarding", /InboxRule|forwardingSmtp|rules/i], ["change mailbox master categories", /masterCategories\s*\.\s*(addAsync|removeAsync)/],
];
itemOps.forEach(([name, re]) => check(!re.test(js), `Does not ${name}`));

/* 4. Unsafe code patterns */
const unsafe = [
  ["eval", /\beval\s*\(/], ["new Function", /new\s+Function\s*\(/], ["string setTimeout", /setTimeout\s*\(\s*["'`]/],
  ["innerHTML / outerHTML", /\.(inner|outer)HTML\s*=/], ["insertAdjacentHTML", /insertAdjacentHTML/], ["document.write", /document\.write/],
  ["localStorage", /localStorage/], ["IndexedDB", /indexedDB/], ["cookies", /document\.cookie/], ["postMessage to other windows", /\.postMessage\s*\(/],
];
unsafe.forEach(([name, re]) => check(!re.test(js), `No ${name}`, "Note text is shown as plain text only."));

/* 5. What it does write – and only where */
const writes = [...js.matchAll(/\b(categories\.(?:addAsync|removeAsync)|props\.set|props\.remove|roamingSettings\.(?:set|remove)|saveAsync|displayNewAppointmentForm|displayMessageForm|displayAppointmentForm)\b/g)].map((m) => m[1]);
const allowed = new Set(["categories.addAsync", "categories.removeAsync", "props.set", "props.remove", "roamingSettings.set", "roamingSettings.remove", "saveAsync", "displayNewAppointmentForm", "displayMessageForm", "displayAppointmentForm"]);
check(writes.every((w) => allowed.has(w)), "Only writes: note on the open item, its category tag, and this add-in's own settings",
  "Plus opening an unsaved calendar form (user must press Save) and opening an item the user clicks.");
const catNames = [...js.matchAll(/category:\s*"([^"]+)"/g)].map((m) => m[1]);
check(catNames.length > 0 && catNames.every((n) => /^Note( - .+)?$/.test(n)), "Only adds/removes its own 'Note…' categories", catNames.join(", "));

/* 6. Page loading rules (Content Security Policy) */
const csp = (html.match(/http-equiv="Content-Security-Policy" content="([^"]+)"/) || [])[1] || "";
check(/default-src 'none'/.test(csp), "CSP: everything blocked by default", csp);
check(/script-src 'self' https:\/\/appsforoffice\.microsoft\.com;/.test(csp), "CSP: scripts only from this site + Microsoft Office.js");
check(/connect-src https:\/\/appsforoffice\.microsoft\.com;/.test(csp), "CSP: network calls only to Microsoft Office.js CDN");
check(/form-action 'none'/.test(csp) && /frame-src 'none'/.test(csp) && /object-src 'none'/.test(csp) && /base-uri 'none'/.test(csp), "CSP: no forms, frames, plugins or base-URL tricks");
const scripts = [...html.matchAll(/<script[^>]*src="([^"]+)"/g)].map((m) => m[1]);
check(scripts.every((s) => s === "https://appsforoffice.microsoft.com/lib/1/hosted/office.js" || /^taskpane\.js(\?v=[\d.]+)?$/.test(s)),
  "HTML loads only Office.js (Microsoft) and its own taskpane.js", scripts.join(", "));
check(!/<script(?![^>]*src=)[^>]*>/.test(html), "No inline scripts in the HTML");
check(!/<(iframe|object|embed|form)\b/i.test(html), "No iframes, plugins or forms in the HTML");

/* 7. Manifest: least privilege, no extra domains */
manifests.forEach((m) => {
  const x = read(m);
  check(/<Permissions>ReadWriteItem<\/Permissions>/.test(x) && !/ReadWriteMailbox/.test(x), `${m}: permission is ReadWriteItem (open item only), not ReadWriteMailbox`);
  check(!/<AppDomains>/.test(x), `${m}: no extra AppDomains`);
  check(!/WebApplicationInfo|<Scopes>|ExtendedPermissions|EventBasedActivation|LaunchEvent/.test(x), `${m}: no sign-in, no background/on-send events`);
  const hosts = new Set([...x.matchAll(/https:\/\/([^/"\s<>]+)/g)].map((h) => h[1]).filter((h) => !/schemas\.microsoft\.com|w3\.org/.test(h)));
  check(hosts.size === 1, `${m}: every URL points to one host`, [...hosts].join(", "));
});

/* 8. No executables or unexpected file types */
const files = [];
(function walk(d) { for (const f of readdirSync(d)) { if (f.startsWith(".")) continue; const p = join(d, f); statSync(p).isDirectory() ? walk(p) : files.push(p); } })(root);
const okTypes = new Set([".html", ".js", ".mjs", ".css", ".png", ".xml", ".md", ".json", ".txt"]);
const rel = (f) => relative(root, f).split("\\").join("/");
const isAdminScript = (f) => /^admin\/[^/]+\.ps1$/.test(rel(f));
const odd = files.filter((f) => !okTypes.has(extname(f).toLowerCase()) && !isAdminScript(f));
check(odd.length === 0, "No executables or unexpected file types", odd.map(rel).join(", "));
const adminScripts = files.filter(isAdminScript);
check(adminScripts.length > 0 && adminScripts.every((f) => !read(rel(f)).match(/Invoke-WebRequest|Invoke-RestMethod|DownloadString|DownloadFile|Start-BitsTransfer|iex\b|Invoke-Expression/i)),
  "Admin scripts (run by IT, never by the add-in) don't download or execute remote code", adminScripts.map(rel).join(", "));

/* report */
const pad = (s) => (s + " ".repeat(70)).slice(0, 70);
let failed = 0;
for (const r of results) { if (!r.ok) failed++; console.log(`${r.ok ? "PASS" : "FAIL"}  ${pad(r.title)} ${r.ok ? "" : r.detail}`); }
console.log(`\n${results.length - failed}/${results.length} security checks passed` + (failed ? ` – ${failed} FAILED` : ""));
process.exit(failed ? 1 : 0);
