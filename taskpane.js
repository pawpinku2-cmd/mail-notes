/* My Mail Notes – private notes on emails and meetings (Outlook add-in).
 *
 * How it works:
 *  - The note is saved as a custom property on the item, in your own mailbox
 *    (Office.js CustomProperties). Custom properties are not sent to recipients, and
 *    are only readable through this add-in.
 *  - Items with a note get an Outlook category ("Note" or a colour category such as
 *    "Note - Urgent") so they stand out in the inbox.
 *  - Quick-note presets and the "All notes" index (subject, sender, first 160 characters of
 *    each note) are stored in this add-in's roaming settings, which also live in your mailbox.
 *  - No network calls, no sign-in, no server, no storage outside your mailbox.
 *  - Note text is only ever rendered with .value / .textContent.
 */
(function () {
  "use strict";

  const PROP = "mailnote";        // custom property name
  const MAX_TEXT = 2000;          // characters of note text
  const MAX_JSON = 2500;          // Outlook limit for all custom properties on one item
  const AUTOSAVE_MS = 1200;

  // Colour -> Outlook category. "none" uses the plain marker.
  const COLOURS = [
    { id: "none",     label: "None",      category: "Note",             outlook: "any colour" },
    { id: "urgent",   label: "Urgent",    category: "Note - Urgent",    outlook: "Red" },
    { id: "followup", label: "Follow-up", category: "Note - Follow-up", outlook: "Orange" },
    { id: "info",     label: "Info",      category: "Note - Info",      outlook: "Blue" },
    { id: "done",     label: "Done",      category: "Note - Done",      outlook: "Green" },
  ];
  const ALL_NOTE_CATEGORIES = COLOURS.map(function (c) { return c.category; });
  const colourById = function (id) { return COLOURS.find(function (c) { return c.id === id; }) || COLOURS[0]; };

  const IDX_KEY = "idx";          // roaming-settings key for the All-notes index
  const IDX_BUDGET = 29000;       // roaming settings cap is 32 KB per add-in; leave room for quick notes
  const SNIPPET = 160;            // characters of note text kept in the index
  const QUICK_KEY = "quickNotes";
  const QUICK_MAX = 12;
  const QUICK_LEN = 60;
  const QUICK_DEFAULTS = [
    "Called – no answer",
    "Left voicemail",
    "Waiting on reply",
    "Waiting on PO",
    "Sent to accounts",
    "Escalated",
    "Done – no further action",
  ];

  const state = {
    seq: 0,            // bumps on every item change so late callbacks are ignored
    item: null,
    props: null,       // CustomProperties for the current item
    saved: "",         // last saved text
    savedColour: "none",
    colour: "none",    // colour selected in the UI
    savedMeta: null,   // { created, modified }
    timer: null,
    saving: false,
    ctx: null,         // { compose, kind, subject, sender, key, itemId, shared }
    view: "item",
  };

  const ADDIN_VERSION = "1.2.2";
  const env = { mobile: false, canOpen: false, canRemind: false, account: "", itemEvents: "not registered", lastError: null, maxSet: "" };

  /* Outlook on the web / new Outlook / mobile allow only 3 async Office calls at once.
     Queue them (max 2 in flight) and time out calls that never answer, so the pane can't get stuck. */
  const officeQueue = { running: 0, waiting: [] };
  const OFFICE_MAX = 2, OFFICE_TIMEOUT_MS = 15000;
  function callOffice(op, start) {
    return new Promise(function (resolve) {
      officeQueue.waiting.push({ op: op, start: start, resolve: resolve });
      pumpOffice();
    });
  }
  function pumpOffice() {
    while (officeQueue.running < OFFICE_MAX && officeQueue.waiting.length) {
      const job = officeQueue.waiting.shift();
      officeQueue.running++;
      let done = false;
      const finish = function (r) {
        if (done) return;
        done = true;
        clearTimeout(timer);
        officeQueue.running--;
        if (!r || r.status !== Office.AsyncResultStatus.Succeeded) noteError(job.op, r);
        try { job.resolve(r || { status: Office.AsyncResultStatus.Failed, error: { message: "No response" } }); }
        finally { setTimeout(pumpOffice, 0); }
      };
      const timer = setTimeout(function () {
        finish({ status: Office.AsyncResultStatus.Failed, error: { name: "Timeout", message: "Outlook didn't respond in time." } });
      }, OFFICE_TIMEOUT_MS);
      try { job.start(finish); }
      catch (e) { finish({ status: Office.AsyncResultStatus.Failed, error: { name: e && e.name || "Error", message: e && e.message || String(e) } }); }
    }
  }
  function noteError(op, r) {
    if (op === "shared") return;   // "not shared" is the normal answer for most items
    const err = (r && r.error) || {};
    env.lastError = { op: op, name: err.name || err.code || "", message: err.message || "", at: new Date().toISOString() };
    renderDiagnostics();
  }

  const $ = (id) => document.getElementById(id);

  /* ------------------------------------------------------------ startup */

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) {
      fatal("This page is the My Mail Notes add-in. It only runs inside Outlook.");
      return;
    }
    if (!Office.context.requirements.isSetSupported("Mailbox", "1.5")) {
      fatal("This version of Outlook is too old for My Mail Notes. Update Outlook or use Outlook on the web.");
      return;
    }
    // Each step is isolated: one failure must not stop notes from loading or saving.
    [detectEnv, buildColourButtons, buildFilter, bindUi, bindDashboard, bindDiagnostics, renderQuickList, renderAll].forEach(safely);
    safely(function registerItemEvents() {
      if (!Office.context.requirements.isSetSupported("Mailbox", "1.5")) { env.itemEvents = "not supported"; return; }
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged, function (r) {
        env.itemEvents = r && r.status === Office.AsyncResultStatus.Succeeded ? "registered"
          : "failed: " + ((r && r.error && r.error.message) || "unknown");
        renderDiagnostics();
      });
    });
    safely(loadItem);
    safely(startItemWatcher);
  });

  function safely(fn) {
    try { fn(); }
    catch (e) {
      env.lastError = { op: "start:" + (fn.name || "step"), name: e && e.name || "", message: e && e.message || String(e), at: new Date().toISOString() };
      try { renderDiagnostics(); } catch (e2) { /* ignore */ }
    }
  }

  function detectEnv() {
    const d = Office.context.diagnostics || {};
    env.mobile = /^(ios|android)$/i.test(String(d.platform || d.hostName || ""));
    const mb = Office.context.mailbox;
    env.canOpen = !env.mobile && typeof mb.displayMessageForm === "function";
    env.canRemind = !env.mobile && typeof mb.displayNewAppointmentForm === "function";
    document.documentElement.classList.toggle("mobile", env.mobile);
    env.account = (mb.userProfile && mb.userProfile.accountType) || "";
    const req = Office.context.requirements;
    for (let v = 16; v >= 1; v--) { if (req.isSetSupported("Mailbox", "1." + v)) { env.maxSet = "1." + v; break; } }
  }

  /* Fallback for clients that keep the pane open but don't send ItemChanged (seen in Outlook.com):
     check every second whether Outlook is now pointing at a different item. */
  let lastSeenItem = null;
  function currentItemSignature() {
    const it = Office.context.mailbox.item;
    if (!it) return "none";
    return (it.itemId || "") + "|" + (typeof it.subject === "string" ? it.subject : "") + "|" + (it.internetMessageId || "");
  }
  function startItemWatcher() {
    lastSeenItem = currentItemSignature();
    setInterval(function () {
      const sig = currentItemSignature();
      if (sig === lastSeenItem) return;
      lastSeenItem = sig;
      if (Office.context.mailbox.item !== state.item || sig === "none") {
        env.itemEvents = env.itemEvents.indexOf("(watcher used)") === -1 ? env.itemEvents + " (watcher used)" : env.itemEvents;
        onItemChanged();
      }
    }, 1000);
  }

  function onItemChanged() {
    lastSeenItem = currentItemSignature();
    // Flush a pending autosave for the previous item before switching.
    const pending = state.timer !== null;
    clearTimeout(state.timer);
    state.timer = null;
    if (pending && state.props && isDirty()) {
      saveNow(true).finally(loadItem);
    } else {
      loadItem();
    }
  }

  /* --------------------------------------------------------------- load */

  function isSupportedItem(item) {
    if (!item) return false;
    const t = Office.MailboxEnums.ItemType;
    return item.itemType === t.Message || item.itemType === t.Appointment;
  }

  function isMeeting(item) {
    return item && item.itemType === Office.MailboxEnums.ItemType.Appointment;
  }

  function loadItem() {
    const seq = ++state.seq;
    const item = Office.context.mailbox.item;
    state.item = item;
    state.props = null;
    state.saved = "";
    state.savedColour = "none";
    state.colour = "none";
    state.savedMeta = null;
    hideBanner();
    hideConfirm();
    setLoading(true);

    state.ctx = null;
    $("shared-line").hidden = true;

    if (!isSupportedItem(item)) {
      setLoading(false);
      $("content").hidden = true;
      $("context-line").textContent = "Select an email or meeting to see or add a note.";
      renderAll();
      if (state.view === "item") showView("all");   // nothing selected: show the dashboard
      return;
    }
    const compose = isCompose(item);
    const kind = isMeeting(item) ? "meeting" : "email";
    const ctx = state.ctx = {
      compose: compose,
      kind: kind,
      subject: compose ? "" : subjectOf(item),
      sender: compose ? "" : senderOf(item),
      key: null,
      itemId: compose ? null : (item.itemId || null),
      shared: null,
    };
    setContextLine();
    // Outlook.com has no pinnable panes: the pane stays on the email it was opened from.
    const bound = $("bound-line");
    bound.hidden = env.account !== "outlookCom";
    if (!bound.hidden) {
      bound.replaceChildren(
        document.createTextNode("Outlook.com keeps this pane on the email it was opened from"),
        document.createTextNode(ctx.subject ? ": " : "."),
      );
      if (ctx.subject) { const b = document.createElement("strong"); b.textContent = ctx.subject; bound.append(b, document.createTextNode(".")); }
      bound.append(document.createTextNode(" To note a different email, close this pane (X) and open My Mail Notes again from that email."));
    }
    $("note").placeholder = compose
      ? "Type a private note. It stays on your copy after you send or save."
      : "Type a private note for this " + kind + "…";
    $("remind-btn").hidden = compose || !env.canRemind;
    if (compose && item.subject && typeof item.subject.getAsync === "function") {
      callOffice("subject", function (cb) { item.subject.getAsync(cb); }).then(function (r) {
        if (seq !== state.seq || r.status !== Office.AsyncResultStatus.Succeeded) return;
        ctx.subject = r.value || "";
        setContextLine();
      });
    }
    if (typeof item.getSharedPropertiesAsync === "function" && (env.account === "office365" || env.account === "enterprise")) {
      try {
        callOffice("shared", function (cb) { item.getSharedPropertiesAsync(cb); }).then(function (r) {
          if (seq !== state.seq || r.status !== Office.AsyncResultStatus.Succeeded || !r.value) return;
          ctx.shared = String(r.value.owner || "").toLowerCase() || "shared";
          const line = $("shared-line");
          line.textContent = "Shared mailbox" + (r.value.owner ? " (" + r.value.owner + ")" : "") +
            ": everyone with access to this mailbox can see this note.";
          line.hidden = false;
        });
      } catch (e) { /* not shared or not supported */ }
    }
    if (!compose) {
      itemKey(item).then(function (k) {
        if (seq !== state.seq) return;
        ctx.key = k;
        if (state.savedMeta && state.saved) indexUpsertFromState(true);
        renderAll();
      });
    }

    callOffice("load", function (cb) { item.loadCustomPropertiesAsync(cb); }).then(function (r) {
      if (seq !== state.seq) return;
      setLoading(false);
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        $("content").hidden = true;
        showBanner("error", "Couldn't load the note for this item. " + errText(r));
        return;
      }
      state.props = r.value;
      renderDiagnostics();
      const data = parse(state.props.get(PROP));
      state.saved = data ? data.text : "";
      state.savedColour = data ? data.colour : "none";
      state.colour = state.savedColour;
      state.savedMeta = data;
      $("note").value = state.saved;
      $("content").hidden = false;
      updateUi();
      if (data && data.text) syncCategories(item, data.colour, seq, true);
      if (data && data.text && ctx.key) indexUpsertFromState(true);
      if (state.view === "all") showView("item");
    });
  }

  function parse(v) {
    if (!v) return null;
    try {
      const o = typeof v === "string" ? JSON.parse(v) : v;
      if (o && typeof o.text === "string") {
        return {
          text: o.text,
          colour: COLOURS.some(function (c) { return c.id === o.colour; }) ? o.colour : "none",
          created: o.created || null,
          modified: o.modified || null,
        };
      }
    } catch (e) { /* ignore corrupt value */ }
    return null;
  }

  /* --------------------------------------------------------------- save */

  function isDirty() {
    return clean($("note").value) !== state.saved || (state.saved && state.colour !== state.savedColour);
  }

  function scheduleSave() {
    clearTimeout(state.timer);
    state.timer = setTimeout(function () { state.timer = null; saveNow(false); }, AUTOSAVE_MS);
  }

  function saveNow(silent) {
    return new Promise(function (resolve) {
      clearTimeout(state.timer);
      state.timer = null;
      const props = state.props;
      const item = state.item;
      const seq = state.seq;
      if (!props) return resolve();

      const ctxSnapshot = state.ctx;
      const text = clean($("note").value);
      const colour = state.colour;
      if (text === state.saved && (colour === state.savedColour || !text)) { updateUi(); return resolve(); }
      if (!text) { deleteNote(silent).then(resolve); return; }

      const now = new Date().toISOString();
      const value = {
        v: 2,
        text: text,
        colour: colour,
        created: (state.savedMeta && state.savedMeta.created) || now,
        modified: now,
      };
      const size = JSON.stringify({ [PROP]: value }).length;
      if (size > MAX_JSON) {
        setStatus("error", "Too long to save: Outlook allows about " + MAX_JSON +
          " characters including formatting. Shorten the note.");
        return resolve();
      }

      state.saving = true;
      setStatus("", "Saving…");
      props.set(PROP, value);
      callOffice("save", function (cb) { props.saveAsync(cb); }).then(function (r) {
        state.saving = false;
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          if (seq === state.seq) {
            setStatus("error", "Not saved: " + errText(r) + " Try again.");
            if (env.account === "gmail") {
              showBanner("error", "This looks like a Gmail account added to Outlook. Outlook may not let add-ins store data on Gmail messages. Try a Microsoft 365 or Outlook.com mailbox.");
            }
          }
          return resolve();
        }
        if (seq === state.seq) {
          state.saved = text;
          state.savedColour = colour;
          state.savedMeta = value;
          updateUi();
        }
        syncCategories(item, colour, seq, silent);
        if (ctxSnapshot && ctxSnapshot.key) indexUpsert(entryFrom(ctxSnapshot, value));
        resolve();
      });
    });
  }

  function deleteNote(silent) {
    return new Promise(function (resolve) {
      const props = state.props;
      const item = state.item;
      const seq = state.seq;
      const ctxSnapshot = state.ctx;
      if (!props) return resolve();
      props.remove(PROP);
      callOffice("save", function (cb) { props.saveAsync(cb); }).then(function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          if (seq === state.seq) setStatus("error", "Couldn't delete: " + errText(r));
          return resolve();
        }
        if (seq === state.seq) {
          state.saved = "";
          state.savedColour = "none";
          state.colour = "none";
          state.savedMeta = null;
          $("note").value = "";
          updateUi();
          if (!silent) showBanner("ok", "Note deleted.", 2500);
        }
        syncCategories(item, null, seq, true);
        if (ctxSnapshot && ctxSnapshot.key) indexRemove(ctxSnapshot.key);
        resolve();
      });
    });
  }

  /* ------------------------------------------- inbox marker (categories) */

  // colour = colour id to show, or null to remove all note categories.
  // Adds the wanted category first and only then removes the others, so a missing colour
  // category never leaves the item unmarked.
  function syncCategories(item, colour, seq, silent) {
    // Categories can't be set on drafts in web/new Outlook and aren't available on mobile.
    // They're applied the next time the item is opened in read mode.
    if (!item || !item.categories || isCompose(item)) return;
    const want = colour ? colourById(colour).category : null;
    callOffice("categories.get", function (cb) { item.categories.getAsync(cb); }).then(function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return;
      const current = (r.value || []).map(function (c) { return c.displayName; });
      const keepOnly = function (kept) {
        const remove = current.filter(function (n) { return ALL_NOTE_CATEGORIES.indexOf(n) !== -1 && n !== kept; });
        if (remove.length) callOffice("categories.remove", function (cb) { item.categories.removeAsync(remove, cb); });
      };
      const tell = function (text) { if (!silent && seq === state.seq) showBanner("info", text); };

      if (!want) { keepOnly(null); return; }
      if (current.indexOf(want) !== -1) { keepOnly(want); return; }

      callOffice("categories.add", function (cb) { item.categories.addAsync([want], cb); }).then(function (a) {
        if (a.status === Office.AsyncResultStatus.Succeeded) { keepOnly(want); return; }
        if (want === "Note") {
          tell('Saved. To mark items that have notes, create an Outlook category called "Note" ' +
            "(Settings > Mail > Categories). After that, items get tagged automatically.");
          return;
        }
        // Colour category missing from this mailbox: fall back to the plain "Note" marker.
        const c = colourById(colour);
        const fallbackMsg = 'Saved. To colour this in your inbox, create an Outlook category called "' +
          c.category + '" (' + c.outlook + ') in Settings > Mail > Categories. For now it\'s tagged "Note".';
        if (current.indexOf("Note") !== -1) { keepOnly("Note"); tell(fallbackMsg); return; }
        callOffice("categories.add", function (cb) { item.categories.addAsync(["Note"], cb); }).then(function (b) {
          if (b.status === Office.AsyncResultStatus.Succeeded) { keepOnly("Note"); tell(fallbackMsg); return; }
          tell('Saved. To mark items that have notes, create Outlook categories called "Note" and "' +
            c.category + '" in Settings > Mail > Categories.');
        });
      });
    });
  }

  /* ---------------------------------------------------------- colours */

  function buildColourButtons() {
    const box = $("colours");
    COLOURS.forEach(function (c) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "swatch";
      b.dataset.c = c.id;
      b.textContent = c.label;
      b.title = c.id === "none" ? 'Tag as "Note"' : 'Tag as "' + c.category + '"';
      b.setAttribute("aria-pressed", "false");
      b.addEventListener("click", function () { setColour(c.id); });
      box.append(b);
    });
  }

  function setColour(id) {
    state.colour = id;
    renderColours();
    if (clean($("note").value)) saveNow(false);
    else updateUi();
  }

  function renderColours() {
    document.querySelectorAll("#colours .swatch").forEach(function (b) {
      b.setAttribute("aria-pressed", String(b.dataset.c === state.colour));
    });
  }

  /* ------------------------------------------------------- quick notes */

  function getQuickNotes() {
    try {
      const v = Office.context.roamingSettings.get(QUICK_KEY);
      if (Array.isArray(v) && v.length) return v.filter(function (s) { return typeof s === "string" && s.trim(); }).slice(0, QUICK_MAX);
    } catch (e) { /* fall through */ }
    return QUICK_DEFAULTS.slice();
  }

  function renderQuickList() {
    const box = $("quick-list");
    box.replaceChildren();
    getQuickNotes().forEach(function (q) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "chip";
      b.textContent = q;
      b.title = "Add “" + q + "” with today's date and time";
      b.addEventListener("click", function () { addQuickNote(q); });
      box.append(b);
    });
  }

  function addQuickNote(text) {
    const ta = $("note");
    const line = stampText() + text;
    const cur = ta.value.replace(/\s+$/, "");
    const next = cur ? cur + "\n" + line : line;
    if (next.length > MAX_TEXT) {
      setStatus("error", "Note is full. Remove some text before adding a quick note.");
      return;
    }
    ta.value = next;
    ta.scrollTop = ta.scrollHeight;
    updateUi();
    saveNow(false);
  }

  function openQuickEditor() {
    $("quick-text").value = getQuickNotes().join("\n");
    $("quick-editor").hidden = false;
    $("quick-list").hidden = true;
    $("quick-text").focus();
  }

  function closeQuickEditor() {
    $("quick-editor").hidden = true;
    $("quick-list").hidden = false;
  }

  function saveQuickNotes(list) {
    try {
      if (list) Office.context.roamingSettings.set(QUICK_KEY, list);
      else Office.context.roamingSettings.remove(QUICK_KEY);
    } catch (e) {
      showBanner("error", "Couldn't save the quick-note list: " + e.message);
      return;
    }
    callOffice("settings.save", function (cb) { Office.context.roamingSettings.saveAsync(cb); }).then(function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        showBanner("error", "Couldn't save the quick-note list. " + errText(r));
        return;
      }
      renderQuickList();
      closeQuickEditor();
      showBanner("ok", "Quick notes saved.", 2000);
    });
  }

  /* --------------------------------------------------------- reminder */

  function openReminder() {
    const item = state.item;
    if (!item || !env.canRemind || (state.ctx && state.ctx.compose)) return;
    const start = new Date();
    start.setDate(start.getDate() + 1);
    start.setHours(9, 0, 0, 0);
    const end = new Date(start.getTime() + 15 * 60000);
    const subject = subjectOf(item) || "(no subject)";
    const who = senderOf(item);
    const note = clean($("note").value);
    const body = [
      "Follow-up reminder created from My Mail Notes.",
      "",
      (isMeeting(item) ? "Meeting: " : "Email: ") + subject,
      who ? (isMeeting(item) ? "Organiser: " : "From: ") + who : "",
      item.dateTimeCreated ? "Received: " + fmtFull(item.dateTimeCreated) : "",
      "",
      note ? "Note:\n" + note : "",
    ].join("\n").replace(/\n{3,}/g, "\n\n").trim();
    try {
      Office.context.mailbox.displayNewAppointmentForm({
        subject: truncate("Follow up: " + subject, 255),
        start: start,
        end: end,
        body: truncate(body, 30000),
      });
    } catch (e) {
      showBanner("error", "Couldn't open a new calendar event here. " + (e && e.message ? e.message : ""));
    }
  }

  /* ------------------------------------------------- All-notes index (dashboard) */

  const index = { entries: [], evicted: 0, timer: null };

  function loadIndex() {
    let v = null;
    try { v = Office.context.roamingSettings.get(IDX_KEY); } catch (e) { v = null; }
    const list = v && Array.isArray(v.e) ? v.e : [];
    index.entries = list.filter(function (e) { return e && typeof e.k === "string" && typeof e.t === "string"; });
    index.evicted = v && typeof v.x === "number" ? v.x : 0;
  }

  function entryFrom(ctx, value) {
    return {
      k: ctx.key,
      i: ctx.itemId || "",
      y: ctx.kind === "meeting" ? "m" : "e",
      s: truncate(ctx.subject || "(no subject)", 80),
      f: truncate(ctx.sender || "", 60),
      c: value.colour || "none",
      t: truncate(value.text || "", SNIPPET),
      m: value.modified || new Date().toISOString(),
      o: ctx.shared || "",
    };
  }

  // Refresh the entry for the open item (e.g. new item ID after a move, or note pre-dates the index).
  function indexUpsertFromState(onlyIfChanged) {
    const c = state.ctx;
    if (!c || !c.key || !state.saved || !state.savedMeta) return;
    const e = entryFrom(c, { text: state.saved, colour: state.savedColour, modified: state.savedMeta.modified });
    if (onlyIfChanged) {
      const cur = index.entries.find(function (x) { return x.k === e.k; });
      if (cur && JSON.stringify(cur) === JSON.stringify(e)) return;
    }
    indexUpsert(e);
  }

  function indexUpsert(e) {
    index.entries = index.entries.filter(function (x) { return x.k !== e.k; });
    index.entries.unshift(e);
    enforceBudget();
    saveIndexSoon();
    renderAll();
  }

  function indexRemove(k) {
    const before = index.entries.length;
    index.entries = index.entries.filter(function (x) { return x.k !== k; });
    if (index.entries.length !== before) { saveIndexSoon(); renderAll(); }
  }

  // Keep the index under the roaming-settings budget: drop oldest "Done" notes first, then oldest.
  function enforceBudget() {
    index.entries.sort(function (a, b) { return a.m < b.m ? 1 : a.m > b.m ? -1 : 0; });
    const size = function () { return JSON.stringify({ v: 1, x: index.evicted, e: index.entries }).length; };
    while (index.entries.length && size() > IDX_BUDGET) {
      let at = -1;
      for (let i = index.entries.length - 1; i >= 0; i--) { if (index.entries[i].c === "done") { at = i; break; } }
      if (at === -1) at = index.entries.length - 1;
      index.entries.splice(at, 1);
      index.evicted++;
    }
  }

  function saveIndexSoon() {
    clearTimeout(index.timer);
    index.timer = setTimeout(function () {
      try {
        Office.context.roamingSettings.set(IDX_KEY, { v: 1, x: index.evicted, e: index.entries });
        callOffice("settings.save", function (cb) { Office.context.roamingSettings.saveAsync(cb); }).then(function (r) {
          if (r.status !== Office.AsyncResultStatus.Succeeded) {
            showBanner("error", "Couldn't update the All notes list. " + errText(r));
          }
        });
      } catch (e) {
        showBanner("error", "Couldn't update the All notes list: " + (e && e.message ? e.message : e));
      }
    }, 600);
  }

  function buildFilter() {
    const sel = $("filter");
    COLOURS.forEach(function (c) {
      const o = document.createElement("option");
      o.value = c.id;
      o.textContent = c.id === "none" ? "No colour" : c.label;
      sel.append(o);
    });
  }

  function bindDashboard() {
    loadIndex();
    $("tab-item").addEventListener("click", function () { showView("item"); });
    $("tab-all").addEventListener("click", function () { showView("all"); });
    $("search").addEventListener("input", renderAll);
    $("filter").addEventListener("change", renderAll);
    $("export-btn").addEventListener("click", openExport);
    $("export-close").addEventListener("click", function () { $("export-panel").hidden = true; });
    $("export-copy").addEventListener("click", copyExport);
    $("export-download").addEventListener("click", downloadExport);
  }

  function showView(v) {
    state.view = v;
    $("tab-item").setAttribute("aria-selected", String(v === "item"));
    $("tab-all").setAttribute("aria-selected", String(v === "all"));
    $("view-item").hidden = v !== "item";
    $("view-all").hidden = v !== "all";
    if (v === "all") renderAll();
  }

  function filteredEntries(ignoreColour) {
    const terms = $("search").value.toLowerCase().split(/\s+/).filter(Boolean);
    const colour = ignoreColour ? "" : $("filter").value;
    return index.entries.filter(function (e) {
      if (colour && e.c !== colour) return false;
      const hay = (e.s + " " + e.f + " " + e.t + " " + e.o).toLowerCase();
      return terms.every(function (t) { return hay.indexOf(t) !== -1; });
    });
  }

  function renderAll() {
    $("all-count").textContent = index.entries.length ? "(" + index.entries.length + ")" : "";
    if ($("view-all").hidden) return;
    const terms = $("search").value.split(/\s+/).filter(Boolean);
    const list = filteredEntries();
    renderSummary(filteredEntries(true));
    const ul = $("all-list");
    ul.replaceChildren();
    list.forEach(function (e) {
      const li = document.createElement("li");
      li.className = "all-item" + (state.ctx && state.ctx.key === e.k ? " current" : "");
      li.tabIndex = 0;
      li.setAttribute("role", "button");
      li.title = env.canOpen ? "Open this " + (e.y === "m" ? "meeting" : "email") : "";
      const top = document.createElement("div");
      top.className = "all-top";
      const subj = document.createElement("span");
      subj.className = "all-subj";
      const dot = document.createElement("span");
      dot.className = "dotc";
      dot.dataset.c = e.c;
      subj.append(dot);
      highlight(subj, e.s, terms);
      if (e.y === "m") subj.append(badge("Meeting"));
      if (e.o) subj.append(badge("Shared"));
      const when = document.createElement("span");
      when.className = "all-when";
      when.textContent = fmt(e.m);
      top.append(subj, when);
      const from = document.createElement("div");
      from.className = "all-from";
      highlight(from, e.f || (e.o ? e.o : ""), terms);
      const text = document.createElement("div");
      text.className = "all-text";
      highlight(text, e.t, terms);
      li.append(top, from, text);
      li.addEventListener("click", function () { openEntry(e); });
      li.addEventListener("keydown", function (ev) { if (ev.key === "Enter") openEntry(e); });
      ul.append(li);
    });
    $("all-empty").hidden = list.length > 0;
    $("all-empty").textContent = index.entries.length
      ? "No notes match your search."
      : "No notes yet. Notes you add (or open) appear here.";
    const pct = Math.round(JSON.stringify({ v: 1, x: index.evicted, e: index.entries }).length / IDX_BUDGET * 100);
    $("capacity").textContent = index.entries.length + " note" + (index.entries.length === 1 ? "" : "s") + " · list " + pct + "% full";
    const hints = [];
    hints.push(env.canOpen
      ? "Click a note to open it. If an email doesn't open, it was moved: find it in Outlook and open it once to refresh the link."
      : "Open the email from your inbox to edit its note (search Outlook for the subject).");
    if (index.evicted) hints.push("The list keeps your most recent notes. Older notes are still saved on their emails.");
    $("all-hint").textContent = hints.join(" ");
    $("export-btn").disabled = list.length === 0;
  }

  /* ------------------------------------------------ summary: donut + legend */

  // Segment order is chosen so neighbouring colours stay distinguishable for colour-blind users.
  const CHART_ORDER = ["urgent", "info", "followup", "none", "done"];
  const SVGNS = "http://www.w3.org/2000/svg";

  function setFilter(id) {
    const sel = $("filter");
    sel.value = sel.value === id ? "" : id;
    renderAll();
  }

  function renderSummary(entries) {
    const box = $("summary");
    box.hidden = index.entries.length === 0;
    if (box.hidden) return;
    const counts = {};
    CHART_ORDER.forEach(function (id) { counts[id] = 0; });
    entries.forEach(function (e) { counts[e.c in counts ? e.c : "none"]++; });
    const total = entries.length;
    const active = $("filter").value;

    // Legend (doubles as the data table: label, count, share)
    const legend = $("legend");
    legend.replaceChildren();
    CHART_ORDER.forEach(function (id) {
      const c = colourById(id);
      const n = counts[id];
      const li = document.createElement("li");
      const b = document.createElement("button");
      b.type = "button";
      b.dataset.c = id;
      b.setAttribute("aria-pressed", String(active === id));
      b.disabled = n === 0 && active !== id;
      b.title = n ? "Show only " + c.label + " notes" : "No " + c.label + " notes";
      const sw = document.createElement("span");
      sw.className = "sw";
      sw.style.background = "var(--c-" + id + ")";
      const name = document.createElement("span");
      name.textContent = id === "none" ? "No colour" : c.label;
      const num = document.createElement("span");
      num.className = "n";
      num.textContent = String(n);
      const pct = document.createElement("span");
      pct.className = "pct";
      pct.textContent = total ? Math.round(n / total * 100) + "%" : "–";
      b.append(sw, name, num, pct);
      b.addEventListener("click", function () { setFilter(id); });
      li.append(b);
      legend.append(li);
    });

    // Donut
    const svg = $("donut");
    Array.from(svg.querySelectorAll(":scope > :not(title)")).forEach(function (n) { n.remove(); });
    svg.classList.toggle("has-focus", !!active);
    const cx = 60, cy = 60, R = 56, r = 38;
    if (!total) {
      const ring = document.createElementNS(SVGNS, "circle");
      ring.setAttribute("cx", cx); ring.setAttribute("cy", cy); ring.setAttribute("r", (R + r) / 2);
      ring.setAttribute("fill", "none"); ring.setAttribute("stroke", "var(--line)"); ring.setAttribute("stroke-width", R - r);
      svg.append(ring);
    }
    let a0 = -Math.PI / 2;
    CHART_ORDER.forEach(function (id) {
      const n = counts[id];
      if (!n) return;
      const sweep = n / total * Math.PI * 2;
      const path = document.createElementNS(SVGNS, "path");
      path.setAttribute("d", arcPath(cx, cy, R, r, a0, a0 + sweep));
      path.setAttribute("class", "seg" + (active === id ? " on" : ""));
      path.setAttribute("fill", "var(--c-" + id + ")");
      path.setAttribute("stroke", "var(--card)");     // 2px surface gap between segments
      path.setAttribute("stroke-width", "2");
      path.setAttribute("stroke-linejoin", "round");
      path.dataset.c = id;
      const label = (id === "none" ? "No colour" : colourById(id).label) + " · " + n + " note" + (n === 1 ? "" : "s") +
        " (" + Math.round(n / total * 100) + "%)";
      path.addEventListener("mousemove", function (ev) { showTip(ev, label); });
      path.addEventListener("mouseleave", hideTip);
      path.addEventListener("click", function () { hideTip(); setFilter(id); });
      svg.append(path);
      a0 += sweep;
    });
    const t1 = document.createElementNS(SVGNS, "text");
    t1.setAttribute("x", cx); t1.setAttribute("y", cy + 3); t1.setAttribute("text-anchor", "middle");
    t1.setAttribute("class", "total");
    t1.textContent = String(total);
    const t2 = document.createElementNS(SVGNS, "text");
    t2.setAttribute("x", cx); t2.setAttribute("y", cy + 17); t2.setAttribute("text-anchor", "middle");
    t2.setAttribute("class", "total-sub");
    t2.textContent = total === 1 ? "note" : "notes";
    svg.append(t1, t2);
    $("donut-title").textContent = "Notes by colour: " + CHART_ORDER.filter(function (id) { return counts[id]; })
      .map(function (id) { return (id === "none" ? "No colour" : colourById(id).label) + " " + counts[id]; }).join(", ");
  }

  function arcPath(cx, cy, R, r, a0, a1) {
    if (a1 - a0 >= Math.PI * 2 - 1e-6) {          // full ring: draw as two halves
      const m = a0 + Math.PI;
      return arcPath(cx, cy, R, r, a0, m) + " " + arcPath(cx, cy, R, r, m, a1);
    }
    const large = a1 - a0 > Math.PI ? 1 : 0;
    const p = function (rad, a) { return (cx + rad * Math.cos(a)).toFixed(2) + " " + (cy + rad * Math.sin(a)).toFixed(2); };
    return "M " + p(R, a0) + " A " + R + " " + R + " 0 " + large + " 1 " + p(R, a1) +
      " L " + p(r, a1) + " A " + r + " " + r + " 0 " + large + " 0 " + p(r, a0) + " Z";
  }

  function showTip(ev, text) {
    const tip = $("donut-tip");
    const box = $("donut").getBoundingClientRect();
    tip.textContent = text;
    tip.style.left = "60px";   // centred over the donut so it never runs off a narrow pane
    tip.style.top = (ev.clientY - box.top) + "px";
    tip.hidden = false;
  }
  function hideTip() { $("donut-tip").hidden = true; }

  function badge(text) {
    const b = document.createElement("span");
    b.className = "badge";
    b.textContent = text;
    return b;
  }

  // Appends text to el, wrapping search terms in <mark>. Uses text nodes only (no HTML parsing).
  function highlight(el, text, terms) {
    text = String(text || "");
    if (!terms.length) { el.append(document.createTextNode(text)); return; }
    const re = new RegExp("(" + terms.map(function (t) { return t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }).join("|") + ")", "ig");
    let last = 0;
    text.replace(re, function (m, _g, off) {
      if (off > last) el.append(document.createTextNode(text.slice(last, off)));
      const mk = document.createElement("mark");
      mk.textContent = m;
      el.append(mk);
      last = off + m.length;
      return m;
    });
    if (last < text.length) el.append(document.createTextNode(text.slice(last)));
  }

  function openEntry(e) {
    if (!env.canOpen || !e.i) {
      showBanner("info", 'Open it from Outlook: search for "' + e.s + '".');
      return;
    }
    try {
      if (e.y === "m") Office.context.mailbox.displayAppointmentForm(e.i);
      else Office.context.mailbox.displayMessageForm(e.i);
    } catch (err) {
      showBanner("info", 'Couldn\'t open it from here. Search Outlook for "' + e.s + '".');
    }
  }

  /* ------------------------------------------------------------- CSV export */

  function csvCell(v) {
    let s = String(v == null ? "" : v);
    if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;          // stop spreadsheet formula injection
    return '"' + s.replace(/"/g, '""') + '"';
  }

  function buildCsv(list) {
    const rows = [["Subject", "From", "Type", "Colour", "Note (first " + SNIPPET + " characters)", "Last edited", "Shared mailbox"]];
    list.forEach(function (e) {
      rows.push([e.s, e.f, e.y === "m" ? "Meeting" : "Email", colourById(e.c).label, e.t, fmtFull(e.m), e.o]);
    });
    return rows.map(function (r) { return r.map(csvCell).join(","); }).join("\r\n");
  }

  let exportCsv = "";   // keep the exact CSV (textarea values normalise CRLF to LF)
  function openExport() {
    exportCsv = buildCsv(filteredEntries());
    $("export-text").value = exportCsv;
    $("export-panel").hidden = false;
  }

  function copyExport() {
    const ta = $("export-text");
    const done = function () { showBanner("ok", "Copied. Paste into Excel or Notepad.", 2500); };
    const fallback = function () {
      ta.focus(); ta.select();
      let ok = false;
      try { ok = document.execCommand("copy"); } catch (e) { ok = false; }
      if (ok) done(); else showBanner("info", "Press Ctrl+C (Cmd+C on Mac) to copy the selected text.");
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(exportCsv).then(done, fallback);
    } else {
      fallback();
    }
  }

  function downloadExport() {
    try {
      const blob = new Blob(["\ufeff" + exportCsv], { type: "text/csv;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "my-mail-notes-" + new Date().toISOString().slice(0, 10) + ".csv";
      document.body.append(a);
      a.click();
      a.remove();
      setTimeout(function () { URL.revokeObjectURL(url); }, 10000);
    } catch (e) {
      showBanner("info", "Download isn't allowed here. Use Copy instead.");
    }
  }

  /* ------------------------------------------------------------ diagnostics */

  function diagnosticsText() {
    const d = Office.context.diagnostics || {};
    const it = state.item;
    const lines = [
      "My Mail Notes " + ADDIN_VERSION,
      "Outlook: " + (d.hostName || "?") + " " + (d.hostVersion || "") + " on " + (d.platform || "?"),
      "Account type: " + (env.account || "unknown"),
      "Highest Mailbox API set: " + (env.maxSet || "?"),
      "Item change events: " + env.itemEvents,
      "Current item: " + (it ? (it.itemType || "?") + (state.ctx && state.ctx.compose ? " (draft)" : "") : "none"),
      "Note loaded: " + (state.props ? "yes" : "no") + ", saved text: " + (state.saved ? state.saved.length + " chars" : "none"),
      "Categories API: " + (it && it.categories ? "available" : "not available"),
      "All notes list: " + index.entries.length + " entries",
      "Last error: " + (env.lastError ? env.lastError.op + " – " + [env.lastError.name, env.lastError.message].filter(Boolean).join(": ") + " (" + env.lastError.at + ")" : "none"),
    ];
    return lines.join("\n");
  }

  function renderDiagnostics() {
    const box = $("diag-text");
    if (box && !$("diag").hidden) box.textContent = diagnosticsText();
  }

  function bindDiagnostics() {
    $("diag-toggle").addEventListener("click", function () {
      const d = $("diag");
      d.hidden = !d.hidden;
      renderDiagnostics();
    });
    $("diag-copy").addEventListener("click", function () {
      const t = diagnosticsText();
      const ok = function () { showBanner("ok", "Diagnostics copied.", 2000); };
      if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(t).then(ok, function () { selectText($("diag-text")); });
      else selectText($("diag-text"));
    });
  }

  function selectText(el) {
    const range = document.createRange();
    range.selectNodeContents(el);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    showBanner("info", "Press Ctrl+C (Cmd+C on Mac) to copy.");
  }

  /* ------------------------------------------------------------------ UI */

  function bindUi() {
    const ta = $("note");
    ta.maxLength = MAX_TEXT;
    ta.addEventListener("input", function () { updateUi(); scheduleSave(); });
    ta.addEventListener("blur", function () { if (state.timer) saveNow(false); });
    ta.addEventListener("keydown", function (e) {
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); saveNow(false); }
      if (e.key.toLowerCase() === "d" && e.shiftKey && (e.ctrlKey || e.metaKey)) { e.preventDefault(); insertStamp(); }
    });
    $("stamp-btn").addEventListener("click", insertStamp);
    $("remind-btn").addEventListener("click", openReminder);
    $("save-btn").addEventListener("click", function () { saveNow(false); });
    $("delete-btn").addEventListener("click", function () { $("confirm-delete").hidden = false; });
    $("confirm-no").addEventListener("click", hideConfirm);
    $("confirm-yes").addEventListener("click", function () { hideConfirm(); deleteNote(false); });
    $("quick-edit").addEventListener("click", openQuickEditor);
    $("quick-cancel").addEventListener("click", closeQuickEditor);
    $("quick-reset").addEventListener("click", function () { saveQuickNotes(null); });
    $("quick-save").addEventListener("click", function () {
      const list = $("quick-text").value.split("\n")
        .map(function (s) { return clean(s).slice(0, QUICK_LEN); })
        .filter(Boolean)
        .slice(0, QUICK_MAX);
      saveQuickNotes(list.length ? list : null);
    });
  }

  function stampText() {
    const now = new Date();
    return "[" + now.toLocaleDateString("en-AU") + " " +
      now.toLocaleTimeString("en-AU", { hour: "numeric", minute: "2-digit" }) + "] ";
  }

  /* Insert "[22/10/2026 2:15 pm] " at the cursor, on its own line, then autosave. */
  function insertStamp() {
    const ta = $("note");
    const stamp = stampText();
    const start = ta.selectionStart, end = ta.selectionEnd;
    const before = ta.value.slice(0, start);
    const prefix = before && !before.endsWith("\n") ? "\n" : "";
    const insert = prefix + stamp;
    if (ta.value.length - (end - start) + insert.length > MAX_TEXT) {
      setStatus("error", "Note is full. Remove some text before adding a date/time.");
      return;
    }
    ta.setRangeText(insert, start, end, "end");
    ta.focus();
    updateUi();
    scheduleSave();
  }

  function renderMeta() {
    const m = state.savedMeta;
    const show = !!(state.saved && m && m.created);
    $("meta").hidden = !show;
    if (!show) return;
    $("meta-added").textContent = fmtFull(m.created);
    const edited = m.modified && m.modified !== m.created;
    $("meta-edited-row").hidden = !edited;
    $("meta-edited").textContent = edited ? fmtFull(m.modified) : "";
  }

  function updateUi() {
    const current = $("note").value;
    $("char-count").textContent = current.length + " / " + MAX_TEXT;
    $("delete-btn").hidden = !state.saved;
    renderColours();
    renderMeta();
    if (state.saving) return;
    if (isDirty()) {
      setStatus("unsaved", "Unsaved changes");
    } else if (state.saved && state.savedMeta && state.savedMeta.modified) {
      setStatus("saved", "Saved " + fmt(state.savedMeta.modified));
    } else {
      setStatus("", state.saved ? "Saved" : "No note yet");
    }
  }

  function setStatus(kind, text) {
    const s = $("status");
    s.className = "hint " + kind;
    s.textContent = text;
  }

  function hideConfirm() { $("confirm-delete").hidden = true; }

  // While an item is loading, lock the editor so typing can't be overwritten by the incoming note.
  function setLoading(on) {
    $("loading").hidden = !on;
    $("app").setAttribute("aria-busy", on ? "true" : "false");
    $("content").inert = on;
    $("note").readOnly = on;
    $("content").classList.toggle("busy", on);
  }

  let bannerTimer = null;
  function showBanner(kind, text, autoHideMs) {
    const b = $("banner");
    b.className = "banner " + kind;
    b.textContent = text;
    b.hidden = false;
    clearTimeout(bannerTimer);
    if (autoHideMs) bannerTimer = setTimeout(hideBanner, autoHideMs);
  }
  function hideBanner() { $("banner").hidden = true; }

  function fatal(text) {
    setLoading(false);
    $("content").hidden = true;
    showBanner("error", text);
  }

  /* ------------------------------------------------------------- helpers */

  function isCompose(item) {
    return !!(item && item.subject && typeof item.subject !== "string" && typeof item.subject.getAsync === "function");
  }

  function setContextLine() {
    const c = state.ctx;
    if (!c) return;
    const what = c.compose ? (c.kind === "meeting" ? "this meeting (draft)" : "this draft") : "this " + c.kind;
    $("context-line").textContent = c.subject
      ? "Private note for " + (c.compose ? what : c.kind) + ": " + c.subject
      : "Private note for " + what;
  }

  // Stable key for the index: survives moving the email between folders.
  function itemKey(item) {
    const basis = (item.internetMessageId && "m:" + item.internetMessageId) || (item.itemId && "i:" + item.itemId) || "";
    if (!basis) return Promise.resolve(null);
    return sha256(basis).then(function (h) { return h.slice(0, 20); });
  }

  function sha256(text) {
    if (window.crypto && crypto.subtle && window.TextEncoder) {
      return crypto.subtle.digest("SHA-256", new TextEncoder().encode(text)).then(function (buf) {
        return Array.from(new Uint8Array(buf), function (b) { return b.toString(16).padStart(2, "0"); }).join("");
      }).catch(function () { return fnv(text); });
    }
    return Promise.resolve(fnv(text));
  }

  function fnv(text) { // fallback only
    let h1 = 0x811c9dc5, h2 = 0x01000193;
    for (let i = 0; i < text.length; i++) {
      h1 = Math.imul(h1 ^ text.charCodeAt(i), 16777619) >>> 0;
      h2 = Math.imul(h2 ^ text.charCodeAt(i), 2246822519) >>> 0;
    }
    return ("00000000" + h1.toString(16)).slice(-8) + ("00000000" + h2.toString(16)).slice(-8) + "0000";
  }

  function subjectOf(item) {
    const s = item.normalizedSubject || item.subject;
    return typeof s === "string" ? s : "";
  }

  function senderOf(item) {
    const p = isMeeting(item) ? item.organizer : (item.from || item.sender);
    if (!p) return "";
    return p.displayName && p.emailAddress ? p.displayName + " <" + p.emailAddress + ">" : (p.emailAddress || p.displayName || "");
  }

  function clean(s) {
    return String(s || "")
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, "")
      .trim()
      .slice(0, MAX_TEXT);
  }

  function truncate(s, n) {
    return s.length > n ? s.slice(0, n - 1) + "…" : s;
  }

  function errText(r) {
    return (r && r.error && r.error.message) ? r.error.message : "Unknown error.";
  }

  /* "Wed 22 Oct 2026, 2:15 pm" in the user's local time zone. */
  function fmtFull(iso) {
    const d = new Date(iso);
    if (isNaN(d)) return "";
    return d.toLocaleString("en-AU", {
      weekday: "short", day: "numeric", month: "short", year: "numeric", hour: "numeric", minute: "2-digit",
    });
  }

  function fmt(iso) {
    const d = new Date(iso);
    if (isNaN(d)) return "";
    const diff = (Date.now() - d.getTime()) / 1000;
    if (diff < 60) return "just now";
    if (diff < 3600) return Math.floor(diff / 60) + " min ago";
    return d.toLocaleString("en-AU", { day: "numeric", month: "short", hour: "numeric", minute: "2-digit" });
  }
})();
