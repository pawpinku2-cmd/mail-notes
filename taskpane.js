/* My Mail Notes – private notes on emails (Outlook add-in).
 *
 * How it works:
 *  - The note is saved as a custom property on the email item, in your own mailbox
 *    (Office.js CustomProperties). Custom properties are not sent to recipients, and
 *    are only readable through this add-in.
 *  - Emails that have a note get the Outlook category "Note" so they stand out in the inbox.
 *  - No network calls, no sign-in, no server, no storage outside your mailbox.
 *  - Note text is only ever rendered with .value / .textContent (never innerHTML).
 */
(function () {
  "use strict";

  const PROP = "mailnote";        // custom property name
  const CATEGORY = "Note";        // Outlook category used as the inbox marker
  const MAX_TEXT = 2000;          // characters of note text
  const MAX_JSON = 2500;          // Outlook limit for all custom properties on one item
  const AUTOSAVE_MS = 1200;

  const state = {
    seq: 0,          // bumps on every item change so late callbacks are ignored
    item: null,
    props: null,     // CustomProperties for the current item
    saved: "",       // last saved text
    savedMeta: null, // { created, modified }
    timer: null,
    saving: false,
  };

  const $ = (id) => document.getElementById(id);

  /* ------------------------------------------------------------ startup */

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) {
      fatal("This page is the My Mail Notes add-in. It only runs inside Outlook.");
      return;
    }
    if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
      fatal("This version of Outlook is too old for My Mail Notes. Update Outlook or use Outlook on the web.");
      return;
    }
    bindUi();
    if (Office.context.requirements.isSetSupported("Mailbox", "1.5")) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onItemChanged);
    }
    loadItem();
  });

  function onItemChanged() {
    // Flush a pending autosave for the previous email before switching.
    const pending = state.timer !== null;
    clearTimeout(state.timer);
    state.timer = null;
    if (pending && state.props && $("note").value.trim() !== state.saved) {
      saveNow(true).finally(loadItem);
    } else {
      loadItem();
    }
  }

  /* --------------------------------------------------------------- load */

  function loadItem() {
    const seq = ++state.seq;
    const item = Office.context.mailbox.item;
    state.item = item;
    state.props = null;
    state.saved = "";
    state.savedMeta = null;
    hideBanner();
    hideConfirm();
    setLoading(true);

    if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message) {
      setLoading(false);
      $("content").hidden = true;
      $("context-line").textContent = "Select an email to see or add a note.";
      return;
    }
    const subject = item.normalizedSubject || item.subject || "";
    $("context-line").textContent = subject ? "Private note for: " + subject : "Private note for this email";

    item.loadCustomPropertiesAsync(function (r) {
      if (seq !== state.seq) return;
      setLoading(false);
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        $("content").hidden = true;
        showBanner("error", "Couldn't load the note for this email. " + errText(r));
        return;
      }
      state.props = r.value;
      const data = parse(state.props.get(PROP));
      state.saved = data ? data.text : "";
      state.savedMeta = data;
      $("note").value = state.saved;
      $("content").hidden = false;
      updateUi();
      if (data && data.text) ensureCategory(item, seq);
    });
  }

  function parse(v) {
    if (!v) return null;
    try {
      const o = typeof v === "string" ? JSON.parse(v) : v;
      if (o && typeof o.text === "string") {
        return { text: o.text, created: o.created || null, modified: o.modified || null };
      }
    } catch (e) { /* ignore corrupt value */ }
    return null;
  }

  /* --------------------------------------------------------------- save */

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

      const text = clean($("note").value);
      if (text === state.saved) { updateUi(); return resolve(); }
      if (!text) { deleteNote(silent).then(resolve); return; }

      const now = new Date().toISOString();
      const value = {
        v: 1,
        text: text,
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
      props.saveAsync(function (r) {
        state.saving = false;
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          if (seq === state.seq) setStatus("error", "Not saved: " + errText(r) + " Try again.");
          return resolve();
        }
        if (seq === state.seq) {
          state.saved = text;
          state.savedMeta = value;
          updateUi();
        }
        ensureCategory(item, seq, silent);
        resolve();
      });
    });
  }

  function deleteNote(silent) {
    return new Promise(function (resolve) {
      const props = state.props;
      const item = state.item;
      const seq = state.seq;
      if (!props) return resolve();
      props.remove(PROP);
      props.saveAsync(function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          if (seq === state.seq) setStatus("error", "Couldn't delete: " + errText(r));
          return resolve();
        }
        if (seq === state.seq) {
          state.saved = "";
          state.savedMeta = null;
          $("note").value = "";
          updateUi();
          if (!silent) showBanner("ok", "Note deleted.", 2500);
        }
        removeCategory(item);
        resolve();
      });
    });
  }

  /* --------------------------------------------------- inbox marker (category) */

  function ensureCategory(item, seq, silent) {
    if (!item || !item.categories) return;
    item.categories.getAsync(function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return;
      const has = (r.value || []).some(function (c) { return c.displayName === CATEGORY; });
      if (has) return;
      item.categories.addAsync([CATEGORY], function (a) {
        if (a.status === Office.AsyncResultStatus.Succeeded || silent || seq !== state.seq) return;
        showBanner("info", 'Note saved. To mark emails that have notes, create an Outlook category called "' +
          CATEGORY + '" (Settings > Mail > Categories). After that, emails get tagged automatically.');
      });
    });
  }

  function removeCategory(item) {
    if (!item || !item.categories) return;
    item.categories.removeAsync([CATEGORY], function () { /* best effort */ });
  }

  /* ------------------------------------------------------------------ UI */

  function bindUi() {
    const ta = $("note");
    ta.maxLength = MAX_TEXT;
    ta.addEventListener("input", function () { updateUi(); scheduleSave(); });
    ta.addEventListener("blur", function () { if (state.timer) saveNow(false); });
    ta.addEventListener("keydown", function (e) {
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); saveNow(false); }
    });
    ta.addEventListener("keydown", function (e) {
      if (e.key.toLowerCase() === "d" && e.shiftKey && (e.ctrlKey || e.metaKey)) { e.preventDefault(); insertStamp(); }
    });
    $("stamp-btn").addEventListener("click", insertStamp);
    $("save-btn").addEventListener("click", function () { saveNow(false); });
    $("delete-btn").addEventListener("click", function () { $("confirm-delete").hidden = false; });
    $("confirm-no").addEventListener("click", hideConfirm);
    $("confirm-yes").addEventListener("click", function () { hideConfirm(); deleteNote(false); });
  }

  /* Insert "[22/10/2026 2:15 pm] " at the cursor, on its own line, then autosave. */
  function insertStamp() {
    const ta = $("note");
    const now = new Date();
    const stamp = "[" + now.toLocaleDateString("en-AU") + " " +
      now.toLocaleTimeString("en-AU", { hour: "numeric", minute: "2-digit" }) + "] ";
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
    renderMeta();
    if (state.saving) return;
    if (clean(current) !== state.saved) {
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

  function setLoading(on) {
    $("loading").hidden = !on;
    $("app").setAttribute("aria-busy", on ? "true" : "false");
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

  function clean(s) {
    return String(s || "")
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, "")
      .trim()
      .slice(0, MAX_TEXT);
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
