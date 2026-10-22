# My Mail Notes – roadmap

Positioning: private, Microsoft 365-native notes on emails. Nothing leaves the mailbox, and an admin can deploy it centrally.

Rule: each phase is tested in real Outlook and signed off before the next phase starts.

| Phase | Status |
|---|---|
| 0 – Private notes MVP | ✅ Done (22 Oct 2026) |
| 1 – Quick wins | ✅ Tested by owner in Outlook on the web (22 Oct 2026). Narrow-pane layout bug found from the test screenshot and fixed |
| 2 – Differentiators | 🧪 Built and deployed (v1.2, 22 Oct 2026). Real-Outlook test pending |
| 3 – Optional / on demand | ⏸ Not started |

---

## Phase 0 – Private notes MVP ✅

- [x] One private note per email, stored on the email (Office.js custom properties)
- [x] Autosave, delete, 2,000-character limit
- [x] Inbox marker via the `Note` category
- [x] Added / Last edited date and time, plus a "+ Date/time" stamp button (Ctrl+Shift+D)
- [x] Works in Outlook on the web, new Outlook and classic Outlook
- [x] No server, no sign-in, no network calls. CSP locked down. Least-privilege `ReadWriteItem`
- [x] Hosted on GitHub Pages: https://pawpinku2-cmd.github.io/mail-notes/

## Phase 1 – Quick wins ✅

| # | Feature | Notes |
|---|---|---|
| 1.1 | **Colour-coded notes**: None / Urgent / Follow-up / Info / Done | Each maps to an Outlook category (`Note`, `Note - Urgent`, `Note - Follow-up`, `Note - Info`, `Note - Done`). The user creates them once. If one is missing, the email falls back to `Note`. |
| 1.2 | **Quick-note buttons** | One click adds a date-stamped line, e.g. "Called – no answer". The list can be edited per user and is stored in the add-in's mailbox settings (roams with the user). |
| 1.3 | **Follow-up reminder** | Opens a new calendar event pre-filled with the note, subject and sender (tomorrow 9:00–9:15). The user adjusts the time and saves. Not available on mobile (Microsoft limitation). |
| 1.4 | **Notes on meetings (attendee view)** | Same pane on calendar meetings you've been invited to. Sent Items already work. |

Phase 1 test checklist: see README "Phase 1 tests". Tested by the owner on 22 Oct 2026 before Phase 2 started.

Fixed after test: at Outlook's default pane width (~300 px), the button row overflowed ("+ Date/time" clipped on the left, "/ 2000" and "Edit" on the right). Rows now wrap. Checked at 280, 300, 320 and 360 px.

Known limits:
- The colour categories must exist in each mailbox. Creating them automatically needs `ReadWriteMailbox`, which is deliberately not used. Phase 3 adds an admin script.
- A meeting you organised opens in edit mode, so the pane isn't offered there yet (fixed in Phase 2.4).

## Phase 2 – Differentiators 🧪 (v1.2)

| # | Feature | Status / notes |
|---|---|---|
| 2.1 | **All notes dashboard + search** | ✅ Built. "All notes" tab: search subject, sender and note text (first 160 characters); filter by colour; click to open. Index lives in roaming settings (32 KB). **Real capacity: about 60–120 notes**, depending on subject and note length. When full, the oldest "Done" notes drop off the list first; the notes themselves stay on their emails. Notes written before v1.2 appear once each email is opened. Moved emails refresh their link when opened. |
| 2.2 | **Export to CSV** | ✅ Built. Exports what's listed (after search/filter). Download, or Copy as a fallback. Cells starting with = + - @ are neutralised (formula-injection protection). UTF-8 BOM so Excel shows symbols correctly. |
| 2.3 | **Shared mailbox support** | ✅ Built, **needs real test**. `SupportsSharedFolders` on. Pane shows "Shared mailbox (owner): everyone with access can see this note." Delegates can't create categories, so the mailbox owner/admin must create the colour categories in the shared mailbox. |
| 2.4 | **Notes while composing** (emails and organiser meetings) | ✅ Built, **needs real test**. Note saved on the draft and kept on the Sent Items copy. No category or dashboard entry on a draft (Outlook limitation). Both apply when the sent item is opened. Reminder hidden in compose. |
| 2.5 | **Mobile** (Outlook iOS/Android) | ✅ Built, **needs real test**. Mobile button on received emails. Notes, colours (saved, but no inbox tag on mobile), quick notes and dashboard work. Not on mobile (Microsoft limits): reminder, opening items from the dashboard, inbox colour tags. |

Fallback: `manifest-desktop.xml` is the same v1.2 manifest without the mobile section, for use if Outlook rejects `manifest.xml`.

## Phase 3 – Optional / on demand ⏸

| # | Feature | Notes |
|---|---|---|
| 3.1 | Production hosting | Azure Static Web Apps (or similar) on a custom domain with security headers. Private repo. |
| 3.2 | Admin pack | Central deploy via Integrated apps, a script to pre-create colour categories, a usage check |
| 3.3 | AI note suggestions | Only with Azure OpenAI in the client's own tenant (Australia East). Hold unless clients ask. |

## Platform limits (won't change)

- Add-ins can't show note text in the inbox list. Only the coloured category tag shows.
- Outlook's own search can't see note text (2.1 is the workaround).
- Notes don't follow forwards or copies of an email.
