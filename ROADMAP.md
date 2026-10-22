# My Mail Notes – roadmap

Positioning: private, Microsoft 365-native notes on emails. Nothing leaves the mailbox, and an admin can deploy it centrally.

Rule: each phase is tested in real Outlook and signed off before the next phase starts.

| Phase | Status |
|---|---|
| 0 – Private notes MVP | ✅ Done (22 Oct 2026) |
| 1 – Quick wins | 🧪 Built, waiting for real-Outlook test and sign-off |
| 2 – Differentiators | ⏸ Not started (needs Phase 1 sign-off) |
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

## Phase 1 – Quick wins 🧪

| # | Feature | Notes |
|---|---|---|
| 1.1 | **Colour-coded notes**: None / Urgent / Follow-up / Info / Done | Each maps to an Outlook category (`Note`, `Note - Urgent`, `Note - Follow-up`, `Note - Info`, `Note - Done`). The user creates them once. If one is missing, the email falls back to `Note`. |
| 1.2 | **Quick-note buttons** | One click adds a date-stamped line, e.g. "Called – no answer". The list can be edited per user and is stored in the add-in's mailbox settings (roams with the user). |
| 1.3 | **Follow-up reminder** | Opens a new calendar event pre-filled with the note, subject and sender (tomorrow 9:00–9:15). The user adjusts the time and saves. Not available on mobile (Microsoft limitation). |
| 1.4 | **Notes on meetings (attendee view)** | Same pane on calendar meetings you've been invited to. Sent Items already work. |

Phase 1 test checklist: see README "Phase 1 tests".

Known limits:
- The colour categories must exist in each mailbox. Creating them automatically needs `ReadWriteMailbox`, which is deliberately not used. Phase 3 adds an admin script.
- A meeting you organised opens in edit mode, so the pane isn't offered there yet (Phase 2: compose support).

## Phase 2 – Differentiators ⏸ (needs Phase 1 sign-off)

| # | Feature | Notes / risk |
|---|---|---|
| 2.1 | **My Notes dashboard + search** | Lists all noted emails, searchable by note text and filterable by colour; click to open the email. Uses a small index in roaming settings (32 KB cap, so a few hundred notes). Item IDs change when an email is moved, so the dashboard needs a fallback. |
| 2.2 | **Export to CSV** | From the dashboard index |
| 2.3 | **Shared mailbox support** | Notes on reception@ / accounts@ visible to everyone with access. Needs `SupportsSharedFolders` and Mailbox 1.13, plus testing. |
| 2.4 | **Notes while composing** (emails and organiser meetings) | Custom properties in compose mode persist to Sent Items. Categories can't be set in compose mode on web/new Outlook. |
| 2.5 | **Mobile** (Outlook iOS/Android) | Manifest `MobileFormFactor` plus testing. Reminder button not available on mobile. |

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
