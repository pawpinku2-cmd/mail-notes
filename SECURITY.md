# My Mail Notes – security summary

For IT and cyber security reviewers. Version 1.2.3. Created by Vaibhav Girdhar.

**In one line:** an Outlook add-in that stores a private note on the email you have open. It has no server, makes no network calls of its own, asks for the lowest permission that can write to an email, and can't read, send, forward or delete mail.

## What it can and can't do

| | |
|---|---|
| **Outlook permission** | `ReadWriteItem`: only the email or meeting the user has open, only while the pane is open. Not `ReadWriteMailbox`. |
| **Reads** | The open item's subject, sender, date, item ID, message ID and categories. **Never the body or attachments.** |
| **Writes** | (1) A hidden note on the open item (Office.js custom property, max 2,500 characters, not sent to recipients). (2) One of its own categories: `Note`, `Note - Urgent`, `Note - Follow-up`, `Note - Info`, `Note - Done`. (3) Its own settings: the quick-note list and the All notes index (Office.js roaming settings, 32 KB). |
| **Opens** | An unsaved calendar form (the user must press Save), and an email the user clicks in the All notes list. |
| **Never** | Send, reply, forward, delete or move mail. Change recipients, body or headers. Create rules or forwarding. Read other items. Request mailbox tokens, EWS, Graph or any sign-in. |
| **Network** | None from the add-in's code. The only external file is Microsoft's own Office.js. |
| **Data location** | Inside the user's mailbox (Exchange Online / Outlook.com). Nothing goes to the add-in publisher or any third party. |
| **Who can see notes** | The mailbox owner, plus anyone who already has full access to that mailbox (delegates, shared-mailbox members, and admins using compliance tools). |

## How it's enforced (not just promised)

1. **Manifest:** `ReadWriteItem` only, no extra domains, no sign-in, no on-send or background events.
2. **Content Security Policy** in `taskpane.html`: blocks everything by default. Scripts load only from the add-in's own site and `appsforoffice.microsoft.com`. Network calls go only to Microsoft's Office.js CDN. No frames, forms, plugins or inline scripts.
3. **Safe rendering:** note text is shown as plain text only (`textContent` / `.value`). It's never parsed as HTML, so a note can't inject code.
4. **CSV export** neutralises cells starting with `= + - @`, so a note can't become an Excel formula.
5. **Automated proof:** `security-check.mjs` scans the code and manifests on every change (GitHub Actions → "Security check"). It runs 58 checks covering everything in the "Never" row above and more. It was tested by injecting bad code (sending data out, grabbing a mailbox token, reading the body, sending mail, forwarding, unsafe HTML, upgrading to `ReadWriteMailbox`). It caught every one.

Run it yourself:

```bash
node security-check.mjs
```

## How a reviewer can verify it independently (10 minutes)

1. Read `taskpane.js`, the only code that runs (about 1,200 lines, no libraries, not minified). Search it for `fetch`, `XMLHttpRequest`, `getCallbackTokenAsync`, `body.getAsync`, `sendAsync`: none are there.
2. Read `manifest.xml`: `<Permissions>ReadWriteItem</Permissions>`, and every URL points to one host.
3. In Outlook on the web, open the pane, press **F12 → Network**, then add, edit and delete a note. No requests leave Outlook apart from loading the add-in's own files and Office.js.
4. Check that **GitHub → Actions → Security check** is green for the commit being deployed.

## Risks and how they're controlled

| Risk | Control |
|---|---|
| Someone changes the hosted code (e.g. GitHub account compromise) | 2FA on the GitHub account. The Security check flags dangerous code. **Production: host under the organisation's own control** (Azure Static Web Apps, private repo, branch protection). See ADMIN.md. |
| Public test repo exposes the code | No secrets in the code. Being public also makes it easy to audit. Production uses a private repo. |
| Users installing unapproved add-ins | Admins can turn off user sideloading (remove the "My Custom Apps" role from the user role assignment policy) and deploy approved add-ins centrally via Integrated apps. |
| Sensitive information typed into notes | Notes stay in the mailbox and fall under the organisation's existing retention, eDiscovery and access controls. The pane reminds users not to store passwords or card numbers. |
| Data loss if the add-in is removed | Notes stay on the emails. Reinstalling the same manifest (same add-in ID) brings them back. |

## Australian context

- **Privacy Act / APP 11:** no new data processor and no data leaves the Microsoft 365 tenant. Notes are covered by the tenant's existing data residency, retention and access controls.
- **ACSC Essential Eight:** fits application control (admin-deployed, sideloading can be blocked) and restricting admin privileges (users need no admin rights; the add-in needs no consent).
