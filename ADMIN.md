# My Mail Notes – admin guide (Phase 3)

Created by Vaibhav Girdhar.

For IT admins rolling the add-in out to an organisation. The GitHub Pages version stays live for testing. Nothing here changes it until you choose to switch.

| Step | What | When |
|---|---|---|
| 1 | Security review | Before any work or client tenant. See `SECURITY.md` and run `node security-check.mjs`. |
| 2 | Deploy centrally to a pilot group | Instead of each user sideloading |
| 3 | Create the colour categories for everyone | Once per tenant, plus new starters |
| 4 | Move hosting under your control | Before production |
| 5 | Lock down user sideloading (optional) | After central deployment |

---

## 1. Security review

- Send the reviewer `SECURITY.md`.
- Check that **GitHub → Actions → Security check** is green for the commit being deployed.
- The add-in asks for `ReadWriteItem` only. It needs **no admin consent, no Entra app and no Graph permission**.

## 2. Deploy centrally (Microsoft 365 admin centre)

Requires **Global Administrator** or **Exchange Administrator**.

1. https://admin.microsoft.com → **Settings → Integrated apps → Upload custom apps**.
2. App type: **Office Add-in**. Choose **Upload manifest file (.xml) from device** → `manifest.xml`.
3. **Assign users:** pick a pilot group, e.g. `SG-Pilot-MailNotes`. Avoid "Entire organisation" for the pilot.
4. Accept the permissions (read/write the open item) → **Deploy**.
5. It can take **up to 24 hours** to appear in users' Outlook. Users find it under **Apps → My Mail Notes**.
6. Pilot users should **remove their own sideloaded copy** first (https://aka.ms/olksideload → My add-ins → Remove). Notes are kept, because it's the same add-in ID.

To update: upload the new `manifest.xml` in the same place (only needed when the manifest changes). To remove: **Integrated apps → My Mail Notes → Remove app**. Notes stay on the emails.

## 3. Create the colour categories for everyone

The add-in deliberately can't create categories (that would need full-mailbox access). `admin/Set-NoteCategories.ps1` does it once, from IT. It creates `Note` (purple), `Note - Urgent` (red), `Note - Follow-up` (orange), `Note - Info` (blue) and `Note - Done` (green). Safe to re-run, never deletes anything, and `-WhatIf` previews.

### Option A: one person, their own mailbox (no admin)

```powershell
# PowerShell 7, normal user (no Administrator needed)
.\admin\Set-NoteCategories.ps1 -Self
```

### Option B: many mailboxes, app-only, limited to a group (recommended for clients)

This uses Exchange **RBAC for Applications**, so the app can only touch mailboxes in the group you choose, and only their settings (not their mail).

**B1. Create the app (Entra admin centre):** App registrations → New registration → `MyMailNotes-CategorySetup`, single tenant. Under **Certificates & secrets**, upload a certificate (keep the private key on the admin PC only). **Don't** add Graph `MailboxSettings.ReadWrite` in API permissions: the access comes from step B2. Note the **Application (client) ID**, and the **Object ID** from **Enterprise applications** (not App registrations).

**B2. Scope it to a group (Exchange Online PowerShell, Exchange Administrator):**

```powershell
Connect-ExchangeOnline
$dn = (Get-Group -Identity "SG-MailNotes-Users").DistinguishedName
New-ServicePrincipal -AppId <client-id> -ObjectId <enterprise-app-object-id> -DisplayName "MyMailNotes-CategorySetup"
New-ManagementScope -Name "MyMailNotes-Users" -RecipientRestrictionFilter "MemberOfGroup -eq '$dn'"
New-ManagementRoleAssignment -App "MyMailNotes-CategorySetup" -Role "Application MailboxSettings.ReadWrite" -CustomResourceScope "MyMailNotes-Users"
# Check: InScope should be True for a member, False for anyone else
Test-ServicePrincipalAuthorization -Identity "MyMailNotes-CategorySetup" -Resource someone@contoso.com | Format-Table
```

Only **direct** group members are in scope (not nested groups). Changes can take 30 minutes to 2 hours to apply.

**B3. Run it:**

```powershell
# Preview first
.\admin\Set-NoteCategories.ps1 -TenantId contoso.onmicrosoft.com -ClientId <client-id> -CertificateThumbprint <thumbprint> `
    -Users (Import-Csv .\users.csv).UserPrincipalName -WhatIf
# Then for real (remove -WhatIf)
```

`users.csv` needs one column, `UserPrincipalName`. To use `-GroupId` instead of a CSV, the app also needs Graph **GroupMember.Read.All** (application). The CSV route avoids that extra permission.

**New starters:** add them to the group and re-run the script (for example as part of onboarding).

## 4. Move hosting under your control (production)

GitHub Pages under a personal account is fine for testing, not for clients: whoever controls that GitHub account controls the code. Recommended: **Azure Static Web Apps (Free plan)** in a subscription you or the client owns.

1. Azure portal → **Static Web Apps → Create**. Plan: **Free**. Deployment: **Other** (manual).
2. Deploy the contents of this folder (it already includes `staticwebapp.config.json`, which adds HTTPS-only, no-sniff, no-referrer and CSP headers). Example with the SWA CLI from a PC with Node.js:
   ```powershell
   npx @azure/static-web-apps-cli deploy . --deployment-token <token-from-portal> --env production
   ```
3. Optional: add a custom domain, e.g. `mailnotes.yourdomain.com.au`.
4. Point the manifest at the new address:
   ```powershell
   (Get-Content .\manifest.xml -Raw).Replace('https://pawpinku2-cmd.github.io/mail-notes', 'https://mailnotes.yourdomain.com.au') |
       Set-Content .\manifest.xml -Encoding utf8
   ```
5. **Raise the version** in `manifest.xml` (e.g. `1.2.0.0` → `1.3.0.0`), run `node security-check.mjs`, then upload the manifest in Integrated apps (step 2).
6. Keep the source in a **private** repo with branch protection and MFA.

Don't change `<Id>` in the manifest. Notes are tied to it.

## 5. Lock down user sideloading (optional, after step 2)

So users can only run add-ins IT has approved:

```powershell
Connect-ExchangeOnline
Get-ManagementRoleAssignment -RoleAssignee "Default Role Assignment Policy" -Role "My Custom Apps" |
    Remove-ManagementRoleAssignment -Confirm:$false
```

Test this on a pilot role assignment policy first. It removes **Add from file** for everyone on that policy, including for testing.

## Support

- In the pane: **Diagnostics** (bottom) shows version, account type, API support and the last error. Ask users for a screenshot of it.
- **Outlook.com** personal accounts have no pinned panes. Users close and reopen the pane for each email.
- **Mobile:** no colour tags in the inbox, no reminder, and items can't be opened from All notes (Microsoft limits).
