<#
.SYNOPSIS
    Creates the My Mail Notes colour categories in Outlook mailboxes, so users don't have to.

.DESCRIPTION
    Creates (or fixes the colour of) these Outlook categories:
        Note              Purple
        Note - Urgent     Red
        Note - Follow-up  Orange
        Note - Info       Blue
        Note - Done       Green
    Existing categories are left alone unless -FixColours is used. Nothing is deleted.
    Uses Microsoft Graph (outlookCategory). The add-in itself never gets this permission.

    Two ways to run:

    1) -Self  (one person, their own mailbox)
       Signs in as you with delegated MailboxSettings.ReadWrite. Work or school accounts.
       No admin rights needed (you may be asked to consent for yourself).

    2) -Users / -GroupId  (IT admin, many mailboxes, app-only)
       Uses an Entra app with a certificate. The app's access should be limited to the
       target mailboxes with Exchange "RBAC for Applications"
       (role "Application MailboxSettings.ReadWrite"). See ADMIN.md, section 3.

.NOTES
    Created by Vaibhav Girdhar. Part of My Mail Notes.

.REQUIREMENTS
    PowerShell 7+ (recommended) or Windows PowerShell 5.1.
    Module Microsoft.Graph.Authentication (installed for the current user if missing; no local admin needed).

.EXAMPLE
    # Your own mailbox
    .\Set-NoteCategories.ps1 -Self

.EXAMPLE
    # Selected users, app-only (run as IT admin)
    .\Set-NoteCategories.ps1 -TenantId contoso.onmicrosoft.com -ClientId <app-id> -CertificateThumbprint <thumbprint> `
        -Users alice@contoso.com,bob@contoso.com

.EXAMPLE
    # Everyone in a group, preview only
    .\Set-NoteCategories.ps1 -TenantId contoso.onmicrosoft.com -ClientId <app-id> -CertificateThumbprint <thumbprint> `
        -GroupId <group-object-id> -WhatIf
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Self')]
param(
    [Parameter(ParameterSetName = 'Self', Mandatory)] [switch] $Self,

    [Parameter(ParameterSetName = 'Users', Mandatory)] [Parameter(ParameterSetName = 'Group', Mandatory)] [string] $TenantId,
    [Parameter(ParameterSetName = 'Users', Mandatory)] [Parameter(ParameterSetName = 'Group', Mandatory)] [string] $ClientId,
    [Parameter(ParameterSetName = 'Users', Mandatory)] [Parameter(ParameterSetName = 'Group', Mandatory)] [string] $CertificateThumbprint,
    [Parameter(ParameterSetName = 'Users', Mandatory)] [string[]] $Users,
    [Parameter(ParameterSetName = 'Group', Mandatory)] [string] $GroupId,

    # Also correct the colour of categories that already exist with the right name.
    [switch] $FixColours
)

$ErrorActionPreference = 'Stop'

# Graph colour presets: preset0 Red, preset1 Orange, preset4 Green, preset7 Blue, preset8 Purple.
$Categories = @(
    @{ Name = 'Note';             Colour = 'preset8'; Label = 'Purple' }
    @{ Name = 'Note - Urgent';    Colour = 'preset0'; Label = 'Red' }
    @{ Name = 'Note - Follow-up'; Colour = 'preset1'; Label = 'Orange' }
    @{ Name = 'Note - Info';      Colour = 'preset7'; Label = 'Blue' }
    @{ Name = 'Note - Done';      Colour = 'preset4'; Label = 'Green' }
)

function Invoke-Graph {
    param([string]$Method, [string]$Uri, $Body)
    if ($null -ne $Body) {
        return Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body ($Body | ConvertTo-Json -Depth 5) -ContentType 'application/json' -OutputType HashTable
    }
    return Invoke-MgGraphRequest -Method $Method -Uri $Uri -OutputType HashTable
}

function Set-CategoriesFor {
    param([string]$Target, [string]$Label)
    $base = if ($Target -eq 'me') { 'v1.0/me/outlook/masterCategories' } else { "v1.0/users/$([uri]::EscapeDataString($Target))/outlook/masterCategories" }
    $result = [ordered]@{ Mailbox = $Label; Created = 0; Recoloured = 0; AlreadyOk = 0; Error = '' }
    try {
        $existing = @((Invoke-Graph GET $base)['value'])
        foreach ($c in $Categories) {
            $match = $existing | Where-Object { $_['displayName'] -eq $c.Name } | Select-Object -First 1
            if (-not $match) {
                if ($PSCmdlet.ShouldProcess($Label, "Create category '$($c.Name)' ($($c.Label))")) {
                    Invoke-Graph POST $base @{ displayName = $c.Name; color = $c.Colour } | Out-Null
                }
                $result.Created++
            } elseif ($FixColours -and $match['color'] -ne $c.Colour) {
                if ($PSCmdlet.ShouldProcess($Label, "Set colour of '$($c.Name)' to $($c.Label)")) {
                    Invoke-Graph PATCH "$base/$($match['id'])" @{ color = $c.Colour } | Out-Null
                }
                $result.Recoloured++
            } else {
                $result.AlreadyOk++
            }
        }
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match '403|Forbidden|Access') { $msg = "Access denied. Is this mailbox inside the app's RBAC scope? ($msg)" }
        $result.Error = $msg
    }
    [pscustomobject]$result
}

# ---------------------------------------------------------------- module + sign-in
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Host 'Installing Microsoft.Graph.Authentication for the current user...' -ForegroundColor Cyan
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Graph.Authentication

$results = @()
if ($PSCmdlet.ParameterSetName -eq 'Self') {
    Connect-MgGraph -Scopes 'MailboxSettings.ReadWrite' -NoWelcome
    $who = (Get-MgContext).Account
    $results += Set-CategoriesFor -Target 'me' -Label $who
} else {
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
    $targets = @()
    if ($PSCmdlet.ParameterSetName -eq 'Users') {
        $targets = $Users
    } else {
        # Direct members only (matches how RBAC for Applications scopes groups). Needs GroupMember.Read.All on the app.
        $uri = "v1.0/groups/$GroupId/members/microsoft.graph.user?`$select=userPrincipalName,mail&`$top=999"
        while ($uri) {
            $page = Invoke-Graph GET $uri
            $targets += @($page['value'] | ForEach-Object { if ($_['mail']) { $_['mail'] } else { $_['userPrincipalName'] } })
            $uri = $page['@odata.nextLink']
        }
        Write-Host "Group has $($targets.Count) user(s)." -ForegroundColor Cyan
    }
    foreach ($t in $targets) {
        $results += Set-CategoriesFor -Target $t -Label $t
    }
}

Disconnect-MgGraph | Out-Null

$results | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
$failed = @($results | Where-Object { $_.Error })
if ($WhatIfPreference) { Write-Host 'Preview only (-WhatIf): nothing was changed.' -ForegroundColor Yellow }
if ($failed.Count) {
    Write-Host "$($failed.Count) mailbox(es) failed. See the Error column." -ForegroundColor Red
    exit 1
}
Write-Host 'Done. Users see the categories straight away (refresh Outlook).' -ForegroundColor Green
