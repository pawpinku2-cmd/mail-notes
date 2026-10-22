<#
.SYNOPSIS
    Creates the My Mail Notes colour categories using classic Outlook (desktop) on this PC.

.DESCRIPTION
    Creates these categories in the Outlook master category list:
        Note              Purple
        Note - Urgent     Red
        Note - Follow-up  Orange
        Note - Info       Blue
        Note - Done       Green
    For Exchange / M365 mailboxes the list is stored in the mailbox, so the categories
    also appear in Outlook on the web, new Outlook and mobile after sync.
    Existing categories are left alone unless -FixColours is used. Nothing is deleted.
    No Graph, no admin consent, no network calls: it only talks to Outlook on this PC.

.NOTES
    Created by Vaibhav Girdhar. Part of My Mail Notes.
    Run in NORMAL PowerShell as the signed-in user (NOT "Run as Administrator").
    Classic Outlook must be installed. New Outlook does not support this (no COM).

.EXAMPLE
    .\Set-NoteCategories-Classic.ps1                     # your own mailbox
.EXAMPLE
    .\Set-NoteCategories-Classic.ps1 -AllMailboxes       # plus shared/delegate mailboxes open in Outlook
.EXAMPLE
    .\Set-NoteCategories-Classic.ps1 -AllMailboxes -FixColours -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch] $AllMailboxes,
    [switch] $FixColours
)

# OlCategoryColor values (match the Graph presets used by Set-NoteCategories.ps1)
$Categories = @(
    @{ Name = 'Note';             Colour = 9; Label = 'Purple' }
    @{ Name = 'Note - Urgent';    Colour = 1; Label = 'Red' }
    @{ Name = 'Note - Follow-up'; Colour = 2; Label = 'Orange' }
    @{ Name = 'Note - Info';      Colour = 8; Label = 'Blue' }
    @{ Name = 'Note - Done';      Colour = 5; Label = 'Green' }
)

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) {
    Write-Warning 'This window is running as Administrator. Outlook runs as you, so this can fail. Re-run in a normal PowerShell window.'
}

try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace('MAPI')
} catch {
    Write-Host 'Could not start classic Outlook. Is it installed (not just new Outlook) and set up with a profile?' -ForegroundColor Red
    exit 1
}

# 0 = primary mailbox, 1 = delegate mailbox, 4 = additional mailbox (shared). Skips PSTs and public folders.
$stores = if ($AllMailboxes) { @($ns.Stores | Where-Object { $_.ExchangeStoreType -in 0, 1, 4 }) } else { @($ns.DefaultStore) }

$results = foreach ($store in $stores) {
    $r = [ordered]@{ Mailbox = $store.DisplayName; Created = 0; Recoloured = 0; AlreadyOk = 0; Error = '' }
    try {
        $list = $store.Categories
        foreach ($c in $Categories) {
            $match = $null
            foreach ($existing in $list) { if ($existing.Name -eq $c.Name) { $match = $existing; break } }
            if (-not $match) {
                if ($PSCmdlet.ShouldProcess($store.DisplayName, "Create '$($c.Name)' ($($c.Label))")) { [void]$list.Add($c.Name, $c.Colour) }
                $r.Created++
            } elseif ($FixColours -and $match.Color -ne $c.Colour) {
                if ($PSCmdlet.ShouldProcess($store.DisplayName, "Set '$($c.Name)' to $($c.Label)")) { $match.Color = $c.Colour }
                $r.Recoloured++
            } else {
                $r.AlreadyOk++
            }
        }
    } catch {
        $r.Error = "$($_.Exception.Message) (shared/delegate mailboxes need Full Access to change categories)"
    }
    [pscustomobject]$r
}

$results | Format-Table -AutoSize
if ($WhatIfPreference) { Write-Host 'Preview only (-WhatIf): nothing was changed.' -ForegroundColor Yellow }
else { Write-Host 'Done. Categories sync to Outlook on the web and new Outlook shortly.' -ForegroundColor Green }

[void][Runtime.InteropServices.Marshal]::ReleaseComObject($ns)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($outlook)
