
<#
.SYNOPSIS
Exports all shared mailboxes in Exchange Online for a specific SMTP domain with who has access to them into a CSV.

.PARAMETER Domain
SMTP domain to filter shared mailboxes by (matches primary SMTP or any proxy address).

.REQUIREMENTS
- Exchange Online PowerShell V3 module (EXO V3).
  Install-Module ExchangeOnlineManagement -Scope CurrentUser
- Permission to read mailbox and recipient permissions.

.OUTPUT
- CSV: .\SharedMailbox_Permissions_<domain>_<yyyyMMdd_HHmm>.csv
#>

param(
    [string]$OutputFolder = ".",
    [string]$Domain = "microsoft.com",
    [switch]$IncludeInheritedFullAccess
)

# 1) Connect to Exchange Online
if (-not (Get-Module ExchangeOnlineManagement -ListAvailable)) {
    Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Yellow
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
}
Import-Module ExchangeOnlineManagement

try {
    if (-not (Get-ConnectionInformation)) {
        Connect-ExchangeOnline -ShowProgress $true
    }
} catch {
    Connect-ExchangeOnline -ShowProgress $true
}

# 2) Helper: Exclusion filters
$systemPrincipals = @(
    "NT AUTHORITY\SELF",
    "S-1-5-",
    "NT AUTHORITY\ANONYMOUS LOGON",
    "NT AUTHORITY\Authenticated Users",
    "Exchange Servers",
    "FederatedEmail.4c1f4d8b-8179-4148-93bf-00a95fa1e042",
    "Discovery Management"
)

function Test-IsSystemPrincipal {
    param([string]$Identity)
    if ([string]::IsNullOrWhiteSpace($Identity)) { return $true }
    if ($systemPrincipals | Where-Object { $Identity -like ("{0}*" -f $_) -or $Identity -eq $_ }) {
        return $true
    }
    return $false
}

# 3) Get shared mailboxes for the domain
Write-Host "Retrieving shared mailboxes in domain @$Domain ..." -ForegroundColor Cyan

# Pull properties needed to evaluate both primary SMTP and proxy addresses
$allShared = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited `
    -Properties GrantSendOnBehalfTo, PrimarySmtpAddress, DisplayName, EmailAddresses

# Safe regex: end with @domain
$domainPattern = '@' + [regex]::Escape($Domain) + '$'

$sharedMailboxes = $allShared | Where-Object {
    ($_.PrimarySmtpAddress -match $domainPattern) -or
    ($_.EmailAddresses -and ($_.EmailAddresses | Where-Object { $_ -match $domainPattern }))
}

if (-not $sharedMailboxes) {
    Write-Warning "No shared mailboxes found with SMTP addresses in @$Domain."
    # still produce an empty CSV with headers to be consistent
    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmm")
    $safeDomain = $Domain -replace '[^\w.-]','_'
    $csvPath    = Join-Path $OutputFolder ("SharedMailbox_Permissions_{0}_{1}.csv" -f $safeDomain, $timestamp)

    [pscustomobject]@{
        MailboxPrimarySmtpAddress = $null
        MailboxDisplayName        = $null
        PermissionType            = $null
        User                      = $null
        AccessRights              = $null
        IsInherited               = $null
    } | Select-Object * |
        Export-Csv -Path $csvPath -Encoding UTF8 -NoTypeInformation

    Write-Host "`nCreated empty CSV: $csvPath" -ForegroundColor Yellow
    return
}

# 4) Collect permissions
$results = New-Object System.Collections.Generic.List[object]
$idx = 0
$total = $sharedMailboxes.Count

foreach ($mbx in $sharedMailboxes) {
    $idx++
    Write-Progress -Activity "Processing shared mailboxes" -Status "$($mbx.DisplayName) ($idx of $total)" -PercentComplete (($idx / $total) * 100)

    $mailboxAddress = $mbx.PrimarySmtpAddress
    $mailboxName    = $mbx.DisplayName

    # Full Access
    try {
        $faPerms = Get-EXOMailboxPermission -Identity $mailboxAddress -ResultSize Unlimited -ErrorAction Stop
        $faPerms = $faPerms | Where-Object {
            -not (Test-IsSystemPrincipal $_.User) -and
            $_.User -ne $mailboxName -and
            ($IncludeInheritedFullAccess -or ($_.IsInherited -eq $false)) -and
            ($_.AccessRights -contains "FullAccess")
        }
        foreach ($p in $faPerms) {
            $results.Add([pscustomobject]@{
                MailboxPrimarySmtpAddress = $mailboxAddress
                MailboxDisplayName        = $mailboxName
                PermissionType            = "FullAccess"
                User                      = $p.User
                AccessRights              = "FullAccess"
                IsInherited               = $p.IsInherited
            })
        }
    } catch {
        Write-Warning "Failed to get FullAccess for $mailboxAddress : $($_.Exception.Message)"
    }

    # SendAs
    try {
        $saPerms = Get-RecipientPermission -Identity $mailboxAddress -ErrorAction Stop | Where-Object {
            ($_ -ne $null) -and
            -not $_.Deny -and
            ($_.AccessRights -contains "SendAs") -and
            -not (Test-IsSystemPrincipal $_.Trustee)
        }
        foreach ($p in $saPerms) {
            $results.Add([pscustomobject]@{
                MailboxPrimarySmtpAddress = $mailboxAddress
                MailboxDisplayName        = $mailboxName
                PermissionType            = "SendAs"
                User                      = $p.Trustee
                AccessRights              = "SendAs"
                IsInherited               = $null
            })
        }
    } catch {
        Write-Warning "Failed to get SendAs for $mailboxAddress : $($_.Exception.Message)"
    }

    # SendOnBehalf
    try {
        $sobUsers = $mbx.GrantSendOnBehalfTo
        if ($sobUsers) {
            foreach ($u in $sobUsers) {
                $userDisplay = $u.ToString()
                if (-not (Test-IsSystemPrincipal $userDisplay)) {
                    $results.Add([pscustomobject]@{
                        MailboxPrimarySmtpAddress = $mailboxAddress
                        MailboxDisplayName        = $mailboxName
                        PermissionType            = "SendOnBehalf"
                        User                      = $userDisplay
                        AccessRights              = "SendOnBehalf"
                        IsInherited               = $null
                    })
                }
            }
        }
    } catch {
        Write-Warning "Failed to get SendOnBehalf for $mailboxAddress : $($_.Exception.Message)"
    }
}

# 5) Export to CSV (single pipeline line + empty-result handling)
$timestamp = (Get-Date).ToString("yyyyMMdd_HHmm")
$safeDomain = $Domain -replace '[^\w.-]','_'
$csvPath    = Join-Path $OutputFolder ("SharedMailbox_Permissions_{0}_{1}.csv" -f $safeDomain, $timestamp)

if ($results.Count -gt 0) {
    $results | Sort-Object MailboxPrimarySmtpAddress, PermissionType, User |
        Export-Csv -Path $csvPath -Encoding UTF8 -NoTypeInformation
    Write-Host "`nExport complete: $csvPath" -ForegroundColor Green
} else {
    [pscustomobject]@{
        MailboxPrimarySmtpAddress = $null
        MailboxDisplayName        = $null
        PermissionType            = $null
        User                      = $null
        AccessRights              = $null
        IsInherited               = $null
    } | Select-Object * |
        Export-Csv -Path $csvPath -Encoding UTF8 -NoTypeInformation
    Write-Warning "No permissions found for shared mailboxes in @$Domain. Empty CSV created: $csvPath"
}

# 6) Disconnect (optional)
# Disconnect-ExchangeOnline -Confirm:$false
