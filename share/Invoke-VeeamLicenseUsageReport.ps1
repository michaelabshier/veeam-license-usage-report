#Requires -Version 5.1
<#
.SYNOPSIS
    Downloads the aggregated License Usage Report from Veeam Backup Enterprise
    Manager, and can install itself as a scheduled task that runs on the last
    day of every month.

.DESCRIPTION
    Enterprise Manager produces one aggregated License Usage Report covering
    every backup server it manages, with a single company-level summary. This
    script retrieves that report unattended, in the same PDF and JSON formats
    the Download button in the Enterprise Manager UI offers.

    Each run produces, in a month-stamped folder:

      * VeeamLicenseUsageReport_<yyyy-MM>.pdf    the aggregated report
      * VeeamLicenseUsageReport_<yyyy-MM>.json   the same, machine readable
      * LicenseSummary_<yyyy-MM>.json            the licensing summary
      * a row appended to LicenseUsage_History.csv for month-over-month trending

    This is a single self-contained file. There is no configuration file:
    settings are passed as parameters, and -Install writes whichever parameters
    you used into the scheduled task itself.

    IMPORTANT, PLEASE READ

    Enterprise Manager exposes no supported or documented API for this report.
    The endpoints used here are the ones its own web interface calls, and the
    script authenticates by performing the same form sign-in a browser performs.

    That works, and it has been tested, but it is not a contract. Veeam may
    change these endpoints in any update, without notice, and such a change is
    not a defect in the product. After any Veeam upgrade, re-run with
    -TestConnection before relying on the scheduled task.

    A supported alternative exists, with a trade-off. The Veeam cmdlet
    Generate-VBRLicenseUsageReport is official and documented, but it runs on a
    backup server and reports only that server's usage, so covering an estate
    means running it on each server and combining the results yourself.

.PARAMETER Server
    Enterprise Manager FQDN, for example em.contoso.com.

.PARAMETER Port
    Enterprise Manager web port. Defaults to 9443. This is the web application
    port, not the 9398 REST API port.

.PARAMETER CredentialPath
    Where the encrypted Enterprise Manager credential is stored. The file is
    protected with DPAPI, which ties it to one Windows account on one machine,
    so it must be created by the account that will run the scheduled task, on
    the server that will run it.

.PARAMETER OutputFolder
    Root folder for the month folders, logs and the history CSV.

.PARAMETER Formats
    Which formats to download: Pdf, Json, or both. Defaults to both.

.PARAMETER Monthly
    Download Veeam's archived monthly report for the month just ended, instead
    of a current-consumption snapshot.

    Omitted (the default), the report is headed "Current license consumption"
    and reflects usage at the moment it runs, which is why the scheduled task
    runs on the last day of the month.

    Supplied, the report is headed with an explicit interval such as
    "01 October 2026 - 31 October 2026". Enterprise Manager generates that
    report on the first day of each month, so pair this with -ScheduleTime on a
    task you have moved to run early in the month.

.PARAMETER RetentionMonths
    Month folders older than this are deleted at the end of a run. Use 0 to
    keep everything.

.PARAMETER SkipCertificateCheck
    Accept the Enterprise Manager TLS certificate without validating it.
    Required while Enterprise Manager uses its default self-signed certificate.
    Prefer installing a trusted certificate and omitting this.

.PARAMETER SmtpServer
    Supply this, with -MailFrom and -MailTo, to email the report after a run.

.PARAMETER SaveCredential
    Prompt for the Enterprise Manager credential and store it encrypted at
    -CredentialPath. The account needs the Portal Administrator role in
    Enterprise Manager.

.PARAMETER CredentialScope
    Machine (the default) encrypts the credential to this computer, so a task
    running as SYSTEM or as a service account can read it without that account
    ever signing in. This is what allows the report to run with nobody logged
    on. Anything running on this computer can decrypt it, so the file is
    stripped of inherited permissions and granted only to SYSTEM, the local
    Administrators group, the account that saved it, and anything named by
    -GrantAccess.

    User restores the older per-account encryption. Only the account that saved
    the file can then read it, which means signing in as the task account to
    create it.

    Neither form works on another computer.

.PARAMETER GrantAccess
    Extra accounts to grant read access to the credential file, for when the
    task runs as something other than SYSTEM or the account saving it, for
    example 'CONTOSO\svc-veeam'.

.PARAMETER TestConnection
    Sign in, fetch the licensing summary, generate the report and confirm a real
    document comes back. Saves nothing. Run this after setup, and after every
    Veeam upgrade.

.PARAMETER RunAsUser
    The Windows account the scheduled task logs on as. Defaults to SYSTEM.

    This is not the Enterprise Manager login. That one is prompted for by
    -SaveCredential, stored at -CredentialPath, and sent over HTTPS at run
    time. Passing an Enterprise Manager user name here fails with "No mapping
    between account names and security IDs was done", because Windows is being
    asked to log on as an account it has never heard of.

    The account must be one this computer can resolve: a member of its own
    domain, or a local account. Install the task on a server joined to the
    domain of the service account you intend to use. Whichever account you
    choose, run -SaveCredential as that account on that computer, because the
    credential file is DPAPI encrypted per account per machine.

.PARAMETER TaskLogonType
    How the task authenticates as -RunAsUser.

    Auto (the default) stores a password, and so runs whether or not anyone is
    signed in. Password is the same thing, stated explicitly.

    Interactive stores no password and runs only while that account is signed
    in. Use it when Task Scheduler cannot cache the password, which is the case
    for a Microsoft Entra joined workstation, and on anything else treat it as
    a way to try the task out rather than to run it unattended.

.PARAMETER Install
    Register the scheduled task, on the last day of every month. Whichever
    parameters you pass alongside -Install are written into the task, so pass
    the same ones you used for -TestConnection.

.PARAMETER Uninstall
    Remove the scheduled task.

.EXAMPLE
    .\Invoke-VeeamLicenseUsageReport.ps1 -SaveCredential

    Store the credential. Run this first, as the account that will run the task.

.EXAMPLE
    .\Invoke-VeeamLicenseUsageReport.ps1 -TestConnection -Server em.contoso.com -SkipCertificateCheck

    Confirm the whole path works before scheduling anything.

.EXAMPLE
    .\Invoke-VeeamLicenseUsageReport.ps1 -Server em.contoso.com -SkipCertificateCheck

    Run once now, writing to C:\VeeamLicenseReports.

.EXAMPLE
    .\Invoke-VeeamLicenseUsageReport.ps1 -Install -Server em.contoso.com -SkipCertificateCheck -RunAsUser 'CONTOSO\svc-veeam' -ScheduleTime 22:00

    Schedule it for 22:00 on the last day of every month.

.EXAMPLE
    .\Invoke-VeeamLicenseUsageReport.ps1 -Server em.contoso.com -SkipCertificateCheck -SmtpServer smtp.contoso.com -MailFrom veeam@contoso.com -MailTo licensing@contoso.com

    Run once and email the result.

.NOTES
    Requires Windows PowerShell 5.1 or later, network access to the Enterprise
    Manager web port, and an Enterprise Manager account holding the Portal
    Administrator role.

    Exit codes: 0 success, 1 failure. Task Scheduler surfaces this as the task's
    last result, so it can be monitored.
#>
[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(ParameterSetName = 'Run', Mandatory)]
    [Parameter(ParameterSetName = 'TestConnection', Mandatory)]
    [Parameter(ParameterSetName = 'Install', Mandatory)]
    [string]$Server,

    [ValidateRange(1, 65535)]
    [int]$Port = 9443,

    [string]$CredentialPath = 'C:\ProgramData\VeeamLicenseReport\em.cred',

    [string]$OutputFolder = 'C:\VeeamLicenseReports',

    # No ValidateSet here. powershell.exe -File passes every argument as a
    # literal string, so a scheduled task's "-Formats Pdf,Json" arrives as a
    # single element and fails validation against the very set it looks like.
    # Normalised and checked just below instead.
    [string[]]$Formats = @('Pdf', 'Json'),

    [switch]$Monthly,

    [ValidateRange(0, 600)]
    [int]$RetentionMonths = 24,

    [switch]$SkipCertificateCheck,

    [string]$SmtpServer,
    [ValidateRange(1, 65535)]
    [int]$SmtpPort = 25,
    [switch]$SmtpUseSsl,
    [string]$MailFrom,
    [string[]]$MailTo,
    [string]$MailSubject = 'Veeam license usage report, {0}',

    [string]$TaskName = 'Veeam License Usage Report (Monthly)',

    [Parameter(ParameterSetName = 'Install')]
    [ValidatePattern('^([01]\d|2[0-3]):[0-5]\d$')]
    [string]$ScheduleTime = '22:00',

    [Parameter(ParameterSetName = 'Install')]
    [string]$RunAsUser = 'SYSTEM',

    [Parameter(ParameterSetName = 'Install')]
    [ValidateSet('Auto', 'Password', 'Interactive')]
    [string]$TaskLogonType = 'Auto',

    [Parameter(ParameterSetName = 'Run')]
    [switch]$NoEmail,

    [Parameter(ParameterSetName = 'SaveCredential', Mandatory)]
    [switch]$SaveCredential,

    [Parameter(ParameterSetName = 'SaveCredential')]
    [ValidateSet('Machine', 'User')]
    [string]$CredentialScope = 'Machine',

    [Parameter(ParameterSetName = 'SaveCredential')]
    [string[]]$GrantAccess,

    [Parameter(ParameterSetName = 'TestConnection', Mandatory)]
    [switch]$TestConnection,

    [Parameter(ParameterSetName = 'Install', Mandatory)]
    [switch]$Install,

    [Parameter(ParameterSetName = 'Uninstall', Mandatory)]
    [switch]$Uninstall
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# powershell.exe -File does not parse PowerShell syntax. A comma separated
# value arrives as one string, while an interactive run passes a real array.
# Both are normalised to the same shape here, before anything reads them.
function Expand-CommaList {
    param([string[]]$Value)
    if (-not $Value) { return @() }
    return @($Value |
        ForEach-Object { $_ -split ',' } |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ })
}

$Formats = @(Expand-CommaList $Formats | ForEach-Object {
    switch ($_.ToLowerInvariant()) {
        'pdf'   { 'Pdf' }
        'json'  { 'Json' }
        default { throw "-Formats accepts Pdf and Json, separated by a comma. Got '$_'." }
    }
})
if ($Formats.Count -eq 0) { throw '-Formats needs at least one of Pdf or Json.' }
$MailTo = Expand-CommaList $MailTo

$script:LogFile = $null
$script:WebArgs = @{ UseBasicParsing = $true }
$script:BaseUri = 'https://{0}:{1}' -f $Server, $Port

# The query parameters the Enterprise Manager UI sends with the download.
$script:ReportEndpoint = '/api/ZipDownload/DownloadReport'
$script:ReportQuery    = [ordered]@{ isAdjusted = '1' }

# ============================================================================
# Logging
# ============================================================================

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'OK')][string]$Level = 'INFO'
    )
    $line = '{0}  [{1,-5}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'OK'    { Write-Host $line -ForegroundColor Green }
        default { Write-Host $line }
    }
    if ($script:LogFile) {
        try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 } catch { }
    }
}

function Get-SafeProperty {
    # Safe optional-property read, needed because Set-StrictMode is on.
    param($Object, [Parameter(Mandatory)][string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object.PSObject.Properties.Name -contains $Name) {
        $value = $Object.$Name
        if ($null -eq $value) { return $Default }
        return $value
    }
    return $Default
}

# ============================================================================
# Credentials
# ============================================================================

function Protect-CredentialFile {
    # A machine scope blob can be decrypted by anything running on this
    # computer, so the file permissions are what actually limit who can read
    # it. Inheritance is switched off so a permissive ACL further up
    # ProgramData cannot widen it again.
    param(
        [Parameter(Mandatory)][string]$Path,
        [string[]]$GrantTo
    )
    try {
        $sids = New-Object System.Collections.Generic.List[string]
        # SYSTEM, so a task registered to run as SYSTEM can read it.
        [void]$sids.Add('S-1-5-18')
        # Administrators, for maintenance from an elevated session.
        [void]$sids.Add('S-1-5-32-544')
        # Whoever saved it. Without this an unelevated administrator locks
        # themselves out of their own file, because an unelevated token does
        # not carry the Administrators ACE.
        [void]$sids.Add([System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value)
        # The account the task will run as, when it is neither of the above.
        foreach ($name in @($GrantTo)) {
            if ($name) { [void]$sids.Add((Resolve-TaskAccount -Name $name).Sid) }
        }

        $acl = Get-Acl -LiteralPath $Path
        $acl.SetAccessRuleProtection($true, $false)
        foreach ($rule in @($acl.Access)) { [void]$acl.RemoveAccessRule($rule) }

        $granted = New-Object System.Collections.Generic.List[string]
        foreach ($sid in ($sids | Select-Object -Unique)) {
            $id = New-Object System.Security.Principal.SecurityIdentifier($sid)
            $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule(
                $id, 'FullControl', 'Allow')))
            $label = $sid
            try { $label = $id.Translate([System.Security.Principal.NTAccount]).Value } catch { }
            [void]$granted.Add($label)
        }
        Set-Acl -LiteralPath $Path -AclObject $acl
        Write-Log "Restricted the credential file to: $($granted -join ', ')."
    }
    catch {
        Write-Log "Could not restrict permissions on $Path ($($_.Exception.Message)). Review them by hand before relying on this." 'WARN'
    }
}

function Import-StoredCredential {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "No stored credential at $Path. Run this script with -SaveCredential on this computer."
    }

    $raw = Get-Content -LiteralPath $Path -Raw
    if ($raw -match '^\s*\{') {
        # Machine scope payload. Readable by any account on this computer,
        # which is exactly what lets the task run with nobody signed in.
        $data = $raw | ConvertFrom-Json
        Add-Type -AssemblyName System.Security
        $plain = $null
        try {
            $bytes = [System.Security.Cryptography.ProtectedData]::Unprotect(
                [Convert]::FromBase64String($data.Password), $null,
                [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
            $plain = [System.Text.Encoding]::UTF8.GetString($bytes)
            [Array]::Clear($bytes, 0, $bytes.Length)
        }
        catch {
            throw "Could not decrypt $Path. A machine scope credential is bound to the computer that created it ($($data.Machine)), so it does not travel. Run -SaveCredential here. Underlying error: $($_.Exception.Message)"
        }
        return New-Object System.Management.Automation.PSCredential(
            $data.UserName, (ConvertTo-SecureString -String $plain -AsPlainText -Force))
    }

    try {
        return Import-Clixml -LiteralPath $Path
    }
    catch {
        throw "Could not decrypt $Path. This is a user scope file, and DPAPI ties those to one Windows account on one machine, so only the account that saved it can read it. If the task runs as SYSTEM or as a service account, run -SaveCredential again to rewrite it in machine scope, which is the default. Underlying error: $($_.Exception.Message)"
    }
}

function Save-StoredCredential {
    param(
        [Parameter(Mandatory)][string]$Path,
        [ValidateSet('Machine', 'User')][string]$Scope = 'Machine',
        [string[]]$GrantTo
    )
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $cred = Get-Credential -Message 'Enterprise Manager account (needs the Portal Administrator role)'
    if (-not $cred) { throw 'No credential was entered.' }

    if ($Scope -eq 'User') {
        $cred | Export-Clixml -LiteralPath $Path -Force
        Write-Log "Stored the credential at $Path, encrypted for $env:USERDOMAIN\$env:USERNAME on $env:COMPUTERNAME." 'OK'
        Write-Log 'User scope means no other account can read it. A task running as SYSTEM or as a service account will fail unless that account saves its own copy, which means signing in as it.' 'WARN'
        return
    }

    # Machine scope. The point of it is that the task's account never has to
    # sign in to save a copy of its own, which is what lets the report run
    # with nobody logged on.
    Add-Type -AssemblyName System.Security
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($cred.GetNetworkCredential().Password)
    try {
        $blob = [System.Security.Cryptography.ProtectedData]::Protect(
            $bytes, $null, [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
    }
    finally { [Array]::Clear($bytes, 0, $bytes.Length) }

    [ordered]@{
        Format   = 'VeeamLicenseReport.Credential.v1'
        Scope    = 'LocalMachine'
        UserName = $cred.UserName
        Password = [Convert]::ToBase64String($blob)
        Machine  = $env:COMPUTERNAME
        SavedBy  = "$env:USERDOMAIN\$env:USERNAME"
        SavedUtc = (Get-Date).ToUniversalTime().ToString('o')
    } | ConvertTo-Json | Set-Content -LiteralPath $Path -Encoding UTF8 -Force

    Protect-CredentialFile -Path $Path -GrantTo $GrantTo
    Write-Log "Stored the credential at $Path, encrypted for this computer rather than for one account." 'OK'
    Write-Log 'Machine scope is the trade that buys unattended runs: any administrator on this computer can decrypt it, which is why the file is restricted to SYSTEM and Administrators. It is still useless on any other computer.' 'WARN'
}

# ============================================================================
# TLS and certificates
# ============================================================================

function Initialize-WebSettings {
    param([bool]$AllowUntrustedCertificate)

    try {
        [System.Net.ServicePointManager]::SecurityProtocol =
            [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
    } catch { }

    if (-not $AllowUntrustedCertificate) { return }

    if ($PSVersionTable.PSVersion.Major -ge 6) {
        $script:WebArgs['SkipCertificateCheck'] = $true
    }
    else {
        if (-not ('VeeamTrustAllCertsPolicy' -as [type])) {
            Add-Type -TypeDefinition @'
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class VeeamTrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(ServicePoint sp, X509Certificate cert, WebRequest req, int problem) {
        return true;
    }
}
'@
        }
        [System.Net.ServicePointManager]::CertificatePolicy = New-Object VeeamTrustAllCertsPolicy
    }
    Write-Log 'Certificate validation is disabled for this run. Install a trusted certificate on Enterprise Manager and drop -SkipCertificateCheck when you can.' 'WARN'
}

function Get-TimezoneOffsetMinutes {
    # Matches the browser's Date.getTimezoneOffset(), which counts minutes WEST
    # of UTC: UTC-8 reports 480. Enterprise Manager uses it to localise the
    # timestamps printed in the report.
    return [int](-([System.TimeZoneInfo]::Local.GetUtcOffset([datetime]::Now).TotalMinutes))
}

function Get-EmErrorMessage {
    # Enterprise Manager returns errors as a styled HTML page. Strip the markup
    # so the real message reaches the log instead of a bare status code.
    param($ErrorRecord)

    $status = ''
    try { $status = 'HTTP ' + [int]$ErrorRecord.Exception.Response.StatusCode } catch { }

    $text = ''
    try {
        $reader = New-Object System.IO.StreamReader($ErrorRecord.Exception.Response.GetResponseStream())
        $raw = $reader.ReadToEnd()
        $reader.Dispose()
        $text = $raw -replace '(?s)<script.*?</script>', '' `
                     -replace '(?s)<style.*?</style>', '' `
                     -replace '<[^>]+>', ' ' `
                     -replace '\s+', ' '
        $text = $text.Trim()
        $text = $text -replace '^Error\s+Veeam Backup Enterprise Manager\s+', ''
        $text = $text -replace '\s*Here is what you can do:.*$', ''
        if ($text.Length -gt 300) { $text = $text.Substring(0, 300) + '...' }
    } catch { }

    if (-not $text) { $text = $ErrorRecord.Exception.Message }
    return (($status, $text) | Where-Object { $_ }) -join ': '
}

# ============================================================================
# Sign-in
#
# Enterprise Manager's report endpoints are session-cookie authenticated: they
# redirect to the login page even when given a bearer token. So the script
# performs the same form sign-in the browser performs.
# ============================================================================

function Connect-VeeamEmWebSession {
    param([Parameter(Mandatory)][pscredential]$Credential)

    $session = $null
    $loginPageUri = "$script:BaseUri/index"

    Write-Log "Fetching the Enterprise Manager login page from $loginPageUri"
    try {
        $page = Invoke-WebRequest -Uri $loginPageUri -Method Get -SessionVariable session `
            -MaximumRedirection 10 @script:WebArgs
    }
    catch {
        throw "Could not load the Enterprise Manager login page at $loginPageUri. Check the server name and port, that the Veeam Backup Enterprise Manager web service is running, and that the port is open. Underlying error: $($_.Exception.Message)"
    }

    $html = [string]$page.Content

    # The credential form's action carries the OAuth client_id and redirect_uri.
    # Read it from the page rather than hardcoding it, so a change to the client
    # id in a future build does not break this script.
    $formMatch = [regex]::Match($html, '<form[^>]*id="by-credentials-form"[^>]*?action="([^"]+)"', 'IgnoreCase')
    if (-not $formMatch.Success) {
        $formMatch = [regex]::Match($html, '<form[^>]*?action="([^"]*oauth/login[^"]*)"', 'IgnoreCase')
    }
    if (-not $formMatch.Success) {
        throw 'Could not find the credential sign-in form on the Enterprise Manager login page. The sign-in page has probably changed in this Veeam build.'
    }
    $action = [System.Net.WebUtility]::HtmlDecode($formMatch.Groups[1].Value)
    $postUri = if ($action -match '^https?://') { $action } else { $script:BaseUri + $action }

    # ASP.NET Core anti-forgery: a hidden field that must match a cookie the
    # session already holds, so both have to travel together.
    $tokenMatch = [regex]::Match($html, 'name="__RequestVerificationToken"[^>]*?value="([^"]+)"', 'IgnoreCase')
    if (-not $tokenMatch.Success) {
        throw 'Could not find the __RequestVerificationToken field on the login page.'
    }

    $body = @{
        'Input.Username'             = $Credential.UserName
        'Input.Password'             = $Credential.GetNetworkCredential().Password
        'Input.RemainSignedIn'       = 'false'
        '__RequestVerificationToken' = $tokenMatch.Groups[1].Value
    }

    Write-Log "Signing in as $($Credential.UserName)."
    try {
        $response = Invoke-WebRequest -Uri $postUri -Method Post -Body $body -WebSession $session `
            -ContentType 'application/x-www-form-urlencoded' -MaximumRedirection 10 @script:WebArgs
    }
    catch {
        throw "The sign-in request failed: $($_.Exception.Message)"
    }

    # A rejected sign-in returns the login page again with HTTP 200 rather than
    # an error status, so detect it by looking for the password field.
    $result = [string]$response.Content
    if ($result -match 'name="Input\.Password"') {
        $reason = ''
        $summary = [regex]::Match($result, 'validation-summary[^>]*>\s*(?:<[^>]+>\s*)*([^<]{3,200})', 'IgnoreCase')
        if ($summary.Success) { $reason = ' Message from Enterprise Manager: ' + $summary.Groups[1].Value.Trim() }
        throw "Sign-in was rejected for $($Credential.UserName).$reason Check the password, and that the account holds the Portal Administrator role in Enterprise Manager."
    }

    Write-Log 'Signed in; session cookie acquired.' 'OK'
    return $session
}

function Get-EmCsrfToken {
    <#
        The app shell declares its CSRF request token as a JavaScript variable:

            var CSRFToken = 'CfDJ8...';

        Every API call must send it back as an x-csrf-token header, paired with
        the anti-forgery cookie the session holds. Calls are rejected without it.
    #>
    param([Parameter(Mandatory)]$Session)

    $page = Invoke-WebRequest -Uri "$script:BaseUri/index" -Method Get -WebSession $Session `
        -MaximumRedirection 10 @script:WebArgs

    $match = [regex]::Match([string]$page.Content, 'var\s+CSRFToken\s*=\s*[''"]([^''"]+)[''"]')
    if (-not $match.Success) {
        $match = [regex]::Match([string]$page.Content, 'csrf[A-Za-z]*\s*[:=]\s*[''"]([^''"]{20,})[''"]', 'IgnoreCase')
    }
    if (-not $match.Success) {
        throw 'Could not find the CSRF token on the Enterprise Manager application page. The page layout has probably changed in this Veeam build.'
    }

    Write-Log 'CSRF token acquired.' 'OK'
    return $match.Groups[1].Value
}

function Get-EmApiHeader {
    param([Parameter(Mandatory)][string]$CsrfToken, [string]$Referer)
    if (-not $Referer) { $Referer = "$script:BaseUri/index" }
    return @{
        'Accept'           = '*/*'
        'Origin'           = $script:BaseUri
        'Referer'          = $Referer
        'X-Requested-With' = 'XMLHttpRequest'
        'x-csrf-token'     = $CsrfToken
        'timezone-offset'  = [string](Get-TimezoneOffsetMinutes)
    }
}

# ============================================================================
# Licensing summary
# ============================================================================

function Get-EmLicensingInfo {
    # The endpoint behind the Licensing > Summary tab. Returns edition,
    # validity, per-platform instance usage and the licensed/used totals.
    param([Parameter(Mandatory)]$Session, [Parameter(Mandatory)]$Headers)

    $dc = [int64]([datetimeoffset]::UtcNow.ToUnixTimeMilliseconds())
    $uri = "$script:BaseUri/api/Licensing/GetInfo?_dc=$dc"

    Write-Log 'Requesting the licensing summary.'
    try {
        $response = Invoke-WebRequest -Uri $uri -Method Get -WebSession $Session `
            -Headers $Headers -MaximumRedirection 10 @script:WebArgs
    }
    catch {
        throw "The licensing summary request failed. $(Get-EmErrorMessage -ErrorRecord $_)"
    }

    $contentType = ''
    try { $contentType = [string]$response.Headers['Content-Type'] } catch { }
    if ($contentType -notlike '*json*') {
        throw "The licensing summary returned $contentType rather than JSON, which usually means the session was not accepted."
    }

    Write-Log ("Licensing summary received ({0:N0} bytes)." -f $response.RawContentLength) 'OK'
    return ([string]$response.Content | ConvertFrom-Json)
}

function Write-HistoryCsv {
    # Rewrites the history file so that running again inside the same month
    # replaces that month's rows rather than appending a second set. Rows for
    # every other month are carried through untouched.
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][string]$HistoryPath,
        [Parameter(Mandatory)][string]$MonthStamp
    )

    $new = @($Rows)
    if (-not (Test-Path -LiteralPath $HistoryPath)) {
        $new | Export-Csv -LiteralPath $HistoryPath -NoTypeInformation -Encoding UTF8
        Write-Log "Created $HistoryPath with $($new.Count) row(s)." 'OK'
        return
    }

    $existing = @()
    try { $existing = @(Import-Csv -LiteralPath $HistoryPath) }
    catch {
        Write-Log "Could not read $HistoryPath ($($_.Exception.Message)). Keeping the old file as $HistoryPath.bad and starting a new one." 'WARN'
        Move-Item -LiteralPath $HistoryPath -Destination "$HistoryPath.bad" -Force
        $new | Export-Csv -LiteralPath $HistoryPath -NoTypeInformation -Encoding UTF8
        return
    }

    # A row with no Month column is kept rather than discarded, so a file this
    # script did not write is never silently thrown away.
    $kept = @($existing | Where-Object {
        (-not $_.PSObject.Properties['Month']) -or ($_.Month -ne $MonthStamp)
    })
    $replaced = @($existing).Count - $kept.Count
    $all = @($kept) + $new

    # Export-Csv takes its column set from the first object alone, and the
    # licensing schema can change between Veeam versions, so every row is
    # projected onto the union of the columns actually present.
    $columns = New-Object System.Collections.Generic.List[string]
    foreach ($row in $all) {
        foreach ($name in $row.PSObject.Properties.Name) {
            if (-not $columns.Contains($name)) { [void]$columns.Add($name) }
        }
    }
    $all = @($all | Select-Object -Property $columns.ToArray())

    # Written alongside and swapped in, so an interrupted run cannot truncate
    # history that has been accumulating for months.
    $tmp = "$HistoryPath.tmp"
    $all | Export-Csv -LiteralPath $tmp -NoTypeInformation -Encoding UTF8
    Move-Item -LiteralPath $tmp -Destination $HistoryPath -Force

    if ($replaced -gt 0) {
        Write-Log "Wrote $($new.Count) row(s) for $MonthStamp to $HistoryPath, replacing $replaced earlier row(s) for the same month." 'OK'
    }
    else {
        Write-Log "Added $($new.Count) row(s) for $MonthStamp to $HistoryPath." 'OK'
    }
}

function Add-HistoryRow {
    # Flattens the licensing summary into CSV rows for month-over-month
    # trending. Field names were read from a live response, but every lookup
    # tolerates absence so a schema change degrades to blanks rather than
    # failing the run.
    param(
        [Parameter(Mandatory)]$Info,
        [Parameter(Mandatory)][string]$HistoryPath,
        [Parameter(Mandatory)][string]$MonthStamp
    )

    $generated = (Get-Date).ToUniversalTime().ToString('s')
    $status  = Get-SafeProperty $Info 'licenseStatus'
    $details = Get-SafeProperty $Info 'licenseInfo'

    $edition  = Get-SafeProperty $status 'licEdition'
    $daysLeft = Get-SafeProperty $status 'daysLeft'

    $named = @{}
    foreach ($list in 'licenseInformationItems', 'instancesItems', 'socketsItems', 'capacityItems') {
        foreach ($item in @(Get-SafeProperty $details $list @())) {
            $name = [string](Get-SafeProperty $item 'name')
            if ($name) { $named[$name] = [string](Get-SafeProperty $item 'value') }
        }
    }
    $lookup = {
        param([string[]]$Keys)
        foreach ($key in $Keys) {
            foreach ($actual in $named.Keys) {
                if ($actual -like "*$key*") { return $named[$actual] }
            }
        }
        return $null
    }

    # "Instances" reads like "500 (33 used)". Matched by name rather than by
    # scanning values, because "Expiration date" also begins with digits.
    $licensed = $null
    $used = $null
    $instancesValue = [string](& $lookup @('Instances'))
    if ($instancesValue) {
        $m = [regex]::Match($instancesValue, '([\d,]+)\s*\((\d[\d,]*)\s*used\)')
        if ($m.Success) {
            $licensed = $m.Groups[1].Value -replace ',', ''
            $used     = $m.Groups[2].Value -replace ',', ''
        }
        else {
            $m = [regex]::Match($instancesValue, '([\d,]+)')
            if ($m.Success) { $licensed = $m.Groups[1].Value -replace ',', '' }
        }
    }

    # "Expiration date" reads like "3/10/2028 (483 days left)"; keep the date.
    $expiration = [string](& $lookup @('Expiration', 'Expiry'))
    if ($expiration) { $expiration = ($expiration -replace '\s*\(.*$', '').Trim() }

    $common = @{
        Edition           = $edition
        LicenseType       = (& $lookup @('Type'))
        Package           = (& $lookup @('Package'))
        LicensedInstances = $licensed
        UsedInstances     = $used
        DaysLeft          = $daysLeft
        LicensedTo        = (& $lookup @('Licensed to', 'Company'))
        SupportId         = (& $lookup @('Support'))
        ExpirationDate    = $expiration
    }

    $rows = @()
    $usage = Get-SafeProperty $status 'instancesUsage'
    foreach ($platform in @(Get-SafeProperty $usage 'platformUsages' @())) {
        $count = Get-SafeProperty $platform 'count' 0
        $instances = Get-SafeProperty $platform 'instances' 0
        if ($count -eq 0 -and $instances -eq 0) { continue }   # skip unused platforms
        $row = [ordered]@{
            Month        = $MonthStamp
            GeneratedUtc = $generated
            Scope        = 'Platform'
            Name         = (Get-SafeProperty $platform 'platform')
            Count        = $count
            Multiplier   = (Get-SafeProperty $platform 'multiplier')
            Instances    = $instances
        }
        foreach ($k in $common.Keys) { $row[$k] = $common[$k] }
        $rows += [pscustomobject]$row
    }

    $totalRow = [ordered]@{
        Month        = $MonthStamp
        GeneratedUtc = $generated
        Scope        = 'Total'
        Name         = 'ALL'
        Count        = $null
        Multiplier   = $null
        Instances    = $used
    }
    foreach ($k in $common.Keys) { $totalRow[$k] = $common[$k] }
    $rows += [pscustomobject]$totalRow

    Write-HistoryCsv -Rows $rows -HistoryPath $HistoryPath -MonthStamp $MonthStamp
    Write-Log "Licence: $edition, $used of $licensed instances used, $daysLeft days remaining." 'OK'
}

# ============================================================================
# The aggregated report
# ============================================================================

function New-EmReportPreview {
    <#
        This is the step the Report link in the UI performs, and the one without
        which the download answers "Monthly report not found".

            POST /api/Licensing/CreateMonthlyReportPreview
            Content-Type: application/json
            Body: false   (current licence consumption)
                  true    (the archived report for the month just ended)

        The body is a bare JSON boolean, not an object. The response carries the
        rendered report as HTML, which is not needed here: what matters is the
        server-side state it leaves for the download to collect.
    #>
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)]$Headers,
        [bool]$ArchivedMonth = $false
    )

    $body = if ($ArchivedMonth) { 'true' } else { 'false' }
    $label = if ($ArchivedMonth) { 'the archived monthly report' } else { 'current licence consumption' }

    Write-Log "Generating $label."
    try {
        $response = Invoke-WebRequest -Uri "$script:BaseUri/api/Licensing/CreateMonthlyReportPreview" `
            -Method Post -WebSession $Session -Headers $Headers -Body $body `
            -ContentType 'application/json' -MaximumRedirection 10 @script:WebArgs
    }
    catch {
        throw "Could not generate the report. $(Get-EmErrorMessage -ErrorRecord $_)"
    }

    $contentType = ''
    try { $contentType = [string]$response.Headers['Content-Type'] } catch { }
    if ($contentType -notlike '*json*') {
        throw "The report generation call returned $contentType rather than JSON, which usually means the session was not accepted."
    }

    # The heading inside the returned HTML states the reporting interval, which
    # is a useful confirmation of what was actually generated.
    $heading = ''
    try {
        $preview = [string]$response.Content | ConvertFrom-Json
        $html = [string](Get-SafeProperty $preview 'html')
        $m = [regex]::Match($html, '<b>([^<]{3,60})</b>')
        if ($m.Success) { $heading = $m.Groups[1].Value.Trim() }
    } catch { }

    Write-Log ("Report generated{0}." -f $(if ($heading) { ": $heading" } else { '' })) 'OK'
}

function New-ReportUri {
    param([Parameter(Mandatory)][string]$Format)

    $query = [ordered]@{}
    foreach ($key in $script:ReportQuery.Keys) { $query[$key] = $script:ReportQuery[$key] }
    $query['format'] = $Format.ToLower()
    $query['timezoneOffset'] = [string](Get-TimezoneOffsetMinutes)

    $pairs = foreach ($key in $query.Keys) {
        '{0}={1}' -f [uri]::EscapeDataString($key), [uri]::EscapeDataString($query[$key])
    }
    return '{0}{1}?{2}' -f $script:BaseUri, $script:ReportEndpoint, ($pairs -join '&')
}

function Save-EmLicenseReport {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)]$Headers,
        [Parameter(Mandatory)][string]$Format,
        [Parameter(Mandatory)][string]$OutFile
    )

    Write-Log "Downloading the $Format report."
    try {
        $response = Invoke-WebRequest -Uri $Uri -Method Get -WebSession $Session `
            -Headers $Headers -MaximumRedirection 10 @script:WebArgs
    }
    catch {
        $detail = Get-EmErrorMessage -ErrorRecord $_
        if ($detail -match 'report not found') {
            throw "Enterprise Manager answered: $detail. The report was generated but could not be collected, which usually means the generate and download steps did not share a session."
        }
        throw "The $Format download failed. $detail"
    }

    $contentType = ''
    try { $contentType = [string]$response.Headers['Content-Type'] } catch { }

    # A redirect back to the login page returns HTML with HTTP 200, so the
    # content type is the only reliable signal that the session was accepted.
    if ($contentType -like 'text/html*') {
        throw "The $Format download returned HTML rather than a report, which means the session was not accepted."
    }

    $content = $response.Content
    if ($content -is [byte[]]) {
        if ($content.Length -eq 0) { throw "The $Format download returned an empty response." }
        [System.IO.File]::WriteAllBytes($OutFile, $content)
    }
    else {
        $text = [string]$content
        if ([string]::IsNullOrWhiteSpace($text)) { throw "The $Format download returned an empty response." }
        Set-Content -LiteralPath $OutFile -Value $text -Encoding UTF8
    }

    $item = Get-Item -LiteralPath $OutFile
    Write-Log ("Saved {0} ({1:N0} KB)" -f $item.Name, ($item.Length / 1KB)) 'OK'
    return $item.FullName
}

# ============================================================================
# Email and retention
# ============================================================================

function Send-ReportEmail {
    param([Parameter(Mandatory)][string[]]$Attachments, [Parameter(Mandatory)][string]$MonthStamp)

    $body = @(
        "The Veeam license usage report for $MonthStamp is attached.",
        '',
        "Generated on $env:COMPUTERNAME."
    ) -join [Environment]::NewLine

    $params = @{
        SmtpServer  = $SmtpServer
        Port        = $SmtpPort
        From        = $MailFrom
        To          = $MailTo
        Subject     = ($MailSubject -f $MonthStamp)
        Body        = $body
        Attachments = $Attachments
    }
    if ($SmtpUseSsl) { $params['UseSsl'] = $true }

    # Send-MailMessage is obsolete but remains the lowest-dependency option on
    # Windows PowerShell 5.1. Replace with a Microsoft Graph call if basic SMTP
    # submission has been retired in your environment.
    Send-MailMessage @params -WarningAction SilentlyContinue
    Write-Log "Emailed the report to $($MailTo -join ', ')" 'OK'
}

function Remove-ExpiredReport {
    param([Parameter(Mandatory)][string]$Root, [int]$KeepMonths)
    if ($KeepMonths -le 0) { return }
    $cutoff = (Get-Date).AddMonths(-$KeepMonths)
    foreach ($dir in Get-ChildItem -LiteralPath $Root -Directory -ErrorAction SilentlyContinue) {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParseExact($dir.Name, 'yyyy-MM', [cultureinfo]::InvariantCulture, 'None', [ref]$parsed)) {
            if ($parsed -lt $cutoff) {
                Remove-Item -LiteralPath $dir.FullName -Recurse -Force
                Write-Log "Pruned $($dir.Name), past the $KeepMonths month retention window."
            }
        }
    }
}

# ============================================================================
# Scheduled task
# ============================================================================

function Resolve-TaskAccount {
    # The Windows account the scheduled task runs as. This is NOT the Enterprise
    # Manager login, which lives in the credential file.
    #
    # Task Scheduler resolves <UserId> through LookupAccountName, and that fails
    # on a user principal name unless this computer can reach that account's
    # domain. The name is therefore translated to a SID here, and the SID is
    # what goes into the XML, because a SID always maps.
    param([Parameter(Mandatory)][string]$Name)

    $candidates = New-Object System.Collections.Generic.List[string]
    $candidates.Add($Name)
    if ($Name.Contains('@')) {
        $at     = $Name.IndexOf('@')
        $local  = $Name.Substring(0, $at)
        $domain = $Name.Substring($at + 1)
        $candidates.Add(('{0}\{1}' -f $domain, $local))
        if ($domain.Contains('.')) {
            $candidates.Add(('{0}\{1}' -f $domain.Split('.')[0].ToUpperInvariant(), $local))
        }
    }
    # Deliberately no fall back to this computer's own domain. A local or Entra
    # account that happens to share the user name is a different principal, and
    # silently registering the task under it would look like success while
    # running as somebody else.

    foreach ($candidate in $candidates) {
        try {
            $sid = (New-Object System.Security.Principal.NTAccount($candidate)).Translate(
                [System.Security.Principal.SecurityIdentifier])
            $resolved = $candidate
            try { $resolved = $sid.Translate([System.Security.Principal.NTAccount]).Value } catch { }
            return [pscustomobject]@{ Sid = $sid.Value; Name = $resolved; Supplied = $Name }
        }
        catch { }
    }

    $computerDomain = 'WORKGROUP'
    try { $computerDomain = (Get-CimInstance Win32_ComputerSystem).Domain } catch { }
    throw (@(
        "Windows cannot resolve '$Name' to an account, so the task cannot be registered under it.",
        "  Tried: $($candidates -join ', ')",
        "  This computer is $env:COMPUTERNAME in $computerDomain, and it can only resolve accounts from its own domain or its local account database.",
        '',
        '-RunAsUser is the Windows account the scheduled task logs on as. It is not the Enterprise Manager login: that one is held in the credential file and is supplied over HTTPS at run time.',
        '',
        'Either install the task on a server joined to the account domain, which is the usual place for it, or run it here under an account this computer knows, for example:',
        "  -RunAsUser '$env:USERDOMAIN\$env:USERNAME' -TaskLogonType Interactive"
    ) -join [Environment]::NewLine)
}

function Protect-ScriptAvailability {
    # The script can legitimately live in a OneDrive or other cloud synced
    # folder. Files On-Demand may evict an unpinned file's contents to reclaim
    # space, and when the task fires with nobody signed in there is no sync
    # client running to fetch it back, so the run fails with the script simply
    # unreadable. Marking the file pinned, which is what "Always keep on this
    # device" does in Explorer, prevents the eviction. On a path that is not a
    # cloud placeholder this does nothing.
    param([Parameter(Mandatory)][string]$Path)

    $REPARSE_POINT = 0x400
    $PINNED        = 0x80000

    try {
        $attr = [int](Get-Item -LiteralPath $Path -Force).Attributes
        if (-not ($attr -band $REPARSE_POINT)) { return }

        if ($attr -band $PINNED) {
            Write-Log 'The script sits in a cloud synced folder and is already pinned, so it stays on this disk.'
            return
        }

        # +P sets the pinned attribute, -U clears the unpinned one. They are
        # separate flags and a file can carry neither, which is the evictable
        # state this is here to fix.
        & attrib.exe +P -U $Path 2>&1 | Out-Null

        if ([int](Get-Item -LiteralPath $Path -Force).Attributes -band $PINNED) {
            Write-Log 'The script sits in a cloud synced folder, so it has been pinned to this disk. Without that, Files On-Demand can evict it and the task then fails when it runs with nobody signed in.' 'OK'
        }
        else {
            Write-Log "The script sits in a cloud synced folder and could not be pinned. Set 'Always keep on this device' on it in Explorer, or keep the copy the task runs outside the synced folder, otherwise the task will fail once the file is evicted." 'WARN'
        }
    }
    catch {
        Write-Log "Could not confirm that $Path stays available offline ($($_.Exception.Message))." 'WARN'
    }
}

function Install-ReportTask {
    # Rebuilds the current invocation as a command line and writes it into the
    # task, so the task runs with exactly the settings used here.
    Protect-ScriptAvailability -Path $PSCommandPath
    $quoted = {
        param([string]$Value)
        '"' + ($Value -replace '"', '\"') + '"'
    }

    $argList = New-Object System.Collections.Generic.List[string]
    $argList.Add('-NoProfile')
    $argList.Add('-NonInteractive')
    $argList.Add('-ExecutionPolicy')
    $argList.Add('Bypass')
    $argList.Add('-File')
    $argList.Add((& $quoted $PSCommandPath))
    $argList.Add('-Server');          $argList.Add((& $quoted $Server))
    $argList.Add('-Port');            $argList.Add([string]$Port)
    $argList.Add('-CredentialPath');  $argList.Add((& $quoted $CredentialPath))
    $argList.Add('-OutputFolder');    $argList.Add((& $quoted $OutputFolder))
    $argList.Add('-Formats');         $argList.Add(($Formats -join ','))
    $argList.Add('-RetentionMonths'); $argList.Add([string]$RetentionMonths)
    $argList.Add('-TaskName');        $argList.Add((& $quoted $TaskName))
    if ($Monthly)              { $argList.Add('-Monthly') }
    if ($SkipCertificateCheck) { $argList.Add('-SkipCertificateCheck') }
    if ($SmtpServer) {
        $argList.Add('-SmtpServer'); $argList.Add((& $quoted $SmtpServer))
        $argList.Add('-SmtpPort');   $argList.Add([string]$SmtpPort)
        if ($SmtpUseSsl) { $argList.Add('-SmtpUseSsl') }
        if ($MailFrom)   { $argList.Add('-MailFrom'); $argList.Add((& $quoted $MailFrom)) }
        if ($MailTo)     { $argList.Add('-MailTo');   $argList.Add(($MailTo -join ',')) }
        if ($MailSubject) { $argList.Add('-MailSubject'); $argList.Add((& $quoted $MailSubject)) }
    }
    $arguments = $argList -join ' '

    # "Last day of the month" has to be expressed as <Day>Last</Day> in the task
    # XML. The MSFT_TaskMonthlyTrigger CIM class exposes a RunOnLastDayOfMonth
    # property, but the Task Scheduler service rejects a trigger built that way
    # with "The parameter is incorrect", so the XML is written directly.
    $isSystem = $RunAsUser -in @('SYSTEM', 'NT AUTHORITY\SYSTEM', 'S-1-5-18')
    $taskPassword = $null
    $principalUser = 'SYSTEM'
    $principalSid = 'S-1-5-18'
    $logonType = 'ServiceAccount'

    if ($isSystem) {
        Write-Log 'The task will run as SYSTEM, so it needs no stored Windows password and runs with nobody signed in. Save the credential with -SaveCredential in its default machine scope: a user scope file saved by a person cannot be decrypted by SYSTEM.' 'WARN'
    }
    else {
        # Resolved before the password prompt, so an account this computer
        # cannot use fails immediately rather than after typing a password.
        $account = Resolve-TaskAccount -Name $RunAsUser
        $principalUser = $account.Name
        $principalSid = $account.Sid
        if ($account.Name -ne $RunAsUser) {
            Write-Log "Resolved '$RunAsUser' to $($account.Name)."
        }

        $currentSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        if ($account.Sid -ne $currentSid) {
            Write-Log "The task will run as $($account.Name), which is not the account running this install. The credential file is DPAPI encrypted for one account on one computer, so run -SaveCredential while signed in as $($account.Name) on $env:COMPUTERNAME, or the task will not be able to decrypt it." 'WARN'
        }

        $effectiveLogon = $TaskLogonType
        if ($effectiveLogon -eq 'Auto') { $effectiveLogon = 'Password' }

        if ($effectiveLogon -eq 'Interactive') {
            # Runs only while that account is signed in, and stores no password.
            # The only workable form for an account whose password Task
            # Scheduler cannot cache, such as a Microsoft Entra account.
            $logonType = 'InteractiveToken'
            Write-Log "The task will run as $($account.Name) with an interactive token, so it only fires while that account is signed in to $env:COMPUTERNAME." 'WARN'
        }
        else {
            $taskCred = Get-Credential -UserName $account.Name -Message "Windows password for $($account.Name). This is the Windows account the task logs on as, not the Enterprise Manager login."
            if (-not $taskCred) { throw 'A password is required to register the task under a named account. Use -TaskLogonType Interactive to register one that needs no stored password.' }
            $taskPassword = $taskCred.GetNetworkCredential().Password
            $logonType = 'Password'
        }
    }

    $xmlEscape = {
        param([string]$Text)
        $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
    }

    # A SYSTEM task carries no LogonType element at all. The Task Scheduler
    # service rejects <LogonType>ServiceAccount</LogonType> outright, with
    # "The task XML contains a value which is incorrectly formatted or out of
    # range. (24,35):LogonType:ServiceAccount", in every combination of UserId
    # and RunLevel tried. Windows writes its own SYSTEM tasks with the element
    # absent, so this does the same.
    $logonElement = if ($isSystem) { '' } else { "`n      <LogonType>$logonType</LogonType>" }

    $taskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Downloads the aggregated License Usage Report from Veeam Backup Enterprise Manager on the last day of each month.</Description>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>$(Get-Date -Format 'yyyy-MM-dd')T$($ScheduleTime):00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByMonth>
        <DaysOfMonth>
          <Day>Last</Day>
        </DaysOfMonth>
        <Months>
          <January/><February/><March/><April/><May/><June/>
          <July/><August/><September/><October/><November/><December/>
        </Months>
      </ScheduleByMonth>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$principalSid</UserId>$logonElement
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell.exe</Command>
      <Arguments>$(& $xmlEscape $arguments)</Arguments>
      <WorkingDirectory>$(& $xmlEscape (Split-Path -Parent $PSCommandPath))</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@

    $registerArgs = @{ TaskName = $TaskName; Xml = $taskXml; Force = $true }
    if ($taskPassword) {
        $registerArgs['User'] = $principalUser
        $registerArgs['Password'] = $taskPassword
    }

    try {
        Register-ScheduledTask @registerArgs | Out-Null
        $next = $null
        try { $next = (Get-ScheduledTaskInfo -TaskName $TaskName).NextRunTime } catch { }
        Write-Log "Registered '$TaskName'. Next run: $(if ($next) { $next } else { 'see Task Scheduler' })" 'OK'
    }
    catch {
        $registerError = $_.Exception.Message
        Write-Log "Register-ScheduledTask failed ($registerError). Falling back to schtasks.exe." 'WARN'

        # schtasks.exe re-parses /TR, so the quotes inside it have to survive
        # CommandLineToArgvW as literal characters. Without the backslash
        # escape, a path such as "...\OneDrive - Veeam Software..." is split
        # apart and schtasks rejects the stray '-'.
        $trCommand = 'powershell.exe ' + $arguments
        if ($trCommand.Length -gt 261) {
            throw (@(
                "Register-ScheduledTask failed: $registerError",
                "schtasks.exe cannot stand in, because the command line is $($trCommand.Length) characters and schtasks caps /TR at 261.",
                'Shorten the paths, or register the task from the XML this script builds, rather than working around the registration error.'
            ) -join [Environment]::NewLine)
        }

        $schtasksArgs = @(
            '/Create', '/TN', $TaskName,
            '/TR', ($trCommand -replace '"', '\"'),
            '/SC', 'MONTHLY', '/MO', 'LASTDAY', '/M', '*',
            '/ST', $ScheduleTime, '/RL', 'HIGHEST', '/F'
        )
        if ($taskPassword) { $schtasksArgs += @('/RU', $principalUser, '/RP', $taskPassword) }
        elseif ($isSystem) { $schtasksArgs += @('/RU', 'SYSTEM') }
        else { $schtasksArgs += @('/RU', $principalUser, '/IT') }
        & schtasks.exe @schtasksArgs
        if ($LASTEXITCODE -ne 0) { throw "schtasks.exe exited with code $LASTEXITCODE." }
        Write-Log 'Registered the task through schtasks.exe.' 'OK'
    }
}

function Uninstall-ReportTask {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Log "Removed '$TaskName'." 'OK'
    }
    else {
        Write-Log "No scheduled task named '$TaskName' was found." 'WARN'
    }
}

# ============================================================================
# Main
# ============================================================================

function Connect-AndPrepare {
    # Sign in, take a CSRF token, and return everything later calls need.
    $session = Connect-VeeamEmWebSession -Credential (Import-StoredCredential -Path $CredentialPath)
    $csrf = Get-EmCsrfToken -Session $session
    $referer = '{0}/MonthlyReport?isMonthly={1}' -f $script:BaseUri, $Monthly.ToString().ToLower()
    return [pscustomobject]@{
        Session       = $session
        ApiHeaders    = (Get-EmApiHeader -CsrfToken $csrf)
        ReportHeaders = (Get-EmApiHeader -CsrfToken $csrf -Referer $referer)
    }
}

try {
    switch ($PSCmdlet.ParameterSetName) {

        'Uninstall' {
            Uninstall-ReportTask
            exit 0
        }

        'SaveCredential' {
            Save-StoredCredential -Path $CredentialPath -Scope $CredentialScope -GrantTo $GrantAccess
            exit 0
        }

        'Install' {
            Install-ReportTask
            Write-Log 'Run -TestConnection first if you have not already, and again after every Veeam upgrade.' 'WARN'
            exit 0
        }

        'TestConnection' {
            Initialize-WebSettings -AllowUntrustedCertificate $SkipCertificateCheck.IsPresent
            Write-Log "Target: $script:BaseUri"
            Write-Log "Timezone offset this host will send: $(Get-TimezoneOffsetMinutes) minutes west of UTC."

            $ctx = Connect-AndPrepare
            $info = Get-EmLicensingInfo -Session $ctx.Session -Headers $ctx.ApiHeaders
            $status = Get-SafeProperty $info 'licenseStatus'
            Write-Log ("Edition: {0}. {1} {2} days remaining." -f `
                (Get-SafeProperty $status 'licEdition'),
                (Get-SafeProperty $status 'validationStatusComment'),
                (Get-SafeProperty $status 'daysLeft')) 'OK'

            New-EmReportPreview -Session $ctx.Session -Headers $ctx.ReportHeaders -ArchivedMonth $Monthly.IsPresent
            $response = Invoke-WebRequest -Uri (New-ReportUri -Format 'Pdf') -Method Get `
                -WebSession $ctx.Session -Headers $ctx.ReportHeaders -MaximumRedirection 10 @script:WebArgs
            $contentType = ''
            try { $contentType = [string]$response.Headers['Content-Type'] } catch { }
            if ($contentType -like 'text/html*') {
                throw 'The report endpoint returned HTML rather than a document.'
            }
            Write-Log ("Aggregated report available: {0}, {1:N0} bytes." -f $contentType, $response.RawContentLength) 'OK'
            Write-Log 'Connection test passed.' 'OK'
            exit 0
        }
    }

    # ---------------- Normal run ----------------

    if (-not (Test-Path -LiteralPath $OutputFolder)) {
        try {
            New-Item -ItemType Directory -Path $OutputFolder -Force -ErrorAction Stop | Out-Null
        }
        catch {
            throw "Could not create the output folder '$OutputFolder'. Check that the drive exists and that $env:USERDOMAIN\$env:USERNAME may write there. Creating a folder directly under C:\ needs administrative rights. Underlying error: $($_.Exception.Message)"
        }
    }
    $logDir = Join-Path $OutputFolder '_logs'
    if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $script:LogFile = Join-Path $logDir ('run_{0}.log' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

    # The folder and file names carry the month the report COVERS, not the month
    # it was produced in. With -Monthly the report is for the month just ended,
    # so a run on 1 November 2026 files its October report under 2026-10.
    $monthStamp = if ($Monthly) {
        (Get-Date).AddMonths(-1).ToString('yyyy-MM')
    } else {
        Get-Date -Format 'yyyy-MM'
    }
    $monthFolder = Join-Path $OutputFolder $monthStamp
    if (-not (Test-Path -LiteralPath $monthFolder)) { New-Item -ItemType Directory -Path $monthFolder -Force | Out-Null }

    Write-Log "=== Veeam license usage report, $monthStamp ==="
    Write-Log "Host $env:COMPUTERNAME, account $env:USERDOMAIN\$env:USERNAME, PowerShell $($PSVersionTable.PSVersion)"
    Write-Log "Target: $script:BaseUri"

    $today = Get-Date
    $lastDay = [datetime]::DaysInMonth($today.Year, $today.Month)
    if (-not $Monthly -and $today.Day -ne $lastDay) {
        Write-Log "Today is day $($today.Day) of $lastDay, so this is a mid-month snapshot rather than a month-end one. Expected for an ad hoc run." 'WARN'
    }

    Initialize-WebSettings -AllowUntrustedCertificate $SkipCertificateCheck.IsPresent

    $saved = @()

    # 1. The licensing summary. Dependable, so it runs first and is fatal.
    $ctx = Connect-AndPrepare
    $info = Get-EmLicensingInfo -Session $ctx.Session -Headers $ctx.ApiHeaders
    $summaryFile = Join-Path $monthFolder ('LicenseSummary_{0}.json' -f $monthStamp)
    $info | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $summaryFile -Encoding UTF8
    Write-Log "Saved the licensing summary to $summaryFile" 'OK'
    $saved += $summaryFile

    Add-HistoryRow -Info $info -HistoryPath (Join-Path $OutputFolder 'LicenseUsage_History.csv') -MonthStamp $monthStamp

    # 2. The aggregated report: generate, then download, per format. A failed
    #    request invalidates the session, so each format gets a fresh sign-in
    #    and a fresh generation rather than sharing state. A failure here is
    #    reported but does not sink the run, so the trending data is never lost
    #    to a user interface change.
    $first = $true
    foreach ($format in $Formats) {
        try {
            if (-not $first) { $ctx = Connect-AndPrepare }
            $first = $false

            New-EmReportPreview -Session $ctx.Session -Headers $ctx.ReportHeaders -ArchivedMonth $Monthly.IsPresent

            $outFile = Join-Path $monthFolder ('VeeamLicenseUsageReport_{0}.{1}' -f $monthStamp, $format.ToLower())
            $saved += Save-EmLicenseReport -Uri (New-ReportUri -Format $format) -Session $ctx.Session `
                -Headers $ctx.ReportHeaders -Format $format -OutFile $outFile
        }
        catch {
            Write-Log $_.Exception.Message 'WARN'
            Write-Log "Continuing without the $format report; the licensing summary was still collected." 'WARN'
        }
    }

    if (-not $NoEmail -and $SmtpServer) {
        if (-not $MailFrom -or -not $MailTo) {
            Write-Log 'SmtpServer was supplied without MailFrom and MailTo, so no email was sent.' 'WARN'
        }
        else {
            Send-ReportEmail -Attachments $saved -MonthStamp $monthStamp
        }
    }

    Remove-ExpiredReport -Root $OutputFolder -KeepMonths $RetentionMonths

    Write-Log "=== Finished: $(@($saved).Count) file(s) in $monthFolder ===" 'OK'
    exit 0
}
catch {
    Write-Log $_.Exception.Message 'ERROR'
    if ($_.ScriptStackTrace) { Write-Log ($_.ScriptStackTrace | Out-String) 'ERROR' }
    exit 1
}
