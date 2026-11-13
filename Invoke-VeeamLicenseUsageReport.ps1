#Requires -Version 5.1

<#
.SYNOPSIS
    Downloads the aggregated Veeam License Usage Report from Veeam Backup
    Enterprise Manager and installs itself as a scheduled task that runs on the
    last day of every month.

.DESCRIPTION
    Enterprise Manager produces one aggregated License Usage Report covering
    every backup server it manages, with a single company-level summary. That
    is the report this script retrieves, in the same PDF and JSON formats the
    Download button in the Enterprise Manager UI offers.

    How it authenticates, and why:

    Enterprise Manager exposes no supported API for this report. The endpoint
    the UI uses, /api/ZipDownload/DownloadReport, answers 302 to /login even
    when a bearer token is supplied, so it is session-cookie authenticated
    rather than a token API. The script therefore performs the same form login
    the browser performs: it fetches the login page, carries over the
    __RequestVerificationToken anti-forgery pair, posts the credential, follows
    the OAuth redirect chain to pick up the session cookie, and then requests
    the report with that session.

    This is an undocumented interface. It works, but it can change without
    notice on upgrade. Run -TestConnection after any Veeam update. The README
    describes a fully supported alternative built on
    Generate-VBRLicenseUsageReport if you would rather not carry that risk.

.PARAMETER ConfigPath
    Path to the JSON configuration file. Defaults to
    VeeamLicenseReport.config.json beside this script.

.PARAMETER SaveCredential
    Prompts for the Enterprise Manager credential and stores it encrypted at
    EnterpriseManager.CredentialFile. Run this as the same Windows account the
    scheduled task will run under, on the same machine: DPAPI ties the file to
    one user on one computer.

.PARAMETER TestConnection
    Signs in, confirms the report endpoint responds with a real document, and
    reports what it found without saving anything. Use after installation and
    after every Veeam upgrade.

.PARAMETER Install
    Registers the scheduled task described by the Schedule section of the
    configuration file. The trigger fires on the last day of every month.

.PARAMETER Uninstall
    Removes the scheduled task.

.PARAMETER NoEmail
    Runs the report but skips email delivery regardless of configuration.

.EXAMPLE
    .\Invoke-VeeamLicenseUsageReport.ps1 -SaveCredential
    .\Invoke-VeeamLicenseUsageReport.ps1 -TestConnection
    .\Invoke-VeeamLicenseUsageReport.ps1
    .\Invoke-VeeamLicenseUsageReport.ps1 -Install

.NOTES
    The report reflects current license consumption at the moment it runs,
    which is why the task is scheduled for the last day of the month: that is
    what makes the output a month-end snapshot.

    Enterprise Manager separately generates its own official monthly usage
    report on the first day of each month for the month just ended. This script
    does not replace that; it gives you a month-end snapshot you control, keep
    and can trend.
#>

# The blank line between #Requires and the comment block above is load bearing.
# A comment line directly above <# stops Get-Help associating the help with the
# script: it then reports no parameters at all and falls back to the generated
# syntax, so every .PARAMETER becomes invisible to anyone who has not opened the
# file. #Requires is a comment like any other, which is how this went unnoticed.
# Keep a blank line between any comment and the help block, this one included.
[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [string]$ConfigPath,

    [Parameter(ParameterSetName = 'SaveCredential', Mandatory)]
    [switch]$SaveCredential,

    [Parameter(ParameterSetName = 'TestConnection', Mandatory)]
    [switch]$TestConnection,

    [Parameter(ParameterSetName = 'Install', Mandatory)]
    [switch]$Install,

    [Parameter(ParameterSetName = 'Uninstall', Mandatory)]
    [switch]$Uninstall,

    [Parameter(ParameterSetName = 'Run')]
    [switch]$NoEmail
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:TaskName = 'Veeam License Usage Report (Monthly)'
$script:LogFile  = $null
$script:WebArgs  = @{ UseBasicParsing = $true }

if (-not $ConfigPath) {
    $ConfigPath = Join-Path $PSScriptRoot 'VeeamLicenseReport.config.json'
}

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

# ============================================================================
# Configuration
# ============================================================================

function Get-ConfigValue {
    # Safe optional-property read, needed because Set-StrictMode is on.
    param($Object, [Parameter(Mandatory)][string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object.PSObject.Properties.Name -contains $Name) {
        $value = $Object.$Name
        if ($null -eq $value -or ($value -is [string] -and $value -eq '')) { return $Default }
        return $value
    }
    return $Default
}

function Get-ReportConfig {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Configuration file not found: $Path. Copy VeeamLicenseReport.config.sample.json to that name and edit it."
    }
    $cfg = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    foreach ($section in 'EnterpriseManager', 'Output') {
        if ($cfg.PSObject.Properties.Name -notcontains $section) {
            throw "Configuration file is missing the '$section' section."
        }
    }
    if (-not (Get-ConfigValue $cfg.EnterpriseManager 'Server')) {
        throw "EnterpriseManager.Server is not set in $Path."
    }
    if (-not (Get-ConfigValue $cfg.Output 'Folder')) {
        throw "Output.Folder is not set in $Path."
    }
    return $cfg
}

function Import-StoredCredential {
    param([string]$Path)
    if (-not $Path) {
        throw 'EnterpriseManager.CredentialFile is not set in the configuration file.'
    }
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "No stored credential at $Path. Run this script with -SaveCredential, as the account the scheduled task will run under."
    }
    $raw = Get-Content -LiteralPath $Path -Raw
    if ($raw -match '^\s*\{') {
        # Machine scope payload. Readable by any account on this computer,
        # which is what lets the task run with nobody signed in.
        $data = $raw | ConvertFrom-Json
        Add-Type -AssemblyName System.Security
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
        throw "Could not decrypt $Path. This is a user scope file, and DPAPI ties those to one Windows account on one machine. If the task runs as SYSTEM or as a service account, run -SaveCredential again to rewrite it in machine scope, which is the default. Underlying error: $($_.Exception.Message)"
    }
}

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
            if ($name -and $name -notin @('SYSTEM', 'NT AUTHORITY\SYSTEM', 'S-1-5-18')) {
                [void]$sids.Add((Resolve-TaskAccount -Name $name).Sid)
            }
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
        Write-Log 'User scope means no other account can read it, so a task running as SYSTEM or as a service account will fail unless that account saves its own copy.' 'WARN'
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
    Write-Log 'Machine scope is the trade that buys unattended runs: any administrator on this computer can decrypt it, which is why the file is restricted. It is still useless on any other computer.' 'WARN'
}

# ============================================================================
# TLS and certificates
# ============================================================================

function Initialize-WebSettings {
    param([bool]$SkipCertificateCheck)

    try {
        [System.Net.ServicePointManager]::SecurityProtocol =
            [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
    } catch { }

    if (-not $SkipCertificateCheck) { return }

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
    Write-Log 'Certificate validation is disabled for this run. Install a trusted certificate on Enterprise Manager and set SkipCertificateCheck to false when you can.' 'WARN'
}

function Get-EmBaseUri {
    param([Parameter(Mandatory)]$EmConfig)
    return 'https://{0}:{1}' -f (Get-ConfigValue $EmConfig 'Server'), (Get-ConfigValue $EmConfig 'Port' 9443)
}

# ============================================================================
# Sign-in
# ============================================================================

function Connect-VeeamEmWebSession {
    <#
        Reproduces the browser's form login and returns a WebRequestSession
        holding the resulting cookies.

        The login page serves several forms. The credential one carries
        id="by-credentials-form"; its action already contains the OAuth
        client_id, response_type and redirect_uri, so the action is read from
        the page rather than hardcoded. That way a change to the client id in a
        future build does not break the script.
    #>
    param(
        [Parameter(Mandatory)][string]$BaseUri,
        [Parameter(Mandatory)][pscredential]$Credential
    )

    $session = $null
    $loginPageUri = "$BaseUri/index"

    Write-Log "Fetching the Enterprise Manager login page from $loginPageUri"
    try {
        $page = Invoke-WebRequest -Uri $loginPageUri -Method Get -SessionVariable session `
            -MaximumRedirection 10 @script:WebArgs
    }
    catch {
        throw "Could not load the Enterprise Manager login page. Confirm the server is reachable on this port. Underlying error: $($_.Exception.Message)"
    }

    $html = [string]$page.Content

    # The credential form's action, which carries the OAuth parameters.
    $formMatch = [regex]::Match($html, '<form[^>]*id="by-credentials-form"[^>]*?action="([^"]+)"', 'IgnoreCase')
    if (-not $formMatch.Success) {
        $formMatch = [regex]::Match($html, '<form[^>]*?action="([^"]*oauth/login[^"]*)"', 'IgnoreCase')
    }
    if (-not $formMatch.Success) {
        throw 'Could not find the credential login form on the Enterprise Manager login page. The sign-in page has probably changed; re-capture it and update this script.'
    }
    $action = [System.Net.WebUtility]::HtmlDecode($formMatch.Groups[1].Value)
    $postUri = if ($action -match '^https?://') { $action } else { $BaseUri + $action }

    # ASP.NET Core anti-forgery: a hidden field that must match a cookie the
    # session already holds.
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

    # A failed sign-in returns the login page again rather than an error status.
    $result = [string]$response.Content
    if ($result -match 'name="Input\.Password"') {
        $reason = ''
        $summary = [regex]::Match($result, 'validation-summary[^>]*>\s*(?:<[^>]+>\s*)*([^<]{3,200})', 'IgnoreCase')
        if ($summary.Success) { $reason = ' Message from Enterprise Manager: ' + $summary.Groups[1].Value.Trim() }
        throw "Sign-in was rejected for $($Credential.UserName).$reason Confirm the password is current and that the account holds the Portal Administrator role in Enterprise Manager."
    }

    Write-Log 'Signed in; session cookie acquired.' 'OK'
    return $session
}

function Get-EmCsrfToken {
    <#
        The SPA shell at /index declares its CSRF request token as a plain
        JavaScript variable:

            var CSRFToken = 'CfDJ8...';

        Every XHR the UI makes sends it back as an x-csrf-token header, paired
        with the .AspNetCore.Antiforgery.* cookie the session already holds.
        API calls are rejected without it.
    #>
    param(
        [Parameter(Mandatory)][string]$BaseUri,
        [Parameter(Mandatory)]$Session
    )

    $page = Invoke-WebRequest -Uri "$BaseUri/index" -Method Get -WebSession $Session `
        -MaximumRedirection 10 @script:WebArgs

    $match = [regex]::Match([string]$page.Content, 'var\s+CSRFToken\s*=\s*[''"]([^''"]+)[''"]')
    if (-not $match.Success) {
        # Fall back to a looser pattern in case the declaration style changes.
        $match = [regex]::Match([string]$page.Content, 'csrf[A-Za-z]*\s*[:=]\s*[''"]([^''"]{20,})[''"]', 'IgnoreCase')
    }
    if (-not $match.Success) {
        throw 'Could not find the CSRF token on the Enterprise Manager app page. The page layout has probably changed.'
    }

    Write-Log 'CSRF token acquired.' 'OK'
    return $match.Groups[1].Value
}

function Get-EmApiHeader {
    param(
        [Parameter(Mandatory)][string]$BaseUri,
        [Parameter(Mandatory)][string]$CsrfToken,
        [string]$Referer
    )
    if (-not $Referer) { $Referer = "$BaseUri/index" }
    return @{
        'Accept'           = '*/*'
        'Origin'           = $BaseUri
        'Referer'          = $Referer
        'X-Requested-With' = 'XMLHttpRequest'
        'x-csrf-token'     = $CsrfToken
        'timezone-offset'  = [string](Get-TimezoneOffsetMinutes)
    }
}

function Get-EmLicensingInfo {
    <#
        GET /api/Licensing/GetInfo is what the Licensing > Summary tab calls. It
        returns the licence edition, validity, per-platform instance usage and
        the licensed/used instance totals as JSON. This is a documented-by-
        observation endpoint, confirmed working against EM 13.1.
    #>
    param(
        [Parameter(Mandatory)][string]$BaseUri,
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)]$Headers
    )

    $dc = [int64]([datetimeoffset]::UtcNow.ToUnixTimeMilliseconds())
    $uri = "$BaseUri/api/Licensing/GetInfo?_dc=$dc"

    Write-Log "Requesting the licensing summary from $uri"
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

function New-EmReportPreview {
    <#
        This is the step the Report link performs, and the one without which
        DownloadReport answers "Monthly report not found".

            POST /api/Licensing/CreateMonthlyReportPreview
            Content-Type: application/json
            Body: false   (current license consumption)
                  true    (the archived report for the month just ended)

        The body is a bare JSON boolean, not an object. The response carries the
        rendered report as HTML, which is not needed here: the point is the
        server-side state it leaves behind for the download to collect.
    #>
    param(
        [Parameter(Mandatory)][string]$BaseUri,
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)]$Headers,
        [bool]$Monthly = $false
    )

    $body = if ($Monthly) { 'true' } else { 'false' }
    $label = if ($Monthly) { 'the archived monthly report' } else { 'current license consumption' }

    Write-Log "Generating $label (POST /api/Licensing/CreateMonthlyReportPreview, body $body)."
    try {
        $response = Invoke-WebRequest -Uri "$BaseUri/api/Licensing/CreateMonthlyReportPreview" `
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
        $html = [string](Get-ConfigValue $preview 'html')
        $m = [regex]::Match($html, '<b>([^<]{3,60})</b>')
        if ($m.Success) { $heading = $m.Groups[1].Value.Trim() }
    } catch { }

    Write-Log ("Report generated{0} ({1:N0} bytes of preview HTML)." -f `
        $(if ($heading) { ": $heading" } else { '' }), $response.RawContentLength) 'OK'
}

# ============================================================================
# Report download
# ============================================================================

function Get-TimezoneOffsetMinutes {
    # Matches the browser's Date.getTimezoneOffset(), which is minutes WEST of
    # UTC: UTC-8 reports 480. Enterprise Manager uses it to localise the
    # timestamps printed in the report.
    return [int](-([System.TimeZoneInfo]::Local.GetUtcOffset([datetime]::Now).TotalMinutes))
}

function Build-ReportUri {
    param(
        [Parameter(Mandatory)][string]$BaseUri,
        [Parameter(Mandatory)][string]$Endpoint,
        [Parameter(Mandatory)][string]$Format,
        [string]$FormatParameter = 'format',
        $ExtraQuery,
        [bool]$SendTimezoneOffset = $true
    )

    $query = [ordered]@{}
    if ($ExtraQuery) {
        foreach ($property in $ExtraQuery.PSObject.Properties) {
            $query[$property.Name] = [string]$property.Value
        }
    }
    if ($FormatParameter) { $query[$FormatParameter] = $Format.ToLower() }
    if ($SendTimezoneOffset) { $query['timezoneOffset'] = [string](Get-TimezoneOffsetMinutes) }

    $pairs = foreach ($key in $query.Keys) {
        '{0}={1}' -f [uri]::EscapeDataString($key), [uri]::EscapeDataString($query[$key])
    }

    $uri = $BaseUri + $Endpoint
    if ($pairs) {
        # Contains, not -like '*?*': in a wildcard pattern '?' matches any single
        # character, so '*?*' would match every non-empty string.
        $separator = if ($uri.Contains('?')) { '&' } else { '?' }
        $uri = $uri + $separator + ($pairs -join '&')
    }
    return $uri
}

function Get-EmErrorMessage {
    <#
        Enterprise Manager returns errors as a styled HTML page. This strips the
        markup so the actual message, for example "Monthly report not found,
        please try again", reaches the log instead of a bare status code.
    #>
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
        # Drop the boilerplate the error page wraps around the real message.
        $text = $text -replace '^Error\s+Veeam Backup Enterprise Manager\s+', ''
        $text = $text -replace '\s*Here is what you can do:.*$', ''
        if ($text.Length -gt 300) { $text = $text.Substring(0, 300) + '...' }
    } catch { }

    return (($status, $text) | Where-Object { $_ }) -join ': '
}

function Save-EmLicenseReport {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$Format,
        [Parameter(Mandatory)][string]$OutFile,
        $Headers
    )

    Write-Log "Downloading the $Format report from $Uri"
    $callArgs = @{} + $script:WebArgs
    if ($Headers) { $callArgs['Headers'] = $Headers }
    try {
        $response = Invoke-WebRequest -Uri $Uri -Method Get -WebSession $Session `
            -MaximumRedirection 10 @callArgs
    }
    catch {
        $detail = Get-EmErrorMessage -ErrorRecord $_
        if ($detail -match 'report not found') {
            throw @"
Enterprise Manager answered: $detail

This endpoint only retrieves a report that already exists; it does not create
one. In the UI, clicking Report generates the report first and Download then
fetches it. A fresh session has nothing to fetch, hence this error.

The generating request still needs to be captured. In the Enterprise Manager UI,
open F12 > Network, filter to Fetch/XHR, then click Report and note the request
that fires. Add it to this script ahead of the download. See the README section
"The missing generate step".
"@
        }
        throw "The $Format download failed. $detail"
    }

    $contentType = ''
    try { $contentType = [string]$response.Headers['Content-Type'] } catch { }

    # A redirect back to the login page returns HTML with a 200, so the content
    # type is the only reliable signal that the session was still valid.
    if ($contentType -like 'text/html*') {
        throw "The $Format download returned HTML rather than a report, which means the session was not accepted. Re-run with -TestConnection to check the sign-in."
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
    Write-Log ("Saved {0} ({1:N0} KB, {2})" -f $item.Name, ($item.Length / 1KB), $contentType) 'OK'
    return $item.FullName
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
    <#
        Turns the GetInfo response into CSV rows so consumption can be trended
        across months. Field names below were read from a live EM 13.1 response,
        not guessed, but every lookup still tolerates absence so a schema change
        degrades to blanks instead of failing the run.
    #>
    param(
        [Parameter(Mandatory)]$Info,
        [Parameter(Mandatory)][string]$HistoryPath,
        [Parameter(Mandatory)][string]$MonthStamp
    )

    $generated = (Get-Date).ToUniversalTime().ToString('s')
    $status  = Get-ConfigValue $Info 'licenseStatus'
    $details = Get-ConfigValue $Info 'licenseInfo'

    $edition  = Get-ConfigValue $status 'licEdition'
    $daysLeft = Get-ConfigValue $status 'daysLeft'

    # Both name/value lists, flattened into one map. As observed on EM 13.1:
    #   licenseInformationItems: Status, Type, Package, Support ID, Licensed to
    #   instancesItems:          Instances, Expiration date
    $named = @{}
    foreach ($list in 'licenseInformationItems', 'instancesItems', 'socketsItems', 'capacityItems') {
        foreach ($item in @(Get-ConfigValue $details $list @())) {
            $name = [string](Get-ConfigValue $item 'name')
            if ($name) { $named[$name] = [string](Get-ConfigValue $item 'value') }
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

    # "Instances" reads "500 (33 used)". Matched by name rather than by scanning
    # every value, because "Expiration date" also begins with digits.
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

    # "Expiration date" reads "3/10/2028 (483 days left)"; keep just the date.
    $expiration = [string](& $lookup @('Expiration', 'Expiry'))
    if ($expiration) { $expiration = ($expiration -replace '\s*\(.*$', '').Trim() }

    $rows = @()

    # One row per platform, which is the Count / Multiplier / Instances table.
    $usage = Get-ConfigValue $status 'instancesUsage'
    foreach ($platform in @(Get-ConfigValue $usage 'platformUsages' @())) {
        $count = Get-ConfigValue $platform 'count' 0
        $instances = Get-ConfigValue $platform 'instances' 0
        if ($count -eq 0 -and $instances -eq 0) { continue }   # skip unused platforms
        $rows += [pscustomobject]@{
            Month             = $MonthStamp
            GeneratedUtc      = $generated
            Scope             = 'Platform'
            Name              = (Get-ConfigValue $platform 'platform')
            Count             = $count
            Multiplier        = (Get-ConfigValue $platform 'multiplier')
            Instances         = $instances
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
    }

    $rows += [pscustomobject]@{
        Month             = $MonthStamp
        GeneratedUtc      = $generated
        Scope             = 'Total'
        Name              = 'ALL'
        Count             = $null
        Multiplier        = $null
        Instances         = $used
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

    Write-HistoryCsv -Rows $rows -HistoryPath $HistoryPath -MonthStamp $MonthStamp
    Write-Log "Licence: $edition, $used of $licensed instances used, $daysLeft days remaining." 'OK'
}

# ============================================================================
# Email and retention
# ============================================================================

function Send-ReportEmail {
    param(
        [Parameter(Mandatory)]$EmailConfig,
        [Parameter(Mandatory)][string[]]$Attachments,
        [Parameter(Mandatory)][string]$MonthStamp
    )

    $to = @(Get-ConfigValue $EmailConfig 'To' @())
    if (-not $to) { throw 'Email.To is empty.' }

    $body = @(
        "The Veeam license usage report for $MonthStamp is attached.",
        '',
        "Generated by Invoke-VeeamLicenseUsageReport.ps1 on $env:COMPUTERNAME."
    ) -join [Environment]::NewLine

    $params = @{
        SmtpServer  = (Get-ConfigValue $EmailConfig 'SmtpServer')
        Port        = [int](Get-ConfigValue $EmailConfig 'Port' 25)
        From        = (Get-ConfigValue $EmailConfig 'From')
        To          = $to
        Subject     = ((Get-ConfigValue $EmailConfig 'Subject' 'Veeam license usage report, {0}') -f $MonthStamp)
        Body        = $body
        Attachments = $Attachments
    }
    if (Get-ConfigValue $EmailConfig 'UseSsl' $false) { $params['UseSsl'] = $true }

    $credPath = Get-ConfigValue $EmailConfig 'CredentialFile'
    if ($credPath) { $params['Credential'] = Import-StoredCredential -Path $credPath }

    # Send-MailMessage is obsolete but is still the lowest-dependency option on
    # Windows PowerShell 5.1. Replace with a Microsoft Graph call if basic SMTP
    # submission has been retired.
    Send-MailMessage @params -WarningAction SilentlyContinue
    Write-Log "Emailed the report to $($to -join ', ')" 'OK'
}

function Remove-ExpiredReport {
    param([Parameter(Mandatory)][string]$Root, [int]$RetentionMonths)
    if ($RetentionMonths -le 0) { return }
    $cutoff = (Get-Date).AddMonths(-$RetentionMonths)
    foreach ($dir in Get-ChildItem -LiteralPath $Root -Directory -ErrorAction SilentlyContinue) {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParseExact($dir.Name, 'yyyy-MM', [cultureinfo]::InvariantCulture, 'None', [ref]$parsed)) {
            if ($parsed -lt $cutoff) {
                Remove-Item -LiteralPath $dir.FullName -Recurse -Force
                Write-Log "Pruned $($dir.Name), past the $RetentionMonths month retention window."
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
        'Schedule.RunAsUser is the Windows account the scheduled task logs on as. It is not the Enterprise Manager login: that one is held in the credential file and is supplied over HTTPS at run time.',
        '',
        'Either install the task on a server joined to the account domain, which is the usual place for it, or run it here under an account this computer knows, by setting Schedule.RunAsUser to',
        "  $env:USERDOMAIN\$env:USERNAME",
        'and Schedule.LogonType to Interactive.'
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
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$ResolvedConfigPath)

    # The config file matters as much as the script here: the task passes
    # -ConfigPath, so an evicted config fails the run just as surely.
    Protect-ScriptAvailability -Path $PSCommandPath
    Protect-ScriptAvailability -Path $ResolvedConfigPath

    $schedule = Get-ConfigValue $Config 'Schedule'
    $time     = [string](Get-ConfigValue $schedule 'Time' '22:00')
    $runAs    = [string](Get-ConfigValue $schedule 'RunAsUser' 'SYSTEM')
    $wantedLogon = [string](Get-ConfigValue $schedule 'LogonType' 'Auto')

    if ($time -notmatch '^([01]\d|2[0-3]):[0-5]\d$') {
        throw "Schedule.Time must be 24 hour HH:mm, for example 22:00. Got '$time'."
    }
    if ($wantedLogon -notin @('Auto', 'Password', 'Interactive')) {
        throw "Schedule.LogonType must be Auto, Password or Interactive. Got '$wantedLogon'."
    }

    $arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -ConfigPath "{1}"' -f $PSCommandPath, $ResolvedConfigPath

    # "Last day of the month" has to be expressed as <Day>Last</Day> in the task
    # XML. The MSFT_TaskMonthlyTrigger CIM class exposes a RunOnLastDayOfMonth
    # property, but the Task Scheduler service rejects a trigger built that way
    # with "The parameter is incorrect", so the XML is written directly.
    $isSystem = $runAs -in @('SYSTEM', 'NT AUTHORITY\SYSTEM', 'S-1-5-18')
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
        $account = Resolve-TaskAccount -Name $runAs
        $principalUser = $account.Name
        $principalSid = $account.Sid
        if ($account.Name -ne $runAs) {
            Write-Log "Resolved '$runAs' to $($account.Name)."
        }

        $currentSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
        if ($account.Sid -ne $currentSid) {
            Write-Log "The task will run as $($account.Name), which is not the account running this install. The credential file is DPAPI encrypted for one account on one computer, so run -SaveCredential while signed in as $($account.Name) on $env:COMPUTERNAME, or the task will not be able to decrypt it." 'WARN'
        }

        $effectiveLogon = $wantedLogon
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
            if (-not $taskCred) { throw 'A password is required to register the task under a named account. Set Schedule.LogonType to Interactive to register one that needs no stored password.' }
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
    <Description>Downloads the aggregated Veeam License Usage Report from Veeam Backup Enterprise Manager on the last day of each month.</Description>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>$(Get-Date -Format 'yyyy-MM-dd')T$($time):00</StartBoundary>
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
      <WorkingDirectory>$(& $xmlEscape $PSScriptRoot)</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@

    $registerArgs = @{
        TaskName = $script:TaskName
        Xml      = $taskXml
        Force    = $true
    }
    if ($taskPassword) {
        $registerArgs['User'] = $principalUser
        $registerArgs['Password'] = $taskPassword
    }

    try {
        Register-ScheduledTask @registerArgs | Out-Null
        $next = $null
        try { $next = (Get-ScheduledTaskInfo -TaskName $script:TaskName).NextRunTime } catch { }
        Write-Log "Registered '$($script:TaskName)'. Next run: $(if ($next) { $next } else { 'see Task Scheduler' })" 'OK'
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
            '/Create', '/TN', $script:TaskName,
            '/TR', ($trCommand -replace '"', '\"'),
            '/SC', 'MONTHLY', '/MO', 'LASTDAY', '/M', '*',
            '/ST', $time, '/RL', 'HIGHEST', '/F'
        )
        if ($taskPassword) { $schtasksArgs += @('/RU', $principalUser, '/RP', $taskPassword) }
        elseif ($isSystem) { $schtasksArgs += @('/RU', 'SYSTEM') }
        else { $schtasksArgs += @('/RU', $principalUser, '/IT') }
        & schtasks.exe @schtasksArgs
        if ($LASTEXITCODE -ne 0) { throw "schtasks.exe exited with code $LASTEXITCODE." }
        Write-Log 'Registered the task through schtasks.exe.' 'OK'
    }

    Get-ScheduledTask -TaskName $script:TaskName -ErrorAction SilentlyContinue |
        Select-Object TaskName, State |
        Format-List | Out-String | Write-Host
}

function Uninstall-ReportTask {
    if (Get-ScheduledTask -TaskName $script:TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $script:TaskName -Confirm:$false
        Write-Log "Removed '$($script:TaskName)'." 'OK'
    }
    else {
        Write-Log "No scheduled task named '$($script:TaskName)' was found." 'WARN'
    }
}

# ============================================================================
# Main
# ============================================================================

try {
    switch ($PSCmdlet.ParameterSetName) {

        'Uninstall' {
            Uninstall-ReportTask
            exit 0
        }

        'SaveCredential' {
            $config = Get-ReportConfig -Path $ConfigPath
            $em = $config.EnterpriseManager

            $scope = [string](Get-ConfigValue $em 'CredentialScope' 'Machine')
            if ($scope -notin @('Machine', 'User')) {
                throw "EnterpriseManager.CredentialScope must be Machine or User. Got '$scope'."
            }

            # The task's own account is granted automatically, since reading
            # the credential is the whole reason it exists. GrantAccess is for
            # anything else that needs it, such as a second service account.
            $grant = New-Object System.Collections.Generic.List[string]
            [void]$grant.Add([string](Get-ConfigValue (Get-ConfigValue $config 'Schedule') 'RunAsUser' 'SYSTEM'))
            foreach ($extra in @(Get-ConfigValue $em 'GrantAccess' @())) {
                if ($extra) { [void]$grant.Add([string]$extra) }
            }

            Save-StoredCredential -Path (Get-ConfigValue $em 'CredentialFile') `
                -Scope $scope -GrantTo $grant.ToArray()
            exit 0
        }

        'Install' {
            $config = Get-ReportConfig -Path $ConfigPath
            Install-ReportTask -Config $config -ResolvedConfigPath (Resolve-Path -LiteralPath $ConfigPath).Path
            exit 0
        }

        'TestConnection' {
            $config = Get-ReportConfig -Path $ConfigPath
            Initialize-WebSettings -SkipCertificateCheck ([bool](Get-ConfigValue $config.EnterpriseManager 'SkipCertificateCheck' $false))
            $baseUri = Get-EmBaseUri -EmConfig $config.EnterpriseManager
            $session = Connect-VeeamEmWebSession -BaseUri $baseUri `
                -Credential (Import-StoredCredential -Path (Get-ConfigValue $config.EnterpriseManager 'CredentialFile'))

            $csrf = Get-EmCsrfToken -BaseUri $baseUri -Session $session
            $headers = Get-EmApiHeader -BaseUri $baseUri -CsrfToken $csrf
            Write-Log "Timezone offset this host will send: $(Get-TimezoneOffsetMinutes) minutes west of UTC."

            $info = Get-EmLicensingInfo -BaseUri $baseUri -Session $session -Headers $headers
            $status = Get-ConfigValue $info 'licenseStatus'
            Write-Log ("Edition: {0}. {1} {2} days remaining." -f `
                (Get-ConfigValue $status 'licEdition'),
                (Get-ConfigValue $status 'validationStatusComment'),
                (Get-ConfigValue $status 'daysLeft')) 'OK'
            Write-Log 'Licensing summary test passed.' 'OK'

            # The aggregated report: generate, then download, exactly as a run does.
            $reportConfig = Get-ConfigValue $config 'Report'
            $endpoint = [string](Get-ConfigValue $reportConfig 'Endpoint' '/api/ZipDownload/DownloadReport')
            $monthly  = [bool](Get-ConfigValue $reportConfig 'Monthly' $false)
            if ($endpoint) {
                $reportHeaders = Get-EmApiHeader -BaseUri $baseUri -CsrfToken $csrf `
                    -Referer ("{0}/MonthlyReport?isMonthly={1}" -f $baseUri, $monthly.ToString().ToLower())
                New-EmReportPreview -BaseUri $baseUri -Session $session -Headers $reportHeaders -Monthly $monthly

                $uri = Build-ReportUri -BaseUri $baseUri -Endpoint $endpoint -Format 'Pdf' `
                    -FormatParameter ([string](Get-ConfigValue $reportConfig 'FormatParameter' 'format')) `
                    -ExtraQuery (Get-ConfigValue $reportConfig 'ExtraQuery') `
                    -SendTimezoneOffset ([bool](Get-ConfigValue $reportConfig 'SendTimezoneOffset' $true))

                $response = Invoke-WebRequest -Uri $uri -Method Get -WebSession $session `
                    -Headers $reportHeaders -MaximumRedirection 10 @script:WebArgs
                $contentType = ''
                try { $contentType = [string]$response.Headers['Content-Type'] } catch { }
                if ($contentType -like 'text/html*') {
                    throw 'The aggregated report endpoint returned HTML rather than a document.'
                }
                Write-Log ("Aggregated report available: {0}, {1:N0} bytes." -f $contentType, $response.RawContentLength) 'OK'
            }
            Write-Log 'Connection test passed.' 'OK'
            exit 0
        }
    }

    # ---------------- Normal run ----------------

    $config = Get-ReportConfig -Path $ConfigPath

    $root = [string](Get-ConfigValue $config.Output 'Folder')
    if (-not (Test-Path -LiteralPath $root)) { New-Item -ItemType Directory -Path $root -Force | Out-Null }

    $logDir = Join-Path $root '_logs'
    if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $script:LogFile = Join-Path $logDir ('run_{0}.log' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

    $monthStamp = Get-Date -Format 'yyyy-MM'
    $monthFolder = Join-Path $root $monthStamp
    if (-not (Test-Path -LiteralPath $monthFolder)) { New-Item -ItemType Directory -Path $monthFolder -Force | Out-Null }

    Write-Log "=== Veeam license usage report, $monthStamp ==="
    Write-Log "Host $env:COMPUTERNAME, account $env:USERDOMAIN\$env:USERNAME, PowerShell $($PSVersionTable.PSVersion)"

    $today = Get-Date
    $lastDay = [datetime]::DaysInMonth($today.Year, $today.Month)
    if ($today.Day -ne $lastDay) {
        Write-Log "Today is day $($today.Day) of $lastDay, so this is a mid-month snapshot rather than a month-end one. Fine for an ad hoc run." 'WARN'
    }

    $formats = @(Get-ConfigValue $config.Output 'Formats' @('Pdf', 'Json'))
    foreach ($format in $formats) {
        if ($format -notin @('Pdf', 'Json')) {
            throw "Output.Formats may contain only Pdf and Json. Got '$format'."
        }
    }

    Initialize-WebSettings -SkipCertificateCheck ([bool](Get-ConfigValue $config.EnterpriseManager 'SkipCertificateCheck' $false))
    $baseUri = Get-EmBaseUri -EmConfig $config.EnterpriseManager
    $session = Connect-VeeamEmWebSession -BaseUri $baseUri `
        -Credential (Import-StoredCredential -Path (Get-ConfigValue $config.EnterpriseManager 'CredentialFile'))

    $csrf = Get-EmCsrfToken -BaseUri $baseUri -Session $session
    $headers = Get-EmApiHeader -BaseUri $baseUri -CsrfToken $csrf

    $saved = @()

    # 1. The licensing summary. This always works and is the dependable part of
    #    the run, so it happens first and its failure is fatal.
    $info = Get-EmLicensingInfo -BaseUri $baseUri -Session $session -Headers $headers
    $summaryFile = Join-Path $monthFolder ('LicenseSummary_{0}.json' -f $monthStamp)
    $info | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $summaryFile -Encoding UTF8
    Write-Log "Saved the licensing summary to $summaryFile" 'OK'
    $saved += $summaryFile

    Add-HistoryRow -Info $info -HistoryPath (Join-Path $root 'LicenseUsage_History.csv') -MonthStamp $monthStamp

    # 2. The aggregated report. Two calls per format: generate, then download.
    #    A failed request invalidates the session, so each format gets a fresh
    #    sign-in and a fresh generation rather than sharing state.
    $reportConfig = Get-ConfigValue $config 'Report'
    $endpoint = [string](Get-ConfigValue $reportConfig 'Endpoint' '/api/ZipDownload/DownloadReport')
    $monthly  = [bool](Get-ConfigValue $reportConfig 'Monthly' $false)
    $reportHeaders = Get-EmApiHeader -BaseUri $baseUri -CsrfToken $csrf `
        -Referer ("{0}/MonthlyReport?isMonthly={1}" -f $baseUri, $monthly.ToString().ToLower())

    if ($endpoint) {
        $first = $true
        foreach ($format in $formats) {
            try {
                if (-not $first) {
                    # Re-establish, so one format's failure cannot affect the next.
                    $session = Connect-VeeamEmWebSession -BaseUri $baseUri `
                        -Credential (Import-StoredCredential -Path (Get-ConfigValue $config.EnterpriseManager 'CredentialFile'))
                    $csrf = Get-EmCsrfToken -BaseUri $baseUri -Session $session
                    $reportHeaders = Get-EmApiHeader -BaseUri $baseUri -CsrfToken $csrf `
                        -Referer ("{0}/MonthlyReport?isMonthly={1}" -f $baseUri, $monthly.ToString().ToLower())
                }
                $first = $false

                New-EmReportPreview -BaseUri $baseUri -Session $session -Headers $reportHeaders -Monthly $monthly

                $uri = Build-ReportUri -BaseUri $baseUri -Endpoint $endpoint -Format $format `
                    -FormatParameter ([string](Get-ConfigValue $reportConfig 'FormatParameter' 'format')) `
                    -ExtraQuery (Get-ConfigValue $reportConfig 'ExtraQuery') `
                    -SendTimezoneOffset ([bool](Get-ConfigValue $reportConfig 'SendTimezoneOffset' $true))

                $outFile = Join-Path $monthFolder ('VeeamLicenseUsageReport_{0}.{1}' -f $monthStamp, $format.ToLower())
                $saved += Save-EmLicenseReport -Uri $uri -Session $session -Format $format `
                    -OutFile $outFile -Headers $reportHeaders
            }
            catch {
                Write-Log $_.Exception.Message 'WARN'
                Write-Log "Continuing without the $format report; the licensing summary above was still collected." 'WARN'
            }
        }
    }

    $emailConfig = Get-ConfigValue $config 'Email'
    if (-not $NoEmail -and (Get-ConfigValue $emailConfig 'Enabled' $false)) {
        Send-ReportEmail -EmailConfig $emailConfig -Attachments $saved -MonthStamp $monthStamp
    }

    Remove-ExpiredReport -Root $root -RetentionMonths ([int](Get-ConfigValue $config.Output 'RetentionMonths' 24))

    Write-Log "=== Finished: $(@($saved).Count) file(s) in $monthFolder ===" 'OK'
    exit 0
}
catch {
    Write-Log $_.Exception.Message 'ERROR'
    if ($_.ScriptStackTrace) { Write-Log ($_.ScriptStackTrace | Out-String) 'ERROR' }
    exit 1
}
