#Requires -Version 5.1
<#
.SYNOPSIS
    Downloads the aggregated License Usage Report from Veeam Backup Enterprise
    Manager, once, and exits. Contains no scheduling of its own.

.DESCRIPTION
    Enterprise Manager produces one aggregated License Usage Report covering
    every backup server it manages, with a single company-level summary. This
    script retrieves that report once and exits, so you can drive it from
    Windows Task Scheduler, or any other scheduler, on whatever cadence suits
    you.

    If you would rather the script create and manage its own monthly scheduled
    task, use Invoke-VeeamLicenseUsageReport.ps1 instead. This one deliberately
    does nothing but produce the report, so that the schedule stays entirely
    under your control and visible in your own tooling.

    Each run writes into a month-stamped folder under -OutputFolder:

      * VeeamLicenseUsageReport_<yyyy-MM>.pdf    the aggregated report
      * VeeamLicenseUsageReport_<yyyy-MM>.json   the same, machine readable
      * LicenseSummary_<yyyy-MM>.json            the licensing summary
      * a row appended to LicenseUsage_History.csv for month-over-month trending
      * a transcript under _logs

    Exit codes are 0 for success and 1 for failure, which Task Scheduler
    surfaces as the task's last result, so runs can be monitored.

    IMPORTANT, PLEASE READ

    Enterprise Manager exposes no supported or documented API for this report.
    The endpoints used here are the ones its own web interface calls, and the
    script authenticates by performing the same form sign-in a browser performs.

    That works, and it has been tested, but it is not a contract. Veeam may
    change these endpoints in any update, without notice, and such a change is
    not a defect in the product. Re-run this script by hand after any Veeam
    upgrade, before trusting the scheduled run again.

    A supported alternative exists, with a trade-off. The Veeam cmdlet
    Generate-VBRLicenseUsageReport is official and documented, but it runs on a
    backup server and reports only that server's usage, so covering an estate
    means running it on each server and combining the results yourself.

.PARAMETER Server
    Enterprise Manager FQDN, for example em.contoso.com.

.PARAMETER Port
    Enterprise Manager web port. Defaults to 9443. This is the web application
    port, not the 9398 REST API port.

.PARAMETER Credential
    Enterprise Manager credential, for running the script by hand. The account
    needs the Portal Administrator role.

    For an unattended scheduled run use -CredentialPath instead, since a
    scheduled task cannot answer a prompt.

.PARAMETER CredentialPath
    Path to a credential file created with Export-Clixml, for unattended runs.
    Create it once, logged in as the account the scheduled task will run under,
    on the server that will run it:

        $p = 'C:\ProgramData\VeeamLicenseReport\em.cred'
        New-Item -ItemType Directory -Path (Split-Path $p) -Force
        Get-Credential | Export-Clixml -Path $p

    The file is protected with DPAPI, which ties it to one Windows account on
    one machine. A file created under your own login will not decrypt under a
    service account, and will not decrypt on another server.

.PARAMETER OutputFolder
    Root folder for the month folders, logs and the history CSV.

.PARAMETER Formats
    Which formats to download: Pdf, Json, or both. Defaults to both.

.PARAMETER Monthly
    Download Veeam's archived monthly report for the month just ended, instead
    of a current-consumption snapshot.

    Omitted (the default), the report is headed "Current license consumption"
    and reflects usage at the moment it runs, so schedule it for the last day
    of the month to get a month-end figure.

    Supplied, the report is headed with an explicit interval such as
    "01 October 2026 - 31 October 2026". Enterprise Manager generates that
    report on the first day of each month, so schedule this variant for early
    in the month instead. For anything audit-facing this is the better choice,
    because the reporting period is stated on the report itself.

.PARAMETER RetentionMonths
    Delete month folders older than this many months at the end of a run.
    Defaults to 0, which keeps everything. A one-off run does not delete
    anything unless you ask it to.

.PARAMETER SkipCertificateCheck
    Accept the Enterprise Manager TLS certificate without validating it.
    Required while Enterprise Manager uses its default self-signed certificate.
    Prefer installing a trusted certificate and omitting this.

.PARAMETER SmtpServer
    Supply this, with -MailFrom and -MailTo, to email the report after the run.

.PARAMETER PassThru
    Emit an object describing the run, including the paths of the files
    written, so this script can be composed into a larger one.

.EXAMPLE
    .\Export-VeeamLicenseUsageReport.ps1 -Server em.contoso.com -Credential (Get-Credential) -SkipCertificateCheck

    Run by hand, prompting for the credential.

.EXAMPLE
    .\Export-VeeamLicenseUsageReport.ps1 -Server em.contoso.com -CredentialPath C:\ProgramData\VeeamLicenseReport\em.cred -SkipCertificateCheck

    Unattended, as a scheduled task would run it.

.EXAMPLE
    .\Export-VeeamLicenseUsageReport.ps1 -Server em.contoso.com -CredentialPath C:\ProgramData\VeeamLicenseReport\em.cred -Monthly -SkipCertificateCheck

    Fetch the archived report for the month just ended. Schedule early in the month.

.EXAMPLE
    $r = .\Export-VeeamLicenseUsageReport.ps1 -Server em.contoso.com -CredentialPath C:\em.cred -PassThru
    $r.Files

    Capture the paths of the files written.

.EXAMPLE
    schtasks /Create /TN "Veeam License Usage Report" /SC MONTHLY /MO LASTDAY /M * /ST 22:00 /RL HIGHEST /RU "CONTOSO\svc-veeam" /RP * /TR "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File \"C:\Scripts\Export-VeeamLicenseUsageReport.ps1\" -Server em.contoso.com -CredentialPath C:\ProgramData\VeeamLicenseReport\em.cred -SkipCertificateCheck"

    Create the monthly task yourself, running on the last day of every month.
    See the notes below about expressing "last day of the month" correctly.

.NOTES
    SCHEDULING THIS ON THE LAST DAY OF THE MONTH

    Two of the three obvious ways to do this are wrong, so it is worth being
    explicit.

    Works, from the command line:

        schtasks /Create /TN "Veeam License Usage Report" ^
                 /SC MONTHLY /MO LASTDAY /M * /ST 22:00 ...

    Works, in the Task Scheduler GUI:

        New Trigger > Monthly > Months: select all
                    > Days: scroll to the end of the list and tick "Last"

    Works, in task XML:

        <ScheduleByMonth>
          <DaysOfMonth><Day>Last</Day></DaysOfMonth>
          <Months><January/>...<December/></Months>
        </ScheduleByMonth>

    Does NOT work: building the trigger from the MSFT_TaskMonthlyTrigger CIM
    class and setting its RunOnLastDayOfMonth property. The class exposes the
    property, but the Task Scheduler service rejects the resulting trigger with
    "The parameter is incorrect", with or without DaysOfMonth also set. Note
    too that the property is MonthOfYear, singular, not MonthsOfYear.

    Does NOT work: New-ScheduledTaskTrigger, which has no monthly option at all.

    Whichever route you take, set the task to "Run whether user is logged on or
    not", and run it as an account that can decrypt the credential file, which
    means the same account that created it.

    REQUIREMENTS

    Windows PowerShell 5.1 or later, network access to the Enterprise Manager
    web port, and an Enterprise Manager account holding the Portal
    Administrator role.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Server,

    [ValidateRange(1, 65535)]
    [int]$Port = 9443,

    [pscredential]$Credential,

    [string]$CredentialPath,

    [string]$OutputFolder = 'C:\VeeamLicenseReports',

    # No ValidateSet here. powershell.exe -File passes every argument as a
    # literal string, so a scheduled task's "-Formats Pdf,Json" arrives as a
    # single element and fails validation against the very set it looks like.
    # Normalised and checked just below instead.
    [string[]]$Formats = @('Pdf', 'Json'),

    [switch]$Monthly,

    [ValidateRange(0, 600)]
    [int]$RetentionMonths = 0,

    [switch]$SkipCertificateCheck,

    [string]$SmtpServer,
    [ValidateRange(1, 65535)]
    [int]$SmtpPort = 25,
    [switch]$SmtpUseSsl,
    [string]$MailFrom,
    [string[]]$MailTo,
    [string]$MailSubject = 'Veeam license usage report, {0}',

    [switch]$PassThru
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# powershell.exe -File does not parse PowerShell syntax. A comma separated
# value arrives as one string, while an interactive run passes a real array.
# Both are normalised to the same shape here, before anything reads them.
# This matters most for this script, which is built to be driven by an
# external scheduler.
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
# Logging and helpers
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

function Resolve-EmCredential {
    if ($Credential) {
        if ($CredentialPath) {
            Write-Log 'Both -Credential and -CredentialPath were supplied; using -Credential.' 'WARN'
        }
        return $Credential
    }

    if (-not $CredentialPath) {
        throw @'
No credential was supplied. Pass one of:

  -Credential (Get-Credential)      to be prompted, for a run by hand
  -CredentialPath <path>            for an unattended or scheduled run

To create the file for unattended use, logged in as the account the scheduled
task will run under, on the server that will run it:

    $p = 'C:\ProgramData\VeeamLicenseReport\em.cred'
    New-Item -ItemType Directory -Path (Split-Path $p) -Force
    Get-Credential | Export-Clixml -Path $p
'@
    }

    if (-not (Test-Path -LiteralPath $CredentialPath)) {
        throw "No credential file at $CredentialPath. Create it with: Get-Credential | Export-Clixml -Path '$CredentialPath', logged in as the account that will run this script."
    }

    try {
        return Import-Clixml -LiteralPath $CredentialPath
    }
    catch {
        throw "Could not decrypt $CredentialPath. DPAPI ties the file to one Windows account on one machine, so it must be created by the account that runs this script, on this server. Underlying error: $($_.Exception.Message)"
    }
}

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
    param([Parameter(Mandatory)][pscredential]$EmCredential)

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
        'Input.Username'             = $EmCredential.UserName
        'Input.Password'             = $EmCredential.GetNetworkCredential().Password
        'Input.RemainSignedIn'       = 'false'
        '__RequestVerificationToken' = $tokenMatch.Groups[1].Value
    }

    Write-Log "Signing in as $($EmCredential.UserName)."
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
        throw "Sign-in was rejected for $($EmCredential.UserName).$reason Check the password, and that the account holds the Portal Administrator role in Enterprise Manager."
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

function Connect-AndPrepare {
    # Sign in, take a CSRF token, and return everything later calls need.
    param([Parameter(Mandatory)][pscredential]$EmCredential)

    $session = Connect-VeeamEmWebSession -EmCredential $EmCredential
    $csrf = Get-EmCsrfToken -Session $session
    $referer = '{0}/MonthlyReport?isMonthly={1}' -f $script:BaseUri, $Monthly.ToString().ToLower()
    return [pscustomobject]@{
        Session       = $session
        ApiHeaders    = (Get-EmApiHeader -CsrfToken $csrf)
        ReportHeaders = (Get-EmApiHeader -CsrfToken $csrf -Referer $referer)
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

    if (Test-Path -LiteralPath $HistoryPath) {
        $rows | Export-Csv -LiteralPath $HistoryPath -NoTypeInformation -Encoding UTF8 -Append
    }
    else {
        $rows | Export-Csv -LiteralPath $HistoryPath -NoTypeInformation -Encoding UTF8
    }
    Write-Log "Appended $(@($rows).Count) row(s) to $HistoryPath" 'OK'
    Write-Log "Licence: $edition, $used of $licensed instances used, $daysLeft days remaining." 'OK'

    return [pscustomobject]@{
        Edition           = $edition
        LicensedInstances = $licensed
        UsedInstances     = $used
        DaysLeft          = $daysLeft
        LicensedTo        = $common.LicensedTo
        SupportId         = $common.SupportId
        ExpirationDate    = $expiration
    }
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
    return $heading
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
# Main
# ============================================================================

try {
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
    # so a run on 13 November 2026 files its October report under 2026-10.
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
        Write-Log "Today is day $($today.Day) of $lastDay, so this is a mid-month snapshot rather than a month-end one. Schedule this for the last day of the month, or use -Monthly." 'WARN'
    }

    $emCredential = Resolve-EmCredential
    Initialize-WebSettings -AllowUntrustedCertificate $SkipCertificateCheck.IsPresent

    $saved = @()
    $reportFiles = @()
    $heading = $null

    # 1. The licensing summary. Dependable, so it runs first and is fatal.
    $ctx = Connect-AndPrepare -EmCredential $emCredential
    $info = Get-EmLicensingInfo -Session $ctx.Session -Headers $ctx.ApiHeaders
    $summaryFile = Join-Path $monthFolder ('LicenseSummary_{0}.json' -f $monthStamp)
    $info | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $summaryFile -Encoding UTF8
    Write-Log "Saved the licensing summary to $summaryFile" 'OK'
    $saved += $summaryFile

    $licence = Add-HistoryRow -Info $info `
        -HistoryPath (Join-Path $OutputFolder 'LicenseUsage_History.csv') -MonthStamp $monthStamp

    # 2. The aggregated report: generate, then download, per format. A failed
    #    request invalidates the session, so each format gets a fresh sign-in
    #    and a fresh generation rather than sharing state. A failure here is
    #    reported but does not sink the run, so the trending data is never lost
    #    to a user interface change.
    $failures = 0
    $first = $true
    foreach ($format in $Formats) {
        try {
            if (-not $first) { $ctx = Connect-AndPrepare -EmCredential $emCredential }
            $first = $false

            $heading = New-EmReportPreview -Session $ctx.Session -Headers $ctx.ReportHeaders `
                -ArchivedMonth $Monthly.IsPresent

            $outFile = Join-Path $monthFolder ('VeeamLicenseUsageReport_{0}.{1}' -f $monthStamp, $format.ToLower())
            $file = Save-EmLicenseReport -Uri (New-ReportUri -Format $format) -Session $ctx.Session `
                -Headers $ctx.ReportHeaders -Format $format -OutFile $outFile
            $saved += $file
            $reportFiles += $file
        }
        catch {
            $failures++
            Write-Log $_.Exception.Message 'WARN'
            Write-Log "Continuing without the $format report; the licensing summary was still collected." 'WARN'
        }
    }

    if ($SmtpServer) {
        if (-not $MailFrom -or -not $MailTo) {
            Write-Log 'SmtpServer was supplied without MailFrom and MailTo, so no email was sent.' 'WARN'
        }
        else {
            Send-ReportEmail -Attachments $saved -MonthStamp $monthStamp
        }
    }

    Remove-ExpiredReport -Root $OutputFolder -KeepMonths $RetentionMonths

    Write-Log "=== Finished: $(@($saved).Count) file(s) in $monthFolder ===" 'OK'

    if ($PassThru) {
        [pscustomobject]@{
            Month            = $monthStamp
            ReportingPeriod  = $heading
            Folder           = $monthFolder
            Files            = $saved
            ReportFiles      = $reportFiles
            SummaryFile      = $summaryFile
            HistoryFile      = (Join-Path $OutputFolder 'LicenseUsage_History.csv')
            Licence          = $licence
            FormatsRequested = $Formats
            FormatsFailed    = $failures
            LogFile          = $script:LogFile
        }
    }

    # A failure to collect the aggregated report is still a failed run as far as
    # a scheduler is concerned, even though the summary was captured.
    if ($failures -gt 0) { exit 1 }
    exit 0
}
catch {
    Write-Log $_.Exception.Message 'ERROR'
    if ($_.ScriptStackTrace) { Write-Log ($_.ScriptStackTrace | Out-String) 'ERROR' }
    exit 1
}
