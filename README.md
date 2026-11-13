# Monthly Veeam License Usage Report

Downloads the aggregated **License Usage Report** from Veeam Backup Enterprise
Manager on a scheduled task that fires on the **last day of every month**.

This is the same report the UI produces from **Configuration → Licensing →
Summary → Report → Download**: one document covering every managed backup
server, with a single company-level summary. The generated PDF has been verified
byte-comparable in content to one downloaded by hand from the UI, including the
per-backup-server sections and their per-VM detail.

| File | Purpose |
| --- | --- |
| `Invoke-VeeamLicenseUsageReport.ps1` | The script. Report run, credential storage, connection test, task install and removal. |
| `VeeamLicenseReport.config.sample.json` | Copy to `VeeamLicenseReport.config.json` and edit. |

---

## Read this first: the interface is undocumented

Enterprise Manager exposes **no supported API** for this report. This was
checked thoroughly, not assumed:

- Its classic REST API on port 9398 has no licensing resource at all.
- Its OAuth token endpoint refuses the password grant to the web UI's own
  client, so there is no service-to-service path.
- The web endpoints used here answer `302 → /login` when given a bearer token.
  They are session-cookie authenticated UI controllers.

So the script signs in the way a browser does and drives the same calls the UI
makes. **It works, and it is tested, but it is not a contract.** Veeam can change
it in any update. After every Veeam upgrade, run:

```powershell
.\Invoke-VeeamLicenseUsageReport.ps1 -TestConnection
```

If you would rather not carry that risk, see
[Alternative: the fully supported route](#alternative-the-fully-supported-route).

---

## How it works

Four steps, all verified against EM 13.1.0.411.

**1. Form sign-in.** `GET /index` redirects to the OAuth login page. The
credential form is found by `id="by-credentials-form"`, and **its action is read
from the page rather than hardcoded**, because that action carries the
`client_id` and `redirect_uri`. The `__RequestVerificationToken` hidden field is
carried over, ASP.NET Core validates it against a cookie in the same session, and
the post follows redirects through `/Authorize` to collect `.VEEAMAUTH`.

**2. CSRF token.** The authenticated app shell declares it as a plain JavaScript
variable:

```javascript
var CSRFToken = 'CfDJ8...';
```

Every API call must send it back as an `x-csrf-token` header, together with
`x-requested-with: XMLHttpRequest`. Calls are rejected without them.

**3. Generate the report.** This is the step the **Report** link performs, and
the one without which the download fails:

```
POST /api/Licensing/CreateMonthlyReportPreview
Content-Type: application/json

false
```

The body is a bare JSON boolean, not an object:

- `false` → **Current license consumption**, a point-in-time snapshot
- `true` → the **archived monthly report** for the month just ended

The response returns the rendered report as HTML. That is discarded; what
matters is the server-side state it leaves for the download to collect.

**4. Download.**

```
GET /api/ZipDownload/DownloadReport?isAdjusted=1&format=pdf&timezoneOffset=480
```

`timezoneOffset` is computed from the host the way `Date.getTimezoneOffset()`
does, so report timestamps localise wherever it runs.

Also worth knowing: **a failed request invalidates the session.** After any
error, every later request returns the login page with HTTP 200. The script
therefore signs in afresh and regenerates before each format rather than sharing
state, and checks `Content-Type` instead of trusting status codes.

### Current consumption or the archived month?

`Report.Monthly` chooses. With the default `false` and a last-day-of-month
trigger you get a month-end snapshot, which matches the on-demand report the UI
produces.

Set it to `true` and move `Schedule.Time` to the 1st if you would rather capture
Veeam's own official monthly report for the month just ended. That report is
headed with an explicit interval, for example "01 October 2026 - 31 October
2026", so it is the better choice for anything audit-facing. Changing the
trigger day means editing the task XML in `Install-ReportTask`.

---

## Setup

### 1. Copy and edit the configuration

```powershell
Copy-Item .\VeeamLicenseReport.config.sample.json .\VeeamLicenseReport.config.json
```

| Setting | Notes |
| --- | --- |
| `EnterpriseManager.Server` | Enterprise Manager FQDN. |
| `EnterpriseManager.Port` | `9443`, the web app port, not 9398. |
| `EnterpriseManager.CredentialFile` | Where the encrypted credential is stored. |
| `EnterpriseManager.SkipCertificateCheck` | `true` while EM uses a self-signed certificate. |
| `Report.Endpoint` | `/api/ZipDownload/DownloadReport`. |
| `Report.Monthly` | `false` for current consumption, `true` for the archived month. |
| `Report.FormatParameter` | `format`. |
| `Report.ExtraQuery` | `isAdjusted: "1"`, matching what the UI sends. |
| `Report.SendTimezoneOffset` | Adds `timezoneOffset`, computed from this host. |
| `Output.Folder` | Root for the month folders. |
| `Output.Formats` | `Pdf`, `Json`, or both. |
| `Output.RetentionMonths` | Month folders older than this are pruned. `0` disables pruning. |
| `Schedule.Time` | 24 hour `HH:mm`. |
| `Schedule.RunAsUser` | The **Windows** account the task logs on as, not the Enterprise Manager login. A domain account is strongly preferred over `SYSTEM`. |
| `Schedule.LogonType` | `Auto` stores a password and runs unattended. `Interactive` stores none and runs only while that account is signed in. |

### 2. Store the credential

Run this **as the account the scheduled task will run under, on the server that
will run it**. DPAPI ties the file to one user on one machine, so a file created
under your own login will not decrypt under a service account.

```powershell
.\Invoke-VeeamLicenseUsageReport.ps1 -SaveCredential
```

The account needs the **Portal Administrator** role in Enterprise Manager.

### 3. Test

```powershell
.\Invoke-VeeamLicenseUsageReport.ps1 -TestConnection
```

Signs in, fetches the licensing summary, generates the report and confirms a
real document comes back. Saves nothing.

### 4. Run once

```powershell
.\Invoke-VeeamLicenseUsageReport.ps1 -NoEmail
```

It warns when run mid-month, which is expected for an ad hoc test.

### 5. Install the schedule

```powershell
.\Invoke-VeeamLicenseUsageReport.ps1 -Install
```

To remove it:

```powershell
.\Invoke-VeeamLicenseUsageReport.ps1 -Uninstall
```

---

## Output layout

```
C:\VeeamLicenseReports\
├─ 2026-11\
│   ├─ VeeamLicenseUsageReport_2026-11.pdf    <- the aggregated report
│   ├─ VeeamLicenseUsageReport_2026-11.json   <- same, machine readable
│   └─ LicenseSummary_2026-11.json            <- GetInfo snapshot
├─ 2026-12\
├─ LicenseUsage_History.csv                   <- cumulative, for trending
└─ _logs\
    └─ run_20261130_220001.log
```

The script also calls `GET /api/Licensing/GetInfo`, the endpoint behind the
Summary tab, and flattens it into `LicenseUsage_History.csv`: one row per
platform plus a `Total` row, accumulating across months.

| Month | Scope | Name | Count | Instances | Edition | LicenseType | Package | LicensedInstances | UsedInstances | DaysLeft | LicensedTo | SupportId | ExpirationDate |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 2026-11 | Platform | Virtual Machines | 24 | 24.0 | Enterprise Plus | Subscription | Advanced | 500 | 33 | 483 | Contoso Ltd | 01234567 | 3/10/2028 |
| 2026-11 | Total | ALL | | 33 | ... | | | | | | | | |

Platforms with zero usage are skipped. Field names were read from a live
response, but every lookup tolerates absence, so a schema change degrades to
blank columns instead of failing the run.

The summary is collected first and its failure is fatal. A failure of the
aggregated report is logged as a warning and the run still completes, so you
never lose the trending data to a UI change.

---

## What was verified

Against a Veeam Backup Enterprise Manager 13.1 instance managing three backup
servers:

- The generated PDF is content-identical to one downloaded by hand from the UI,
  including every per-backup-server section and the per-VM detail.
- `GetInfo` reproduces the PDF's License Information and Summary sections
  exactly: the per-platform Count / Multiplier / Instances table, the licensed
  and used instance totals, edition, package, expiry, licensed-to and support
  ID all matched the rendered report.
- `Monthly` returns the archived report for the month just ended, headed with
  an explicit interval such as "01 October 2026 - 31 October 2026".
- The 9398 API root advertises only `logonSessions` and `sessionMngr`, API
  versions `v1_5`, `v1_6`, `v1_7`. No licensing resource.
- `/api/ZipDownload/DownloadReport` returns 302 → `/login` both unauthenticated
  and with an `Authorization: Bearer` header.
- Kerberos SSO (`/oauth/gss`) returns 401 from outside the lab domain. From a
  domain-joined server the GSSAPI grant may work, which would remove the stored
  password; the login page does carry a GSS form.
- The scheduled task trigger was registered for real: `<Day>Last</Day>` returned
  a next run time of **30 November 2026 22:00**.

### Traps worth recording

**The run-as account is a Windows account, not an Enterprise Manager one.**
These are two different logins and it is easy to pass the wrong one. The
Enterprise Manager account is prompted for by `-SaveCredential`, stored
encrypted, and sent over HTTPS at run time. The run-as account is what Windows
logs the task on as, so it must be an account the machine registering the task
can resolve. Giving it an Enterprise Manager user name fails with **"No mapping
between account names and security IDs was done. (26,8):UserId:"**, where 26,8
is the line and column of `<UserId>` in the task XML. The script now resolves
the name to a SID up front, writes the SID into the XML rather than the name,
and refuses to quietly substitute a same-named local account.

**A Microsoft Entra joined workstation cannot cache a task password.** There is
no domain to authenticate against, so the usual "run whether the user is logged
on or not" registration has nothing to store. Use `Interactive`, which registers
an `InteractiveToken` principal that needs no password and fires only while that
account is signed in. Treat it as a way to try the task out; unattended runs
belong on a server joined to the service account's domain.

**`schtasks.exe` re-parses `/TR`, so quotes inside it must be escaped.**
PowerShell wraps a native argument containing spaces in quotes but leaves the
quotes already inside it alone, so `CommandLineToArgvW` splits the value apart.
With a path like `...\OneDrive - Veeam Software Corporation\...` the fragment
` - ` becomes its own token and schtasks answers **"Invalid argument/option -
'-'"**. Replace `"` with `\"` in the `/TR` value first. Note also that schtasks
caps `/TR` at 261 characters, which a full command line can easily exceed, so it
is a poor fallback for anything but a short invocation.

**`MSFT_TaskMonthlyTrigger` cannot express "last day of month".** The CIM class
has a `RunOnLastDayOfMonth` property and it looks like the right answer, but the
Task Scheduler service rejects such a trigger with "The parameter is incorrect",
with or without `DaysOfMonth` set. Note also the property is `MonthOfYear`,
singular. The script writes the task XML directly with
`<DaysOfMonth><Day>Last</Day></DaysOfMonth>`, the only accepted form, keeping
`schtasks.exe /SC MONTHLY /MO LASTDAY /M *` as a fallback.

**`-like '*?*'` does not test for a question mark.** In a wildcard pattern `?`
matches any single character, so that test matches every non-empty string. This
initially built URLs with `&` where `?` belonged. Use `.Contains('?')`.

**A failed request kills the session,** and every request after it returns the
login page with HTTP 200, not an error. This makes probing deeply misleading: a
single early failure makes every later variant look broken. Test one request per
session, and judge success by `Content-Type`.

---

## Alternative: the fully supported route

If the undocumented endpoints break and you want something that cannot, there is
a documented path producing the same figures.

`Generate-VBRLicenseUsageReport -Path <string> -Type Html|Pdf|Json` is an
official Veeam cmdlet. It runs on a backup server and produces that server's
License Usage Report.

The per-server sections in the Enterprise Manager report are **not
deduplicated**: on the instance checked, the three per-server figures summed to
exactly the company total, and a workload protected by two backup servers was
counted under both. So running the cmdlet on each managed backup server and
summing gives the same numbers. The trade-off is that the combined document would
be assembled by the script rather than by Veeam, so it would not be identical to
the Enterprise Manager PDF.

---

## Notes and caveats

- **Enterprise Manager already produces its own official monthly usage report**
  on the first day of each month, covering the month just ended. Setting
  `Report.Monthly` to `true` fetches exactly that one.
- **`Send-MailMessage` is obsolete** but remains the lowest-dependency option on
  Windows PowerShell 5.1. Replace it with a Microsoft Graph call if basic SMTP
  submission has been retired.
- **Certificate validation.** `SkipCertificateCheck` exists because Enterprise
  Manager ships with a self-signed certificate. Install a trusted certificate
  and turn it off when you can; the script warns on every run while it is on.
- **Exit codes.** `0` success, `1` failure. Task Scheduler surfaces this as the
  last result, so it can be monitored.
