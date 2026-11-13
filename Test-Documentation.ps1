#Requires -Version 5.1

<#
.SYNOPSIS
    Checks that the documentation still says what the scripts rely on it saying.

.DESCRIPTION
    The failure modes this project learned are recorded once, in README.md, and
    the scripts point at it rather than restating it. That only works while the
    README still contains them, so this asserts it.

    It also checks that Get-Help can actually read each script's comment based
    help, which is not visible from reading the file: a comment line directly
    above the <# block stops Get-Help associating the help with the script, and
    the generated syntax it falls back to looks enough like real help to pass.

    Finally it scans for lab and personal identifiers, since this repository is
    meant to carry none.

.EXAMPLE
    .\Test-Documentation.ps1

    Run from anywhere. Exits 0 when everything holds, 1 otherwise, so it can be
    wired into a pre-commit hook or CI.
#>

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$base = $PSScriptRoot
$failures = New-Object System.Collections.Generic.List[string]

function Test-That {
    param([string]$Name, [bool]$Condition, [string]$Detail)
    if ($Condition) {
        Write-Host ('  PASS  {0}' -f $Name) -ForegroundColor Green
    }
    else {
        Write-Host ('  FAIL  {0}{1}' -f $Name, $(if ($Detail) { ": $Detail" } else { '' })) -ForegroundColor Red
        [void]$failures.Add($Name)
    }
}

# --------------------------------------------------------------------------
# Every failure mode must survive in the README, because that is now the only
# copy. Matching is done against whitespace-normalised text: the README wraps
# these phrases across lines, and matching the raw file reports false
# positives, as the first version of this check did.
# --------------------------------------------------------------------------
$readme = (Get-Content (Join-Path $base 'README.md') -Raw) -replace '\s+', ' '

$traps = [ordered]@{
    'MSFT_TaskMonthlyTrigger cannot express last day' = 'MSFT_TaskMonthlyTrigger'
    'New-ScheduledTaskTrigger has no monthly option'  = 'New-ScheduledTaskTrigger'
    'SYSTEM task rejects LogonType ServiceAccount'    = 'ServiceAccount'
    'schtasks re-parses /TR and mangles quotes'       = 'CommandLineToArgvW'
    'schtasks caps /TR at 261 characters'             = '261'
    '-File does not split comma separated values'     = 'does not parse PowerShell syntax'
    'ValidateSet prints a set identical to the input' = 'does not belong to the set'
    'comment above <# hides the help from Get-Help'   = 'directly above'
    "-like '*?*' does not test for a question mark"   = "like '\*\?\*'"
    'a failed request invalidates the session'        = 'kills the session'
    'run-as account is Windows, not Enterprise Mgr'   = 'No mapping between account names'
    'Entra joined machines cannot cache a password'   = 'Entra joined'
    'S4U cannot decrypt user scope DPAPI'             = 'S4U'
    'Files On-Demand evicts an unpinned script'       = 'Files On-Demand'
    'registering a task is not evidence it works'     = 'not evidence that it works'
    'the three working last-day trigger forms'        = 'LASTDAY'
}

Write-Host "`nFailure modes recorded in README.md" -ForegroundColor Cyan
foreach ($name in $traps.Keys) {
    Test-That $name ($readme -match $traps[$name]) 'not found in README.md'
}

# --------------------------------------------------------------------------
# Each script points at the README, and its help is reachable and complete.
# --------------------------------------------------------------------------
$scripts = @(
    'Invoke-VeeamLicenseUsageReport.ps1'
    'share\Invoke-VeeamLicenseUsageReport.ps1'
    'share\Export-VeeamLicenseUsageReport.ps1'
)

Write-Host "`nScript help" -ForegroundColor Cyan
foreach ($rel in $scripts) {
    $path = Join-Path $base $rel
    $errs = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$errs)
    Test-That "$rel parses" (@($errs).Count -eq 0) "$(@($errs).Count) parse error(s)"

    $declared = @($ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
    $help = Get-Help $path -Full
    $rendered = @($help.parameters.parameter)

    # A null .parameters means Get-Help never found the comment block at all.
    Test-That "$rel help is reachable by Get-Help" ([bool]$help.parameters) 'Get-Help returned no parameters; check for a comment line directly above <#'
    Test-That "$rel documents every parameter" ($rendered.Count -eq $declared.Count) "$($rendered.Count) rendered of $($declared.Count) declared"

    $blank = @($rendered | Where-Object { -not ($_.description | Out-String).Trim() })
    Test-That "$rel has no empty parameter descriptions" ($blank.Count -eq 0) (($blank | ForEach-Object { $_.name }) -join ', ')
    Test-That "$rel references README.md" ((($help | Out-String) -match 'README\.md')) 'no pointer to the single copy of the traps'
}

# --------------------------------------------------------------------------
# Nothing in the repository should name the lab or anybody in particular.
# --------------------------------------------------------------------------
Write-Host "`nIdentifiers" -ForegroundColor Cyan

# Assembled from fragments on purpose. Written out whole, this line matches
# itself and the check fails on its own pattern definition, which is how the
# first version of it behaved. The user profile path is included because a
# hardcoded path is the way one of these actually reached a commit.
$pattern = @(
    'us' + 'demo'
    'veeam' + '\.local'
    'michael' + '\.abshier'
    'OneDrive - Veeam ' + 'Software Corporation'
    '0361' + '6336'
    'Demo' + ' ACME'
    'VSAP' + '-PF'
    'AM' + 'UST'
) -join '|'
$tracked = & git -C $base ls-files 2>$null
$hits = @()
if ($LASTEXITCODE -eq 0 -and $tracked) {
    foreach ($file in $tracked) {
        $full = Join-Path $base $file
        if (-not (Test-Path -LiteralPath $full)) { continue }
        foreach ($line in (Select-String -LiteralPath $full -Pattern $pattern -ErrorAction SilentlyContinue)) {
            $hits += "$file`:$($line.LineNumber)"
        }
    }
    Test-That 'no lab or personal identifiers in tracked files' ($hits.Count -eq 0) ($hits -join ', ')
}
else {
    Write-Host '  SKIP  not a git working tree, so tracked files could not be listed' -ForegroundColor Yellow
}

Write-Host ''
if ($failures.Count) {
    Write-Host "$($failures.Count) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host 'All documentation checks passed.' -ForegroundColor Green
exit 0
