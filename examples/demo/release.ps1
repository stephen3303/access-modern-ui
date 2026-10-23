# Builds a released, locked copy of the kit demo (KitDemo) from a git tag, as a
# showcase of the kit.
#
#   release.ps1 -Version 1.0.0                 build tag kitdemo-v1.0.0 into <repo>\Releases\KitDemo\v1.0.0
#   release.ps1 -Version 1.0.0 -OutRoot <dir>  put the v1.0.0 folder somewhere else
#
# The demo is one file: its fictional items live in it, so there is no shared data
# file. The release holds the locked app (KitDemo.accde), the same app as the full,
# editable program (KitDemo.accdb), the splash bitmap the app writes beside itself
# (KitDemo.bmp, HideAccessSplash), the design folder, RELEASE-NOTES.md and
# SHA256SUMS.txt.
#
# The tag is built from a clean export in TEMP, so this checkout, its uncommitted edits
# and the KitDemo.accdb next to this script are never touched. Steps:
#   1. fetch tags and check the tag kitdemo-v<Version> exists
#   2. export the tag and run its kit checks (tests\native-smoke.ps1, in its own
#      scratch folder)
#   3. build the demo from the export (examples\demo\build.ps1: kit and demo modules,
#      seed data, every page) and check its error log is empty
#   4. copy it to the release folder; write the design folder (every form as JSON and a
#      SQL Server build script for the tables, tools\ReleaseDesign.ps1; only a release
#      build makes these); compile it to KitDemo.accde and keep the .accdb as the full,
#      editable program; lock the .accde (tools\ReleaseLock.ps1: Shift can't bypass the
#      start-up, the navigation pane stays hidden); Compact & Repair both
#   5. write RELEASE-NOTES.md (merged pull requests up to the tag) and SHA256SUMS.txt
# An existing release folder is never overwritten. Any Visual Basic or Access dialog
# stops the release and is reported (tools\AccessWatch.ps1).
#
# The .accde runs only on the bitness of Office it was compiled with (this PC's).
param(
    [Parameter(Mandatory)] [ValidatePattern('^\d+\.\d+\.\d+$')] [string]$Version,
    [string]$OutRoot,
    [int]$TestTimeoutSec = 1800
)
$ErrorActionPreference = 'Stop'
$repo = git -C $PSScriptRoot rev-parse --show-toplevel
if ($LASTEXITCODE -ne 0 -or -not $repo) { throw "Not a git checkout: $PSScriptRoot" }
$repo = (Resolve-Path $repo.Trim()).Path
$tag = "kitdemo-v$Version"
if (-not $OutRoot) { $OutRoot = Join-Path $repo 'Releases\KitDemo' }
$out = Join-Path $OutRoot "v$Version"
if (Test-Path $out) { throw "$out already exists. Remove it first, or release a new version." }

# 1. The tag.
git -C $repo fetch --quiet --tags origin
if ($LASTEXITCODE -ne 0) { Write-Warning 'Could not fetch tags from origin; using the tags on this checkout.' }
git -C $repo rev-parse --quiet --verify "refs/tags/$tag" | Out-Null
if ($LASTEXITCODE -ne 0) { throw "No tag $tag (on this checkout or origin)." }

# 2. Export the tag and run its kit checks. They drive their own hidden Access, which a
# dialog would block, so a run past -TestTimeoutSec is stopped and fails.
$work = Join-Path $env:TEMP "KitDemo-release-$tag"
Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force $work | Out-Null
$zip = Join-Path $env:TEMP "KitDemo-release-$tag.zip"
git -C $repo archive --format=zip -o $zip $tag
if ($LASTEXITCODE -ne 0) { throw "git archive $tag failed." }
Expand-Archive $zip $work
Remove-Item $zip

# Runs a script of the export in its own Windows PowerShell, logged to $log; throws past
# -TestTimeoutSec or on a non-zero exit.
function Invoke-ReleaseStep([string]$what, [string]$log, [string[]]$scriptArgs) {
    $p = Start-Process powershell -PassThru -NoNewWindow -RedirectStandardOutput $log -RedirectStandardError "$log.err" `
        -ArgumentList (@('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File') + $scriptArgs)
    # Windows PowerShell 5.1 only fills in ExitCode when the process handle was opened
    # before the process ended, so open it now.
    $null = $p.Handle
    if (-not $p.WaitForExit($TestTimeoutSec * 1000)) {
        Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
        Get-Process MSACCESS -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -ge $p.StartTime } | Stop-Process -Force
        throw "$what of $tag ran past $TestTimeoutSec s (a dialog?); nothing was released. Log: $log"
    }
    Get-Content $log | Write-Host
    if ($p.ExitCode -ne 0) {
        Get-Content "$log.err" -ErrorAction SilentlyContinue | Write-Host
        throw "$what of $tag failed; nothing was released."
    }
}

$testDir = Join-Path $work 'kit-checks'
Invoke-ReleaseStep 'The kit checks' (Join-Path $work 'kit-checks.log') @(
    "`"$(Join-Path $work 'tests\native-smoke.ps1')`"", '-OutputDirectory', "`"$testDir`"")
$checks = Get-Content (Join-Path $testDir 'result.txt') -ErrorAction SilentlyContinue | Select-Object -First 1
if ("$checks" -notmatch '^\d+ kit checks passed$') { throw "The kit checks of $tag did not pass: $checks" }

# 3. Build the demo from the export.
$demo = Join-Path $work 'examples\demo'
Invoke-ReleaseStep 'Building the demo' (Join-Path $work 'build.log') @("`"$(Join-Path $demo 'build.ps1')`"")
$built = Join-Path $demo 'KitDemo.accdb'
if (-not (Test-Path $built)) { throw "The build of $tag wrote no KitDemo.accdb." }

# 4. The release folder: the built .accdb, its design files, the .accde, compacted.
New-Item -ItemType Directory -Force $out | Out-Null
$accdb = Join-Path $out 'KitDemo.accdb'
$accde = Join-Path $out 'KitDemo.accde'
Copy-Item $built $accdb
Copy-Item (Join-Path $demo 'KitDemo.bmp') $out
. (Join-Path $work 'tools\AccessWatch.ps1')
. (Join-Path $work 'tools\ReleaseDesign.ps1')
. (Join-Path $work 'tools\ReleaseLock.ps1')
$acc = New-Object -ComObject Access.Application
$watch = Start-AccessWatch $acc -TimeoutSec 1800
$failed = $false
$problem = $null
try {
    $acc.Visible = $false
    # The build ran every page's builder; any error it logged stops the release. DAO
    # opens the file read-only, so no start-up code runs.
    $db = $acc.DBEngine.OpenDatabase([string]$accdb, $false, $true)
    try { $logged = [int]$db.OpenRecordset('SELECT COUNT(*) FROM error_log').Fields.Item(0).Value }
    finally { $db.Close() }
    if ($logged) { $problem = "The build of $tag logged $logged error(s) in error_log."; throw $problem }

    # 1 (low): at 3 (force disable) SysCmd 603 returns 0 and writes no file. 603 converts
    # the file without opening it as the current database, so no start-up code runs.
    $acc.AutomationSecurity = 1
    # From the .accdb, which still has its forms' designs. Any problem stops the release.
    $design = Export-ReleaseDesign -Access $acc -FrontEnd $accdb `
        -ThemeFile (Join-Path $demo 'src\themes\light\modTheme.bas') -OutDir (Join-Path $out 'design') -WorkDir $work
    # 603 (make ACCDE) is not in the AcSysCmdAction enum, so PowerShell can't pass it
    # through the typed SysCmd method; call it late-bound instead. Join-Path returns
    # PSObject-wrapped strings, which InvokeMember won't pass as BSTRs, hence [string].
    [void][System.__ComObject].InvokeMember('SysCmd', 'InvokeMethod', $null, $acc, @(603, [string]$accdb, [string]$accde))
    # Shift can't bypass the .accde's start-up, and its navigation pane stays hidden.
    Lock-ReleaseFrontEnd -Access $acc -Path $accde

    # Compact & Repair each file into a temporary copy, then replace it. Sizes go in the
    # release notes.
    $sizes = @()
    if (Test-Path $accde) {
        foreach ($f in @($accde, $accdb)) {
            $tmp = Join-Path $out ('compact-' + (Split-Path -Leaf $f))
            Remove-Item $tmp -ErrorAction SilentlyContinue
            $before = (Get-Item $f).Length
            if (-not $acc.CompactRepair([string]$f, [string]$tmp, $false) -or -not (Test-Path $tmp)) { throw "Compact & Repair of $f failed." }
            Move-Item -Force $tmp $f
            $sizes += '- `{0}`: {1:N1} MB (compacted from {2:N1} MB)' -f (Split-Path -Leaf $f), ((Get-Item $f).Length / 1MB), ($before / 1MB)
        }
    }
}
catch { $failed = $true; if (-not $watch.Killed) { throw } }
finally { $dialogs = Stop-AccessWatch $watch $acc }
if ($problem) { throw $problem }
if ($dialogs -or $failed -or -not (Test-Path $accde)) { throw "Writing the design files, compiling KitDemo.accde, or compacting the release, failed." }

# 5. Release notes and checksums.
$prev = git -C $repo tag --list 'kitdemo-v*' --merged "$tag^" --sort=-v:refname | Select-Object -First 1
$range = if ($prev) { "$prev..$tag" } else { $tag }
# One line per merged pull request, newest first. A merge commit's title is its body, or
# (when that is empty) the subject of the pull request's last commit; a squash merge's
# subject is "<title> (#N)".
$prs = git -C $repo log --first-parent --format='%H' $range | ForEach-Object {
    $h = $_
    $subj = git -C $repo log -1 --format=%s $h
    if ($subj -match '^Merge pull request (#\d+)') {
        $num = $Matches[1]
        $title = git -C $repo log -1 --format=%b $h | Where-Object { $_.Trim() } | Select-Object -First 1
        if (-not $title) { $title = git -C $repo log -1 --format=%s "$h^2" }
    } elseif ($subj -match '^(.+) \((#\d+)\)$') {
        $title = $Matches[1]
        $num = $Matches[2]
    } else { return }
    "- ${num}: $title"
}
$notes = @("# Kit Demo v$Version", '',
    "Built $(Get-Date -Format 'yyyy-MM-dd HH:mm') from tag $tag ($(git -C $repo rev-parse --short "$tag^{commit}")) of access-modern-ui. $checks before the build.", '',
    'A small app that shows the kit at work: cards and progress, tables with filters and a column chooser, steppers and accordions, a dashboard, and a list and detail page, in every palette (Appearance in the sidebar). The items in it are fictional.', '',
    '- `KitDemo.accde`: the locked app. Copy the folder somewhere and open it; until that folder is a Trusted Location, Access blocks its code and the demo shows how to trust it (Folder trust).',
    '- `KitDemo.accdb`: the same app as the full, editable program (forms in design view, VBA).',
    '- `KitDemo.bmp`: the blank splash bitmap the app writes beside itself so Access shows no logo at start-up.',
    "- ``design\``: the $($design.Forms) forms as JSON designs (``design\forms``) and a SQL Server build script for the tables (``design\*.sqlserver.sql``), for rebuilding or porting the app. Only release builds make these.", '',
    '## Files (compacted)', '') + $sizes + @('') + @(if ($design.Skipped) { @('## Not in the SQL build scripts', '') + @($design.Skipped | ForEach-Object { "- $_" }) + @('') }) + @(
    $(if ($prev) { "## Changes since $prev" } else { '## Pull requests included' }), '') + $prs
Set-Content (Join-Path $out 'RELEASE-NOTES.md') $notes -Encoding UTF8
Get-ChildItem $out -File -Recurse | Where-Object Name -ne 'SHA256SUMS.txt' | ForEach-Object {
    '{0}  {1}' -f (Get-FileHash $_.FullName -Algorithm SHA256).Hash.ToLower(), $_.FullName.Substring($out.Length + 1)
} | Set-Content (Join-Path $out 'SHA256SUMS.txt') -Encoding ASCII
Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Released Kit Demo v$Version to $out"
exit 0
