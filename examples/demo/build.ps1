# Builds the kit demo (KitDemo.accdb, next to this script) straight from this
# repository: the kit's modules and forms, examples\modules, and the demo's src.
#
#   build.ps1           build a new KitDemo.accdb with seed data
#   build.ps1 -UiOnly   rebuild code and forms inside the existing KitDemo.accdb
#
# The build stops and reports any Visual Basic or Access dialog (a compile
# error, an error message) instead of hanging behind it.
# Requires Microsoft Access (2016+ / Microsoft 365). Python is only needed to
# regenerate the palette (python ..\..\tools\make_themes.py src).
param([switch]$UiOnly)
$ErrorActionPreference = 'Stop'
$demo = Split-Path -Parent $MyInvocation.MyCommand.Path
$repo = Split-Path -Parent (Split-Path -Parent $demo)
$fe = Join-Path $demo 'KitDemo.accdb'

# SaveForm attaches <srcRoot>\forms\<name>.vba, so stage the demo's forms and
# the kit's forms (subGrid, dlgAppearance, dlgDatePick) in one folder.
$stage = Join-Path $env:TEMP 'KitDemoSrc'
Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force (Join-Path $stage 'forms') | Out-Null
Copy-Item (Join-Path $repo 'kit\forms\*.vba') (Join-Path $stage 'forms')
Copy-Item (Join-Path $demo 'src\forms\*.vba') (Join-Path $stage 'forms')

. (Join-Path $repo 'tools\AccessWatch.ps1')
$acc = New-Object -ComObject Access.Application
$watch = Start-AccessWatch $acc -TimeoutSec 600
$failed = $false
try {
    $acc.Visible = $false
    $acc.AutomationSecurity = 1   # let the build's own VBA run
    if ($UiOnly) {
        if (-not (Test-Path $fe)) { throw "No existing front end at $fe" }
        $acc.OpenCurrentDatabase($fe)
    } else {
        Remove-Item $fe -ErrorAction SilentlyContinue
        $acc.NewCurrentDatabase($fe)
    }

    # The kit first, then the example modules, the demo's modules and the build palette.
    $files = @(Get-ChildItem (Join-Path $repo 'kit\modules') -Filter *.bas) +
             @(Get-ChildItem (Join-Path $repo 'examples\modules') -Filter *.bas) +
             @(Get-ChildItem (Join-Path $demo 'src\modules') -Filter *.bas) +
             @(Get-Item (Join-Path $demo 'src\themes\light\modTheme.bas'))
    foreach ($f in $files) {
        try { $acc.DoCmd.DeleteObject(5, $f.BaseName) } catch {}
        # LoadFromText expects Windows line endings.
        $tmp = Join-Path $env:TEMP ("kitdemo_" + $f.Name)
        $text = [IO.File]::ReadAllText($f.FullName) -replace "`r?`n", "`r`n"
        [IO.File]::WriteAllText($tmp, $text, [Text.Encoding]::ASCII)
        $acc.LoadFromText(5, $f.BaseName, $tmp)
        Remove-Item $tmp
    }

    $bad = Test-AccessCompiles $acc $watch
    if ($bad) { Write-Host "Compile error: $bad"; throw 'compile failed' }

    $a1 = [string]$stage; $a2 = [string]$demo; $a3 = [bool]$UiOnly
    $result = $acc.Run('BuildAll', [ref]$a1, [ref]$a2, [ref]$a3)
    Write-Host "BuildAll: $result"
    if ($result -ne 'OK') { $failed = $true }
}
catch { $failed = $true; if (-not $watch.Killed) { throw } }
finally {
    $dialogs = Stop-AccessWatch $watch $acc
}
if ($dialogs) { Write-Host "Build stopped: Access showed the dialog(s) above."; exit 1 }
if ($failed) { exit 1 }
Write-Host "Built: $fe"
