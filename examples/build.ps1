# Example build script for an app that uses the kit. Copy it to <app>\access\build.ps1
# next to kit\ (or point $kit at this repo's kit folder) and src\.
#
#   build.ps1                 build a new front end (<AppName>_FE.accdb) from source
#   build.ps1 -UiOnly         rebuild code, forms and reports inside the existing front end
#
# The app's src\modules must contain a public BuildAll(srcRoot, outDir, uiOnly) that
# returns "OK" or the step that failed. Requires Microsoft Access (2016+ / Microsoft 365).
param(
    [switch]$UiOnly,
    [string]$AppName = 'MyApp',
    [string]$OutDir
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $root 'src'
$kit = Join-Path $root 'kit'
if (-not $OutDir) { $OutDir = Split-Path -Parent $root }
New-Item -ItemType Directory -Force $OutDir | Out-Null
$fe = Join-Path $OutDir "$($AppName)_FE.accdb"

$acc = New-Object -ComObject Access.Application
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

    # Load standard modules: the kit first, then the app's modules, then the build palette.
    $files = @(Get-ChildItem (Join-Path $kit 'modules') -Filter *.bas) +
             @(Get-ChildItem (Join-Path $src 'modules') -Filter *.bas) +
             @(Get-Item (Join-Path $src 'themes\light\modTheme.bas'))
    foreach ($f in $files) {
        try { $acc.DoCmd.DeleteObject(5, $f.BaseName) } catch {}
        # LoadFromText expects Windows line endings.
        $tmp = Join-Path $env:TEMP ("kit_" + $f.Name)
        $text = [IO.File]::ReadAllText($f.FullName) -replace "`r?`n", "`r`n"
        [IO.File]::WriteAllText($tmp, $text, [Text.Encoding]::ASCII)
        $acc.LoadFromText(5, $f.BaseName, $tmp)
        Remove-Item $tmp
    }

    # Compile first so syntax errors surface before any work runs.
    $acc.DoCmd.RunCommand(126)   # acCmdCompileAndSaveAllModules

    $a1 = [string]$src; $a2 = [string]$OutDir; $a3 = [bool]$UiOnly
    $result = $acc.Run('BuildAll', [ref]$a1, [ref]$a2, [ref]$a3)
    Write-Host "BuildAll: $result"
    if ($result -ne 'OK') { exit 1 }
}
finally {
    try { $acc.CloseCurrentDatabase() } catch {}
    $acc.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($acc) | Out-Null
}
Write-Host "Front end: $fe"
