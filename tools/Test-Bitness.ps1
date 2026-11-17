# Static check that the kit's VBA compiles and runs on both 32-bit and 64-bit Access.
# Needs no Access, so a cloud thread can run it before pushing. Copied from
# SABER-Access access\tools\Test-Bitness.ps1 (PR 75); keep the two in step.
#
#   tools\Test-Bitness.ps1                 check kit and tests
#   tools\Test-Bitness.ps1 -Path <dir>...  check other folders (.bas and .vba files)
#
# Exit code 0 when clean, 1 with one line per problem (file:line  rule  text).
# Rules:
#   PtrSafe     every Declare is PtrSafe (except in the #Else branch of #If VBA7,
#               which only Access 2007 and older compile)
#   Handle      a ByVal parameter named like a handle or pointer (hWnd, hdc, lpX, pX,
#               phX, pszX, pbX, pvX, ...Ptr) is LongPtr, not Long
#   Return      a Windows function that returns a handle or pointer is declared
#               As LongPtr (GetDC, SelectObject, GlobalAlloc, ...)
#   Variable    a Long variable is never given a handle (from one of those functions,
#               .hWnd or hWndAccessApp)
#   LongLong    LongLong and CLngLng appear only inside #If Win64 Then (they don't
#               exist in 32-bit VBA)
#   WindowLong  GetWindowLong / SetWindowLong (not ...Ptr) are declared only as the
#               32-bit branch of #If Win64
#   LenB        a structure size (cbSize, lStructSize, .cb) is set with LenB, not Len
#               (Len ignores the padding 64-bit VBA adds around LongPtr members)
# Windows PowerShell 5.1 compatible.
param([string[]]$Path)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not $Path) { $Path = @((Join-Path $root 'kit'), (Join-Path $root 'tests')) }

# Windows functions (A/W suffix dropped) that return a handle or pointer.
$handleApis = @('CreateWindowEx', 'CreateCompatibleDC', 'CreateCompatibleBitmap', 'CreateSolidBrush', 'CreateFont',
    'CreateFontIndirect', 'CreateDIBSection', 'CreateEvent', 'CreateFile', 'GetDC', 'GetWindowDC', 'GetWindow', 'GetFocus',
    'GetActiveWindow', 'GetForegroundWindow', 'GetParent', 'GetDesktopWindow', 'FindWindow', 'FindWindowEx', 'SelectObject',
    'MonitorFromWindow', 'MonitorFromRect', 'MonitorFromPoint', 'GlobalAlloc', 'GlobalLock', 'GlobalFree', 'LocalAlloc',
    'LocalFree', 'LoadLibrary', 'LoadLibraryEx', 'GetProcAddress', 'GetModuleHandle', 'RegisterEventSource',
    'GetCurrentProcess', 'GetCurrentThread', 'OpenProcess', 'SetClipboardData', 'GetClipboardData', 'SetFocus',
    'GetWindowLongPtr', 'SetWindowLongPtr', 'ShellExecute', 'SHBrowseForFolder', 'CopyImage', 'LoadImage')
$handleParam = '^(h[A-Z]\w*|hwnd|hdc|hWnd|lp[A-Z]\w*|p[A-Z]\w*|ph[A-Z]\w*|psz[A-Z]\w*|pb[A-Z]\w*|pv[A-Z]\w*|\w+Ptr)$'

$problems = New-Object System.Collections.Generic.List[string]
function Add-Problem([string]$File, [int]$Line, [string]$Rule, [string]$Text) {
    $problems.Add(('{0}:{1}  {2}  {3}' -f $File, $Line, $Rule, $Text.Trim()))
}
# True when a Windows function name (with or without its A/W suffix) is in $handleApis.
function Test-HandleApi([string]$Name) {
    if ($handleApis -contains $Name) { return $true }
    return ($handleApis -contains ($Name -creplace '(?<=[a-z])[AW]$', ''))
}

$files = @()
foreach ($p in $Path) { $files += @(Get-ChildItem -LiteralPath $p -Recurse -File | Where-Object { $_.Extension -eq '.bas' -or $_.Extension -eq '.vba' }) }
foreach ($f in $files) {
    $rel = $f.FullName
    if ($rel.StartsWith($root)) { $rel = $rel.Substring($root.Length).TrimStart('\', '/') }
    $raw = [IO.File]::ReadAllLines($f.FullName)
    # Join " _" continuations into logical lines, keeping the first physical line number.
    $lines = New-Object System.Collections.Generic.List[object]
    $buf = ''; $start = 0
    for ($i = 0; $i -lt $raw.Count; $i++) {
        $t = $raw[$i]
        if ($buf -eq '') { $start = $i + 1 }
        if ($t -match '\s_\s*$') { $buf += ($t -replace '\s_\s*$', ' '); continue }
        $lines.Add(@($start, ($buf + $t)))
        $buf = ''
    }
    # Conditional compilation: a stack of @(condition, inElse).
    $cond = New-Object System.Collections.Generic.List[object]
    # Names declared As Long in the current procedure (and at module level).
    $moduleLong = @{}; $procLong = @{}; $inProc = $false
    # Win32 handle APIs this file declares (by the name it calls them).
    $declared = @{}
    foreach ($entry in $lines) {
        $n = $entry[0]; $line = [string]$entry[1]
        # Drop a trailing comment (a ' outside a string).
        $code = $line
        $q = $false
        for ($k = 0; $k -lt $code.Length; $k++) {
            $ch = $code[$k]
            if ($ch -eq '"') { $q = -not $q }
            elseif ($ch -eq "'" -and -not $q) { $code = $code.Substring(0, $k); break }
        }
        $trim = $code.Trim()
        if ($trim -eq '') { continue }
        if ($trim -match '^#If\s+(.+?)\s+Then$') { $cond.Add(@($Matches[1].Trim(), $false)); continue }
        if ($trim -match '^#Else$' -and $cond.Count) { $cond[$cond.Count - 1] = @($cond[$cond.Count - 1][0], $true); continue }
        if ($trim -match '^#End If$' -and $cond.Count) { $cond.RemoveAt($cond.Count - 1); continue }
        $legacy = $false; $win64True = $false; $win64Else = $false
        foreach ($c in $cond) {
            if ($c[0] -match '^VBA7$' -and $c[1]) { $legacy = $true }
            if ($c[0] -match '^Win64$' -and -not $c[1]) { $win64True = $true }
            if ($c[0] -match '^Win64$' -and $c[1]) { $win64Else = $true }
        }
        if ($legacy) { continue }

        if ($trim -match '\b(LongLong|CLngLng)\b' -and -not $win64True) { Add-Problem $rel $n 'LongLong' $trim }
        if ($trim -match '\b(cbSize|lStructSize|\.cb)\s*=\s*Len\s*\(') { Add-Problem $rel $n 'LenB' $trim }

        if ($trim -match '^(Public\s+|Private\s+)?Declare\s' -and $trim -notmatch '^(Public\s+|Private\s+)?Declare\s+PtrSafe\s') {
            Add-Problem $rel $n 'PtrSafe' $trim
        }
        if ($trim -match '^(Public\s+|Private\s+)?Declare\s+(PtrSafe\s+)?(Function|Sub)\s+(\w+)\s+Lib\s+"[^"]*"(\s+Alias\s+"([^"]*)")?\s*\((.*)\)\s*(As\s+(\w+))?\s*$') {
            $kind = $Matches[3]; $name = $Matches[4]; $alias = $Matches[6]
            $params = $Matches[7]; $ret = $Matches[9]
            $real = if ($alias) { $alias } else { $name }
            if (($real -match '^(Get|Set)WindowLong[AW]?$') -and -not $win64Else) { Add-Problem $rel $n 'WindowLong' $trim }
            foreach ($prm in ($params -split ',')) {
                if ($prm -match '^\s*(Optional\s+)?ByVal\s+(\w+)\s+As\s+Long\s*$' -and $Matches[2] -cmatch $handleParam) {
                    Add-Problem $rel $n 'Handle' "$name($($prm.Trim()))"
                }
            }
            if ($kind -eq 'Function' -and (Test-HandleApi $real) -and $real -notmatch '^(Get|Set)WindowLong[AW]?$') {
                $declared[$name] = $true
                if ($ret -ne 'LongPtr') { Add-Problem $rel $n 'Return' $trim }
            }
            continue
        }

        if ($trim -match '^(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s') { $inProc = $true; $procLong = @{}; continue }
        if ($trim -match '^End\s+(Sub|Function|Property)$') { $inProc = $false; $procLong = @{}; continue }
        if ($trim -match '^(Dim|Private|Public|Static)\s+(.+)$' -and $trim -notmatch '^(Private|Public)\s+(Const|Declare|Type|Enum|Sub|Function|Property)\b') {
            foreach ($d in ($Matches[2] -split ',')) {
                if ($d -match '^\s*(WithEvents\s+)?(\w+)(\([^)]*\))?\s+As\s+Long\s*$') {
                    if ($inProc) { $procLong[$Matches[2].ToLower()] = $true } else { $moduleLong[$Matches[2].ToLower()] = $true }
                }
            }
            continue
        }
        if ($trim -match '^(Set\s+)?(\w+)(\(\w+\))?\s*=\s*(.+)$') {
            $var = $Matches[2].ToLower(); $rhs = $Matches[4]
            if ($procLong.ContainsKey($var) -or (-not $procLong.ContainsKey($var) -and $moduleLong.ContainsKey($var))) {
                $hit = $false
                foreach ($api in $declared.Keys) { if ($rhs -match ('^\s*' + [regex]::Escape($api) + '\s*\(')) { $hit = $true } }
                if ($rhs -match '\.hWnd\s*$|hWndAccessApp') { $hit = $true }
                if ($hit) { Add-Problem $rel $n 'Variable' $trim }
            }
        }
    }
}

if ($problems.Count) {
    Write-Host "Bitness check: $($problems.Count) problem(s) that break 32-bit or 64-bit Access:"
    $problems | ForEach-Object { Write-Host "  $_" }
    exit 1
}
Write-Host "Bitness check: $($files.Count) files clean."
exit 0
