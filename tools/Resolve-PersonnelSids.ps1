<#
.SYNOPSIS
    Adds each person's Windows account and SID to a people list, from their
    UPN or DoD ID, for the Personnel page's Import.

.DESCRIPTION
    The apps' Import looks people up in Active Directory by itself. Use this
    script instead when you'd rather prepare the file first, check it, or
    import on a PC that can't reach the directory.

    For every row it takes the user principal name (a "User principal name"
    or "UPN" column), or makes one from a "DoD ID" or "EDIPI" column
    (<DoD ID>.A@smil.mil by default), and asks Active Directory for that
    user: Get-ADUser when the ActiveDirectory module (RSAT) is installed,
    otherwise Windows' own name lookup, which needs no RSAT. It writes the
    rows to -OutFile with "Windows account" (DOMAIN\user) and "Windows SID"
    filled in and the UPN and DoD ID columns left out, so the DoD ID never
    reaches the Personnel file. Rows it can't find are listed at the end and
    written with the two columns blank.

.PARAMETER Path
    The people list: a CSV, such as an Outlook contacts export with a DoD ID
    or UPN column added.

.PARAMETER OutFile
    Where to write the result. Default: beside Path, named "<name> - linked.csv".

.PARAMETER UpnSuffix
    What follows a DoD ID to make its UPN. Default ".A@smil.mil".

.EXAMPLE
    .\Resolve-PersonnelSids.ps1 -Path C:\Temp\people.csv
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $Path,
    [string] $OutFile,
    [string] $UpnSuffix = '.A@smil.mil'
)
$ErrorActionPreference = 'Stop'

function Get-Key([string] $s) { ($s -replace '[\s_-]', '').ToLowerInvariant() }

# The first column whose header matches one of the keys, or $null.
function Find-Column($headers, [string[]] $keys) {
    foreach ($k in $keys) {
        $h = $headers | Where-Object { (Get-Key $_) -eq $k } | Select-Object -First 1
        if ($h) { return $h }
    }
    $null
}

$rows = @(Import-Csv -LiteralPath $Path)
if ($rows.Count -eq 0) { Write-Warning "$Path has no rows."; return }
$headers = $rows[0].PSObject.Properties.Name
$upnCol = Find-Column $headers @('userprincipalname', 'upn')
$idCol = Find-Column $headers @('dodid', 'edipi', 'dodidnumber')
if (-not $upnCol -and -not $idCol) { throw 'The file needs a "User principal name" (UPN) or "DoD ID" column.' }
$haveRsat = [bool](Get-Module -ListAvailable ActiveDirectory)
if ($haveRsat) { Import-Module ActiveDirectory }

if (-not $OutFile) {
    $OutFile = Join-Path (Split-Path -Parent (Resolve-Path -LiteralPath $Path)) `
        ([IO.Path]::GetFileNameWithoutExtension($Path) + ' - linked.csv')
}
$keep = $headers | Where-Object { $_ -ne $upnCol -and $_ -ne $idCol -and (Get-Key $_) -notin @('windowsaccount', 'windowssid') }
$missing = New-Object System.Collections.Generic.List[string]
$n = 1
$out = foreach ($r in $rows) {
    $n++
    $upn = if ($upnCol) { "$($r.$upnCol)".Trim() } else { '' }
    if (-not $upn -and $idCol) {
        $id = "$($r.$idCol)".Trim()
        if ($id -match '^\d{10}$') { $upn = $id + $UpnSuffix }
    }
    $account = ''; $sid = ''
    if ($upn) {
        try {
            if ($haveRsat) {
                $u = Get-ADUser -Filter "UserPrincipalName -eq '$($upn -replace "'", "''")'" -Properties SID
                if ($u) { $sid = $u.SID.Value }
            } else {
                $sid = ([Security.Principal.NTAccount] $upn).Translate([Security.Principal.SecurityIdentifier]).Value
            }
            if ($sid) {
                $account = ([Security.Principal.SecurityIdentifier] $sid).Translate([Security.Principal.NTAccount]).Value
            }
        } catch { $account = ''; $sid = '' }
    }
    if (-not $sid) { $missing.Add("Row $n") }
    $o = [ordered]@{}
    foreach ($h in $keep) { $o[$h] = $r.$h }
    $o['Windows account'] = $account
    $o['Windows SID'] = $sid
    [pscustomobject] $o
}
$out | Export-Csv -LiteralPath $OutFile -NoTypeInformation -Encoding UTF8
Write-Host "Wrote $($rows.Count) rows to $OutFile."
if ($missing.Count -gt 0) {
    Write-Warning "Not found in Active Directory ($($missing.Count)): $($missing -join ', ')"
}
