# Checks ReleaseLock.ps1 against a stand-in for a DAO database, without Access:
#   powershell -NoProfile -File tools\ReleaseLock.Tests.ps1
# (pwsh works too). Prints each check and exits 1 if any fails. Opening a real .accde
# (Lock-ReleaseFrontEnd) needs Access; every app's release.ps1 runs it.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ReleaseLock.ps1')

$failures = 0
function Check([string]$Name, [bool]$Ok) {
    if ($Ok) { Write-Host "ok   $Name" } else { Write-Host "FAIL $Name"; $script:failures++ }
}

# A DAO database's Properties as DAO behaves: Item throws for a name it doesn't have,
# CreateProperty makes one (recording whether it is DDL), Append adds it. -Stuck names a
# property whose value never changes (a setting refused).
function New-FakeDatabase([hashtable]$Existing = @{}, [string]$Stuck = '') {
    $props = @{}
    foreach ($k in $Existing.Keys) { $props[$k] = [pscustomobject]@{ Name = $k; Value = $Existing[$k]; DDL = $false } }
    $collection = [pscustomobject]@{ Props = $props; Stuck = $Stuck }
    $collection | Add-Member ScriptMethod Item {
        param($n)
        if (-not $this.Props.ContainsKey($n)) { throw "Property not found (3270): $n" }
        $p = $this.Props[$n]
        if ($n -eq $this.Stuck) { return [pscustomobject]@{ Name = $n; Value = $p.Value; DDL = $p.DDL } }
        $p
    }
    $collection | Add-Member ScriptMethod Append { param($p) $this.Props[$p.Name] = $p }
    $db = [pscustomobject]@{ Properties = $collection }
    $db | Add-Member ScriptMethod CreateProperty {
        param($n, $type, $value, $ddl)
        [pscustomobject]@{ Name = $n; Value = $value; Type = $type; DDL = [bool]$ddl }
    }
    $db
}

# A new .accde has none of the three: each is created, as a Yes/No, set to False.
$db = New-FakeDatabase
Set-ReleaseLock $db
foreach ($n in 'AllowBypassKey', 'StartUpShowDBWindow', 'AllowSpecialKeys') {
    Check "created $n False" ($db.Properties.Props.ContainsKey($n) -and $db.Properties.Props[$n].Value -eq $false)
    Check "created $n as Yes/No" ($db.Properties.Props[$n].Type -eq 1)
}
Check 'AllowBypassKey is a DDL property' ($db.Properties.Props['AllowBypassKey'].DDL)

# A build that already set them (StartUpShowDBWindow from the app's start-up options,
# AllowSpecialKeys left True) is changed in place, not appended again.
$db = New-FakeDatabase @{ StartUpShowDBWindow = $true; AllowSpecialKeys = $true; AppTitle = 'X' }
Set-ReleaseLock $db
Check 'existing StartUpShowDBWindow set False' ($db.Properties.Props['StartUpShowDBWindow'].Value -eq $false)
Check 'existing AllowSpecialKeys set False' ($db.Properties.Props['AllowSpecialKeys'].Value -eq $false)
Check 'other properties left alone' ($db.Properties.Props['AppTitle'].Value -eq 'X')

# A value that doesn't stick stops the release.
$db = New-FakeDatabase @{ AllowBypassKey = $true } 'AllowBypassKey'
$threw = $false
try { Set-ReleaseLock $db } catch { $threw = $_.Exception.Message -like '*AllowBypassKey*not locked*' }
Check 'a setting that does not stick stops the release' $threw

# No file, no lock.
$threw = $false
try { Lock-ReleaseFrontEnd -Access ([pscustomobject]@{}) -Path (Join-Path $PSScriptRoot 'no-such-file.accde') } catch { $threw = $true }
Check 'a missing front end stops the release' $threw

if ($failures) { Write-Host "$failures check(s) failed"; exit 1 }
Write-Host 'All checks passed'
