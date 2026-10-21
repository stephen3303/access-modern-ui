# Checks the SQL side of ReleaseDesign.ps1 against a made-up schema, without Access:
#   powershell -NoProfile -File tools\ReleaseDesign.Tests.ps1
# (pwsh works too). Prints each check and exits 1 if any fails. Reading a real data
# file (Get-AccessSchema) and exporting forms need Access and run in release.ps1.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ReleaseDesign.ps1')

$failures = 0
function Check([string]$what, [bool]$ok) {
    if ($ok) { Write-Host "ok   $what" } else { Write-Host "FAIL $what"; $script:failures++ }
}
function F($name, $type, [int]$size = 0, [bool]$req = $false, [string]$def = '', [bool]$auto = $false, [bool]$azl = $true) {
    [pscustomobject]@{ name = $name; type = $type; size = $size; required = $req; allowZeroLength = $azl; default = $def; autoIncrement = $auto; precision = 0; scale = 0 }
}
function X($name, [bool]$primary, [bool]$unique, [string[]]$fields) {
    [pscustomobject]@{ name = $name; primary = $primary; unique = $unique; fields = @($fields | ForEach-Object { [pscustomobject]@{ name = $_.TrimEnd('-'); descending = $_.EndsWith('-') } }) }
}

Check 'FormDesign loaded alongside' ([bool](Get-Command Export-FormDesignFolder -ErrorAction SilentlyContinue))
Check 'text type with size' ((ConvertTo-SqlType (F 'a' 10 50)) -eq 'NVARCHAR(50)')
Check 'memo, date, money, guid' ((ConvertTo-SqlType (F 'a' 12)) -eq 'NVARCHAR(MAX)' -and (ConvertTo-SqlType (F 'a' 8)) -eq 'DATETIME2(0)' -and (ConvertTo-SqlType (F 'a' 5)) -eq 'MONEY' -and (ConvertTo-SqlType (F 'a' 15)) -eq 'UNIQUEIDENTIFIER')
$dec = F 'a' 20; $dec.precision = 9; $dec.scale = 2
Check 'decimal keeps precision' ((ConvertTo-SqlType $dec) -eq 'DECIMAL(9, 2)')
Check 'date/time extended is type 26, not dbNumeric 19' ((ConvertTo-SqlType (F 'a' 26)) -eq 'DATETIME2(7)' -and $null -eq (ConvertTo-SqlType (F 'a' 19)))
Check 'attachment has no type' ($null -eq (ConvertTo-SqlType (F 'a' 101)))
Check 'defaults converted' ((ConvertTo-SqlDefault (F 'a' 8 0 $false '=Now()')) -eq 'SYSDATETIME()' -and (ConvertTo-SqlDefault (F 'a' 8 0 $false 'Date()')) -eq 'CAST(GETDATE() AS date)' -and (ConvertTo-SqlDefault (F 'a' 4 0 $false '0')) -eq '0' -and (ConvertTo-SqlDefault (F 'a' 1 0 $false 'Yes')) -eq '1')
Check 'text default quoted and escaped' ((ConvertTo-SqlDefault (F 'a' 10 20 $false '"O''Brien ""x"""')) -eq "N'O''Brien ""x""'")
Check 'unknown expression not converted' ($null -eq (ConvertTo-SqlDefault (F 'a' 10 20 $false '=Environ("USERNAME")')))

$schema = [pscustomobject]@{
    tables = @(
        [pscustomobject]@{ name = 'items'; fields = @(
            (F 'id' 4 4 $false '' $true), (F 'title' 10 120 $true '' $false $false), (F 'code' 10 20), (F 'done' 1 1 $false 'False'),
            (F 'added' 8 8 $false '=Now()'), (F 'owner' 10 50 $false '=Environ("USERNAME")'), (F 'photo' 101), (F 'cat_id' 4 4))
            indexes = @((X 'PrimaryKey' $true $true @('id')), (X 'code' $false $true @('code')), (X 'by_added' $false $false @('added-', 'title'))) },
        [pscustomobject]@{ name = 'cats'; fields = @((F 'id' 4 4 $false '' $true), (F 'name]x' 10 40 $true))
            indexes = @((X 'PrimaryKey' $true $true @('id'))) })
    relations = @(
        [pscustomobject]@{ name = 'catsitems'; table = 'cats'; foreignTable = 'items'; enforced = $true; updateCascade = $false; deleteCascade = $true
                           fields = @([pscustomobject]@{ name = 'id'; foreignName = 'cat_id' }) },
        [pscustomobject]@{ name = 'loose'; table = 'cats'; foreignTable = 'items'; enforced = $false; updateCascade = $false; deleteCascade = $false
                           fields = @([pscustomobject]@{ name = 'id'; foreignName = 'cat_id' }) })
}
$sql = ConvertTo-TableSql -Schema $schema -Title 'Test v1.0.0' -Source 'test_be.accdb'
Check 'title and source' ($sql -match '-- Test v1\.0\.0' -and $sql -match 'Tables of test_be\.accdb')
Check 'autonumber as identity' ($sql -match '\[id\] INT IDENTITY\(1,1\) NOT NULL')
Check 'required text not empty' ($sql -match '\[title\] NVARCHAR\(120\) NOT NULL CHECK \(LEN\(\[title\]\) > 0\)')
Check 'optional text nullable' ($sql -match '\[code\] NVARCHAR\(20\) NULL')
Check 'yes/no not null with default' ($sql -match '\[done\] BIT NOT NULL CONSTRAINT \[DF_items_done\] DEFAULT \(0\)')
Check 'date default' ($sql -match '\[added\] DATETIME2\(0\) NULL CONSTRAINT \[DF_items_added\] DEFAULT \(SYSDATETIME\(\)\)')
Check 'unconverted default noted' ($sql -match '-- \[owner\]: Access default =Environ\("USERNAME"\) not converted\.')
Check 'attachment noted, not created' ($sql -match '-- \[photo\]: Access field type 101' -and $sql -notmatch '\[photo\] ')
Check 'primary key constraint' ($sql -match 'CONSTRAINT \[PK_items\] PRIMARY KEY \(\[id\]\)')
Check 'unique index filtered on nullable column' ($sql -match 'CREATE UNIQUE INDEX \[IX_items_code\] ON \[items\] \(\[code\]\) WHERE \[code\] IS NOT NULL;')
Check 'index with descending column' ($sql -match 'CREATE INDEX \[IX_items_by_added\] ON \[items\] \(\[added\] DESC, \[title\]\);')
Check 'bracket in name escaped' ($sql -match '\[name\]\]x\] NVARCHAR\(40\) NOT NULL')
Check 'foreign key with cascade' ($sql -match 'ALTER TABLE \[items\] ADD CONSTRAINT \[FK_catsitems\] FOREIGN KEY \(\[cat_id\]\) REFERENCES \[cats\] \(\[id\]\) ON DELETE CASCADE;')
Check 'unenforced relation as a comment' ($sql -match '-- ALTER TABLE \[items\] ADD CONSTRAINT \[FK_loose\]')
Check 'foreign keys after every table' ($sql.IndexOf('ALTER TABLE') -gt $sql.LastIndexOf('CREATE TABLE'))

if ($failures) { Write-Host "$failures check(s) failed"; exit 1 }
Write-Host 'All checks passed'
