# Release hand-off files: the design of every form as JSON and a SQL Server script that
# recreates the data file's tables, for rebuilding a released app elsewhere (another
# app, HTML pages, SQL Server). An app's release.ps1 calls these; they are made only for
# a release, never by a normal build.
#
#   . "$PSScriptRoot\tools\ReleaseDesign.ps1"
#   Export-ReleaseFormDesigns -FrontEnd $accdb -OutDir "$out\design\forms" -ThemeFile $theme
#   Export-ReleaseTableSql -DataFile $data -OutFile "$out\design\tables.sql" -Title 'SABER v1.3.0'
#
# Needs Microsoft Access (both open their own hidden instance). The form JSON is written
# by FormDesign.ps1, which an app keeps next to this file with FormSnapshot.ps1 (the kit
# keeps them in tools\review-mcp); see that file for what a design holds.

foreach ($dir in $PSScriptRoot, (Join-Path $PSScriptRoot 'review-mcp')) {
    if (Test-Path (Join-Path $dir 'FormDesign.ps1')) { . (Join-Path $dir 'FormDesign.ps1'); break }
}

# Saves every form of a built front end (the .accdb, before it is compiled to .accde)
# with SaveAsText and writes the JSON designs to OutDir. The forms are copied into a
# blank database first and nothing in it is allowed to run, so the app's start-up code
# (AutoExec, a sign-in form) never runs. Returns the forms written; throws when any form
# couldn't be saved or converted, so a release never ships a partial set.
function Export-ReleaseFormDesigns {
    param([Parameter(Mandatory)][string]$FrontEnd, [Parameter(Mandatory)][string]$OutDir, [string]$ThemeFile)
    if (-not (Get-Command Export-FormDesignFolder -ErrorAction SilentlyContinue)) { throw "FormDesign.ps1 is not next to $PSCommandPath." }
    $work = Join-Path ([IO.Path]::GetTempPath()) ('release-design-' + [guid]::NewGuid().ToString('N'))
    $snap = Join-Path $work 'snapshot'
    New-Item -ItemType Directory -Force $snap | Out-Null
    $acc = New-Object -ComObject Access.Application
    $watch = $null
    if (Get-Command Start-AccessWatch -ErrorAction SilentlyContinue) { $watch = Start-AccessWatch $acc -TimeoutSec 900 }
    try {
        $acc.Visible = $false
        $acc.AutomationSecurity = 3      # force disable: nothing in the copied forms runs
        $acc.NewCurrentDatabase([string](Join-Path $work 'export.accdb'))
        $db = $acc.DBEngine.OpenDatabase([string]$FrontEnd, $false, $true)
        $names = @(foreach ($d in $db.Containers.Item('Forms').Documents) { [string]$d.Name })
        $db.Close()
        foreach ($n in $names) {
            try {
                $acc.DoCmd.TransferDatabase(0, 'Microsoft Access', [string]$FrontEnd, 2, $n, $n)   # acImport, acForm
                $acc.SaveAsText(2, $n, [string](Join-Path $snap "$n.txt"))
            } catch {
                [IO.File]::WriteAllText((Join-Path $snap "$n.failed"), $_.Exception.Message)
            }
        }
        $acc.CloseCurrentDatabase()
    } finally {
        if ($watch) { [void](Stop-AccessWatch $watch $acc) } else { try { $acc.Quit(2) } catch {} }
    }
    try { $r = Export-FormDesignFolder -SnapshotDir $snap -OutDir $OutDir -ThemeFile $ThemeFile }
    finally { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
    if (@($r.Failed).Count) { throw "Could not export these forms:`r`n  " + ((@($r.Failed) | ForEach-Object { "$($_.form): $($_.error)" }) -join "`r`n  ") }
    $r
}

# The local tables of an Access data file as plain objects (no COM objects left in
# them), for ConvertTo-TableSql: each table's fields in order, its indexes and the
# relationships between tables. System, temporary and linked tables are left out.
function Get-AccessSchema {
    param([Parameter(Mandatory)][string]$DataFile)
    $acc = New-Object -ComObject Access.Application
    try {
        $db = $acc.DBEngine.OpenDatabase([string]$DataFile, $false, $true)
        $tables = @(foreach ($t in $db.TableDefs) {
            $tn = [string]$t.Name
            if ($tn -like 'MSys*' -or $tn -like '~*' -or $tn -like 'USys*' -or ($t.Attributes -band 0x80000002) -or [string]$t.Connect) { continue }
            $fields = @(foreach ($f in $t.Fields) {
                $o = [ordered]@{ name = [string]$f.Name; type = [int]$f.Type; size = [int]$f.Size; required = [bool]$f.Required
                                 allowZeroLength = $false; default = [string]$f.DefaultValue; autoIncrement = [bool]($f.Attributes -band 16)
                                 precision = 0; scale = 0; expression = '' }
                try { $o.allowZeroLength = [bool]$f.AllowZeroLength } catch {}
                try { $o.expression = [string]$f.Properties.Item('Expression').Value } catch {}   # calculated field
                if ($o.type -eq 20) {
                    try { $o.precision = [int]$f.Properties.Item('Precision').Value; $o.scale = [int]$f.Properties.Item('Scale').Value } catch {}
                }
                [pscustomobject]$o
            })
            $indexes = @(foreach ($ix in $t.Indexes) {
                if ($ix.Foreign) { continue }   # made by a relationship; the foreign key covers it
                [pscustomobject]@{ name = [string]$ix.Name; primary = [bool]$ix.Primary; unique = [bool]$ix.Unique
                                   fields = @(foreach ($xf in $ix.Fields) { [pscustomobject]@{ name = [string]$xf.Name; descending = [bool]($xf.Attributes -band 1) } }) }
            })
            [pscustomobject]@{ name = $tn; fields = $fields; indexes = $indexes }
        })
        $names = @($tables | ForEach-Object { $_.name })
        $relations = @(foreach ($r in $db.Relations) {
            if ($names -notcontains [string]$r.Table -or $names -notcontains [string]$r.ForeignTable) { continue }
            [pscustomobject]@{ name = [string]$r.Name; table = [string]$r.Table; foreignTable = [string]$r.ForeignTable
                               enforced = -not ($r.Attributes -band 2); updateCascade = [bool]($r.Attributes -band 256); deleteCascade = [bool]($r.Attributes -band 4096)
                               fields = @(foreach ($rf in $r.Fields) { [pscustomobject]@{ name = [string]$rf.Name; foreignName = [string]$rf.ForeignName } }) }
        })
        $db.Close()
        [pscustomobject]@{ tables = @($tables | Sort-Object name); relations = @($relations | Sort-Object name) }
    } finally { try { $acc.Quit(2) } catch {} }
}

function Format-SqlName([string]$Name) { '[' + $Name.Replace(']', ']]') + ']' }

# A constraint or index name made from parts, kept to letters, digits and _.
function Format-SqlConstraint([string[]]$Parts) { Format-SqlName (($Parts -join '_') -replace '[^A-Za-z0-9_]', '_') }

# The SQL Server type for a DAO field type. Attachment and multi-value fields (101 and
# up) have no column equivalent and return $null.
function ConvertTo-SqlType($Field) {
    switch ($Field.type) {
        1 { 'BIT' }                     # Yes/No
        2 { 'TINYINT' }                 # Byte
        3 { 'SMALLINT' }                # Integer
        4 { 'INT' }                     # Long Integer
        5 { 'MONEY' }                   # Currency
        6 { 'REAL' }                    # Single
        7 { 'FLOAT' }                   # Double
        8 { 'DATETIME2(0)' }            # Date/Time (whole seconds, as Access keeps them)
        9 { "VARBINARY($([Math]::Max(1, $Field.size)))" }
        10 { "NVARCHAR($([Math]::Max(1, $Field.size)))" }   # Short Text
        11 { 'VARBINARY(MAX)' }         # OLE Object
        12 { 'NVARCHAR(MAX)' }          # Long Text (Memo, Hyperlink)
        15 { 'UNIQUEIDENTIFIER' }       # Replication ID
        16 { 'BIGINT' }                 # Large Number
        26 { 'DATETIME2(7)' }           # Date/Time Extended (dbDateTime2; 19 is dbNumeric)
        20 { if ($Field.precision -gt 0) { "DECIMAL($($Field.precision), $($Field.scale))" } else { 'DECIMAL(18, 0)' } }
        default { $null }
    }
}

# An Access default value as a SQL Server one, or $null when it can't be carried over
# (an expression that calls a function SQL Server doesn't have).
function ConvertTo-SqlDefault($Field) {
    $d = ([string]$Field.default).Trim()
    if ($d.StartsWith('=')) { $d = $d.Substring(1).Trim() }
    if (-not $d) { return $null }
    if ($Field.type -eq 1) {
        if ($d -match '^(True|Yes|On|-1)$') { return '1' }
        if ($d -match '^(False|No|Off|0)$') { return '0' }
        return $null
    }
    if ($d -match '^Now\(\)$') { return 'SYSDATETIME()' }
    if ($d -match '^Date\(\)$') { return 'CAST(GETDATE() AS date)' }
    if ($d -match '^-?\d+(\.\d+)?$') { return $d }
    if ($d -match '^"(.*)"$' -or $d -match "^'(.*)'$") { return "N'" + $Matches[1].Replace('""', '"').Replace("'", "''") + "'" }
    $null
}

# The schema (Get-AccessSchema) as a T-SQL script: CREATE TABLE for each table with its
# primary key and defaults, then its other indexes, then the relationships as foreign
# keys (at the end, so table order doesn't matter). What can't be carried over is
# written as a comment where it would have gone.
function ConvertTo-TableSql {
    param([Parameter(Mandatory)]$Schema, [string]$Title, [string]$Source)
    $sb = New-Object Text.StringBuilder
    $line = { param($s) [void]$sb.AppendLine($s) }
    if ($Title) { & $line "-- $Title" }
    if ($Source) { & $line "-- Tables of $Source, for SQL Server (T-SQL). Generated by ReleaseDesign.ps1." }
    & $line '-- A unique index on a column that may be empty becomes a filtered index (WHERE ... IS NOT NULL),'
    & $line '-- because Access allows any number of empty values in a unique index and SQL Server allows one.'
    & $line ''
    foreach ($t in $Schema.tables) {
        $cols = New-Object System.Collections.Generic.List[string]
        $notes = New-Object System.Collections.Generic.List[string]
        foreach ($f in $t.fields) {
            if ($f.expression) { $notes.Add("-- $(Format-SqlName $f.name): calculated field ($($f.expression)) not created; add it as a computed column if needed."); continue }
            $type = ConvertTo-SqlType $f
            if (-not $type) { $notes.Add("-- $(Format-SqlName $f.name): Access field type $($f.type) (attachment or multi-value) has no SQL Server column; not created."); continue }
            $c = "    $(Format-SqlName $f.name) $type"
            if ($f.autoIncrement) { $c += ' IDENTITY(1,1)' }
            # Access keeps Yes/No fields filled (never empty), so they are NOT NULL.
            $c += $(if ($f.required -or $f.autoIncrement -or $f.type -eq 1) { ' NOT NULL' } else { ' NULL' })
            $def = if ($f.autoIncrement) { $null } else { ConvertTo-SqlDefault $f }
            if ($def) { $c += " CONSTRAINT $(Format-SqlConstraint 'DF', $t.name, $f.name) DEFAULT ($def)" }
            elseif ($f.type -eq 1 -and -not $f.autoIncrement) { $c += " CONSTRAINT $(Format-SqlConstraint 'DF', $t.name, $f.name) DEFAULT (0)" }
            elseif ([string]$f.default -and -not $f.autoIncrement) { $notes.Add("-- $(Format-SqlName $f.name): Access default $($f.default) not converted.") }
            if ($f.type -eq 10 -and $f.required -and -not $f.allowZeroLength) { $c += " CHECK (LEN($(Format-SqlName $f.name)) > 0)" }
            $cols.Add($c)
        }
        $pk = @($t.indexes | Where-Object { $_.primary }) | Select-Object -First 1
        if ($pk) { $cols.Add("    CONSTRAINT $(Format-SqlConstraint 'PK', $t.name) PRIMARY KEY ($((@($pk.fields | ForEach-Object { Format-SqlName $_.name })) -join ', '))") }
        foreach ($n in $notes) { & $line $n }
        & $line "CREATE TABLE $(Format-SqlName $t.name) ("
        & $line (($cols -join ("," + [Environment]::NewLine)))
        & $line ');'
        foreach ($ix in @($t.indexes | Where-Object { -not $_.primary })) {
            $cols2 = @($ix.fields | ForEach-Object { (Format-SqlName $_.name) + $(if ($_.descending) { ' DESC' } else { '' }) }) -join ', '
            $kind = if ($ix.unique) { 'CREATE UNIQUE INDEX' } else { 'CREATE INDEX' }
            $where = ''
            if ($ix.unique) {
                $nullable = @(foreach ($xf in $ix.fields) {
                    $fd = @($t.fields | Where-Object { $_.name -eq $xf.name }) | Select-Object -First 1
                    if ($fd -and -not ($fd.required -or $fd.autoIncrement -or $fd.type -eq 1)) { (Format-SqlName $xf.name) + ' IS NOT NULL' }
                })
                if ($nullable.Count) { $where = ' WHERE ' + ($nullable -join ' AND ') }
            }
            & $line "$kind $(Format-SqlConstraint 'IX', $t.name, $ix.name) ON $(Format-SqlName $t.name) ($cols2)$where;"
        }
        & $line ''
    }
    foreach ($r in $Schema.relations) {
        $fk = @($r.fields | ForEach-Object { Format-SqlName $_.foreignName }) -join ', '
        $pkc = @($r.fields | ForEach-Object { Format-SqlName $_.name }) -join ', '
        $stmt = "ALTER TABLE $(Format-SqlName $r.foreignTable) ADD CONSTRAINT $(Format-SqlConstraint 'FK', $r.name) FOREIGN KEY ($fk) REFERENCES $(Format-SqlName $r.table) ($pkc)"
        if ($r.deleteCascade) { $stmt += ' ON DELETE CASCADE' }
        if ($r.updateCascade) { $stmt += ' ON UPDATE CASCADE' }
        if (-not $r.enforced) { & $line "-- Not enforced in Access (no referential integrity); kept as a comment:"; & $line "-- $stmt;" }
        else { & $line "$stmt;" }
    }
    $sb.ToString()
}

# Writes the T-SQL script for a data file's tables. Returns the table and relationship
# counts.
function Export-ReleaseTableSql {
    param([Parameter(Mandatory)][string]$DataFile, [Parameter(Mandatory)][string]$OutFile, [string]$Title)
    $schema = Get-AccessSchema -DataFile $DataFile
    $sql = ConvertTo-TableSql -Schema $schema -Title $Title -Source ([IO.Path]::GetFileName($DataFile))
    New-Item -ItemType Directory -Force (Split-Path $OutFile) | Out-Null
    [IO.File]::WriteAllText($OutFile, $sql, (New-Object Text.UTF8Encoding($false)))
    [pscustomobject]@{ File = $OutFile; Tables = @($schema.tables).Count; Relations = @($schema.relations).Count }
}
