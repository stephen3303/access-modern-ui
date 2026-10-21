# Release design files: what a release build writes beside the locked front end so the
# app can be rebuilt or ported without opening Access.
#
#   design\forms\<form>.json   each form of the front end as a JSON design
#                              (FormDesign.ps1, schema access-form-design/1), with
#                              index.json and, given the build palette, tokens.json
#   design\<database>.sql      a build script for each database's own tables: CREATE
#                              TABLE, CREATE INDEX and ALTER TABLE ... FOREIGN KEY
#
# Only a release build makes these (release.ps1 calls Export-ReleaseDesign before it
# compiles the front end), so they always match a tagged version. They are not kept in
# the repository.
#
# Dot-source it next to FormDesign.ps1 and FormSnapshot.ps1 (in this repository they are
# in review-mcp\; an app's copy keeps all three in one folder):
#
#   . "$PSScriptRoot\tools\ReleaseDesign.ps1"
#   Export-ReleaseDesign -Access $acc -FrontEnd $fe -DataFiles $be -ThemeFile $theme -OutDir "$out\design"
#
# The SQL is Access SQL in ANSI-92 mode, the mode ADO uses (CurrentProject.Connection,
# or an OLE DB connection to the file), which DEFAULT needs. Run each statement on its
# own, in order: Access runs one statement at a time. Each script is checked by running
# it into an empty database and reading that back (Test-SchemaSqlRebuild); the release
# stops if the rebuilt tables differ.
#
# Converting a schema to SQL needs no Access, so it can be tested anywhere PowerShell
# runs: see ReleaseDesign.Tests.ps1.

$script:ReleaseDesignDir = $PSScriptRoot
if (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'FormDesign.ps1')) { . (Join-Path $PSScriptRoot 'FormDesign.ps1') }
else { . (Join-Path $PSScriptRoot 'review-mcp\FormDesign.ps1') }

# Tables that are Access's own, not the app's: system tables, temporary tables and the
# hidden tables behind multi-value and attachment fields.
function Test-SchemaSystemTable([string]$Name) {
    $Name -match '^(MSys|USysRibbons$|~)' -or $Name -match '^f_[0-9A-Fa-f]{32}_'
}

# Reads a database's own tables (not linked ones), their indexes and the relationships
# between them through DAO into plain objects, for ConvertTo-SchemaSql. Opens the file
# read-only with the DBEngine of an Access.Application, so no start-up code runs.
function Read-AccessSchema {
    param([Parameter(Mandatory)]$Access, [Parameter(Mandatory)][string]$Path)
    $db = $Access.DBEngine.OpenDatabase([string]$Path, $false, $true)
    try {
        $tables = New-Object System.Collections.Generic.List[object]
        $tds = $db.TableDefs
        for ($i = 0; $i -lt $tds.Count; $i++) {
            $t = $tds.Item($i)
            $name = [string]$t.Name
            if ((Test-SchemaSystemTable $name) -or [string]$t.Connect) { continue }
            $fields = New-Object System.Collections.Generic.List[object]
            for ($j = 0; $j -lt $t.Fields.Count; $j++) {
                $f = $t.Fields.Item($j)
                $expr = ''
                try { $expr = [string]$f.Properties.Item('Expression').Value } catch {}
                $precision = $null; $scale = $null
                if ([int]$f.Type -eq 20) {
                    try { $precision = [int]$f.Properties.Item('Precision').Value; $scale = [int]$f.Properties.Item('Scale').Value } catch {}
                }
                $fields.Add([pscustomobject]@{
                    Name = [string]$f.Name; Ordinal = [int]$f.OrdinalPosition; Type = [int]$f.Type; Size = [int]$f.Size
                    Required = [bool]$f.Required; Default = [string]$f.DefaultValue
                    AutoIncrement = (([int]$f.Attributes -band 16) -ne 0)   # dbAutoIncrField
                    Expression = $expr; Precision = $precision; Scale = $scale
                })
            }
            $indexes = New-Object System.Collections.Generic.List[object]
            for ($j = 0; $j -lt $t.Indexes.Count; $j++) {
                $ix = $t.Indexes.Item($j)
                $cols = New-Object System.Collections.Generic.List[object]
                for ($k = 0; $k -lt $ix.Fields.Count; $k++) {
                    $c = $ix.Fields.Item($k)
                    $cols.Add([pscustomobject]@{ Name = [string]$c.Name; Descending = (([int]$c.Attributes -band 1) -ne 0) })   # dbDescending
                }
                $indexes.Add([pscustomobject]@{
                    Name = [string]$ix.Name; Primary = [bool]$ix.Primary; Unique = [bool]$ix.Unique
                    Required = [bool]$ix.Required; IgnoreNulls = [bool]$ix.IgnoreNulls; Foreign = [bool]$ix.Foreign
                    Fields = $cols.ToArray()
                })
            }
            $tables.Add([pscustomobject]@{ Name = $name; Fields = @($fields | Sort-Object Ordinal); Indexes = $indexes.ToArray() })
        }
        $relations = New-Object System.Collections.Generic.List[object]
        $rels = $db.Relations
        for ($i = 0; $i -lt $rels.Count; $i++) {
            $r = $rels.Item($i)
            $pairs = New-Object System.Collections.Generic.List[object]
            for ($k = 0; $k -lt $r.Fields.Count; $k++) {
                $c = $r.Fields.Item($k)
                $pairs.Add([pscustomobject]@{ Name = [string]$c.Name; ForeignName = [string]$c.ForeignName })
            }
            $relations.Add([pscustomobject]@{
                Name = [string]$r.Name; Table = [string]$r.Table; ForeignTable = [string]$r.ForeignTable
                Attributes = [int]$r.Attributes; Fields = $pairs.ToArray()
            })
        }
        [pscustomobject]@{ Tables = $tables.ToArray(); Relations = $relations.ToArray() }
    } finally { $db.Close() }
}

function Format-SqlName([string]$Name) { "[$Name]" }

# A field's type as ANSI-92 Access SQL, or $null with -Reason set when SQL can't make it.
function ConvertTo-SqlType($Field, [ref]$Reason) {
    if ($Field.Expression) { $Reason.Value = 'calculated field'; return $null }
    switch ($Field.Type) {
        1 { 'YESNO' }
        2 { 'BYTE' }
        3 { 'SMALLINT' }
        4 { if ($Field.AutoIncrement) { 'COUNTER' } else { 'INTEGER' } }
        5 { 'CURRENCY' }
        6 { 'REAL' }
        7 { 'FLOAT' }
        8 { 'DATETIME' }
        9 { "BINARY($($Field.Size))" }
        10 { "TEXT($($Field.Size))" }
        11 { 'LONGBINARY' }
        12 { 'LONGTEXT' }
        15 { 'UNIQUEIDENTIFIER' }
        16 { 'BIGINT' }
        20 {
            if ($null -eq $Field.Precision) { $Reason.Value = 'decimal field whose precision DAO does not report'; return $null }
            "DECIMAL($($Field.Precision),$($Field.Scale))"
        }
        default {
            # 101 attachment, 102-109 multi-value lookups: DDL can't create them.
            $Reason.Value = if ($Field.Type -eq 101) { 'attachment field' } elseif ($Field.Type -gt 101) { 'multi-value field' } else { "field type $($Field.Type)" }
            $null
        }
    }
}

# A field's DefaultValue (as DAO returns it) as an ANSI-92 DEFAULT. Text in either kind
# of quotes becomes single-quoted, Yes/No words become -1/0 and a leading = is dropped,
# so a default reads the same however it was set.
function ConvertTo-SqlDefault([string]$Value) {
    $v = $Value.Trim()
    if ($v.StartsWith('=')) { $v = $v.Substring(1).Trim() }
    if (-not $v) { return $null }
    if ($v -match '^"((?:[^"]|"")*)"$') { return "'" + ($Matches[1].Replace('""', '"').Replace("'", "''")) + "'" }
    if ($v -match "^'((?:[^']|'')*)'$") { return $v }
    if ($v -match '^(?i)(yes|true|on)$') { return '-1' }
    if ($v -match '^(?i)(no|false|off)$') { return '0' }
    $v
}

# Turns Read-AccessSchema's objects into a build script. Returns Text (the whole script,
# with a header saying what it leaves out), Statements (each statement on its own, to
# run in order) and Skipped (fields, indexes and relationships SQL can't make).
# Tables, indexes and relationships are sorted by name so the same schema always gives
# the same script.
function ConvertTo-SchemaSql {
    param([Parameter(Mandatory)]$Schema, [string]$Title = 'database')
    $statements = New-Object System.Collections.Generic.List[string]
    $skipped = New-Object System.Collections.Generic.List[string]
    $made = @{}      # table -> names of the fields the script makes
    $tables = @($Schema.Tables | Sort-Object Name)
    foreach ($t in $tables) {
        $cols = New-Object System.Collections.Generic.List[string]
        $made[$t.Name] = @{}
        foreach ($f in $t.Fields) {
            $why = $null
            $type = ConvertTo-SqlType $f ([ref]$why)
            if (-not $type) { $skipped.Add("$(Format-SqlName $t.Name).$(Format-SqlName $f.Name): $why"); continue }
            $col = "    $(Format-SqlName $f.Name) $type"
            # An AutoNumber's default is the next number; DAO reports none.
            $def = if ($f.AutoIncrement) { $null } else { ConvertTo-SqlDefault $f.Default }
            if ($def) { $col += " DEFAULT $def" }
            if ($f.Required) { $col += ' NOT NULL' }
            $cols.Add($col)
            $made[$t.Name][$f.Name] = $true
        }
        if (-not $cols.Count) { $skipped.Add("$(Format-SqlName $t.Name): no field SQL can make"); continue }
        $statements.Add("CREATE TABLE $(Format-SqlName $t.Name) (`r`n" + ($cols -join ",`r`n") + "`r`n)")
    }
    foreach ($t in $tables) {
        if (-not $made.ContainsKey($t.Name) -or -not $made[$t.Name].Count) { continue }
        # The primary key first, so a foreign key can point at it.
        foreach ($ix in @($t.Indexes | Where-Object { -not $_.Foreign } | Sort-Object @{ Expression = { -not $_.Primary } }, Name)) {
            $missing = @($ix.Fields | Where-Object { -not $made[$t.Name].ContainsKey($_.Name) })
            if ($missing.Count) { $skipped.Add("index $(Format-SqlName $ix.Name) on $(Format-SqlName $t.Name): uses a field SQL can't make"); continue }
            $cols = ($ix.Fields | ForEach-Object { (Format-SqlName $_.Name) + $(if ($_.Descending) { ' DESC' } else { '' }) }) -join ', '
            $unique = if ($ix.Unique -and -not $ix.Primary) { 'UNIQUE ' } else { '' }
            # One WITH option at most: PRIMARY covers the other two.
            $with = if ($ix.Primary) { ' WITH PRIMARY' } elseif ($ix.Required) { ' WITH DISALLOW NULL' } elseif ($ix.IgnoreNulls) { ' WITH IGNORE NULL' } else { '' }
            $statements.Add("CREATE ${unique}INDEX $(Format-SqlName $ix.Name) ON $(Format-SqlName $t.Name) ($cols)$with")
        }
    }
    foreach ($r in @($Schema.Relations | Sort-Object Name)) {
        if ($r.Attributes -band 4) { continue }   # dbRelationInherited: belongs to a linked database
        if ((Test-SchemaSystemTable $r.Table) -or (Test-SchemaSystemTable $r.ForeignTable)) { continue }
        $label = "relationship $(Format-SqlName $r.Name) ($(Format-SqlName $r.ForeignTable) to $(Format-SqlName $r.Table))"
        if (-not $made.ContainsKey($r.Table) -or -not $made.ContainsKey($r.ForeignTable)) { $skipped.Add("${label}: a table is not in this database"); continue }
        if ($r.Attributes -band 2) { $skipped.Add("${label}: not enforced, which SQL can't make"); continue }   # dbRelationDontEnforce
        $missing = @($r.Fields | Where-Object { -not $made[$r.Table].ContainsKey($_.Name) -or -not $made[$r.ForeignTable].ContainsKey($_.ForeignName) })
        if ($missing.Count) { $skipped.Add("${label}: uses a field SQL can't make"); continue }
        $fk = ($r.Fields | ForEach-Object { Format-SqlName $_.ForeignName }) -join ', '
        $pk = ($r.Fields | ForEach-Object { Format-SqlName $_.Name }) -join ', '
        $s = "ALTER TABLE $(Format-SqlName $r.ForeignTable) ADD CONSTRAINT $(Format-SqlName $r.Name) FOREIGN KEY ($fk) REFERENCES $(Format-SqlName $r.Table) ($pk)"
        if ($r.Attributes -band 256) { $s += ' ON UPDATE CASCADE' }    # dbRelationUpdateCascade
        if ($r.Attributes -band 4096) { $s += ' ON DELETE CASCADE' }   # dbRelationDeleteCascade
        $statements.Add($s)
    }
    $head = @(
        "-- Tables of $Title as Access SQL (ANSI-92, the mode ADO uses).",
        '-- Run each statement on its own, in order, in an empty database, for example with',
        '-- CurrentProject.Connection.Execute. Statements end with a semicolon.',
        '-- Not included: data, captions, descriptions, formats, input masks, lookups,',
        '-- validation rules, Allow Zero Length and Unicode compression.'
    )
    if ($skipped.Count) {
        $head += '--'
        $head += "-- Not made by this script ($($skipped.Count)):"
        $head += @($skipped | ForEach-Object { "--   $_" })
    }
    $text = (($head -join "`r`n") + "`r`n`r`n" + (($statements | ForEach-Object { "$_;" }) -join "`r`n`r`n") + "`r`n")
    [pscustomobject]@{ Text = $text; Statements = $statements.ToArray(); Skipped = $skipped.ToArray() }
}

# Runs a build script's statements into a new, empty database in -WorkDir, reads that
# database back and returns the statements that differ (in the script but not rebuilt,
# '-', or rebuilt but not in the script, '+'). Nothing returned means the script
# rebuilds the tables exactly as far as SQL can describe them. Closes any database the
# Access instance has open first.
function Test-SchemaSqlRebuild {
    param([Parameter(Mandatory)]$Access, [Parameter(Mandatory)]$Sql, [Parameter(Mandatory)][string]$WorkDir)
    $tmp = Join-Path $WorkDir ('rebuild-' + [guid]::NewGuid().ToString('N') + '.accdb')
    try { $Access.CloseCurrentDatabase() } catch {}
    $Access.NewCurrentDatabase([string]$tmp)
    try {
        $cn = $Access.CurrentProject.Connection
        foreach ($st in $Sql.Statements) {
            try { [void]$cn.Execute($st) }
            catch { throw "Rebuilding from the SQL failed at: $st`r`n$($_.Exception.Message)" }
        }
    } finally { $Access.CloseCurrentDatabase() }
    try {
        $again = ConvertTo-SchemaSql (Read-AccessSchema -Access $Access -Path $tmp)
    } finally { Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue }
    @(Compare-Object @($Sql.Statements) @($again.Statements) | ForEach-Object {
        $(if ($_.SideIndicator -eq '<=') { '- ' } else { '+ ' }) + ($_.InputObject -replace "`r`n\s*", ' ')
    })
}

# Saves every form of a front end with SaveAsText into -SnapshotDir (FormSnapshot's
# layout: <form>.txt, or <form>.failed with the error). Works on a copy with its
# StartUpForm removed and the code off, so nothing in the app runs.
function Save-ReleaseFormText {
    param([Parameter(Mandatory)]$Access, [Parameter(Mandatory)][string]$FrontEnd, [Parameter(Mandatory)][string]$SnapshotDir, [Parameter(Mandatory)][string]$WorkDir)
    $copy = Join-Path $WorkDir ('forms-' + [guid]::NewGuid().ToString('N') + '.accdb')
    Copy-Item -LiteralPath $FrontEnd -Destination $copy
    $db = $Access.DBEngine.OpenDatabase([string]$copy)
    try { $db.Properties.Delete('StartUpForm') } catch {}   # error 3265 when there is none
    $db.Close()
    New-Item -ItemType Directory -Force -Path $SnapshotDir | Out-Null
    $security = $Access.AutomationSecurity
    $failed = New-Object System.Collections.Generic.List[string]
    try { $Access.CloseCurrentDatabase() } catch {}
    $Access.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: no VBA runs
    try {
        $Access.OpenCurrentDatabase([string]$copy)
        try {
            $all = $Access.CurrentProject.AllForms
            $names = @(for ($i = 0; $i -lt $all.Count; $i++) { [string]$all.Item($i).Name })
            foreach ($n in $names) {
                if ($n -match '[\\/:*?"<>|]') { $failed.Add("${n}: the name has a character that can't be used in a file name"); continue }
                $path = Join-Path $SnapshotDir "$n.txt"
                try { $Access.SaveAsText(2, $n, [string]$path) }   # 2 = acForm
                catch {
                    [IO.File]::WriteAllText((Join-Path $SnapshotDir "$n.failed"), $_.Exception.Message, (New-Object Text.UTF8Encoding($false)))
                    $failed.Add("${n}: $($_.Exception.Message)")
                }
            }
        } finally { $Access.CloseCurrentDatabase() }
    } finally {
        $Access.AutomationSecurity = $security
        Remove-Item -LiteralPath $copy -ErrorAction SilentlyContinue
    }
    [pscustomobject]@{ Forms = $names.Count; Failed = $failed.ToArray() }
}

# Writes a release's design folder (see the top of this file): the front end's forms as
# JSON in <OutDir>\forms and a build script <OutDir>\<database>.sql for the front end's
# own tables (when it has any) and each data file. Each script is rebuilt into an empty
# database and read back. Throws if a form can't be exported or a script doesn't rebuild
# its tables, so a release never ships a partial design. Returns what it wrote, with
# what SQL can't make (attachment fields and the like) in Skipped.
function Export-ReleaseDesign {
    param(
        [Parameter(Mandatory)]$Access,
        [Parameter(Mandatory)][string]$FrontEnd,
        [string[]]$DataFiles,
        [string]$ThemeFile,
        [Parameter(Mandatory)][string]$OutDir,
        [string]$WorkDir = $env:TEMP
    )
    if (Test-Path -LiteralPath $OutDir) { throw "$OutDir already exists." }
    $work = Join-Path $WorkDir ('release-design-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $work, $OutDir | Out-Null
    try {
        $snap = Join-Path $work 'forms'
        $saved = Save-ReleaseFormText -Access $Access -FrontEnd $FrontEnd -SnapshotDir $snap -WorkDir $work
        if ($saved.Failed.Count) { throw "Could not export these forms:`r`n  " + ($saved.Failed -join "`r`n  ") }
        $designs = Export-FormDesignFolder -SnapshotDir $snap -OutDir (Join-Path $OutDir 'forms') -ThemeFile $ThemeFile
        if ($designs.Failed.Count) { throw "Could not write these form designs:`r`n  " + (($designs.Failed | ForEach-Object { "$($_.form): $($_.error)" }) -join "`r`n  ") }

        $scripts = @(); $skipped = @()
        foreach ($file in @($FrontEnd) + @($DataFiles | Where-Object { $_ })) {
            $schema = Read-AccessSchema -Access $Access -Path $file
            if (-not $schema.Tables.Count) { continue }   # a front end with linked tables only
            $leaf = Split-Path -Leaf $file
            $sql = ConvertTo-SchemaSql $schema -Title $leaf
            $diff = @(Test-SchemaSqlRebuild -Access $Access -Sql $sql -WorkDir $work)
            if ($diff.Count) { throw "The SQL for $leaf does not rebuild its tables:`r`n  " + ($diff -join "`r`n  ") }
            $name = [IO.Path]::GetFileNameWithoutExtension($leaf) + '.sql'
            [IO.File]::WriteAllText((Join-Path $OutDir $name), $sql.Text, (New-Object Text.UTF8Encoding($false)))
            $scripts += [pscustomobject]@{ File = $name; Tables = @($schema.Tables).Count; Statements = $sql.Statements.Count }
            $skipped += @($sql.Skipped | ForEach-Object { "${leaf}: $_" })
        }
        [pscustomobject]@{ Forms = $designs.Written.Count; Scripts = $scripts; Skipped = $skipped }
    } finally { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue }
}
