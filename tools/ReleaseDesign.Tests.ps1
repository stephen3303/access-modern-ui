# Checks ReleaseDesign.ps1's SQL build script against a small schema, without Access:
#   powershell -NoProfile -File tools\ReleaseDesign.Tests.ps1
# (pwsh works too). Prints each check and exits 1 if any fails. Reading a real database
# and rebuilding it (Read-AccessSchema, Test-SchemaSqlRebuild) needs Access; a release
# build runs both on every database it ships.

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ReleaseDesign.ps1')

$failures = 0
function Check([string]$Name, [bool]$Ok) {
    if ($Ok) { Write-Host "ok   $Name" } else { Write-Host "FAIL $Name"; $script:failures++ }
}

function F($Name, $Type, $Ordinal, [hashtable]$More = @{}) {
    $f = [ordered]@{ Name = $Name; Ordinal = $Ordinal; Type = $Type; Size = 0; Required = $false; AllowZeroLength = $false; Default = ''; AutoIncrement = $false; Expression = ''; Precision = $null; Scale = $null }
    foreach ($k in $More.Keys) { $f[$k] = $More[$k] }
    [pscustomobject]$f
}
function I($Name, $Fields, [hashtable]$More = @{}) {
    $i = [ordered]@{ Name = $Name; Primary = $false; Unique = $false; Required = $false; IgnoreNulls = $false; Foreign = $false
        Fields = @($Fields | ForEach-Object { if ($_ -like '-*') { [pscustomobject]@{ Name = $_.Substring(1); Descending = $true } } else { [pscustomobject]@{ Name = $_; Descending = $false } } }) }
    foreach ($k in $More.Keys) { $i[$k] = $More[$k] }
    [pscustomobject]$i
}

# Defaults read the same however they were set.
Check 'default: double-quoted text' ((ConvertTo-SqlDefault '"Draft"') -eq "'Draft'")
Check 'default: quotes inside text' ((ConvertTo-SqlDefault '"O''Brien ""Jr"""') -eq "'O''Brien ""Jr""'")
Check 'default: single-quoted text kept' ((ConvertTo-SqlDefault "'Draft'") -eq "'Draft'")
Check 'default: empty text' ((ConvertTo-SqlDefault '""') -eq "''")
Check 'default: yes/no words' ((ConvertTo-SqlDefault 'Yes') -eq '-1' -and (ConvertTo-SqlDefault 'False') -eq '0' -and (ConvertTo-SqlDefault '-1') -eq '-1')
Check 'default: leading = dropped' ((ConvertTo-SqlDefault '=Now()') -eq 'Now()' -and (ConvertTo-SqlDefault 'Date()') -eq 'Date()')
Check 'default: none' ($null -eq (ConvertTo-SqlDefault '') -and $null -eq (ConvertTo-SqlDefault '  '))
Check 'default: text that looks quoted but is an expression' ((ConvertTo-SqlDefault '"a" & "b"') -eq '"a" & "b"')

$schema = [pscustomobject]@{
    Tables = @(
        [pscustomobject]@{ Name = 'Messages'; Fields = @(
                (F 'id' 4 0 @{ AutoIncrement = $true; Required = $false }),
                (F 'templateId' 4 1 @{ Required = $true }),
                (F 'subject' 10 2 @{ Size = 120; Required = $true; Default = '""' }),
                (F 'body' 12 3),
                (F 'files' 101 4),
                (F 'created' 8 5 @{ Default = '=Now()' }),
                (F 'final' 1 6 @{ Default = 'No' }),
                (F 'words' 4 7 @{ Expression = 'Len([body])' }),
                (F 'amount' 20 8 @{ Precision = 18; Scale = 4 }))
            Indexes = @(
                (I 'templateId' @('templateId')),
                (I 'PrimaryKey' @('id') @{ Primary = $true; Unique = $true; Required = $true }),
                (I 'Messages_Templates' @('templateId') @{ Foreign = $true }),
                (I 'byFiles' @('files')),
                (I 'bySubject' @('subject', '-created') @{ Unique = $true; IgnoreNulls = $true })) },
        [pscustomobject]@{ Name = 'Templates'; Fields = @(
                (F 'id' 4 0 @{ AutoIncrement = $true }),
                (F 'code' 10 1 @{ Size = 20; Default = "'X'" }),
                (F 'weight' 7 2 @{ Default = '0.5' }),
                (F 'key' 15 3),
                (F 'small' 3 4),
                (F 'tiny' 2 5),
                (F 'real' 6 6),
                (F 'money' 5 7),
                (F 'blob' 11 8))
            Indexes = @(
                (I 'PrimaryKey' @('id') @{ Primary = $true; Unique = $true; Required = $true }),
                (I 'code' @('code') @{ Required = $true })) }
    )
    Relations = @(
        [pscustomobject]@{ Name = 'TemplatesMessages'; Table = 'Templates'; ForeignTable = 'Messages'; Attributes = (256 + 4096); Fields = @([pscustomobject]@{ Name = 'id'; ForeignName = 'templateId' }) },
        [pscustomobject]@{ Name = 'Loose'; Table = 'Templates'; ForeignTable = 'Messages'; Attributes = 2; Fields = @([pscustomobject]@{ Name = 'id'; ForeignName = 'templateId' }) },
        [pscustomobject]@{ Name = 'Linked'; Table = 'People'; ForeignTable = 'Messages'; Attributes = 4; Fields = @([pscustomobject]@{ Name = 'id'; ForeignName = 'templateId' }) },
        [pscustomobject]@{ Name = 'ToMissing'; Table = 'People'; ForeignTable = 'Messages'; Attributes = 0; Fields = @([pscustomobject]@{ Name = 'id'; ForeignName = 'templateId' }) },
        [pscustomobject]@{ Name = 'Sys'; Table = 'MSysObjects'; ForeignTable = 'Messages'; Attributes = 0; Fields = @([pscustomobject]@{ Name = 'Id'; ForeignName = 'templateId' }) }
    )
}

$sql = ConvertTo-SchemaSql $schema -Title 'Test_be.accdb'
$s = $sql.Statements

Check 'tables first, sorted by name' ($s[0] -like 'CREATE TABLE `[Messages`] (*' -and $s[1] -like 'CREATE TABLE `[Templates`] (*')
$messages = @($s[0] -split "`r`n" | ForEach-Object { $_.TrimEnd(",") })   # commas between columns dropped
Check 'autonumber, no default or NOT NULL' ($messages[1] -eq '    [id] COUNTER')
Check 'required long' ($messages[2] -eq '    [templateId] INTEGER NOT NULL')
Check 'text size, default before NOT NULL' ($messages[3] -eq "    [subject] TEXT(120) DEFAULT '' NOT NULL")
Check 'memo' ($messages[4] -eq '    [body] LONGTEXT')
Check 'date default' ($messages[5] -eq '    [created] DATETIME DEFAULT Now()')
Check 'yes/no default' ($messages[6] -eq '    [final] YESNO DEFAULT 0')
Check 'decimal' ($messages[7] -eq '    [amount] DECIMAL(18,4)')
Check 'attachment and calculated fields left out' ($s[0] -notmatch 'files|words' -and $messages[-1] -eq ')')
$templates = @($s[1] -split "`r`n" | ForEach-Object { $_.TrimEnd(",") })
Check 'other types' (($templates[1..9] -join '|') -eq "    [id] COUNTER|    [code] TEXT(20) DEFAULT 'X'|    [weight] FLOAT DEFAULT 0.5|    [key] UNIQUEIDENTIFIER|    [small] SMALLINT|    [tiny] BYTE|    [real] REAL|    [money] CURRENCY|    [blob] LONGBINARY")

$ix = @($s | Where-Object { $_ -like 'CREATE *INDEX*' })
Check 'primary key first, then by name' ($ix[0] -eq 'CREATE INDEX [PrimaryKey] ON [Messages] ([id]) WITH PRIMARY' -and $ix[1] -like '*`[bySubject`]*' -and $ix[2] -like '*`[templateId`]*')
Check 'unique, descending, ignore nulls' ($ix[1] -eq 'CREATE UNIQUE INDEX [bySubject] ON [Messages] ([subject], [created] DESC) WITH IGNORE NULL')
Check 'plain index' ($ix[2] -eq 'CREATE INDEX [templateId] ON [Messages] ([templateId])')
Check 'required index' ($ix[4] -eq 'CREATE INDEX [code] ON [Templates] ([code]) WITH DISALLOW NULL')
Check 'foreign-key index left to its relationship' (-not ($s | Where-Object { $_ -like '*Messages_Templates*' }))
Check 'index on a field SQL cannot make is left out' (-not ($s | Where-Object { $_ -like '*byFiles*' }))

$fk = @($s | Where-Object { $_ -like 'ALTER TABLE*' })
Check 'one foreign key' ($fk.Count -eq 1)
Check 'foreign key with cascades' ($fk[0] -eq 'ALTER TABLE [Messages] ADD CONSTRAINT [TemplatesMessages] FOREIGN KEY ([templateId]) REFERENCES [Templates] ([id]) ON UPDATE CASCADE ON DELETE CASCADE')
Check 'statements in order: tables, indexes, keys' ([array]::IndexOf($s, $fk[0]) -eq $s.Count - 1 -and [array]::IndexOf($s, $ix[0]) -eq 2)

$sk = $sql.Skipped -join '|'
Check 'skipped: attachment, calculated, index, unenforced, missing table' ($sk -match '\[Messages\]\.\[files\]: attachment field' -and $sk -match '\[words\]: calculated field' -and $sk -match 'index \[byFiles\]' -and $sk -match '\[Loose\].*not enforced' -and $sk -match '\[ToMissing\].*not in this database')
Check 'skipped: inherited and system relationships not listed' ($sk -notmatch 'Linked|Sys\]')
Check 'decimal without precision is skipped' ((ConvertTo-SchemaSql ([pscustomobject]@{ Tables = @([pscustomobject]@{ Name = 'T'; Fields = @((F 'a' 4 0), (F 'd' 20 1)); Indexes = @() }); Relations = @() })).Skipped[0] -match 'decimal field')
Check 'table with no field SQL can make is skipped' ((ConvertTo-SchemaSql ([pscustomobject]@{ Tables = @([pscustomobject]@{ Name = 'T'; Fields = @((F 'a' 101 0)); Indexes = @() }); Relations = @() })).Statements.Count -eq 0)

Check 'text: header names the database and what is left out' ($sql.Text -like '-- Tables of Test_be.accdb as Access SQL*' -and $sql.Text -match 'Not made by this script \(5\):')
Check 'text: each statement ends with a semicolon' (($sql.Text -split ";`r`n").Count -eq $s.Count + 1)
$again = ConvertTo-SchemaSql $schema -Title 'Test_be.accdb'
Check 'same schema, same script' ($again.Text -eq $sql.Text)
$shuffled = [pscustomobject]@{ Tables = @($schema.Tables[1], $schema.Tables[0]); Relations = @($schema.Relations | Sort-Object Name -Descending) }
Check 'order of tables and relationships does not matter' ((ConvertTo-SchemaSql $shuffled -Title 'Test_be.accdb').Text -eq $sql.Text)

Check 'Date/Time Extended has no Access SQL type' ((ConvertTo-SchemaSql ([pscustomobject]@{ Tables = @([pscustomobject]@{ Name = 'T'; Fields = @((F 'a' 4 0), (F 'x' 26 1)); Indexes = @() }); Relations = @() })).Skipped[0] -match 'Date/Time Extended')

# SQL Server (T-SQL) from the same schema.
$ss = ConvertTo-SqlServerSql $schema -Title 'Test_be.accdb'
$t = $ss.Text
Check 'sqlserver: header' ($t -like '-- Tables of Test_be.accdb for SQL Server (T-SQL).*')
Check 'sqlserver: autonumber as identity' ($t -match '\[id\] INT IDENTITY\(1,1\) NOT NULL')
Check 'sqlserver: required text not empty' ($t -match "\[subject\] NVARCHAR\(120\) NOT NULL CONSTRAINT \[DF_Messages_subject\] DEFAULT \(N''\) CHECK \(LEN\(\[subject\]\) > 0\)")
Check 'sqlserver: date default' ($t -match '\[created\] DATETIME2\(0\) NULL CONSTRAINT \[DF_Messages_created\] DEFAULT \(SYSDATETIME\(\)\)')
Check 'sqlserver: yes/no not null with default' ($t -match '\[final\] BIT NOT NULL CONSTRAINT \[DF_Messages_final\] DEFAULT \(0\)')
Check 'sqlserver: decimal, guid, money, memo' ($t -match '\[amount\] DECIMAL\(18, 4\) NULL' -and $t -match '\[key\] UNIQUEIDENTIFIER NULL' -and $t -match '\[money\] MONEY NULL' -and $t -match '\[body\] NVARCHAR\(MAX\) NULL')
Check 'sqlserver: primary key constraint' ($t -match 'CONSTRAINT \[PK_Messages\] PRIMARY KEY \(\[id\]\)')
Check 'sqlserver: unique index filtered on nullable columns' ($t -match 'CREATE UNIQUE INDEX \[IX_Messages_bySubject\] ON \[Messages\] \(\[subject\], \[created\] DESC\) WHERE \[created\] IS NOT NULL;')
Check 'sqlserver: attachment and calculated noted, not created' ($t -match '-- \[Messages\]\.\[files\]: attachment field, not created\.' -and $t -match '\[words\]: calculated field' -and $t -notmatch '    \[files\] ')
Check 'sqlserver: foreign key with cascades' ($t -match 'ALTER TABLE \[Messages\] ADD CONSTRAINT \[FK_TemplatesMessages\] FOREIGN KEY \(\[templateId\]\) REFERENCES \[Templates\] \(\[id\]\) ON DELETE CASCADE ON UPDATE CASCADE;')
Check 'sqlserver: unenforced relationship as a comment' ($t -match '-- ALTER TABLE \[Messages\] ADD CONSTRAINT \[FK_Loose\]')
Check 'sqlserver: foreign keys after every table' ($t.IndexOf('ALTER TABLE') -gt $t.LastIndexOf('CREATE TABLE'))
Check 'sqlserver: skipped list' (($ss.Skipped -join '|') -match '\[files\]: attachment field' -and ($ss.Skipped -join '|') -match '\[Loose\].*not enforced')
Check 'sqlserver: Date/Time Extended' ((ConvertTo-SqlServerSql ([pscustomobject]@{ Tables = @([pscustomobject]@{ Name = 'T'; Fields = @((F 'x' 26 0)); Indexes = @() }); Relations = @() })).Text -match '\[x\] DATETIME2\(7\) NULL')
Check 'sqlserver: defaults' ((ConvertTo-SqlServerDefault (F 'a' 8 0 @{ Default = 'Date()' })) -eq 'CAST(GETDATE() AS date)' -and (ConvertTo-SqlServerDefault (F 'a' 1 0 @{ Default = 'Yes' })) -eq '1' -and (ConvertTo-SqlServerDefault (F 'a' 10 0 @{ Default = '"O''B"' })) -eq "N'O''B'" -and $null -eq (ConvertTo-SqlServerDefault (F 'a' 10 0 @{ Default = '=Environ("USERNAME")' })))
Check 'sqlserver: unconverted default noted' ((ConvertTo-SqlServerSql ([pscustomobject]@{ Tables = @([pscustomobject]@{ Name = 'T'; Fields = @((F 'u' 10 0 @{ Size = 20; Default = '=Environ("USERNAME")' })); Indexes = @() }); Relations = @() })).Text -match '-- \[T\]\.\[u\]: Access default =Environ\("USERNAME"\) not converted\.')
Check 'sqlserver: foreign key already indexed by its own index' ($t -notmatch 'IX_Messages_TemplatesMessages')

# Cascade routes: Event cascades to Area, Area to Grp, and Event straight to Grp (a
# shortcut, a second route to Grp). Loop and Back cascade to each other (a loop).
function Tbl($Name, [string[]]$Keys) { [pscustomobject]@{ Name = $Name; Fields = @(@((F 'id' 4 0 @{ AutoIncrement = $true })) + @(foreach ($k in $Keys) { F $k 4 1 })); Indexes = @((I 'PrimaryKey' @('id') @{ Primary = $true; Unique = $true; Required = $true })) } }
function Rel($Name, $Parent, $Child, $Key, $Attr) { [pscustomobject]@{ Name = $Name; Table = $Parent; ForeignTable = $Child; Attributes = $Attr; Fields = @([pscustomobject]@{ Name = 'id'; ForeignName = $Key }) } }
$routes = ConvertTo-SqlServerSql ([pscustomobject]@{
        Tables = @((Tbl 'Event'), (Tbl 'Area' @('eventId')), (Tbl 'Grp' @('eventId', 'areaId')), (Tbl 'Loop' @('backId')), (Tbl 'Back' @('loopId')))
        Relations = @((Rel 'a_EventGrp' 'Event' 'Grp' 'eventId' (4096 + 256)), (Rel 'b_AreaGrp' 'Area' 'Grp' 'areaId' 4096), (Rel 'c_EventArea' 'Event' 'Area' 'eventId' 4096),
            (Rel 'd_LoopBack' 'Loop' 'Back' 'loopId' 4096), (Rel 'e_BackLoop' 'Back' 'Loop' 'backId' 4096)) })
$rt = $routes.Text
Check 'sqlserver: cascade chain kept' ($rt -match '\[FK_c_EventArea\] FOREIGN KEY \(\[eventId\]\) REFERENCES \[Event\] \(\[id\]\) ON DELETE CASCADE;' -and $rt -match '\[FK_b_AreaGrp\] FOREIGN KEY \(\[areaId\]\) REFERENCES \[Area\] \(\[id\]\) ON DELETE CASCADE;')
Check 'sqlserver: shortcut loses its delete cascade only' ($rt -match '\[FK_a_EventGrp\] FOREIGN KEY \(\[eventId\]\) REFERENCES \[Event\] \(\[id\]\) ON UPDATE CASCADE;' -and $rt -match '-- \[a_EventGrp\]: ON DELETE CASCADE in Access, left off here')
Check 'sqlserver: a loop keeps one cascade' (([regex]::Matches($rt, 'REFERENCES \[(Loop|Back)\] \(\[id\]\) ON DELETE CASCADE;')).Count -eq 1)
Check 'sqlserver: cascades left off are listed' (@($routes.Skipped | Where-Object { $_ -match 'left off' }).Count -eq 2)
Check 'sqlserver: each foreign key indexed once' ($rt -match 'CREATE INDEX \[IX_Grp_a_EventGrp\] ON \[Grp\] \(\[eventId\]\);' -and $rt -match 'CREATE INDEX \[IX_Grp_b_AreaGrp\] ON \[Grp\] \(\[areaId\]\);' -and ([regex]::Matches($rt, 'CREATE INDEX')).Count -eq 5)
Check 'sqlserver: unenforced relationship gets no index' ($t -notmatch 'IX_Messages_Loose')
Check 'sqlserver: bracket in a name escaped' ((Format-SqlName 'a]b') -eq '[a]]b]')

Check 'system tables recognised' ((Test-SchemaSystemTable 'MSysObjects') -and (Test-SchemaSystemTable '~TMPCLP1') -and (Test-SchemaSystemTable 'f_0123456789ABCDEF0123456789ABCDEF_files') -and -not (Test-SchemaSystemTable 'Messages'))
$dialect = (Get-Command Export-ReleaseDesign).ScriptBlock.Ast.Body.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'Dialect' }
Check 'release SQL defaults to SQL Server only' ($dialect.DefaultValue.Extent.Text -eq "@('SqlServer')")
Check 'form designs loaded from review-mcp' ((Get-Command Export-FormDesignFolder -ErrorAction SilentlyContinue) -ne $null)

if ($failures) { Write-Host "$failures check(s) failed"; exit 1 }
Write-Host 'All checks passed'
