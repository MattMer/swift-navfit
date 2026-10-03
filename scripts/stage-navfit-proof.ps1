param(
    [string]$SourcePath=(Join-Path $PSScriptRoot '..\TEST_DB.accdb'),
    [string]$OutputDirectory=(Join-Path $PSScriptRoot '..\build\navfit-staging')
)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$SourcePath=[IO.Path]::GetFullPath($SourcePath)
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
$expected=Get-Content -LiteralPath (Join-Path $root 'planning\evidence\navfit\schema-inventory.json') -Raw | ConvertFrom-Json
$hash=(Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash
[void](New-Item -ItemType Directory -Force -Path $OutputDirectory)
$batch=Join-Path $OutputDirectory $hash
if(Test-Path -LiteralPath $batch) {throw 'This source hash already has a staging attempt. Review that batch before retrying.'}
[void](New-Item -ItemType Directory -Path $batch)
$scratch=Join-Path $batch 'source-copy.accdb'
$target=Join-Path $batch 'staged-values.accdb'
Copy-Item -LiteralPath $SourcePath -Destination $scratch
$source=$null; $stage=$null
try {
    $engine=New-Object -ComObject DAO.DBEngine.120
    $source=$engine.OpenDatabase($scratch,$false,$true)
    $names=@('Folders','Reports','Summary')
    $actualNames=@()
    for($i=0;$i -lt $source.TableDefs.Count;$i++) {
        $t=$source.TableDefs.Item($i)
        if(-not $t.Name.StartsWith('MSys')) {$actualNames+=$t.Name}
    }
    if((($actualNames | Sort-Object) -join '|') -cne (($names | Sort-Object) -join '|')) {throw 'Unrecognized data tables.'}
    foreach($name in $names) {
        $t=$source.TableDefs.Item($name)
        $e=@($expected.Tables | Where-Object Name -CEQ $name)[0]
        if($t.Attributes -ne 0 -or $t.Fields.Count -ne $e.Fields.Count) {throw "Unrecognized table: $name"}
        for($i=0;$i -lt $t.Fields.Count;$i++) {
            $f=$t.Fields.Item($i); $ef=$e.Fields[$i]
            if($f.Name -cne $ef.Name -or $f.Type -ne $ef.Type -or $f.Size -ne $ef.Size) {throw "Unrecognized field: $name.$($f.Name)"}
        }
        if($t.Indexes.Count -ne $e.Indexes.Count) {throw "Unrecognized indexes: $name"}
        for($i=0;$i -lt $t.Indexes.Count;$i++) {
            $ix=$t.Indexes.Item($i); $ei=$e.Indexes[$i]
            if($ix.Name -cne $ei.Name -or $ix.Primary -ne $ei.Primary -or $ix.Unique -ne $ei.Unique -or $ix.Fields.Count -ne $ei.Fields.Count) {throw "Unrecognized index: $name"}
            for($j=0;$j -lt $ix.Fields.Count;$j++) {if($ix.Fields.Item($j).Name -cne $ei.Fields[$j]) {throw "Unrecognized index fields: $name"}}
        }
    }
    if($source.Relations.Count -ne 0 -or $source.QueryDefs.Count -ne 0) {throw 'Unexpected relationships or saved queries.'}
    # Preserve native field precision and properties in a quarantined database copy.
    # DAO does not run source macros or VBA.
    Copy-Item -LiteralPath $scratch -Destination $target
    $stage=$engine.OpenDatabase($target)
    foreach($name in $names) {$stage.TableDefs.Item($name).Name="Raw$name"}
    $checks=@(); $cells=0
    foreach($name in $names) {
        $e=@($expected.Tables | Where-Object Name -CEQ $name)[0]
        $primary=@($e.Indexes | Where-Object Primary); $order=if($primary.Count){" ORDER BY [$($primary[0].Fields[0])]"}else{""}
        $a=$source.OpenRecordset("SELECT * FROM [$name] $order",4)
        $b=$stage.OpenRecordset("SELECT * FROM [Raw$name] $order",4)
        $rows=0
        while(-not $a.EOF) {
            if($b.EOF) {throw "Missing staged row: $name"}
            for($i=0;$i -lt $a.Fields.Count;$i++) {
                $af=$a.Fields.Item($i); $bf=$b.Fields.Item($i)
                if($af.Name -cne $bf.Name -or $af.Type -ne $bf.Type) {throw "Staged type mismatch: $name"}
                $av=$af.Value; $bv=$bf.Value
                $an=$null -eq $av -or $av -is [DBNull]; $bn=$null -eq $bv -or $bv -is [DBNull]
                if($an -ne $bn) {throw "Null mismatch: $name row $rows field $($af.Name)"}
                if(-not $an) {
                    if($av -is [string]) {$same=[string]::Equals($av,$bv,[StringComparison]::Ordinal)} else {$same=$av.Equals($bv)}
                    if(-not $same) {throw "Value mismatch: $name row $rows field $($af.Name)"}
                }
                $cells++
            }
            $rows++; $a.MoveNext(); $b.MoveNext()
        }
        if(-not $b.EOF) {throw "Extra staged row: $name"}
        $a.Close(); $b.Close(); $checks+=@{Table=$name;Rows=$rows;AllTypedValuesEqual=$true}
    }
    $stage.Execute('CREATE TABLE ReviewQueue (SourceReportID LONG, CandidateKind TEXT(12), ReviewStatus TEXT(100))',128)
    $stage.Execute("INSERT INTO ReviewQueue (SourceReportID,CandidateKind,ReviewStatus) SELECT ReportID, IIf(ReportType='Eval','EVAL',IIf(ReportType='Chief','CHIEFEVAL',IIf(ReportType='FitRep','FITREP',Null))), IIf(ReportType In ('Eval','Chief','FitRep'),'Pending review - parent encoding unresolved','Quarantined - unknown or null report type') FROM RawReports",128)
    $stage.Close(); $stage=$null
    $source.Close(); $source=$null
    $stage=$engine.OpenDatabase($target,$false,$true)
    $queue=$stage.OpenRecordset('SELECT ReviewStatus, Count(*) AS N FROM ReviewQueue GROUP BY ReviewStatus',4)
    $statuses=@(); while(-not $queue.EOF) {$statuses+=@{Status=$queue.Fields.Item(0).Value;Count=$queue.Fields.Item(1).Value};$queue.MoveNext()}; $queue.Close()
    $stage.Close(); $stage=$null
    if((Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash -ne $hash) {throw 'Original source changed.'}
    @{Status='PASS';SourceSHA256=$hash;SourceUnchanged=$true;Tables=$checks;CellsCompared=$cells;ReviewQueue=$statuses;Reopened=$true;Scope='Isolated development staging only. No Swift draft import, NAVFIT export or compatibility claim.'} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $batch 'verification.json') -Encoding utf8
    Write-Output "PASS: $cells typed cells preserved. Evidence: $batch"
} catch {
    @{Status='FAIL';Reason=$_.Exception.Message} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $batch 'verification.json') -Encoding utf8
    throw
} finally {
    if($null -ne $stage) {$stage.Close()}
    if($null -ne $source) {$source.Close()}
}
