param([string]$SourcePath=(Join-Path $PSScriptRoot '..\TEST_DB.accdb'),[string]$DestinationPath=(Join-Path $PSScriptRoot '..\build\navfit-proof\NAVFIT_NEXT_FIELDS_PROBE.accdb'))
$ErrorActionPreference='Stop'
$SourcePath=[IO.Path]::GetFullPath($SourcePath)
$DestinationPath=[IO.Path]::GetFullPath($DestinationPath)
if(Test-Path -LiteralPath $DestinationPath){throw 'Destination already exists; refusing overwrite.'}
$hash=(Get-FileHash -LiteralPath $SourcePath).Hash
[void](New-Item -ItemType Directory -Force -Path (Split-Path $DestinationPath -Parent))
Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath
$source=$null; $target=$null
try {
    $engine=New-Object -ComObject DAO.DBEngine.120
    $source=$engine.OpenDatabase($SourcePath,$false,$true)
    $target=$engine.OpenDatabase($DestinationPath)
    $r=$target.OpenRecordset('SELECT * FROM Reports ORDER BY ReportID',2)
    $expected=@{}
    while(-not $r.EOF){
        $kind=$r.Fields.Item('ReportType').Value
        if($kind -is [string] -and $kind -cin @('Eval','Chief','FitRep')){
            $markers=[ordered]@{
                RRSFI='Z';RRSMI='Q';RRSLastName='SWIFT RRS TEST';RRSGrade='CDR';RRSCommand='SWIFT RRS CMD';RRSUIC='12345'
                Achievements=('SWIFT ACHIEVEMENTS TEST - '+$kind)
                Qualifications=('SWIFT QUALIFICATIONS TEST - '+$kind)
                RSAddress=('SWIFT RS ADDRESS TEST - '+$kind)
                RSInfo=('SWIFT RS INFO TEST - '+$kind)
                UserComments=('SWIFT USER COMMENTS TEST - '+$kind)
                Regular=$false;Concurrent=$true;IsValidated=$false
            }
            $expected[[int]$r.Fields.Item('ReportID').Value]=$markers
            $r.Edit()
            foreach($field in $markers.Keys){if($markers[$field] -isnot [bool]){$r.Fields.Item($field).Value=$markers[$field]}}
            $r.Update()
            $target.Execute('UPDATE Reports SET Regular=False,Concurrent=True,IsValidated=False WHERE ReportID='+[int]$r.Fields.Item('ReportID').Value,128)
        }
        $r.MoveNext()
    }
    $r.Close()
    if($expected.Count -ne 3){throw 'Expected exactly three recognized synthetic reports.'}
    $cells=0; $changes=0
    foreach($name in @('Folders','Reports','Summary')){
        $order=if($name -eq 'Folders'){' ORDER BY FolderID'}elseif($name -eq 'Reports'){' ORDER BY ReportID'}else{''}
        $a=$source.OpenRecordset("SELECT * FROM [$name]$order",4)
        $b=$target.OpenRecordset("SELECT * FROM [$name]$order",4)
        while(-not $a.EOF){
            if($b.EOF){throw 'Missing target row.'}
            for($i=0;$i -lt $a.Fields.Count;$i++){
                $af=$a.Fields.Item($i);$bf=$b.Fields.Item($i); $av=$af.Value;$bv=$bf.Value
                if($name -eq 'Reports' -and $expected.ContainsKey([int]$a.Fields.Item('ReportID').Value) -and $expected[[int]$a.Fields.Item('ReportID').Value].Contains($af.Name)){
                    if($bv -cne $expected[[int]$a.Fields.Item('ReportID').Value][$af.Name]){throw 'Marker mismatch.'}; $changes++
                }else{
                    $an=$null -eq $av -or $av -is [DBNull];$bn=$null -eq $bv -or $bv -is [DBNull]
                    if($an -ne $bn){throw 'Unexpected null change.'}
                    if(-not $an){if($av -is [string]){$same=[string]::Equals($av,$bv,[StringComparison]::Ordinal)}else{$same=$av.Equals($bv)};if(-not $same){throw "Unexpected value change in $name.$($af.Name)"}}
                }
                if($af.Type -ne $bf.Type){throw 'Unexpected type change.'}
                $cells++
            }
            $a.MoveNext();$b.MoveNext()
        }
        if(-not $b.EOF){throw 'Unexpected extra row.'};$a.Close();$b.Close()
    }
    $target.Close();$target=$null;$source.Close();$source=$null
    if((Get-FileHash -LiteralPath $SourcePath).Hash -ne $hash){throw 'Source changed.'}
    @{Status='PASS';CellsChecked=$cells;IntentionalFieldChanges=$changes;SourceUnchanged=$true;SourceSHA256=$hash;CandidateSHA256=(Get-FileHash -LiteralPath $DestinationPath).Hash;NAVFITResult='Pending user inspection; synthetic concurrent-report mapping probe, not Swift export';ExpectedMarkers=@($expected.Keys | Sort-Object | ForEach-Object { @{ReportID=$_;Markers=$expected[$_]} })} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath ($DestinationPath+'.verification.json')
    Write-Output "PASS: $cells cells checked; only $changes intentional field changes. $DestinationPath"
} finally {if($null -ne $target){$target.Close()};if($null -ne $source){$source.Close()}}
