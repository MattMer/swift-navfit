param([string]$BuildPath=(Join-Path $PSScriptRoot '..\output\Swift_NAVFIT_0.2.3-tested.accdb'), [string]$SourcePath="")
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$source=Join-Path $root 'build\navfit-staging-verified4\E732EE41A384495AC37BC009C59516A73461C7DB99E251B5384CE917652386E7\staged-values.accdb'
if($SourcePath){$source=[IO.Path]::GetFullPath($SourcePath)}
$hash=(Get-FileHash -LiteralPath $source).Hash
$scratch=Join-Path $root ('build\tests\review-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null; $checks=@()
function Assert($ok,$label) {if(-not $ok){throw $label}; $script:checks+=$label}
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    [void]$app.Eval('ConfigureWorkspace("SYNTHETIC REVIEW","Synthetic Tester")')
    $app.LoadFromText(5,'ReviewHarness',(Join-Path $root 'tests\ReviewHarness.bas'))
    $result=$app.Eval('TryReview("'+$source.Replace('"','""')+'")')
    Assert ($result -eq 'PASS') ('Review opens: '+$result)
    Assert ($app.Screen.ActiveForm.Name -eq 'frmImportReview') 'Review becomes active form'
    $f=$app.Forms.Item('frmImportReview')
    Assert ($f.Controls.Item('lstImportReports').ListCount -eq 4) 'All four source rows visible'
    $f.Controls.Item('lstImportReports').Value=2
    Assert ($app.Eval('ReviewReportFields()')) 'Report selection succeeds'
    Assert ($f.Controls.Item('lstImportFields').ListCount -eq 128) 'All 128 report fields visible'
    $db=$app.CurrentDb()
    $engine=New-Object -ComObject DAO.DBEngine.120
    $raw=$engine.OpenDatabase($source,$false,$true)
    $table=if($SourcePath){'Reports'}else{'RawReports'}
    $fields=$raw.TableDefs.Item($table).Fields
    $parent=-1; $type=-1
    for($i=0;$i -lt $fields.Count;$i++) {if($fields.Item($i).Name -eq 'Parent'){$parent=$i};if($fields.Item($i).Name -eq 'ReportType'){$type=$i}}
    $f.Controls.Item('lstImportFields').Value=$parent
    [void]$app.Eval('ReviewFieldValue()')
    Assert ($f.Controls.Item('txtFieldValue').Value -ceq 'a 2') 'Encoded parent displayed verbatim'
    $f.Controls.Item('lstImportReports').Value=1
    [void]$app.Eval('ReviewReportFields()')
    $f.Controls.Item('lstImportFields').Value=$type
    [void]$app.Eval('ReviewFieldValue()')
    Assert ($f.Controls.Item('txtFieldValue').Value -eq '[NULL]') 'Null type displayed explicitly'
    $r=$db.OpenRecordset('SELECT Count(*) FROM Report'); Assert ($r.Fields.Item(0).Value -eq 0) 'Review creates no draft reports';$r.Close()
    $raw.Close(); $db.Close()
    $bad=Join-Path $root ('build\tests\invalid-review-'+[guid]::NewGuid()+'.accdb')
    Copy-Item -LiteralPath $source -Destination $bad
    $badDb=$engine.OpenDatabase($bad)
     $badDb.TableDefs.Item($table).Fields.Item('Parent').Name='ChangedParent'
    $badDb.Close()
    $rejected=$app.Eval('TryReview("'+$bad.Replace('"','""')+'")')
    Assert ($rejected -ne 'PASS' -and $rejected -like '*Unrecognized staging field*') 'Altered staging schema rejected'
    Assert ($app.Eval('LeaveImportReview()')) 'Close review succeeds'
    Assert ($app.Screen.ActiveForm.Name -eq 'frmHome') 'Close returns to home'
    $app.CloseCurrentDatabase()
    Assert ((Get-FileHash -LiteralPath $source).Hash -eq $hash) 'Staged source hash unchanged'
    $resultFile=if($SourcePath){'import-review-direct-results.json'}else{'import-review-staged-results.json'}
    @{Status='PASS';Checks=$checks;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash} | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $root ('build\tests\'+$resultFile))
    Write-Output "PASS: $($checks.Count) import review checks"
} finally {if($null -ne $app){$app.Quit(2)}}
