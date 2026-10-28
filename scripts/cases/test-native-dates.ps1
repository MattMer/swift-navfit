param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$dir=Join-Path $root ('build/tests/native-dates-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $dir)
$source=Join-Path $dir 'synthetic-times.accdb'; & (Join-Path (Split-Path $PSScriptRoot -Parent) 'make-navfit-full-fixture.ps1') -DestinationPath $source
$engine=New-Object -ComObject DAO.DBEngine.120;$db=$engine.OpenDatabase($source)
try{
 foreach($table in @('Reports','Folders')){
  $assignments=@();foreach($field in $db.TableDefs.Item($table).Fields){if($field.Type -eq 8){$assignments+=('['+$field.Name+"]=DateAdd('s',47655,["+$field.Name+'])')}}
  $db.Execute('UPDATE ['+$table+'] SET '+($assignments -join ','),128)
 }
 $db.Execute("UPDATE Folders SET FromDate=#1/1/2026 1:14:15 PM#, ToDate=#9/30/2026 1:14:15 PM#",128)
 $db.Execute("UPDATE Reports SET SeniorRaterDate=Null WHERE ReportType='Eval'",128)
}finally{$db.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)}
$hash=(Get-FileHash -LiteralPath $source).Hash
$scratch=Join-Path $dir 'test.accdb';Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 $app.LoadFromText(5,'NativeDatesHarness',(Join-Path $root 'tests/NativeDatesHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestNativeDates("'+$source+'")');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
if((Get-FileHash -LiteralPath $source).Hash -ne $hash){throw 'Import modified source'}
Write-Output 'PASS: original source hash unchanged'
# Exercise the temporary SQL cleanup on a separate disposable native copy.
$clean=Join-Path $dir 'workaround-copy.accdb';Copy-Item -LiteralPath $source -Destination $clean
$engine=New-Object -ComObject DAO.DBEngine.120;$db=$engine.OpenDatabase($clean)
try{
 foreach($table in @('Reports','Folders')){
  $assignments=@();foreach($field in $db.TableDefs.Item($table).Fields){if($field.Type -eq 8){$assignments+=('['+$field.Name+']=Int(['+$field.Name+'])')}}
  $db.Execute('UPDATE ['+$table+'] SET '+($assignments -join ','),128)
  $r=$db.OpenRecordset('SELECT * FROM ['+$table+']')
  while(-not $r.EOF){foreach($field in $r.Fields){if($field.Type -eq 8 -and $field.Value -is [datetime] -and ([datetime]$field.Value).TimeOfDay.TotalSeconds -ne 0){throw 'Temporary cleanup retained a time'}};$r.MoveNext()};$r.Close()
 }
 $r=$db.OpenRecordset("SELECT SeniorRaterDate FROM Reports WHERE ReportType='Eval'");if($r.Fields('SeniorRaterDate').Value -isnot [DBNull]){throw 'Temporary cleanup changed a null date'};$r.Close()
}finally{$db.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)}
Write-Output 'PASS: temporary Reports/Folders Int() cleanup removes times and preserves null dates'
