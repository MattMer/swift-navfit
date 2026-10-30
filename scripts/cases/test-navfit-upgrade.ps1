param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$dir=Join-Path $root ('build/tests/navfit-v33-'+[guid]::NewGuid());[void](New-Item -ItemType Directory $dir)
$modern=Join-Path $dir 'modern.accdb'
& (Join-Path (Split-Path $PSScriptRoot -Parent) 'make-navfit-full-fixture.ps1') -DestinationPath $modern
$patch=@(@('Reports','DoDID',10),@('Reports','RSDoDID',10),@('Folders','RSAddress1',30),@('Folders','RSAddress2',30),@('Folders','RSCity',15),@('Folders','RSState',2),@('Folders','RSZipCd',9),@('Folders','RSPhone',10),@('Folders','RSDSN',7))
$engine=New-Object -ComObject DAO.DBEngine.120
try{
 foreach($kind in @('legacy','partial','readonly','bad-size','missing-core','unknown-column','failed-init','stale')){
  $path=Join-Path $dir ($kind+'.accdb');Copy-Item $modern $path
  $db=$engine.OpenDatabase($path)
  try{
   foreach($field in $patch){
    if(($kind -in @('partial','failed-init')) -and $field[1] -eq 'RSPhone'){continue}
    $db.Execute('ALTER TABLE ['+$field[0]+'] DROP COLUMN ['+$field[1]+']',128)
   }
   if($kind -eq 'partial'){$db.Execute("UPDATE Folders SET RSPhone='5551234567'",128)}
   if($kind -eq 'bad-size'){$db.Execute('ALTER TABLE Reports ADD COLUMN DoDID TEXT(9)',128)}
   if($kind -eq 'missing-core'){$db.Execute('ALTER TABLE Reports DROP COLUMN Comments',128)}
   if($kind -eq 'unknown-column'){$db.Execute('ALTER TABLE Reports ADD COLUMN Surprise TEXT(10)',128)}
   if($kind -eq 'failed-init'){$db.Execute('UPDATE Folders SET RSPhone=NULL',128);$db.TableDefs.Item('Folders').Fields.Item('RSPhone').ValidationRule='Is Null'}
  }finally{$db.Close()}
  if($kind -eq 'readonly'){(Get-Item $path).IsReadOnly=$true}
 }
}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)}
$before=@{};Get-ChildItem $dir -Filter *.accdb|ForEach-Object {$before[$_.Name]=(Get-FileHash $_.FullName).Hash}
$scratch=Join-Path $dir 'swift.accdb';Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 $app.LoadFromText(5,'NAVFITUpgradeHarness',(Join-Path $root 'tests/NAVFITUpgradeHarness.bas'))
 $app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestNAVFITUpgrade("'+$dir+'","'+(Join-Path $root 'FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb')+'")');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 foreach($name in $before.Keys){if($name -eq 'stale.accdb'){continue};if((Get-FileHash (Join-Path $dir $name)).Hash -ne $before[$name]){throw ('Original fixture changed: '+$name)}}
 [ordered]@{Result=$result;BuildSHA256=(Get-FileHash $BuildPath).Hash;FixtureDirectory=$dir;OriginalsPreserved=$true}|ConvertTo-Json|Set-Content (Join-Path $root 'build/tests/navfit-upgrade-results.json')
}finally{try{[void]$app.Eval('CloseReviewFile()')}catch{};$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
