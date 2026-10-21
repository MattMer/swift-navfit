param([Parameter(Mandatory=$true)][string]$BuildPath,
 [string]$BaselinePath,[string]$DemoPath,[string]$SourcePath,[switch]$SkipBaseline)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
if(-not $BaselinePath){$BaselinePath=Join-Path $root 'output/Swift_NAVFIT_Blank_2026-10-20-r123a.accdb'}
if(-not $DemoPath){$DemoPath=Join-Path $root 'output/Swift_NAVFIT_Demo_2026-10-20-r123a.accdb'}
$folder=Join-Path $root ('build/tests/engine-import-'+[guid]::NewGuid());[void](New-Item -ItemType Directory $folder)
$source=Join-Path $folder 'large synthetic provenance.accdb';if($SourcePath){Copy-Item $SourcePath $source}else{Copy-Item $DemoPath $source}
$e=New-Object -ComObject DAO.DBEngine.120;$db=$e.OpenDatabase($source)
try{
 if($SourcePath){$r=$db.OpenRecordset('SELECT Count(*) FROM ImportRawValue');$count=[int]$r.Fields.Item(0).Value;$r.Close()}else{
 $receiptGuid=[guid]::NewGuid().ToString('B')
 $db.Execute("INSERT INTO ImportReceipt (ReceiptGuid,SourceHash,SourcePath,ExternalReportID,ReportID,AuthorProfileID,ImportedAt,AdapterVersion) SELECT TOP 1 '$receiptGuid','"+('0'*64)+"','SYNTHETIC provenance benchmark',1,ReportID,1,Now(),'Synthetic' FROM Report",128)
 $r=$db.OpenRecordset('SELECT Max(ImportReceiptID) FROM ImportReceipt');$receipt=[int]$r.Fields.Item(0).Value;$r.Close()
 $r=$db.OpenRecordset('ImportRawValue',2);$r.AddNew();$r.Fields.Item('ImportReceiptID').Value=$receipt;$r.Fields.Item('SourceTable').Value='SyntheticReport';$r.Fields.Item('SourceRow').Value=1;$r.Fields.Item('FieldOrdinal').Value=1;$r.Fields.Item('FieldKey').Value='Synthetic raw value';$r.Fields.Item('DAOType').Value=12;$r.Fields.Item('RawValue').Value=('x'*220)+[char]0x03A9+[char]0x6F22+" ' quotation `r`n trailing ";$r.Fields.Item('IsNullValue').Value=$false;$r.Update();$r.Close()
 $count=1
 for($i=0;$i -lt 18;$i++){
  $db.Execute('INSERT INTO ImportRawValue (ImportReceiptID,SourceTable,SourceRow,FieldOrdinal,FieldKey,DAOType,RawValue,IsNullValue) SELECT ImportReceiptID,SourceTable,SourceRow+'+$count+',FieldOrdinal,FieldKey,DAOType,RawValue,IsNullValue FROM ImportRawValue',128)
  $count*=2
 }
 $db.Execute('UPDATE ImportRawValue SET RawValue=Null,IsNullValue=True WHERE SourceRow Mod 37=0',128)
 }
}finally{$db.Close()}
$sourceHash=(Get-FileHash $source).Hash;$buildHash=(Get-FileHash $BuildPath).Hash
$app=New-Object -ComObject Access.Application;$results=@()
try{
 $kinds=if($SkipBaseline){@('engine')}else{@('engine','old-vba')}
 foreach($kind in $kinds){
  $ui=$null
  $template=if($kind -eq 'engine'){$BuildPath}else{$BaselinePath}
  $target=Join-Path $folder ($kind+'.accdb');Copy-Item $template $target
  $app.AutomationSecurity=1;$app.OpenCurrentDatabase($target,$true)
  foreach($name in @($app.CurrentProject.AllForms|Where-Object {$_.IsLoaded}|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$name,2)}
  $app.LoadFromText(5,'WorkspaceImportHarness',(Join-Path $root 'tests/WorkspaceImportHarness.bas'))
  $app.LoadFromText(5,'EngineVerificationHarness',(Join-Path $root 'tests/EngineVerificationHarness.bas'))
  [void]$app.Eval('ConfigureWorkspace("SYNTHETIC ENGINE BENCHMARK","Tester")')
  if($kind -eq 'engine'){
   $check=$app.Eval('TestEngineVerification()');if($check -notlike 'PASS:*'){throw $check};Write-Output $check
  }
  if(-not $app.Eval('UXPreviewSwiftWorkspace("'+$source+'")')){throw 'Preview failed'}
  if($kind -eq 'engine'){[void]$app.Eval('PrepareEngineImportUI()')}
  $clock=[Diagnostics.Stopwatch]::StartNew();$ok=$app.Eval('UXApplySwiftWorkspace()');$clock.Stop()
  $status=$app.Eval('Forms!frmUXWorkspaceImport!lblStatus.Caption');if(-not $ok){throw $status}
  $log=$app.Eval('Forms!frmUXWorkspaceImport!txtProgress.Value');$log|Set-Content (Join-Path $folder ($kind+'-progress.log'))
  if($app.Eval('DCount("*","ImportRawValue")') -ne $count){throw 'Raw value count changed'}
  if($kind -eq 'engine'){
   $ui=$app.Eval('EngineImportUIResult()');Write-Output $ui;if($ui -notlike 'PASS:*'){throw $ui}
   if($app.Eval('Forms!frmUXWorkspaceImport.Modal') -or $app.Eval('Forms!frmUXWorkspaceImport!cmdApply.Enabled') -or -not $app.Eval('Forms!frmUXWorkspaceImport!cmdClose.Enabled')){throw 'Final controls/modality not restored'}
   if($log -notlike '*database engine*' -or $log -notlike '*Import committed successfully*'){throw 'Engine progress missing'}
  }
  $results+=[pscustomobject]@{Engine=$kind;Seconds=$clock.Elapsed.TotalSeconds;RawRecords=$count;Target=$target;Status=$status;UI=$ui}
  Write-Output ('PASS '+$kind+': '+[Math]::Round($clock.Elapsed.TotalSeconds,3)+' seconds; '+$count+' provenance rows')
  $app.CloseCurrentDatabase()
 }
 if((Get-FileHash $source).Hash -ne $sourceHash -or (Get-FileHash $BuildPath).Hash -ne $buildHash){throw 'Source or template changed'}
 $new=($results|Where-Object Engine -eq 'engine').Seconds;$old=($results|Where-Object Engine -eq 'old-vba').Seconds
 @{Status='PASS';Source=$source;SourceSHA256=$sourceHash;BuildSHA256=$buildHash;Rows=$count;Runs=$results;Speedup=$(if($old){$old/$new}else{$null});ExactAdversarialChecks=16;SourceUnchanged=$true}|ConvertTo-Json -Depth 7|Set-Content (Join-Path $folder 'verification.json')
 Write-Output ('PASS: large import and responsive guarded UI; '+$folder)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
