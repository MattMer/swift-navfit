param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$InstallSource,[string]$LegacyPath,[string]$PDFPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$folder=Join-Path $root ('build/tests/wizard-state-'+[guid]::NewGuid())
[void](New-Item -ItemType Directory -Path $folder)
$hash=(Get-FileHash -LiteralPath $BuildPath).Hash
$scratch=Join-Path $folder 'fixture.accdb';Copy-Item -LiteralPath $BuildPath -Destination $scratch
$pdfInput='';$pdfOriginalHash=''
if($PDFPath){$pdfOriginalHash=(Get-FileHash -LiteralPath $PDFPath).Hash;$pdfInput=Join-Path $folder 'synthetic-input.pdf';Copy-Item -LiteralPath $PDFPath -Destination $pdfInput}
$app=$null;$cleanupFailure='';$assertionTotal=0;$phases=[Collections.Generic.List[string]]::new()
try {
 $app=New-Object -ComObject Access.Application
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 Write-Output ('Controller: focused wizard runner; database: '+$scratch)
 if($InstallSource){
  $schema=Get-Content (Join-Path $root 'src/major-schema.json') -Raw|ConvertFrom-Json
  $db=$app.CurrentDb()
  try{foreach($sql in @($schema.SQL|Where-Object {$_ -match '(TABLE WizardPreview|INDEX UX_WizardPreview)'})){$db.Execute($sql,128)};$db.Execute("UPDATE AppMetadata SET AppVersion='0.31.0-dev',SchemaVersion=30",128)}finally{$db.Close()}
  foreach($module in @('modSwift','modMigration','modWorkspaceImport','modNAVFITExchange','modUXNAVFIT','modPDFImport','modUXImportFlow')){$app.LoadFromText(5,$module,(Join-Path $root ('src/'+$module+'.bas')))}
  foreach($binding in @(@('frmNAVFITExchange','OnUnload','=NAVFITClearReturned(False)'),@('frmPDFImport','OnUnload','=PDFClearPreview(False)'),@('frmUXWorkspaceImport','OnLoad','=UXWorkspacePreviewStart()'))){
   $app.DoCmd.OpenForm($binding[0],1);$form=$app.Forms.Item($binding[0]);$form.Properties.Item($binding[1]).Value=$binding[2];$app.DoCmd.Close(2,$binding[0],1)
  }
 }
 $app.LoadFromText(5,'WizardStateHarness',(Join-Path $root 'tests/WizardStateHarness.bas'))
 Write-Output 'Compile wizard source'
 $app.DoCmd.RunCommand(126)
 $app.CloseCurrentDatabase()
 $blank=Join-Path $folder 'blank.accdb';Copy-Item -LiteralPath $scratch -Destination $blank
 $app.OpenCurrentDatabase($scratch,$true)
 $result=$app.Eval('PrepareWizardState("'+$folder+'","'+(Join-Path $root 'FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb')+'","'+$pdfInput+'")')
 Write-Output $result;if($result -notlike 'PASS:*'){throw $result};$phases.Add($result);$assertionTotal+=[int]([regex]::Match($result,'PASS: (\d+)').Groups[1].Value)
 $app.Visible=$true
 $accessHandle=[int64]$app.hWndAccessApp()
 $ownedProcess=@(Get-Process MSACCESS|Where-Object {$_.MainWindowHandle.ToInt64() -eq $accessHandle})
 if($ownedProcess.Count -ne 1){throw 'Cannot identify owned Access process for the synthetic error test'}
 # The classic VBE error dialog has no UI Automation Invoke pattern. Supervise
 # this one literal synthetic error using a bounded, PID/start-time checked job;
 # one Alt+E ends it. Never click/dismiss arbitrary Access or user dialogs.
 $supervisor=Start-Job -ArgumentList $ownedProcess[0].Id,$ownedProcess[0].StartTime.ToUniversalTime().Ticks -ScriptBlock {
  param($ownedPID,$startedTicks)
  Add-Type -AssemblyName UIAutomationClient
  $deadline=[datetime]::UtcNow.AddSeconds(15)
  while([datetime]::UtcNow -lt $deadline){
   $process=Get-Process -Id $ownedPID -ErrorAction Stop
   if($process.StartTime.ToUniversalTime().Ticks -ne $startedTicks){throw 'Owned test process changed'}
   $window=[Windows.Automation.AutomationElement]::FromHandle($process.MainWindowHandle)
   if($window){
    $names=@($window.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.Condition]::TrueCondition)|ForEach-Object {$_.Current.Name})
    if(($names -join ' ') -like '*SYNTHETIC unhandled wizard failure*'){
     $keyboard=New-Object -ComObject WScript.Shell
     try{if(-not $keyboard.AppActivate($ownedPID)){throw 'Owned synthetic dialog could not be activated'};$keyboard.SendKeys('%e')}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($keyboard)}
     return 'Ended verified synthetic VBA failure'
    }
   }
   Start-Sleep -Milliseconds 500
  }
  throw 'Expected synthetic error dialog did not appear; no alternate control attempted'
 }
 $failureObserved=$false
 try{[void]$app.Eval('WizardUnhandledFailure()')}catch{$failureObserved=$true;Write-Output ('Expected unhandled VBA failure: '+$_.Exception.Message)}
 $supervision=Receive-Job -Job $supervisor -Wait -AutoRemoveJob -ErrorAction Stop
 if(-not $failureObserved -or $supervision -ne 'Ended verified synthetic VBA failure'){throw 'Unhandled failure fixture was not verified'}
 # A real VBA project reset, not a helper clearing selected module variables.
 $reset=@($app.VBE.CommandBars.Item('Run').Controls|Where-Object {$_.Caption.Replace('&','') -eq 'Reset'})
 if($reset.Count -ne 1){throw 'Native VBA Reset command is not available; reset coverage cannot be claimed'}
 Write-Output ('Execute native VBA Reset; command ID: '+$reset[0].Id)
 $reset[0].Execute()
 $result=$app.Eval('VerifyWizardReset('+$(if($PDFPath){'True'}else{'False'})+')');Write-Output $result;if($result -notlike 'PASS:*'){throw $result};$phases.Add($result);$assertionTotal+=[int]([regex]::Match($result,'PASS: (\d+)').Groups[1].Value)
 $app.CloseCurrentDatabase();$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null
 $app=New-Object -ComObject Access.Application
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 $result=$app.Eval('VerifyWizardRestart()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result};$phases.Add($result);$assertionTotal+=[int]([regex]::Match($result,'PASS: (\d+)').Groups[1].Value)
 $result=$app.Eval('SeedWizardMigrationSource("'+$folder+'","'+(Join-Path $root 'FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb')+'")');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
 $expectedStage=$result.Substring(6)
 $app.CloseCurrentDatabase()
 $source=Join-Path $folder 'migration-source.accdb';Copy-Item -LiteralPath $scratch -Destination $source
 $target=Join-Path $folder 'migration-target.accdb';Copy-Item -LiteralPath $blank -Destination $target
 $app.OpenCurrentDatabase($target,$true)
 if(-not $app.Eval('PrepareWorkspaceResume("'+$source+'")')){throw 'Workspace reset preparation failed'}
 $app.VBE.CommandBars.FindControl(1,228).Execute()
 $result=$app.Eval('TestWizardMigration("'+$source+'","'+$expectedStage+'")');Write-Output $result;if($result -notlike 'PASS:*'){throw $result};$phases.Add($result);$assertionTotal+=[int]([regex]::Match($result,'PASS: (\d+)').Groups[1].Value)
 if($PDFPath){$result=$app.Eval('TestPDFWizardApply("'+$pdfInput+'")');Write-Output $result;if($result -notlike 'PASS:*'){throw $result};$phases.Add($result);$assertionTotal+=[int]([regex]::Match($result,'PASS: (\d+)').Groups[1].Value)}
 $app.CloseCurrentDatabase()
 if($LegacyPath){
  $legacyOriginalHash=(Get-FileHash -LiteralPath $LegacyPath).Hash
  $legacy=Join-Path $folder 'legacy29-source.accdb';Copy-Item -LiteralPath $LegacyPath -Destination $legacy
  $app.OpenCurrentDatabase($legacy,$true)
  if($app.Eval('DLookup("SchemaVersion","AppMetadata")') -ne 29){throw 'Legacy proof requires a real schema-29 fixture'}
  [void]$app.Eval('ConfigureWorkspace("SYNTHETIC SCHEMA 29","Legacy tester")')
  [void]$app.Eval('CommitDraft(0,"EVAL","SYNTHETIC LEGACY","SYNTHETIC LEGACY",Null,Null,"SYNTHETIC PRIVATE","Legacy seed","SYNTHETIC ORIGINAL")')
  $app.CloseCurrentDatabase()
  $legacyHash=(Get-FileHash -LiteralPath $legacy).Hash
  $legacyTarget=Join-Path $folder 'legacy29-target.accdb';Copy-Item -LiteralPath $blank -Destination $legacyTarget
  $app.OpenCurrentDatabase($legacyTarget,$true)
  [void]$app.Eval('ConfigureWorkspace("SYNTHETIC LEGACY BOOTSTRAP","Bootstrap tester")')
  if(-not $app.Eval('UXPreviewSwiftWorkspace("'+$legacy+'")') -or -not $app.Eval('UXApplySwiftWorkspace()')){throw 'Schema-29 Settings import failed'}
  if($app.Eval('DCount("*","Report")') -ne 1 -or $app.Eval('DCount("*","WizardPreview")') -ne 0){throw 'Schema-29 import lost records or retained target staging'}
  $app.CloseCurrentDatabase()
  if((Get-FileHash -LiteralPath $legacy).Hash -ne $legacyHash -or (Get-FileHash -LiteralPath $LegacyPath).Hash -ne $legacyOriginalHash){throw 'Legacy import changed source bytes'}
  Write-Output 'PASS: real schema-29 to schema-30 Settings import, original-field verification and source hashes'
 }
}catch{$primary=$_;throw $primary}
finally{if($app){try{$app.Quit(2)}catch{$cleanupFailure=$_.Exception.Message;Write-Warning ('Owned wizard cleanup: '+$cleanupFailure)};try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{Write-Warning $_.Exception.Message}}}
if($cleanupFailure){throw ('Owned test cleanup failed: '+$cleanupFailure)}
if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $hash){throw 'Wizard tests changed their supplied artifact'}
if($PDFPath -and (Get-FileHash -LiteralPath $PDFPath).Hash -ne $pdfOriginalHash){throw 'Wizard tests changed supplied PDF bytes'}
[pscustomobject]@{Status='PASS';Assertions=$assertionTotal;BuildPath=$BuildPath;BuildSHA256=$hash;InstalledSource=[bool]$InstallSource;NativeReset=$true;UnhandledFailure=$failureObserved;NewAccessSessionRestart=$true;Schema29Import=[bool]$LegacyPath;PDFResumeUIAndApply=[bool]$PDFPath;SourceHashesUnchanged=$true;Phases=@($phases.ToArray());Scope='Focused development proof; schema 29/30 only, no release or physical/accessibility acceptance'}|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $folder 'verification.json')
Write-Output ('PASS: wizard reset/restart and source isolation; evidence: '+$folder)
