param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$AdversarialPreviews)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$dir=Join-Path $root ('build/tests/native-groups-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $dir)
$source=Join-Path $dir 'synthetic.accdb'; & (Join-Path (Split-Path $PSScriptRoot -Parent) 'make-navfit-full-fixture.ps1') -DestinationPath $source
$bad=Join-Path $dir 'bad-last-group.accdb';Copy-Item $source $bad
$engine=New-Object -ComObject DAO.DBEngine.120;$db=$engine.OpenDatabase($bad)
try{$db.Execute("UPDATE Reports SET SSN='ABCDEF123' WHERE ReportType='Eval'",128)}finally{$db.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)}
$scratch=Join-Path $dir 'test.accdb';Copy-Item -LiteralPath $BuildPath -Destination $scratch
$hash=(Get-FileHash -LiteralPath $BuildPath).Hash
$app=New-Object -ComObject Access.Application;$primary=$null;$cleanup=@();$results=@()
try{
 $app.AutomationSecurity=1
 $cases=if($AdversarialPreviews){@('Rollback','Committed','Display','Uncertain','Reset')}else{@('Normal')}
 foreach($fault in $cases){
  $casePath=$scratch
  if($AdversarialPreviews){$casePath=Join-Path $dir ($fault+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $casePath}
  Write-Output ('Controller: native-group preview runner; disposable database: '+$casePath)
  $app.OpenCurrentDatabase($casePath,$true)
  if($fault -in @('Committed','Display','Uncertain')){
   $module=if($fault -eq 'Committed'){'modImportReview'}else{'modUXNativeGroups'}
   $code=[IO.File]::ReadAllText((Join-Path $root ('src/'+$module+'.bas')))
   $anchor=if($fault -eq 'Committed'){'    w.CommitTrans: trans = False: outcome = "Committed": Set ImportNAVFITGroups = result: Exit Function'}else{'    Set result = ImportNAVFITGroups(picks, previewHash, previewCentury, gradeOverrides, outcome)'}
   if(([regex]::Matches($code,[regex]::Escape($anchor))).Count -ne 1){throw 'Expected native-group fault boundary missing or duplicated'}
   $injection='    Err.Raise vbObjectError + 9691, "Synthetic native group apply", "SYNTHETIC finalization failure"'
   $replacement=if($fault -eq 'Committed'){'    w.CommitTrans: trans = False: outcome = "Committed"'+"`r`n"+$injection+"`r`n"+'    Set ImportNAVFITGroups = result: Exit Function'}elseif($fault -eq 'Uncertain'){'    outcome = "Rollback could not be confirmed"'+"`r`n"+$injection}else{$anchor+"`r`n"+$injection}
   $instrumented=Join-Path $dir ($fault+'-'+$module+'.bas');[IO.File]::WriteAllText($instrumented,$code.Replace($anchor,$replacement));$app.LoadFromText(5,$module,$instrumented)
  }
  $app.LoadFromText(5,'UXNativeGroupsHarness',(Join-Path $root 'tests/UXNativeGroupsHarness.bas'));$app.DoCmd.RunCommand(126)
  if($fault -eq 'Normal'){$result=$app.Eval('TestUXNativeGroups("'+$source+'","'+$bad+'","'+(Join-Path $root 'FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb')+'","'+$dir+'")')}
  elseif($fault -eq 'Reset'){
   $prepared=$app.Eval('PrepareOwnedPreviewReset("'+$source+'")');if($prepared -notlike 'PASS:*'){throw $prepared};Write-Output $prepared
   $app.VBE.CommandBars.FindControl(1,228).Execute()
   $result=$app.Eval('VerifyOwnedPreviewReset()')
  }else{$faultInput=if($fault -eq 'Rollback'){$bad}else{$source};$result=$app.Eval('TestNativeGroupApplyFault("'+$faultInput+'","'+$fault+'")')}
  Write-Output $result;if($result -notlike 'PASS:*'){throw $result};$results+=$result
  $app.CloseCurrentDatabase()
  if($AdversarialPreviews){
   $app.OpenCurrentDatabase($casePath,$true);$expected=if($fault -in @('Committed','Display','Reset')){1}else{0};$receipts=if($fault -in @('Committed','Display')){1}else{0}
   if($app.Eval('DCount("*","Report")') -ne $expected -or $app.Eval('DCount("*","SummaryGroup")') -ne $expected -or $app.Eval('DCount("*","ImportReceipt")') -ne $receipts){throw 'Native-group outcome counts changed after reopen'}
   $app.CloseCurrentDatabase();Write-Output ('PASS: '+$fault+' report/group/receipt counts persisted across reopen')
  }
 }
}catch{$primary=$_}finally{
 try{$app.Quit(2)}catch{$cleanup+=$_.Exception.Message}
 try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{$cleanup+=$_.Exception.Message}
 if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $hash){$cleanup+='Source build changed'}
}
if($primary){if($cleanup){Write-Warning ($cleanup -join '; ')};throw $primary}
if($cleanup){throw ($cleanup -join '; ')}
@{Status='PASS';Results=$results;BuildSHA256=$hash;AdversarialPreviews=[bool]$AdversarialPreviews;Scope='Synthetic faults and real reset on owned copies; uncertain rollback is simulated'}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $dir 'verification.json')
Write-Output ('Receipt: '+(Join-Path $dir 'verification.json'))
