param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$AdversarialPreviews)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/group-correctness-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$hash=(Get-FileHash -LiteralPath $BuildPath).Hash;$folder=$scratch+'.cases';[void](New-Item -ItemType Directory -Path $folder)
$app=$null;$primary=$null;$cleanup=@();$results=@()
try{
 $cases=if($AdversarialPreviews){@('Rollback','Committed','Display','Uncertain','Reset')}else{@('Normal')}
 foreach($fault in $cases){
  $target=$scratch
  if($AdversarialPreviews){$target=Join-Path $folder ($fault+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $target}
  Write-Output ('Controller: directory batch runner; disposable database: '+$target)
  $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1
  $app.OpenCurrentDatabase($target,$true)
  if($fault -in @('Committed','Display','Uncertain')){
   $module=if($fault -eq 'Committed'){'modDirectoryBatch'}else{'modUXDirectory'}
   $code=[IO.File]::ReadAllText((Join-Path $root ('src/'+$module+'.bas')))
   $anchor=if($fault -eq 'Committed'){'    w.CommitTrans: trans = False: outcome = "Committed": ClearDirectoryReportPreview: Exit Function'}else{'    count = ApplyDirectoryReports(outcome): f!cmdApply.Enabled = False: f!txtStamp = Null'}
   if(([regex]::Matches($code,[regex]::Escape($anchor))).Count -ne 1){throw 'Expected directory batch fault boundary missing or duplicated'}
   $injection='    Err.Raise vbObjectError + 9692, "Synthetic directory batch apply", "SYNTHETIC batch display failure"'
   $replacement=if($fault -eq 'Committed'){'    w.CommitTrans: trans = False: outcome = "Committed"'+"`r`n"+$injection+"`r`n"+'    ClearDirectoryReportPreview: Exit Function'}elseif($fault -eq 'Uncertain'){'    outcome = "Rollback could not be confirmed"'+"`r`n"+$injection}else{$anchor+"`r`n"+$injection}
   $instrumented=Join-Path $folder ($fault+'-'+$module+'.bas');[IO.File]::WriteAllText($instrumented,$code.Replace($anchor,$replacement));$app.LoadFromText(5,$module,$instrumented)
  }
  $app.LoadFromText(5,'DirectoryBatchHarness',(Join-Path $root 'tests/DirectoryBatchHarness.bas'));$app.DoCmd.RunCommand(126)
  if($fault -eq 'Normal'){$result=$app.Eval('TestDirectoryBatch()')}
  elseif($fault -eq 'Reset'){
   $prepared=$app.Eval('PrepareDirectoryBatchState()');Write-Output $prepared;if($prepared -notlike 'PASS:*'){throw $prepared}
   $app.VBE.CommandBars.FindControl(1,228).Execute();$result=$app.Eval('VerifyDirectoryBatchReset()')
  }else{$result=$app.Eval('TestDirectoryBatchOutcome("'+$fault+'")')}
  Write-Output $result;if($result -notlike 'PASS:*'){throw $result};$results+=$result
  $app.CloseCurrentDatabase()
  if($AdversarialPreviews){
   $app.OpenCurrentDatabase($target,$true);$expected=if($fault -in @('Committed','Display')){2}else{0}
   if($app.Eval('DCount("*","Report")') -ne $expected -or $app.Eval('DCount("*","GroupMember")') -ne $expected -or $app.Eval('DCount("*","MemberReport")') -ne $expected){throw 'Directory batch report/link counts changed after reopen'}
   $app.CloseCurrentDatabase();Write-Output ('PASS: '+$fault+' report/group-member/directory-link counts persisted across reopen')
  }
  $app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null
 }
}catch{$primary=$_}finally{
 if($app){try{$app.Quit(2)}catch{$cleanup+=$_.Exception.Message};try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{$cleanup+=$_.Exception.Message}}
 if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $hash){$cleanup+='Source artifact changed'}
}
if($primary){if($cleanup){Write-Warning ($cleanup -join '; ')};throw $primary};if($cleanup){throw ($cleanup -join '; ')}
@{Status='PASS';Results=$results;BuildSHA256=$hash;AdversarialPreviews=[bool]$AdversarialPreviews;Scope='Owned synthetic copies; explicit postcommit/display instrumentation and simulated uncertain outcome; native VBE reset'}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $folder 'verification.json')
Write-Output ('Receipt: '+(Join-Path $folder 'verification.json'))
