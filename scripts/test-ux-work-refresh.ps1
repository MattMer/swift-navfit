param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/queue-owner-'+[guid]::NewGuid()+'.accdb')
$module=Join-Path $root ('build/tests/queue-owner-'+[guid]::NewGuid()+'.bas')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$text=[IO.File]::ReadAllText((Join-Path $root 'src/modUXMyWork.bas'))
$text=$text.Replace('Option Explicit',"Option Explicit`r`nPublic QueueTestRefreshes As Long`r`nPublic Function QueueTestReset() As Boolean`r`n    QueueTestRefreshes = 0: QueueTestReset = True`r`nEnd Function`r`nPublic Function QueueTestCount() As Long`r`n    QueueTestCount = QueueTestRefreshes`r`nEnd Function")
$text=$text.Replace('    selected = Nz(f!lstWork, "")',"    QueueTestRefreshes = QueueTestRefreshes + 1`r`n    selected = Nz(f!lstWork, `"`")")
[IO.File]::WriteAllText($module,$text)
$app=New-Object -ComObject Access.Application
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 foreach($name in @($app.CurrentProject.AllForms|Where-Object {$_.IsLoaded}|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$name,2)}
 $app.LoadFromText(5,'modUXMyWork',$module);$app.DoCmd.RunCommand(126)
 [void]$app.Eval('ConfigureWorkspace("SYNTHETIC QUEUE OWNER","Tester")')
 foreach($action in @('Open','UXLoadMyWork()','LoadMyWork()','UXMyWorkMode("Assignments")','UXWorkBack()')){
  [void]$app.Eval('QueueTestReset()')
  if($action -eq 'Open'){$app.DoCmd.OpenForm('frmMyWork')}else{[void]$app.Eval($action)}
  $expected=1;if($action -eq 'Open'){$expected=2} # Initialization plus preserved activation freshness refresh.
  if($app.Eval('QueueTestCount()') -ne $expected){throw ('Expected one refresh: '+$action+'; actual='+$app.Eval('QueueTestCount()'))}
 }
 $app.DoCmd.Close(2,'frmMyWork',2);[void]$app.Eval('QueueTestReset()');[void]$app.Eval('UXWorkBack()')
 if($app.Eval('QueueTestCount()') -ne 2){throw ('Back to closed My Work expected 2 refreshes; actual='+$app.Eval('QueueTestCount()'))}
 $app.DoCmd.Close(2,'frmMyWork',2);[void]$app.Eval('QueueTestReset()');[void]$app.Eval('OpenMyWork()')
 if($app.Eval('QueueTestCount()') -ne 2){throw ('Legacy open expected 2 refreshes; actual='+$app.Eval('QueueTestCount()'))}
 Write-Output 'PASS: seven queue refresh ownership checks (instrumented disposable copy)'
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
