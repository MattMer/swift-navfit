param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/startup-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$false)
 if(-not @($app.CurrentProject.AllMacros|ForEach-Object {$_.Name}).Contains('AutoExec')){throw 'Startup macro missing'}
 $expected=if($app.Eval('DCount("*","Workspace")') -eq 0){'frmSetup'}else{'frmHome'}
 if(-not $app.CurrentProject.AllForms.Item($expected).IsLoaded){throw ('Normal database open did not display '+$expected)}
 Write-Output ('PASS: normal database open displays '+$expected+' automatically; AutoExec present')
 $app.CloseCurrentDatabase()
} finally {$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
