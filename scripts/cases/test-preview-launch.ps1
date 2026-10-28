param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
# Use the destination folder and normal Office security, rather than bypassing it.
$scratch=Join-Path (Split-Path ([IO.Path]::GetFullPath($BuildPath)) -Parent) ('launch-check-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.OpenCurrentDatabase($scratch,$true)
    $trusted=$app.CurrentProject.IsTrusted
    $app.DoCmd.RunMacro('AutoExec')
    $expected=if(-not $trusted){'frmEnableContent'}elseif($app.Eval('DCount("*","Workspace")') -eq 0){'frmSetup'}else{'frmContributorStart'}
    if(-not $app.Eval('CurrentProject.AllForms("'+$expected+'").IsLoaded')){throw 'AutoExec did not open the expected destination'}
    $app.CloseCurrentDatabase()
    $app.OpenCurrentDatabase($scratch,$true)
    $app.DoCmd.RunMacro('AutoExec')
    if(-not $app.Eval('CurrentProject.AllForms("'+$expected+'").IsLoaded')){throw 'Reopened preview failed startup'}
    $app.CloseCurrentDatabase()
    'PASS: normal Office security, AutoExec, destination and repeat launch ('+$expected+')'
    $app.AutomationSecurity=1
    $app.OpenCurrentDatabase($scratch,$true)
    $app.DoCmd.RunMacro('AutoExec')
    $expected=if($app.Eval('DCount("*","Workspace")') -eq 0){'frmSetup'}else{'frmContributorStart'}
    if(-not $app.Eval('CurrentProject.AllForms("'+$expected+'").IsLoaded')){throw 'Enabled-content startup failed'}
    'PASS: enabled-content AutoExec ('+$expected+')'
} finally {
    if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    Remove-Item -LiteralPath $scratch -ErrorAction SilentlyContinue
}
