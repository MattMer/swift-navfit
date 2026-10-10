param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$LoadCurrentSource)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/trait-retry-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 foreach($opened in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$opened,2)}
 if($LoadCurrentSource){$app.LoadFromText(5,'modUXTraitGrid',(Join-Path $root 'src/modUXTraitGrid.bas'))}
 $app.LoadFromText(5,'UXTraitRetryHarness',(Join-Path $root 'tests/UXTraitRetryHarness.bas'))
 $app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestTraitRetry()');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
} finally {$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
