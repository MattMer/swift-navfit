param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/resource-tools-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 $app.LoadFromText(5,'ResourceToolsHarness',(Join-Path $root 'tests/ResourceToolsHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestResourceTools()');Write-Output $result
 [ordered]@{Result=$result;BuildSHA256=(Get-FileHash $BuildPath).Hash;Scratch=$scratch}|ConvertTo-Json|Set-Content (Join-Path $root 'build/tests/resource-tools-results.json')
 if($result -notlike 'PASS:*'){throw $result}
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
