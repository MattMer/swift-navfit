param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/direct-defaults-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 $app.LoadFromText(5,'RecursiveDefaultsHarness',(Join-Path $root 'tests/RecursiveDefaultsHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestRecursiveDefaults()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
 [ordered]@{Result=$result;BuildSHA256=(Get-FileHash $BuildPath).Hash;Scratch=$scratch}|ConvertTo-Json|Set-Content (Join-Path $root 'build/tests/recursive-defaults-results.json')
}finally{try{[void]$app.Eval('UXCloseBlocks()')}catch{};$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
