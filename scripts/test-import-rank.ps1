param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/import-rank-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 $app.LoadFromText(5,'ImportRankHarness',(Join-Path $root 'tests/ImportRankHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestImportRank("'+(Join-Path $root 'output/pdf/synthetic-eval.pdf')+'")');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
