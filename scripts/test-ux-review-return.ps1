param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/review-return-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true);$app.LoadFromText(5,'UXReviewReturnHarness',(Join-Path $root 'tests/UXReviewReturnHarness.bas'));$app.DoCmd.RunCommand(126)
$result=$app.Eval('TestReviewReturn()');if(-not $result.StartsWith('PASS:')){throw $result};Write-Output $result
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
