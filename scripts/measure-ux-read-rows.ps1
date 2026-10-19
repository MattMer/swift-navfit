param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/read-scale-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 $app.LoadFromText(5,'UXReadRowsScaleHarness',(Join-Path $root 'tests/UXReadRowsScaleHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestReadRowsScale()');if($result -notlike 'PASS:*'){throw $result}
 $result.Substring(7)|Set-Content (Join-Path $root 'build/batch2-read-scale.csv') -Encoding utf8
 Write-Output $result
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
