param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/group-read-scale-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 $app.LoadFromText(5,'GroupReadScaleHarness',(Join-Path $root 'tests/GroupReadScaleHarness.bas'))
 $app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestGroupReadScale()');if($result -notlike 'PASS:*'){throw $result}
 $result.Substring(7)|Set-Content (Join-Path $root 'build/batch3-group-scale.csv') -Encoding utf8
 Write-Output $result
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
