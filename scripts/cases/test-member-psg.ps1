param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'test-pdf-fixtures.ps1');$pdfs=NewSwiftPdfFixtures 'expanded'
$scratch=Join-Path $root ('build/tests/member-psg-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 $app.LoadFromText(5,'MemberPSGHarness',(Join-Path $root 'tests/MemberPSGHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestMemberPSG("'+$pdfs+'")');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 $result=$app.Eval('TestRUADReportRefresh()');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
