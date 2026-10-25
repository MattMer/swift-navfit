param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$dir=Join-Path $root ('build/tests/group-discovery-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $dir)
$source=Join-Path $dir 'missing-rates.accdb'
& (Join-Path $PSScriptRoot 'make-navfit-full-fixture.ps1') -DestinationPath $source
$engine=New-Object -ComObject DAO.DBEngine.120;$db=$engine.OpenDatabase($source)
try{$db.Execute("UPDATE Reports SET Rate='' WHERE ReportType Is Not Null",128)}finally{$db.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)}
$before=(Get-FileHash -LiteralPath $source).Hash
$scratch=Join-Path $dir 'test.accdb';Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 $app.LoadFromText(5,'NativeGroupDiscoveryHarness',(Join-Path $root 'tests/NativeGroupDiscoveryHarness.bas'));$app.DoCmd.RunCommand(126)
 $original=Join-Path $dir 'original.accdb'; & (Join-Path $PSScriptRoot 'make-navfit-full-fixture.ps1') -DestinationPath $original;$originalHash=(Get-FileHash -LiteralPath $original).Hash
 $result=$app.Eval('TestOriginalGroupDiscovery("'+$original+'")');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
 if((Get-FileHash -LiteralPath $original).Hash -ne $originalHash){throw 'Original fixture changed'}
 $result=$app.Eval('TestNativeGroupDiscovery("'+$source+'")');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
if((Get-FileHash -LiteralPath $source).Hash -ne $before){throw 'Source file changed'}
Write-Output 'PASS: source database hash unchanged'
