param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$dir=Join-Path $root ('build/tests/navfit-ui-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $dir)
$source=Join-Path $dir 'synthetic.accdb'; & (Join-Path $PSScriptRoot 'make-navfit-full-fixture.ps1') -DestinationPath $source
$test=Join-Path $dir 'test.accdb';Copy-Item -LiteralPath $BuildPath -Destination $test
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($test,$true)
 $app.LoadFromText(5,'UXNAVFITHarness',(Join-Path $root 'tests/UXNAVFITHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestUXNAVFIT("'+$source+'","'+(Join-Path $root 'NAVFIT_BLANK.accdb')+'")');if($result -notlike 'PASS:*'){throw $result}
 $app.CloseCurrentDatabase()
 @{Status=$result;TestPath=$test;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash}|ConvertTo-Json|Set-Content (Join-Path $dir 'verification.json')
 Write-Output $result
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
