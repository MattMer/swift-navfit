param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$hash=(Get-FileHash -LiteralPath $BuildPath -Algorithm SHA256).Hash
$scratch=Join-Path $root ('build/tests/build-cleanup-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try{
 $app=New-Object -ComObject Access.Application
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 $app.LoadFromText(5,'BuildCleanupHarness',(Join-Path $root 'tests/BuildCleanupHarness.bas'))
 $app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestBuildCleanup()')
 if($result -notlike 'PASS:*'){throw $result}
 Write-Output $result
}finally{
 if($null -ne $app){try{[void]$app.Eval('UXLeaveEditor("Discard")')}catch{};$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 if((Get-FileHash -LiteralPath $BuildPath -Algorithm SHA256).Hash -ne $hash){throw 'Build cleanup test changed its source artifact'}
}
