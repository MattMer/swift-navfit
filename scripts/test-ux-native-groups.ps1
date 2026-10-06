param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$dir=Join-Path $root ('build/tests/native-groups-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $dir)
$source=Join-Path $dir 'synthetic.accdb'; & (Join-Path $PSScriptRoot 'make-navfit-full-fixture.ps1') -DestinationPath $source
$bad=Join-Path $dir 'bad-last-group.accdb';Copy-Item $source $bad
$engine=New-Object -ComObject DAO.DBEngine.120;$db=$engine.OpenDatabase($bad)
try{$db.Execute("UPDATE Reports SET SSN='ABCDEF123' WHERE ReportType='Eval'",128)}finally{$db.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)}
$scratch=Join-Path $dir 'test.accdb';Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true);$app.LoadFromText(5,'UXNativeGroupsHarness',(Join-Path $root 'tests/UXNativeGroupsHarness.bas'));$app.DoCmd.RunCommand(126)
$result=$app.Eval('TestUXNativeGroups("'+$source+'","'+$bad+'","'+(Join-Path $root 'NAVFIT_BLANK.accdb')+'","'+$dir+'")');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
}finally{try{[void]$app.Eval('UXLeaveEditor("Discard")')}catch{};$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
