param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$DisplayFaultOnly)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/group-correctness-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
if($DisplayFaultOnly){
 $code=[IO.File]::ReadAllText((Join-Path $root 'src/modUXGroupOrganization.bas'));$anchor='    committed = True'
 if(([regex]::Matches($code,[regex]::Escape($anchor))).Count -ne 1){throw 'Expected move display fault boundary missing or duplicated'}
 $instrumented=$scratch+'.bas';[IO.File]::WriteAllText($instrumented,$code.Replace($anchor,$anchor+"`r`n"+'    Err.Raise 5, "Synthetic move display", "SYNTHETIC move display failure"'));$app.LoadFromText(5,'modUXGroupOrganization',$instrumented)
}
$app.LoadFromText(5,'UXOrganizationHarness',(Join-Path $root 'tests/UXOrganizationHarness.bas'));$app.DoCmd.RunCommand(126)
$expression=if($DisplayFaultOnly){'TestMoveDisplayFailure()'}else{'TestUXOrganization()'}
$result=$app.Eval($expression);Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
}finally{try{[void]$app.Eval('UXLeaveEditor("Discard")')}catch{};$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
