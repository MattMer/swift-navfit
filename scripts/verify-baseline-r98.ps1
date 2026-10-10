$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$package=Join-Path $root 'output/Swift_NAVFIT_Baseline_2026-10-10-r98'
$blank=Join-Path $package 'Swift_NAVFIT_Deployment_2026-10-10-r98-blank.accdb'
$demo=Join-Path $package 'Swift_NAVFIT_Deployment_2026-10-10-r98-DEMO.accdb'
$checks=@(
 @{Script='test-ux-startup.ps1';Arguments=@('-BuildPath',$blank)},
 @{Script='test-ux-startup.ps1';Arguments=@('-BuildPath',$demo)},
 @{Script='test-ux-trait-retry.ps1';Arguments=@('-BuildPath',$blank)},
 @{Script='test-trait-events.ps1';Arguments=@('-BuildPath',$blank)},
 @{Script='test-trait-queue.ps1';Arguments=@('-BuildPath',$blank)},
 @{Script='test-ruad-import.ps1';Arguments=@('-BuildPath',$blank,'-SamplePath','C:/Users/Matt/Downloads/87114_SMART_RUAD - 1OCT2026 - SPLIT.csv')},
 @{Script='test-ui-author-baseline.ps1';Arguments=@('-BuildPath',$demo,'-AllowGroupDirectoryLayout')},
 @{Script='test-navfit-group-export.ps1';Arguments=@('-BuildPath',$blank)}
)
$results=@()
foreach($check in $checks){
 $arguments=$check.Arguments
 $lines=@(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot $check.Script) @arguments 2>&1)
 $code=$LASTEXITCODE;$lines|ForEach-Object{Write-Output $_}
 $results+=@{Script=$check.Script;Arguments=$arguments;ExitCode=$code;Output=($lines -join "`n")}
 $results|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $package 'verification-progress.json') -Encoding UTF8
 if($code -ne 0){throw ('Baseline verification failed: '+$check.Script)}
}
$audit=Get-Content (Join-Path $package 'source-audit.json') -Raw|ConvertFrom-Json
$artifacts=@(foreach($file in Get-Item -LiteralPath $blank,$demo){@{Name=$file.Name;Bytes=$file.Length;SHA256=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash}})
$verification=@{Baseline='r98';Date='2026-10-10';AppVersion='0.25.0-dev';SchemaVersion=25;Status='Targeted artifact verification passed; development prerelease';SourceAudit=$audit;Checks=$results;Artifacts=$artifacts;Pending=@('Narrative geometry/overflow and physical printed fit','Full Swift/native round trips and returned-group reconciliation','Keyboard/screen-reader/physical scaling acceptance','Schema 20-to-25 deployment transfer and target PDF dependencies')}
$verification|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $root 'planning/BASELINE_R98_VERIFICATION.json') -Encoding UTF8
Copy-Item -LiteralPath (Join-Path $root 'planning/BASELINE_R98_VERIFICATION.json') -Destination $package
$verification|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $package 'MANIFEST.json') -Encoding UTF8
Write-Output 'PASS: r98 deployment artifact verification complete'
