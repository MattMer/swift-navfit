$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'build-plan.ps1')
$manifestPath=Join-Path $PSScriptRoot 'build-manifest.json'
$original=Get-Content $manifestPath -Raw
$folder=Join-Path $root ('build/tests/build-plan-'+[guid]::NewGuid());[void](New-Item -ItemType Directory $folder)
$checks=0
function Check([bool]$ok,[string]$message){if(-not $ok){throw $message};$script:checks++}
function RejectPlan($manifest,[string]$expected){
 $path=Join-Path $folder ([guid]::NewGuid().ToString()+'.json');$manifest|ConvertTo-Json -Depth 12|Set-Content $path
 $failure=''
 try{[void](GetSwiftBuildPlan $root $path)}catch{$failure=$_.Exception.Message}
 Check ($failure -like ('*'+$expected+'*')) ('Expected '+$expected+'; got '+$failure)
}
$plan=GetSwiftBuildPlan $root $manifestPath
Check ($plan.Manifest.sourceModules.Count -eq (Get-ChildItem (Join-Path $root 'src') -Filter '*.bas').Count) 'Source module coverage differs'
Check ($plan.AppVersion -eq '0.26.1-dev' -and $plan.SchemaVersion -eq 26) 'Identity differs'
$m=$original|ConvertFrom-Json;$m.stages[0].file='scripts/does-not-exist.ps1';RejectPlan $m 'Missing build dependency'
$m=$original|ConvertFrom-Json;$m.stages[1].id=$m.stages[0].id;RejectPlan $m 'Duplicate build stage'
$m=$original|ConvertFrom-Json;$m.stages[0].requires=@($m.stages[-1].id);RejectPlan $m 'requires an earlier stage'
$m=$original|ConvertFrom-Json;$m.stages[0].phase='unknown';RejectPlan $m 'Invalid build phase'
$m=$original|ConvertFrom-Json;$m.inputs+=('../outside.txt');RejectPlan $m 'escapes repository'
$m=$original|ConvertFrom-Json;$m.sourceModules=@($m.sourceModules|Select-Object -Skip 1);RejectPlan $m 'inventory differs'
$badSchema=Join-Path $folder 'bad-schema.json';@{SchemaVersion=999}|ConvertTo-Json|Set-Content $badSchema
$m=$original|ConvertFrom-Json;$m.schemaSource=$badSchema.Substring($root.Length+1).Replace('\','/');RejectPlan $m 'Schema definition and application constant disagree'
$stamp=Join-Path $folder 'input-stamp.txt';'before'|Set-Content $stamp
$m=$original|ConvertFrom-Json;$m.inputs+=($stamp.Substring($root.Length+1).Replace('\','/'));$path=Join-Path $folder 'changed-input.json';$m|ConvertTo-Json -Depth 12|Set-Content $path
$plan=GetSwiftBuildPlan $root $path;'after'|Set-Content $stamp
$rejected=$false;try{AssertSwiftBuildInputs $plan}catch{$rejected=$_.Exception.Message -like '*input changed*'};Check $rejected 'Input drift was not rejected'
# The real entry point must reject an invalid plan before creating any database.
$m=$original|ConvertFrom-Json;$m.stages[0].file='scripts/does-not-exist.ps1';$path=Join-Path $folder 'entry-rejection.json';$m|ConvertTo-Json -Depth 12|Set-Content $path
$out=Join-Path $folder 'must-not-exist.accdb';$rejected=$false
try{& (Join-Path $PSScriptRoot 'build.ps1') -ManifestPath $path -OutputPath $out}catch{$rejected=$_.Exception.Message -like '*Missing build dependency*'}
Check ($rejected -and -not(Test-Path $out)) 'Preflight wrote an output before rejection'
Write-Output ('PASS: '+$checks+' build-plan preflight, dependency, identity and input-drift checks')
