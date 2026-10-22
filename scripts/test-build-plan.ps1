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
$identity=Get-Content (Join-Path $root $plan.Manifest.identitySource) -Raw
$expectedVersion=[regex]::Match($identity,'Public Const APP_VERSION As String = "([^"]+)"').Groups[1].Value
$expectedSchema=[int](Get-Content (Join-Path $root $plan.Manifest.schemaSource) -Raw|ConvertFrom-Json).SchemaVersion
Check ($plan.AppVersion -eq $expectedVersion -and $plan.SchemaVersion -eq $expectedSchema) 'Identity differs from canonical source'
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
# Receipt invalidation is file-only; no fake Access application or database is run.
. (Join-Path $PSScriptRoot 'workflow-common.ps1')
$plan=GetSwiftBuildPlan $root $manifestPath
$artifact=[pscustomobject]@{Path=(Join-Path $folder 'receipt-fixture.accdb');SHA256=('A'*64)}
$receiptPath=$artifact.Path+'.build.json'
$validReceipt=[ordered]@{
 Output=$artifact.Path;SHA256=$artifact.SHA256;AppVersion=$plan.AppVersion;SchemaVersion=$plan.SchemaVersion
 Status='Built and structurally verified; workflow tests required'
 Inputs=$plan.InputHashes
 Verification=@{SourceModules=$plan.Manifest.sourceModules.Count;Forms=1;Identity=$true;NativeUnloadGuards=$true;RetiredFormsAbsent=$true}
}
$validJSON=$validReceipt|ConvertTo-Json -Depth 8
Check (-not(TestSwiftBuildReceipt $artifact $plan)) 'Missing receipt was accepted'
'{'|Set-Content -LiteralPath $receiptPath
Check (-not(TestSwiftBuildReceipt $artifact $plan)) 'Malformed receipt was accepted'
$validJSON|Set-Content -LiteralPath $receiptPath
Check (TestSwiftBuildReceipt $artifact $plan) 'Matching complete receipt was rejected'
$changedArtifact=[pscustomobject]@{Path=$artifact.Path;SHA256=('B'*64)}
Check (-not(TestSwiftBuildReceipt $changedArtifact $plan)) 'Changed artifact was accepted'
foreach($key in @('src/modSwift.bas','scripts/apply-contributor-startup.ps1','scripts/verify-built-application.ps1','scripts/vba-source-normalization.ps1')){
 $receipt=$validJSON|ConvertFrom-Json;$receipt.Inputs.$key='changed'
 $receipt|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $receiptPath
 Check (-not(TestSwiftBuildReceipt $artifact $plan)) ('Changed source/generator/checker was accepted: '+$key)
}
foreach($kind in @('missing-input','extra-input','incomplete','string-flag','wrong-version','wrong-path','failed-status','missing-modules','missing-forms')){
 $receipt=$validJSON|ConvertFrom-Json
 switch($kind){
  'missing-input' {$receipt.Inputs.PSObject.Properties.Remove('src/modSwift.bas')}
  'extra-input' {$receipt.Inputs|Add-Member -NotePropertyName 'unexpected' -NotePropertyValue 'hash'}
  'incomplete' {$receipt.Verification.Identity=$false}
  'string-flag' {$receipt.Verification.Identity='true'}
  'wrong-version' {$receipt.SchemaVersion=999}
  'wrong-path' {$receipt.Output=Join-Path $folder 'different.accdb'}
  'failed-status' {$receipt.Status='Failed'}
  'missing-modules' {$receipt.Verification.SourceModules=0}
  'missing-forms' {$receipt.Verification.Forms=0}
 }
 $receipt|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $receiptPath
 Check (-not(TestSwiftBuildReceipt $artifact $plan)) ('Invalid receipt was accepted: '+$kind)
}
Write-Output ('PASS: '+$checks+' build-plan preflight, dependency, identity and input-drift checks')
