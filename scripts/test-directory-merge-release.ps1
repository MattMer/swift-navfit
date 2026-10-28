param([Parameter(Mandatory=$true)][string]$BlankPath,[Parameter(Mandatory=$true)][string]$DemoPath,[switch]$CurrentFixturesOnly)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$logFolder=Join-Path $root ('build/tests/directory-release-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $logFolder)
$results=@()
function NativeCheck([string]$script,[string[]]$arguments){
 $log=Join-Path $logFolder ($script+'.log')
 try{$ErrorActionPreference='Continue'; & powershell.exe -NoProfile -File (Join-Path $root ('scripts/'+$script+'.ps1')) @arguments *> $log; $nativeExit=$LASTEXITCODE}finally{$ErrorActionPreference='Stop'}
 if($nativeExit -ne 0){Get-Content -LiteralPath $log -Tail 35;throw ($script+' failed: '+$log)}
 Get-Content -LiteralPath $log | Where-Object {$_ -like 'PASS*'} | Select-Object -Last 3
 $script:results+=@{Script=$script;Status='PASS';Log=$log}
}
NativeCheck 'test-directory-merge' @('-BuildPath',$BlankPath)
$fixture=(Get-Content (Join-Path $root 'build/directory-merge-verification.json') -Raw|ConvertFrom-Json).Fixture
foreach($script in @('test-directory-history','test-directory-batch','test-ux-directory','test-context-help','smoke-test')){NativeCheck $script @('-BuildPath',$BlankPath)}
NativeCheck 'verify-build-artifact' @('-BuildPath',$BlankPath,'-RequireBlank','-Tier','extended')
foreach($script in @('test-ux-demo','test-report-tools')){NativeCheck $script @('-BuildPath',$DemoPath)}
if(-not $CurrentFixturesOnly){NativeCheck 'test-workspace-upgrade' @('-TemplatePath',$BlankPath,'-SourceWorkspacePath',$DemoPath)}
NativeCheck 'test-directory-merge-transfer' @('-TemplatePath',$BlankPath,'-SourcePath',$fixture)
if($CurrentFixturesOnly){
 NativeCheck 'test-swift-data-import' @('-BuildPath',$BlankPath,'-CurrentSchemaPath',$fixture,'-CurrentFixturesOnly')
}else{
 NativeCheck 'test-swift-data-import' @('-BuildPath',$BlankPath,'-CurrentSchemaPath',$fixture)
}
@{Status='PASS';Blank=$BlankPath;BlankSHA256=(Get-FileHash -LiteralPath $BlankPath).Hash;Demo=$DemoPath;DemoSHA256=(Get-FileHash -LiteralPath $DemoPath).Hash;MergeHistoryFixture=$fixture;Tests=$results}|ConvertTo-Json -Depth 7|Set-Content (Join-Path $root 'build/directory-merge-release-verification.json')
Write-Output ('PASS: directory merge release gates; '+$logFolder)
