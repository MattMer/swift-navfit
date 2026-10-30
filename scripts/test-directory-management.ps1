param([string]$BuildPath,[ValidateSet('core','extended','release')][string]$Tier='core',[switch]$DisplayFaultOnly,[switch]$AdversarialPreviews,[string]$SourcePath,[string]$CandidatePath,[switch]$AllowDemoSeed,[string]$BlankPath,[string]$DemoPath,[switch]$CurrentFixturesOnly)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
if($Tier -eq 'release'){
    if(-not $BlankPath -or -not $DemoPath){throw 'Directory release mode needs -BlankPath and -DemoPath'}
    $root=Split-Path $PSScriptRoot -Parent
    $blankHash=(Get-FileHash -LiteralPath $BlankPath).Hash
    $demoHash=(Get-FileHash -LiteralPath $DemoPath).Hash
    $logFolder=Join-Path $root ('build/tests/directory-release-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $logFolder)
    $results=@()
    function NativeCheck([string]$script,[string[]]$arguments){
     $log=Join-Path $logFolder ($script+'.log')
     $before=$script:suiteFailures.Count
     if(Test-Path (Join-Path $PSScriptRoot ('cases/'+$script+'.ps1'))){
      Invoke-CaseScript $script $arguments *> $log
     }else{
      Invoke-SuiteCase $script {
       & powershell.exe -NoProfile -File (Join-Path $PSScriptRoot ($script+'.ps1')) @arguments *> $log
       if($LASTEXITCODE -ne 0){throw ($script+' exited '+$LASTEXITCODE)}
      }
     }
     if($script:suiteFailures.Count -gt $before){Get-Content -LiteralPath $log -Tail 35;return}
     Get-Content -LiteralPath $log | Where-Object {$_ -like 'PASS*'} | Select-Object -Last 3
     $script:results+=@{Script=$script;Status='PASS';Log=$log}
    }
    NativeCheck 'test-directory-merge' @('-BuildPath',$BlankPath)
    if($script:suiteFailures.Count){Complete-SuiteRun 'directory-management-release' $BlankPath $blankHash}
    $fixture=(Get-Content (Join-Path $root 'build/directory-merge-verification.json') -Raw|ConvertFrom-Json).Fixture
    foreach($script in @('test-directory-history','test-directory-batch','test-ux-directory','test-report-navigation')){NativeCheck $script @('-BuildPath',$BlankPath)}
    NativeCheck 'verify-build-artifact' @('-BuildPath',$BlankPath,'-RequireBlank','-Tier','extended')
    NativeCheck 'test-workspace-startup' @('-BuildPath',$BlankPath,'-Tier','extended','-DemoPath',$DemoPath)
    NativeCheck 'test-report-review-history' @('-BuildPath',$DemoPath)
    if(-not $CurrentFixturesOnly){NativeCheck 'test-workspace-upgrade' @('-TemplatePath',$BlankPath,'-SourceWorkspacePath',$DemoPath)}
    NativeCheck 'test-directory-merge-transfer' @('-TemplatePath',$BlankPath,'-SourcePath',$fixture)
    if($CurrentFixturesOnly){
     NativeCheck 'test-swift-data-import' @('-BuildPath',$BlankPath,'-CurrentSchemaPath',$fixture,'-CurrentFixturesOnly')
    }else{
     NativeCheck 'test-swift-data-import' @('-BuildPath',$BlankPath,'-CurrentSchemaPath',$fixture)
    }
    if((Get-FileHash -LiteralPath $DemoPath).Hash -ne $demoHash){$script:suiteFailures+='suite changed its demo artifact'}
    Complete-SuiteRun 'directory-management-release' $BlankPath $blankHash
    @{Status='PASS';Blank=$BlankPath;BlankSHA256=(Get-FileHash -LiteralPath $BlankPath).Hash;Demo=$DemoPath;DemoSHA256=(Get-FileHash -LiteralPath $DemoPath).Hash;MergeHistoryFixture=$fixture;Tests=$results}|ConvertTo-Json -Depth 7|Set-Content (Join-Path $root 'build/directory-merge-release-verification.json')
    Write-Output ('PASS: directory merge release gates; '+$logFolder)
    return
}
if(-not $BuildPath){throw 'Directory core/extended mode needs -BuildPath'}
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('DisplayFaultOnly')){if($DisplayFaultOnly){$caseArgs+='-DisplayFaultOnly'}}
Invoke-CaseScript 'test-directory-merge' $caseArgs
if($Tier -eq 'extended'){
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('AdversarialPreviews')){if($AdversarialPreviews){$caseArgs+='-AdversarialPreviews'}}
Invoke-CaseScript 'test-directory-batch' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-directory-history' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-directory' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-group-directory' $caseArgs
if($SourcePath -and $CandidatePath){
$caseArgs=@()
if($PSBoundParameters.ContainsKey('SourcePath')){$caseArgs+=@('-SourcePath',$SourcePath)}
if($PSBoundParameters.ContainsKey('CandidatePath')){$caseArgs+=@('-CandidatePath',$CandidatePath)}
if($PSBoundParameters.ContainsKey('AllowDemoSeed')){if($AllowDemoSeed){$caseArgs+='-AllowDemoSeed'}}
Invoke-CaseScript 'test-ux-group-directory-data' $caseArgs
}else{Write-Output 'SKIP: test-ux-group-directory-data needs -SourcePath -CandidatePath'}
}
Complete-SuiteRun 'directory-management' $BuildPath $buildHash
