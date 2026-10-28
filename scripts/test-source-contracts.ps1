param([string]$BuildPath,[ValidateSet('core','extended')][string]$Tier='core')
# Source and build-contract suite. Without -BuildPath: the static checks (build plan
# and runner inventory, VBA normalization, canonical field definitions). With
# -BuildPath: the architecture views against that build; -Tier extended adds the
# native field-catalogue recovery and the diagnostics contracts.
# Each case is a script in scripts/cases, run in its own process.
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$cases=Join-Path $PSScriptRoot 'cases'
function Invoke-CaseScript([string]$Name,[string[]]$Arguments=@()){
    Invoke-SuiteCase $Name {
        & powershell.exe -NoProfile -File (Join-Path $cases ($Name+'.ps1')) @Arguments
        if($LASTEXITCODE -ne 0){throw ($Name+' exited '+$LASTEXITCODE)}
    }
}
if(-not $BuildPath){
    Invoke-CaseScript 'test-build-plan'
    Invoke-CaseScript 'test-vba-source-normalization'
    Invoke-CaseScript 'test-field-definitions'
    if($script:suiteFailures.Count){throw ('source-contracts: '+($script:suiteFailures -join ' | '))}
    Write-Output ('PASS: source-contracts suite, '+$script:suiteCases+' static cases')
    exit 0
}
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
Invoke-CaseScript 'test-architecture-views' @('-BuildPath',$BuildPath)
if($Tier -eq 'extended'){
    Invoke-CaseScript 'test-field-definitions' @('-BuildPath',$BuildPath)
    Invoke-CaseScript 'test-diagnostics-contracts' @('-BuildPath',$BuildPath)
}
Complete-SuiteRun 'source-contracts' $BuildPath $buildHash
