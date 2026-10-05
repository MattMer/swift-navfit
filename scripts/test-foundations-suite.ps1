param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$results=[Collections.Generic.List[object]]::new()
$scripts=@('smoke-test.ps1','test-administration.ps1','test-traits.ps1','test-trait-summary.ps1','test-assessment.ps1','test-promotion-status.ps1','test-reporting-senior.ps1','test-report-context.ps1','test-member-details.ps1','test-service-details.ps1','test-decision-details.ps1','test-report-people.ps1','test-party-names.ps1','test-routing.ps1','test-review-comments.ps1','test-provenance.ps1','test-narratives.ps1','test-revision-compare.ps1','test-recovery.ps1','test-import-review.ps1','test-import-commit.ps1','test-major.ps1','test-rules.ps1','test-pdf-import.ps1')
foreach($name in $scripts){
    $log=Join-Path $root ('build/tests/foundations-'+$name.Replace('.ps1','.log'))
    & powershell.exe -NoProfile -File (Join-Path $PSScriptRoot $name) -BuildPath $BuildPath *> $log
    $status=if($LASTEXITCODE -eq 0){'PASS'}else{'FAIL'}
    $results.Add([pscustomobject]@{Suite=$name;Status=$status;Log=$log})
    Write-Output ($status+': '+$name)
    if($status -eq 'FAIL'){Get-Content -LiteralPath $log -Tail 12;throw ('Suite failed: '+$name)}
}
$results|ConvertTo-Json -Depth 3|Set-Content -LiteralPath (Join-Path $root 'build/tests/foundations-suite-results.json') -Encoding utf8
