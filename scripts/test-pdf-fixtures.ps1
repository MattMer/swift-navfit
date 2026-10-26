# Dot-sourced by PDF runners. NewSwiftPdfFixtures generates a synthetic PDF
# fixture set into a new folder from the blank NAVPERS forms in
# FORMS_REFERENCES_TEMPLATES (read only) and returns the folder. Sets:
# synthetic (synthetic-eval/-chiefeval/-fitrep.pdf), expanded (<kind>-expanded
# and -merge.pdf), counseling (<kind>-NOT-REQ/-NOT-PERF.pdf, built on expanded).
# Needs the development .venv with pikepdf; see docs/Development_Workflow.md.
function NewSwiftPdfFixtures([ValidateSet('synthetic','expanded','counseling')][string]$Set,[string]$Directory){
    $root=Split-Path $PSScriptRoot -Parent
    $python=Join-Path $root '.venv/Scripts/python.exe'
    if(-not (Test-Path -LiteralPath $python)){throw 'PDF fixtures need the development .venv with pikepdf 10.17.0 (docs/Development_Workflow.md).'}
    if(-not $Directory){$Directory=Join-Path $root ('build/tests/pdf-fixtures-'+$Set+'-'+[guid]::NewGuid())}
    $run={param($script,$arguments) $output=& $python (Join-Path $PSScriptRoot $script) @arguments 2>&1; if($LASTEXITCODE -ne 0){throw ($script+' failed: '+($output -join ' '))}}
    switch($Set){
        'synthetic'{& $run 'make-pdf-fixtures.py' @('--output',$Directory)}
        'expanded'{& $run 'make-pdf-expanded-fixtures.py' @('--output',$Directory)}
        'counseling'{$expanded=$Directory+'-expanded';& $run 'make-pdf-expanded-fixtures.py' @('--output',$expanded);& $run 'make-pdf-counseling-status-fixtures.py' @('--source',$expanded,'--output',$Directory)}
    }
    return $Directory
}
