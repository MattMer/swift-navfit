param([Parameter(Mandatory=$true)][string]$BuildPath,[Parameter(Mandatory=$true)][string]$LegacyTemplate)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$folder=Join-Path $root ('build/tests/persistence-transfer-'+[guid]::NewGuid())
[void](New-Item -ItemType Directory -Path $folder)
$originals=@{}
foreach($path in @($BuildPath,$LegacyTemplate)){$originals[[IO.Path]::GetFullPath($path)]=(Get-FileHash -LiteralPath $path).Hash}
$results=@()
foreach($template in @($LegacyTemplate,$BuildPath)){
    $tag=if($template -eq $LegacyTemplate){'legacy27'}else{'current28'}
    $source=Join-Path $folder ($tag+'-source.accdb')
    Copy-Item -LiteralPath $template -Destination $source
    $app=New-Object -ComObject Access.Application
    try {
        $app.AutomationSecurity=1;$app.OpenCurrentDatabase($source,$true)
        $schema=$app.Eval('DLookup("SchemaVersion","AppMetadata")')
        if(($tag -eq 'legacy27' -and $schema -ne 27) -or ($tag -eq 'current28' -and $schema -ne 28)){throw 'Use real schema-27 and schema-28 templates'}
        $app.LoadFromText(5,'PersistenceSeedHarness',(Join-Path $root 'tests/PersistenceSeedHarness.bas'))
        $app.DoCmd.RunCommand(126)
        $seed=$app.Eval('SeedPersistenceTransfer()');Write-Output $seed
        if($seed -notlike 'PASS:*'){throw $seed}
    } finally {$app.CloseCurrentDatabase();$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null;[GC]::Collect();[GC]::WaitForPendingFinalizers()}
    $hash=(Get-FileHash -LiteralPath $source).Hash
    $target=Join-Path $folder ($tag+'-target.accdb');Copy-Item -LiteralPath $BuildPath -Destination $target
    $app=New-Object -ComObject Access.Application
    try {
        $app.AutomationSecurity=1;$app.OpenCurrentDatabase($target,$true)
        [void]$app.Eval('ConfigureWorkspace("SYNTHETIC BOOTSTRAP","Tester")')
        $app.LoadFromText(5,'WorkspaceImportHarness',(Join-Path $root 'tests/WorkspaceImportHarness.bas'))
        $app.DoCmd.RunCommand(126)
        if(-not $app.Eval('UXPreviewSwiftWorkspace("'+$source+'")')){throw 'Settings import preview failed'}
        if($tag -eq 'legacy27'){
            $db=$app.CurrentDb();$db.TableDefs('RevisionDate').Fields('DateState').ValidationRule="='Reject migration facts'";$db.Close()
            $failure=$app.Eval('TryWorkspaceImport("'+$source+'")')
            if($failure -notlike 'REJECTED:*' -or $app.Eval('DCount("*","Report")') -ne 0 -or $app.Eval('DCount("*","GroupCompletion")') -ne 0 -or $app.Eval('DLookup("CommandName","Workspace")') -ne 'SYNTHETIC BOOTSTRAP'){throw ('Late migration rollback failed: '+$failure)}
            $db=$app.CurrentDb();$db.TableDefs('RevisionDate').Fields('DateState').ValidationRule='';$db.Close()
        }
        if(-not $app.Eval('UXApplySwiftWorkspace()')){throw $app.Eval('Forms!frmUXWorkspaceImport!lblStatus.Caption')}
        $status=$app.Eval('Forms!frmUXWorkspaceImport!lblStatus.Caption')
        if($app.Eval('DLookup("DateState","RevisionDate","FieldKey=''DateReported'' AND DateState=''Unknown''")') -ne 'Unknown'){throw 'Unknown legacy date was not flagged'}
        if($app.Eval('DLookup("ChoiceValue","RevisionChoice","ChoiceKind=''Retention'' AND ChoiceValue=''Conflict''")') -ne 'Conflict'){throw 'Conflicting raw flags were coerced'}
        if($app.Eval('DLookup("StartEntry","RecoveryDraft")') -ne 'raw pending date'){throw 'Raw recovery values changed'}
        if($app.Eval('DCount("*","Report","FormKind=''LEGACY'' AND WorkflowState Is Null")') -ne 1){throw 'Unknown form kind/null workflow was coerced or lost'}
        if($app.Eval('DLookup("WorkflowState","GroupLifecycle")') -ne 'In-Work'){throw 'Reopened workflow state lost'}
        if($app.Eval('GroupPolicyValue(1,"RSCAGroupSum")') -ne '28' -or $app.Eval('GroupPolicyValue(1,"RSCAGroupCount")') -ne '7'){throw 'Posting totals changed'}
        $count=$app.Eval('DCount("*","GroupCompletion")')
        [void]$app.Run('MigrateRSAccounting');[void]$app.Run('MigrateReportState')
        if($app.Eval('DCount("*","GroupCompletion")') -ne $count){throw 'Migration retry duplicated a posting'}
        $reconciliation=$app.Eval('RSReconcileAccounting()')
        if($reconciliation -notlike 'PASS:*'){throw $reconciliation}
        $app.CloseCurrentDatabase();$app.OpenCurrentDatabase($target,$true)
        if($app.Eval('DCount("*","ReportHead")') -ne 2){throw 'Reopened authoritative state missing'}
        if((Get-FileHash -LiteralPath $source).Hash -ne $hash){throw 'Production import changed source'}
        $results+=@{Schema=$schema;Status=$status;Source=$source;SHA256=$hash;Target=$target;Reconciliation=$reconciliation;Reopened=$true;LateRollbackTested=($tag -eq 'legacy27')}
        Write-Output ('PASS: schema '+$schema+' Settings import, exact original-value verification, raw recovery, typed review flags, idempotency and reopen')
    } finally {$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
foreach($path in $originals.Keys){if((Get-FileHash -LiteralPath $path).Hash -ne $originals[$path]){throw 'A supplied template changed'}}
$results|ConvertTo-Json -Depth 8|Set-Content (Join-Path $folder 'verification.json') -Encoding UTF8
Write-Output ('PASS: schema-27/current migration and rollback; evidence: '+$folder)
