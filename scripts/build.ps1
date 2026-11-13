param([string]$OutputPath,[string]$ManifestPath = (Join-Path $PSScriptRoot 'build-manifest.json'))
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'build-plan.ps1')
$swiftBuildPlan = GetSwiftBuildPlan $root $ManifestPath
if(-not $OutputPath){$OutputPath=Join-Path $root ('output/Swift_NAVFIT_'+$swiftBuildPlan.AppVersion+'.accdb')}
$swiftBuildTrace=[Collections.Generic.List[object]]::new()
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
if (Test-Path -LiteralPath $OutputPath) { throw "Refusing to overwrite existing database: $OutputPath" }
[void](New-Item -ItemType Directory -Force -Path (Split-Path $OutputPath -Parent))
$app = $null
$script:swiftManagedBuild=$true
try {
    $app = New-Object -ComObject Access.Application
    $app.NewCurrentDatabase($OutputPath, 12)
    $swiftBaseDatabase = $app.CurrentDb()
    $canonicalFields=@((Get-Content (Join-Path $root 'src/field-definitions.json') -Raw -Encoding UTF8|ConvertFrom-Json).fields|Sort-Object ordinal)
    $majorSchema=Get-Content (Join-Path $root 'src/major-schema.json') -Raw | ConvertFrom-Json
    #Every table is defined in major-schema.json: BaseSQL first, then SQL. The Report and
    #RecoveryDraft field columns come from the field catalogue.
    $reportColumns=($canonicalFields | ForEach-Object { '['+$_.key+'] '+$_.storageType+$(if($_.key -eq 'FormKind'){' NOT NULL'}) }) -join ','
    $recoveryColumns=($canonicalFields | ForEach-Object { '['+$_.recoveryKey+'] '+$(if($_.key -eq 'FormKind'){'TEXT(12)'}else{'LONGTEXT'}) }) -join ','
    $definitions = @($majorSchema.BaseSQL | ForEach-Object { $_.Replace('{ReportFieldColumns}',$reportColumns).Replace('{RecoveryFieldColumns}',$recoveryColumns) })
    $definitions += "INSERT INTO AppMetadata VALUES (1,'$($swiftBuildPlan.AppVersion)',$($swiftBuildPlan.SchemaVersion),'Development candidate - native validation/printing remains external')"
    $definitions += $majorSchema.SQL
    foreach ($sql in $definitions) { try { $swiftBaseDatabase.Execute($sql,128) } catch { throw ($sql + "`n" + $_.Exception.Message) } }
    foreach($definition in $majorSchema.Queries){
        $query=$swiftBaseDatabase.CreateQueryDef($definition.Name,$definition.SQL)
        $query.Properties.Append($query.CreateProperty('Description',10,$definition.Description))
        $query.Close()
    }
    $swiftBaseDatabase.TableDefs('RevisionDate').ValidationRule = "[FieldKey] In ('DateReported','CounselingDate','RaterDate','SeniorRaterDate','ReadinessDate') And (([DateState]='Date' And [DateValue] Is Not Null) Or ([DateState] In ('Missing','Blank','Unknown') And [DateValue] Is Null) Or ([FieldKey]='CounselingDate' And [DateState] In ('NOT REQ','NOT PERF') And [DateValue] Is Null))"
    $swiftBaseDatabase.TableDefs('RevisionChoice').ValidationRule = "[ChoiceKind] In ('Retention','Statement') And [ChoiceValue] In ('Yes','No','Unspecified','Conflict','Unknown')"
    $auditPointers=@('RevisionSourceID','ReportEventSourceID','FolderMoveSourceID','CommentEventSourceID','TaskEventSourceID','JournalSourceID','MessageSourceID')
    $auditKinds=@('Revision','Report event','Folder move','Comment event','Review event','Operation','Message')
    $auditCases=for($ai=0;$ai -lt $auditPointers.Count;$ai++){"([SourceKind]='"+$auditKinds[$ai]+"' And ["+$auditPointers[$ai]+'] Is Not Null)'}
    $auditCount=($auditPointers|ForEach-Object {'IIf(['+$_+'] Is Null,0,1)'}) -join '+'
    $swiftBaseDatabase.TableDefs('AuditEvent').ValidationRule='[SourceVersion]>=0 And ('+$auditCount+')=1 And ('+($auditCases -join ' Or ')+')'
    $swiftBaseDatabase.TableDefs('Report').ValidationRule = "([FormKind] In ('EVAL','CHIEFEVAL','FITREP') Or ([LegacyFormKind] Is Not Null And [FormKind]=[LegacyFormKind])) And IIf([WorkflowState] Is Null,[LegacyWorkflowNull]=True,([WorkflowState] In ('Draft','Approved','Signed') Or ([LegacyWorkflowState] Is Not Null And [WorkflowState]=[LegacyWorkflowState])))"
    $swiftBaseDatabase.TableDefs('FieldDefaultLink').ValidationRule = "([LinkState]='Unresolved' And [CommandWorkspaceID] Is Null And [FolderID] Is Null And [GroupID] Is Null And [ReportID] Is Null) Or ([LinkState]='Verified' And (IIf([CommandWorkspaceID] Is Null,0,1)+IIf([FolderID] Is Null,0,1)+IIf([GroupID] Is Null,0,1)+IIf([ReportID] Is Null,0,1))=1)"
    $swiftBaseDatabase.TableDefs('GroupCompletion').ValidationRule = '[TraitCount]>=0 And [TraitSum]>=0'
    $swiftBaseDatabase.TableDefs('GroupLifecycle').ValidationRule = "[WorkflowState] In ('Initialized','In-Work','Ready for Signature','Completed') Or ([LegacyState] Is Not Null And [WorkflowState]=[LegacyState])"
    $swiftBaseDatabase.TableDefs('GroupCompletionReport').ValidationRule = '([RevisionID] Is Null And [RevisionOwnerReportID] Is Null And [TraitCount] Is Null And [TraitSum] Is Null) Or ([RevisionID] Is Not Null And [RevisionOwnerReportID]=[ReportID] And [TraitCount]>=0 And [TraitSum]>=[TraitCount] And [TraitSum]<=5*[TraitCount])'
    $swiftBaseDatabase.TableDefs('ReportSignoff').ValidationRule = "[PriorState] In ('Draft','Approved') And (([ReopenedBy] Is Null And [ReopenedAt] Is Null And [ReopenReason] Is Null) Or ([ReopenedBy] Is Not Null And [ReopenedAt] Is Not Null))"
    $swiftBaseDatabase.TableDefs('MemberCounseling').ValidationRule = "[Outcome] In ('Completed','Not performed') And ([Outcome]<>'Completed' Or ([CounselorName] Is Not Null And [NotPerformedReason] Is Null)) And ([Outcome]<>'Not performed' Or [NotPerformedReason] Is Not Null) And (([WorksheetDisposition] Is Null And [DispositionOn] Is Null) Or ([WorksheetDisposition] In ('Given to member','Destroyed') And [DispositionOn] Is Not Null))"
    for($ti=0;$ti -lt $swiftBaseDatabase.TableDefs.Count;$ti++) {
        $table=$swiftBaseDatabase.TableDefs.Item($ti)
        if($table.Name.StartsWith('MSys')) {continue}
        for($fi=0;$fi -lt $table.Fields.Count;$fi++) {
            $field=$table.Fields.Item($fi)
            if(($field.Type -eq 10 -or $field.Type -eq 12) -and -not $field.Required) {$field.AllowZeroLength=$true}
        }
    }
    $swiftBaseDatabase.Close(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($swiftBaseDatabase); $swiftBaseDatabase=$null
    [void][IO.Directory]::CreateDirectory((Join-Path $root 'build'))
    foreach($sourceModule in $swiftBuildPlan.Manifest.sourceModules){
        $app.LoadFromText(5,[IO.Path]::GetFileNameWithoutExtension($sourceModule),(Join-Path $root $sourceModule))
    }
    $navTraitModule=Join-Path $root ('build/trait-catalogue-'+[guid]::NewGuid()+'.bas')
    & (Join-Path $PSScriptRoot 'write-trait-catalogue.ps1') -SourcePath (Join-Path $root 'src/trait-catalogue.json') -OutputPath $navTraitModule -Application $app
    $app.LoadFromText(5,'modTraitCatalogue',$navTraitModule)
    # Resources: the generated modules, report and properties that are not forms.
    foreach($swiftBuildStage in @($swiftBuildPlan.Manifest.stages)){
        $swiftStageTimer=[Diagnostics.Stopwatch]::StartNew()
        Write-Output ('Build resource: '+$swiftBuildStage.id)
        try{. (Join-Path $root $swiftBuildStage.file)}catch{throw ('Build resource '+$swiftBuildStage.id+' failed: '+$_.Exception.Message+"`n"+$_.ScriptStackTrace)}
        $swiftBuildTrace.Add([pscustomobject]@{Id=$swiftBuildStage.id;Phase=$swiftBuildStage.phase;File=$swiftBuildStage.file;Seconds=$swiftStageTimer.Elapsed.TotalSeconds;Status='PASS'})
    }
    # Forms are source: every form loads from its normalized text in src/forms.
    . (Join-Path $PSScriptRoot 'install-form-text.ps1')
    $swiftFormTimer=[Diagnostics.Stopwatch]::StartNew()
    $swiftFormCount=InstallSwiftFormText
    $swiftBuildTrace.Add([pscustomobject]@{Id='src/forms';Phase='forms';File='src/forms';Seconds=$swiftFormTimer.Elapsed.TotalSeconds;Status='PASS';Forms=$swiftFormCount})
    $identityDB=$app.CurrentDb()
    $identityDB.Execute("UPDATE AppMetadata SET SchemaVersion=$($swiftBuildPlan.SchemaVersion),AppVersion='$($swiftBuildPlan.AppVersion)'",128)
    $identityDB.Close()
    $macro=Join-Path $root 'src\AutoExec.txt'
    $app.LoadFromText(4,'AutoExec',$macro)
    $app.DoCmd.RunCommand(126)
    # Hidden Access skips past a compile error without saying so; fail the build instead.
    if(-not $app.IsCompiled){throw 'VBA source does not compile. Open the build in Access and use Debug > Compile to find the line.'}
    $app.CloseCurrentDatabase()
    $app.OpenCurrentDatabase($OutputPath,$true)
    . (Join-Path $PSScriptRoot 'verify-built-application.ps1')
    $app.CloseCurrentDatabase()

} finally {
    $script:swiftManagedBuild=$false
    if ($null -ne $app) { $app.Quit(2) }
    # Generated forms/modules and CurrentDb clones can outlive Application.Quit.
    # Release only COM references owned by this build's script scope before hashing.
    function ReleaseSwiftBuildCom([object]$Value){
        if($null -eq $Value){return}
        if([Runtime.InteropServices.Marshal]::IsComObject($Value)){
            try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)}catch{}
        }elseif($Value -is [Array]){
            for($index=0;$index -lt $Value.Length;$index++){ReleaseSwiftBuildCom -Value $Value.GetValue($index)}
        }
    }
    foreach($buildVariable in @(Get-Variable -Scope Script)){
        if($buildVariable.Name -eq 'app'){continue}
        ReleaseSwiftBuildCom -Value $buildVariable.Value
    }
    if ($null -ne $app) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
}
$builtHash=$null
for($attempt=0;$attempt -lt 120;$attempt++){
    try{$builtHash=(Get-FileHash -LiteralPath $OutputPath -Algorithm SHA256 -ErrorAction Stop).Hash;break}
    catch{if($attempt -eq 119){throw};Start-Sleep -Milliseconds 250}
}
AssertSwiftBuildInputs $swiftBuildPlan
@{Output=$OutputPath;SHA256=$builtHash;AppVersion=$swiftBuildPlan.AppVersion;SchemaVersion=$swiftBuildPlan.SchemaVersion;Inputs=$swiftBuildPlan.InputHashes;Stages=@($swiftBuildTrace.ToArray());Verification=$swiftBuildVerification;Status='Built and structurally verified; workflow tests required'} | ConvertTo-Json -Depth 8 | Set-Content ($OutputPath+'.build.json')
[pscustomobject]@{Output=$OutputPath;SHA256=$builtHash;Receipt=($OutputPath+'.build.json');Status='Built and structurally verified; workflow tests required'} | ConvertTo-Json
