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
    $definitions = @(
        'CREATE TABLE AppMetadata (MetadataID LONG CONSTRAINT PK_AppMetadata PRIMARY KEY, AppVersion TEXT(40), SchemaVersion LONG, CompatibilityStatus TEXT(120))',
        'CREATE TABLE UserProfile (UserProfileID COUNTER CONSTRAINT PK_UserProfile PRIMARY KEY, ProfileGuid TEXT(38) NOT NULL, DisplayName TEXT(120) NOT NULL)',
        'CREATE TABLE Workspace (WorkspaceID LONG CONSTRAINT PK_Workspace PRIMARY KEY, WorkspaceGuid TEXT(38) NOT NULL, CommandName TEXT(120) NOT NULL, ActiveProfileID LONG NOT NULL)',
        'CREATE TABLE Folder (FolderID COUNTER CONSTRAINT PK_Folder PRIMARY KEY, FolderGuid TEXT(38) NOT NULL, FolderName TEXT(80) NOT NULL, ParentFolderID LONG)',
        ('CREATE TABLE Report (ReportID COUNTER CONSTRAINT PK_Report PRIMARY KEY, ReportGuid TEXT(38) NOT NULL, WorkflowState TEXT(40), IsArchived BIT, CreatedAt DATETIME, ModifiedAt DATETIME, CurrentRevisionID LONG, FolderID LONG,' + (($canonicalFields | ForEach-Object { '['+$_.key+'] '+$_.storageType+$(if($_.key -eq 'FormKind'){' NOT NULL'}) }) -join ',') + ')'),
        'CREATE TABLE FolderMove (FolderMoveID COUNTER CONSTRAINT PK_FolderMove PRIMARY KEY, ReportID LONG NOT NULL, OldFolderID LONG, NewFolderID LONG, AuthorProfileID LONG NOT NULL, MovedAt DATETIME)',
        'CREATE TABLE ReportRevision (RevisionID COUNTER CONSTRAINT PK_ReportRevision PRIMARY KEY, ReportID LONG NOT NULL, RevisionNumber LONG NOT NULL, ParentRevisionID LONG, AuthorProfileID LONG NOT NULL, CommittedAt DATETIME, Reason TEXT(255), SnapshotSchema LONG)',
        'CREATE TABLE RevisionValue (RevisionValueID COUNTER CONSTRAINT PK_RevisionValue PRIMARY KEY, RevisionID LONG NOT NULL, FieldKey TEXT(64) NOT NULL, ValueText LONGTEXT, IsNullValue BIT)',
        'CREATE TABLE RevisionChange (RevisionChangeID COUNTER CONSTRAINT PK_RevisionChange PRIMARY KEY, RevisionID LONG NOT NULL, FieldKey TEXT(64) NOT NULL, OldValue LONGTEXT, NewValue LONGTEXT, OldIsNull BIT, NewIsNull BIT)',
        'CREATE TABLE ReportEvent (ReportEventID COUNTER CONSTRAINT PK_ReportEvent PRIMARY KEY, ReportID LONG NOT NULL, AuthorProfileID LONG NOT NULL, EventAt DATETIME, EventKind TEXT(40))',
        'CREATE TABLE OperationJournal (OperationID COUNTER CONSTRAINT PK_OperationJournal PRIMARY KEY, OperationGuid TEXT(38) NOT NULL, AuthorProfileID LONG NOT NULL, OperationKind TEXT(40), OperationState TEXT(20), SourceDescription LONGTEXT, StartedAt DATETIME, FinishedAt DATETIME, Outcome LONGTEXT)',
        'CREATE TABLE NarrativeSnippet (SnippetID COUNTER CONSTRAINT PK_NarrativeSnippet PRIMARY KEY, SnippetGuid TEXT(38) NOT NULL, CreatedBy LONG NOT NULL, SnippetTitle TEXT(120), SnippetBody LONGTEXT, Attribution TEXT(255), FormKind TEXT(12), ModifiedAt DATETIME)',
        'CREATE TABLE LibraryPreference (LibraryPreferenceID COUNTER CONSTRAINT PK_LibraryPreference PRIMARY KEY, AuthorProfileID LONG NOT NULL, ViewPreset TEXT(12))',
        'CREATE TABLE FolderBatchChange (FolderBatchChangeID COUNTER CONSTRAINT PK_FolderBatchChange PRIMARY KEY, OperationID LONG NOT NULL, ReportID LONG NOT NULL, OldFolderID LONG, NewFolderID LONG, AppliedMoveID LONG NOT NULL, BaseRevisionID LONG NOT NULL, IsUndone BIT)',
        'ALTER TABLE FolderBatchChange ADD CONSTRAINT FK_BatchOperation FOREIGN KEY (OperationID) REFERENCES OperationJournal (OperationID)',
        'ALTER TABLE FolderBatchChange ADD CONSTRAINT FK_BatchReport FOREIGN KEY (ReportID) REFERENCES Report (ReportID)',
        'ALTER TABLE FolderBatchChange ADD CONSTRAINT FK_BatchMove FOREIGN KEY (AppliedMoveID) REFERENCES FolderMove (FolderMoveID)',
        'ALTER TABLE FolderBatchChange ADD CONSTRAINT FK_BatchRevision FOREIGN KEY (BaseRevisionID) REFERENCES ReportRevision (RevisionID)',
        'CREATE UNIQUE INDEX UX_LibraryProfile ON LibraryPreference (AuthorProfileID)',
        'ALTER TABLE LibraryPreference ADD CONSTRAINT FK_LibraryProfile FOREIGN KEY (AuthorProfileID) REFERENCES UserProfile (UserProfileID)',
        'ALTER TABLE NarrativeSnippet ADD CONSTRAINT FK_SnippetProfile FOREIGN KEY (CreatedBy) REFERENCES UserProfile (UserProfileID)',
        'ALTER TABLE OperationJournal ADD CONSTRAINT FK_OperationProfile FOREIGN KEY (AuthorProfileID) REFERENCES UserProfile (UserProfileID)',
        'CREATE TABLE ImportReceipt (ImportReceiptID COUNTER CONSTRAINT PK_ImportReceipt PRIMARY KEY, ReceiptGuid TEXT(38) NOT NULL, SourceHash TEXT(64) NOT NULL, SourcePath LONGTEXT, ExternalReportID LONG NOT NULL, ReportID LONG NOT NULL, AuthorProfileID LONG NOT NULL, ImportedAt DATETIME, AdapterVersion TEXT(40), ValueEncoding TEXT(80), BeforeRevisionID LONG, AppliedRevisionID LONG, OperationID LONG, IsUndone BIT)',
        'CREATE TABLE ImportRawValue (ImportRawValueID COUNTER CONSTRAINT PK_ImportRawValue PRIMARY KEY, ImportReceiptID LONG NOT NULL, SourceTable TEXT(40) NOT NULL, SourceRow LONG NOT NULL, FieldOrdinal LONG NOT NULL, FieldKey TEXT(64) NOT NULL, DAOType LONG NOT NULL, RawValue LONGTEXT, IsNullValue BIT)',
        'CREATE TABLE ReviewComment (CommentID COUNTER CONSTRAINT PK_ReviewComment PRIMARY KEY, CommentGuid TEXT(38) NOT NULL, ReportID LONG NOT NULL, RevisionID LONG NOT NULL, ThreadID LONG, AuthorProfileID LONG NOT NULL, CommentAt DATETIME, CommentBody LONGTEXT)',
        'CREATE UNIQUE INDEX UX_CommentGuid ON ReviewComment (CommentGuid)',
        'ALTER TABLE ReviewComment ADD CONSTRAINT FK_CommentReport FOREIGN KEY (ReportID) REFERENCES Report (ReportID)',
        'ALTER TABLE ReviewComment ADD CONSTRAINT FK_CommentRevision FOREIGN KEY (RevisionID) REFERENCES ReportRevision (RevisionID)',
        'ALTER TABLE ReviewComment ADD CONSTRAINT FK_CommentThread FOREIGN KEY (ThreadID) REFERENCES ReviewComment (CommentID)',
        'ALTER TABLE ReviewComment ADD CONSTRAINT FK_CommentAuthor FOREIGN KEY (AuthorProfileID) REFERENCES UserProfile (UserProfileID)',
        'CREATE TABLE ReviewCommentEvent (CommentEventID COUNTER CONSTRAINT PK_ReviewCommentEvent PRIMARY KEY, ThreadID LONG NOT NULL, AuthorProfileID LONG NOT NULL, EventAt DATETIME, EventKind TEXT(20))',
        'ALTER TABLE ReviewCommentEvent ADD CONSTRAINT FK_CommentEventThread FOREIGN KEY (ThreadID) REFERENCES ReviewComment (CommentID)',
        'ALTER TABLE ReviewCommentEvent ADD CONSTRAINT FK_CommentEventAuthor FOREIGN KEY (AuthorProfileID) REFERENCES UserProfile (UserProfileID)',
        'CREATE TABLE ReviewTask (TaskID COUNTER CONSTRAINT PK_ReviewTask PRIMARY KEY, TaskGuid TEXT(38) NOT NULL, ReportID LONG NOT NULL, RevisionID LONG NOT NULL, RequestedBy LONG NOT NULL, AssignedTo LONG NOT NULL, RequestedAt DATETIME, DueDate DATETIME, Instructions LONGTEXT, TaskState TEXT(20), FinishedAt DATETIME)',
        'CREATE UNIQUE INDEX UX_TaskGuid ON ReviewTask (TaskGuid)',
        'ALTER TABLE ReviewTask ADD CONSTRAINT FK_TaskReport FOREIGN KEY (ReportID) REFERENCES Report (ReportID)',
        'ALTER TABLE ReviewTask ADD CONSTRAINT FK_TaskRevision FOREIGN KEY (RevisionID) REFERENCES ReportRevision (RevisionID)',
        'ALTER TABLE ReviewTask ADD CONSTRAINT FK_TaskRequester FOREIGN KEY (RequestedBy) REFERENCES UserProfile (UserProfileID)',
        'ALTER TABLE ReviewTask ADD CONSTRAINT FK_TaskAssignee FOREIGN KEY (AssignedTo) REFERENCES UserProfile (UserProfileID)',
        'CREATE TABLE ReviewTaskEvent (TaskEventID COUNTER CONSTRAINT PK_ReviewTaskEvent PRIMARY KEY, TaskID LONG NOT NULL, AuthorProfileID LONG NOT NULL, EventAt DATETIME, EventKind TEXT(20), Notes LONGTEXT)',
        'ALTER TABLE ReviewTaskEvent ADD CONSTRAINT FK_TaskEventTask FOREIGN KEY (TaskID) REFERENCES ReviewTask (TaskID)',
        'ALTER TABLE ReviewTaskEvent ADD CONSTRAINT FK_TaskEventAuthor FOREIGN KEY (AuthorProfileID) REFERENCES UserProfile (UserProfileID)',
        ('CREATE TABLE RecoveryDraft (RecoveryDraftID COUNTER CONSTRAINT PK_RecoveryDraft PRIMARY KEY, SessionGuid TEXT(38) NOT NULL, AuthorProfileID LONG NOT NULL, ReportID LONG, BaseRevisionID LONG, CheckpointAt DATETIME,' + (($canonicalFields | ForEach-Object { '['+$_.recoveryKey+'] '+$(if($_.key -eq 'FormKind'){'TEXT(12)'}else{'LONGTEXT'}) }) -join ',') + ')'),
        'CREATE UNIQUE INDEX UX_RecoverySession ON RecoveryDraft (SessionGuid)',
        'ALTER TABLE RecoveryDraft ADD CONSTRAINT FK_RecoveryAuthor FOREIGN KEY (AuthorProfileID) REFERENCES UserProfile (UserProfileID)',
        'ALTER TABLE RecoveryDraft ADD CONSTRAINT FK_RecoveryReport FOREIGN KEY (ReportID) REFERENCES Report (ReportID)',
        'ALTER TABLE RecoveryDraft ADD CONSTRAINT FK_RecoveryRevision FOREIGN KEY (BaseRevisionID) REFERENCES ReportRevision (RevisionID)',
        'ALTER TABLE ImportReceipt ADD CONSTRAINT FK_ImportOperation FOREIGN KEY (OperationID) REFERENCES OperationJournal (OperationID)',
        'ALTER TABLE ImportReceipt ADD CONSTRAINT FK_ImportBefore FOREIGN KEY (BeforeRevisionID) REFERENCES ReportRevision (RevisionID)',
        'ALTER TABLE ImportReceipt ADD CONSTRAINT FK_ImportApplied FOREIGN KEY (AppliedRevisionID) REFERENCES ReportRevision (RevisionID)',
        'CREATE UNIQUE INDEX UX_ImportSource ON ImportReceipt (SourceHash, ExternalReportID)',
        'ALTER TABLE ImportReceipt ADD CONSTRAINT FK_ImportReport FOREIGN KEY (ReportID) REFERENCES Report (ReportID)',
        'ALTER TABLE ImportReceipt ADD CONSTRAINT FK_ImportAuthor FOREIGN KEY (AuthorProfileID) REFERENCES UserProfile (UserProfileID)',
        'ALTER TABLE ImportRawValue ADD CONSTRAINT FK_ImportRawReceipt FOREIGN KEY (ImportReceiptID) REFERENCES ImportReceipt (ImportReceiptID)',
        'CREATE UNIQUE INDEX UX_ImportRaw ON ImportRawValue (ImportReceiptID,SourceTable,SourceRow,FieldOrdinal)',
        'CREATE UNIQUE INDEX UX_ReportGuid ON Report (ReportGuid)',
        'CREATE UNIQUE INDEX UX_ProfileGuid ON UserProfile (ProfileGuid)',
        'CREATE UNIQUE INDEX UX_WorkspaceGuid ON Workspace (WorkspaceGuid)',
        'CREATE UNIQUE INDEX UX_ReportRevision ON ReportRevision (ReportID, RevisionNumber)',
        'CREATE UNIQUE INDEX UX_RevisionValue ON RevisionValue (RevisionID, FieldKey)',
        'ALTER TABLE Workspace ADD CONSTRAINT FK_WorkspaceProfile FOREIGN KEY (ActiveProfileID) REFERENCES UserProfile (UserProfileID)',
        'ALTER TABLE ReportRevision ADD CONSTRAINT FK_RevisionReport FOREIGN KEY (ReportID) REFERENCES Report (ReportID)',
        'ALTER TABLE ReportRevision ADD CONSTRAINT FK_RevisionAuthor FOREIGN KEY (AuthorProfileID) REFERENCES UserProfile (UserProfileID)',
        'ALTER TABLE ReportRevision ADD CONSTRAINT FK_RevisionParent FOREIGN KEY (ParentRevisionID) REFERENCES ReportRevision (RevisionID)',
        'ALTER TABLE RevisionValue ADD CONSTRAINT FK_ValueRevision FOREIGN KEY (RevisionID) REFERENCES ReportRevision (RevisionID)',
        'ALTER TABLE RevisionChange ADD CONSTRAINT FK_ChangeRevision FOREIGN KEY (RevisionID) REFERENCES ReportRevision (RevisionID)',
        'ALTER TABLE ReportEvent ADD CONSTRAINT FK_EventReport FOREIGN KEY (ReportID) REFERENCES Report (ReportID)',
        'ALTER TABLE ReportEvent ADD CONSTRAINT FK_EventAuthor FOREIGN KEY (AuthorProfileID) REFERENCES UserProfile (UserProfileID)',
        'CREATE UNIQUE INDEX UX_FolderGuid ON Folder (FolderGuid)',
        'ALTER TABLE Folder ADD CONSTRAINT FK_FolderParent FOREIGN KEY (ParentFolderID) REFERENCES Folder (FolderID)',
        'ALTER TABLE Report ADD CONSTRAINT FK_ReportFolder FOREIGN KEY (FolderID) REFERENCES Folder (FolderID)',
        'ALTER TABLE FolderMove ADD CONSTRAINT FK_MoveReport FOREIGN KEY (ReportID) REFERENCES Report (ReportID)',
        'ALTER TABLE FolderMove ADD CONSTRAINT FK_MoveAuthor FOREIGN KEY (AuthorProfileID) REFERENCES UserProfile (UserProfileID)',
        'ALTER TABLE FolderMove ADD CONSTRAINT FK_MoveOldFolder FOREIGN KEY (OldFolderID) REFERENCES Folder (FolderID)',
        'ALTER TABLE FolderMove ADD CONSTRAINT FK_MoveNewFolder FOREIGN KEY (NewFolderID) REFERENCES Folder (FolderID)',
        "INSERT INTO AppMetadata VALUES (1,'$($swiftBuildPlan.AppVersion)',$($swiftBuildPlan.SchemaVersion),'Development candidate - native validation/printing remains external')"
    )
    $majorSchema=Get-Content (Join-Path $root 'src/major-schema.json') -Raw | ConvertFrom-Json
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
    $swiftBaseDatabase.TableDefs('Report').ValidationRule = "([FormKind] In ('EVAL','CHIEFEVAL','FITREP') Or ([LegacyFormKind] Is Not Null And [FormKind]=[LegacyFormKind])) And IIf([WorkflowState] Is Null,[LegacyWorkflowNull]=True,([WorkflowState] In ('Draft','Approved') Or ([LegacyWorkflowState] Is Not Null And [WorkflowState]=[LegacyWorkflowState])))"
    $swiftBaseDatabase.TableDefs('FieldDefaultLink').ValidationRule = "([LinkState]='Unresolved' And [CommandWorkspaceID] Is Null And [FolderID] Is Null And [GroupID] Is Null And [ReportID] Is Null) Or ([LinkState]='Verified' And (IIf([CommandWorkspaceID] Is Null,0,1)+IIf([FolderID] Is Null,0,1)+IIf([GroupID] Is Null,0,1)+IIf([ReportID] Is Null,0,1))=1)"
    $swiftBaseDatabase.TableDefs('GroupCompletion').ValidationRule = '[TraitCount]>=0 And [TraitSum]>=0'
    $swiftBaseDatabase.TableDefs('GroupLifecycle').ValidationRule = "[WorkflowState] In ('Initialized','In-Work','Ready for Signature','Completed') Or ([LegacyState] Is Not Null And [WorkflowState]=[LegacyState])"
    $swiftBaseDatabase.TableDefs('GroupCompletionReport').ValidationRule = '([RevisionID] Is Null And [RevisionOwnerReportID] Is Null And [TraitCount] Is Null And [TraitSum] Is Null) Or ([RevisionID] Is Not Null And [RevisionOwnerReportID]=[ReportID] And [TraitCount]>=0 And [TraitSum]>=[TraitCount] And [TraitSum]<=5*[TraitCount])'
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
    . (Join-Path $PSScriptRoot 'build-resources.ps1')
    # Trial: forms come from tracked text in src/forms instead of CreateControl
    # calls and presentation-stage patches. Text is UTF-8 in git; Access loads UTF-16.
    $formStaging=Join-Path $root ('build/forms-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $formStaging)
    foreach($formFile in Get-ChildItem -LiteralPath (Join-Path $root 'src/forms') -Filter '*.form.txt'){
        $formName=$formFile.Name.Substring(0,$formFile.Name.Length-'.form.txt'.Length)
        $staged=Join-Path $formStaging $formFile.Name
        [IO.File]::WriteAllText($staged,[IO.File]::ReadAllText($formFile.FullName,[Text.Encoding]::UTF8),[Text.Encoding]::Unicode)
        $app.LoadFromText(2,$formName,$staged)
    }
    Remove-Item -LiteralPath $formStaging -Recurse
    # Historical standalone stages stamp their own schema. The fresh build already
    # created the complete canonical schema; restore its identity before any reopen.
    $identityDB=$app.CurrentDb()
    $identityDB.Execute("UPDATE AppMetadata SET SchemaVersion=$($swiftBuildPlan.SchemaVersion),AppVersion='$($swiftBuildPlan.AppVersion)'",128)
    $identityDB.Close()
    $macro=Join-Path $root 'src\AutoExec.txt'
    $app.LoadFromText(4,'AutoExec',$macro)
    $app.DoCmd.RunCommand(126)
    $app.CloseCurrentDatabase()
    # Reopen before persisting native event modules; the named final phase owns it.
    $app.OpenCurrentDatabase($OutputPath,$true)
    # Trial: finalize stages edit forms, and their edits are already in src/forms.
    $identityDB=$app.CurrentDb()
    $identityDB.Execute("UPDATE AppMetadata SET SchemaVersion=$($swiftBuildPlan.SchemaVersion),AppVersion='$($swiftBuildPlan.AppVersion)'",128)
    $identityDB.Close()
    $app.DoCmd.RunCommand(126)
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
