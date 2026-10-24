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
    $workerLines=Get-Content -LiteralPath (Join-Path $root 'scripts/workspace-backup.ps1')
    $workerSource=@('Attribute VB_Name = "modBackupWorker"','Option Explicit','Public Function BackupWorkerScript() As String','    Dim result As String')
    foreach($line in $workerLines){$workerSource+='    result = result & "'+$line.Replace('"','""')+'" & vbCrLf'}
    $workerSource+=@('    BackupWorkerScript = result','End Function')
    $workerPath=Join-Path $root 'build/modBackupWorker.bas'
    [IO.File]::WriteAllLines($workerPath,$workerSource,[Text.UTF8Encoding]::new($false))
    $app.LoadFromText(5,'modBackupWorker',$workerPath)
    foreach($sourceModule in $swiftBuildPlan.Manifest.sourceModules){
        $app.LoadFromText(5,[IO.Path]::GetFileNameWithoutExtension($sourceModule),(Join-Path $root $sourceModule))
    }
    $navTraitModule=Join-Path $root ('build/trait-catalogue-'+[guid]::NewGuid()+'.bas')
    & (Join-Path $PSScriptRoot 'write-trait-catalogue.ps1') -SourcePath (Join-Path $root 'src/trait-catalogue.json') -OutputPath $navTraitModule -Application $app
    $app.LoadFromText(5,'modTraitCatalogue',$navTraitModule)
    function Control($f,$type,$name,$x,$y,$w,$h,$caption='',$event='') {
        $c = $app.CreateControl($f.Name,$type,0,'','',$x,$y,$w,$h)
        $c.Name=$name
        if ($type -eq 100 -or $type -eq 104) { $c.Caption=$caption }
        if ($type -notin @(101,106,112)) { $c.FontName='Segoe UI'; $c.FontSize=10 }
        if ($event) { $c.OnClick=$event }
        $script:swiftControl = $c
    }
    function NewScreen($name,$caption,$height=7600) {
        $app.LoadFromText(2,$name,(Join-Path $root 'src\BlankForm.txt'))
        $app.DoCmd.OpenForm($name,1)
        $f=$app.Forms.Item($name); $f.Caption=$caption; $f.Width=13800; $f.Section(0).Height=$height
        $f.RecordSelectors=$false; $f.NavigationButtons=$false; $f.DividingLines=$false
        $f.AutoCenter=$true; $f.ScrollBars=0; $f.Section(0).BackColor=15921906
        $f.CloseButton=$false
        Control $f 100 'lblHeading' 400 250 12800 500 $caption
        $heading=$script:swiftControl
        $heading.FontSize=22; $heading.ForeColor=6177536
        $script:swiftForm = $f
    }
    function SaveScreen($f,$name) {
        $phase='Save'
        try { $app.DoCmd.Save(2,$name); $phase='Close'; $app.DoCmd.Close(2,$name,2) } catch { throw "Unable to persist form ${name} during ${phase}: $($_.Exception.Message). Build location: $($_.ScriptStackTrace)" }
    }
    NewScreen 'frmSetup' 'Set up your Swift NAVFIT workspace' 5300; $f=$script:swiftForm
    [void](Control $f 100 'lblInfo' 400 1000 12600 650 'One authoritative local file. Profiles identify contributors; they do not authenticate users. Separate copies do not synchronize.')
    [void](Control $f 100 'lblCommand' 400 1950 2800 350 'Command name')
    [void](Control $f 109 'txtCommand' 3300 1900 7000 420)
    [void](Control $f 100 'lblUser' 400 2750 2800 350 'Contributor name')
    [void](Control $f 109 'txtProfile' 3300 2700 7000 420)
    [void](Control $f 104 'cmdSetup' 3300 3550 4000 550 'Create workspace' '=SaveSetup()')
    SaveScreen $f 'frmSetup'
    NewScreen 'frmHome' 'Swift NAVFIT' 10200; $f=$script:swiftForm
    [void](Control $f 100 'lblCommand' 400 1000 12400 450 '')
    [void](Control $f 100 'lblProfile' 400 1550 12400 450 '')
    [void](Control $f 100 'lblPrototype' 400 2200 12400 800 'Local draft prototype. Reviewed changes, saved-revision routes, frozen groups and structured PDF import are available. Full Navy policy validation and NAVFIT exchange remain under development.')
    [void](Control $f 104 'cmdLibrary' 400 3300 4000 650 'Open draft library' '=OpenScreen("frmLibrary")')
    [void](Control $f 104 'cmdSupport' 4900 3300 4000 650 'Support and limitations' '=OpenScreen("frmSupport")')
    [void](Control $f 100 'lblUse' 400 4400 12400 700 'Use one person at a time. Open this file exclusively. Close Access before making a backup copy. This prototype does not enforce exclusive use.')
    [void](Control $f 104 'cmdImportReview' 9400 3300 3500 650 'Review NAVFIT database' '=ChooseReviewFile()')
    [void](Control $f 104 'cmdRecovery' 400 5250 5500 550 'Recover pending edits' '=OpenScreen("frmRecovery")')
    [void](Control $f 104 'cmdSnippets' 400 6100 5500 550 'Narrative snippets' '=OpenScreen("frmSnippets")')
    [void](Control $f 104 'cmdOperations' 6500 6100 6300 550 'Operations and import undo' '=OpenScreen("frmOperations")')
    [void](Control $f 104 'cmdMyWork' 400 6800 12600 550 'My Work and local contributors' '=OpenMyWork()')
    [void](Control $f 104 'cmdMajor' 400 7500 12600 550 'Summary groups, batches, routes and meetings' '=OpenScreen("frmMajorWorkbench")')
    [void](Control $f 104 'cmdPDF' 400 8250 12600 550 'Import structured PDF' '=OpenScreen("frmPDFImport")')
    [void](Control $f 104 'cmdBackup' 400 8950 12600 550 'Verified backup and restore' '=OpenScreen("frmBackup")')
    $f.OnLoad='=LoadHome()'; $f.OnActivate='=LoadHome()'; SaveScreen $f 'frmHome'
    NewScreen 'frmLibrary' 'Report workspace' 8900; $f=$script:swiftForm
    [void](Control $f 100 'lblFolders' 400 1050 2800 350 'Organizational folders')
    Control $f 110 'lstFolders' 400 1500 3000 4200; $c=$script:swiftControl
    $c.ColumnCount=2; $c.ColumnWidths='0;2800'; $c.BoundColumn=1; $c.AfterUpdate='=RefreshLibrary()'
    [void](Control $f 100 'lblSearch' 3750 1050 1200 350 'Find draft')
    Control $f 109 'txtSearch' 5000 1000 3500 450; $c=$script:swiftControl; $c.AfterUpdate='=RefreshLibrary()'
    Control $f 111 'cboArchive' 8750 1000 1950 450; $c=$script:swiftControl
    $c.RowSourceType='Value List'; $c.RowSource='Active;Archived;All'; $c.LimitToList=$true; $c.AfterUpdate='=RefreshLibrary()'
    [void](Control $f 104 'cmdFind' 11000 1000 1900 450 'Find / refresh' '=RefreshLibrary()')
    Control $f 110 'lstReports' 3750 1650 9150 4050; $c=$script:swiftControl
    $c.ColumnCount=6; $c.ColumnHeads=$true; $c.BoundColumn=1; $c.ColumnWidths='0;1300;2900;2700;1300;800'; $c.OnDblClick='=EditSelected()'
    [void](Control $f 100 'lblCount' 3750 5800 5600 350 '0 reports shown')
    [void](Control $f 104 'cmdNew' 3750 6400 1650 550 'New draft' '=EditSelected(True)')
    [void](Control $f 104 'cmdOpen' 5600 6400 1650 550 'Open draft' '=EditSelected()')
    [void](Control $f 104 'cmdDuplicate' 7450 6400 1650 550 'Duplicate' '=LibraryAction("Duplicate")')
    [void](Control $f 104 'cmdArchive' 9300 6400 1650 550 'Archive' '=LibraryAction("Archive")')
    [void](Control $f 104 'cmdRecover' 11150 6400 1650 550 'Recover' '=LibraryAction("Recover")')
    [void](Control $f 104 'cmdAddFolder' 400 6000 3000 550 'Create folder' '=AddFolderUI()')
    [void](Control $f 100 'lblMove' 400 6750 3000 350 'Destination folder')
    Control $f 111 'cboDestination' 400 7200 3000 450; $c=$script:swiftControl
    $c.ColumnCount=2; $c.ColumnWidths='0;2800'; $c.BoundColumn=1; $c.LimitToList=$true
    [void](Control $f 104 'cmdMoveReport' 3750 7200 2200 500 'Move report' '=MoveReportUI()')
    [void](Control $f 104 'cmdMoveFolder' 6200 7200 2200 500 'Move folder' '=MoveFolderUI()')
    [void](Control $f 104 'cmdFolderHelp' 8650 7200 2200 500 'Folder help' '=ShowHelp("Folders")')
    [void](Control $f 104 'cmdHome' 400 8050 2800 500 'Home' '=OpenScreen("frmHome")')
    [void](Control $f 100 'lblGroupScope' 3750 8100 7100 450 'Folder moves change organization only; official groups are separate.')
    [void](Control $f 104 'cmdChecklist' 11150 7200 1800 500 'Draft checklist' '=ChecklistFromLibrary()')
    Control $f 111 'cboView' 9700 5800 3150 450; $c=$script:swiftControl
    $c.RowSourceType='Value List'; $c.RowSource='Drafting;Dates;History'; $c.LimitToList=$true; $c.AfterUpdate='=ChangeLibraryViewUI()'
    [void](Control $f 104 'cmdBulkFolders' 11000 8050 1900 500 'Bulk moves' '=OpenScreen("frmBulkFolders")')
    $f.OnLoad='=LoadLibrary()'; SaveScreen $f 'frmLibrary'
    NewScreen 'frmEditor' 'Draft editor' 31550; $f=$script:swiftForm
    $f.Width=13800
    [void](Control $f 104 'cmdOccasions' 400 900 4000 500 'Report occasions' '=OpenReportContext("Occasions")')
    [void](Control $f 104 'cmdCategories' 4700 900 4000 500 'Report categories' '=OpenReportContext("Categories")')
    [void](Control $f 104 'cmdDuties' 9000 900 4000 500 'Duty descriptions' '=OpenReportContext("Duties")')
    Control $f 109 'txtReportID' 10 10 100 100; $hidden=$script:swiftControl; $hidden.Visible=$false
    [void](Control $f 100 'lblType' 400 1550 1600 350 'Report type')
    Control $f 111 'cboKind' 2100 1500 2900 400; $c=$script:swiftControl
    $c.RowSourceType='Value List'; $c.RowSource='EVAL;CHIEFEVAL;FITREP'; $c.LimitToList=$true; $c.AfterUpdate='=EditorTypeChanged()'
    [void](Control $f 100 'lblTitle' 5500 1550 1300 350 'Draft title')
    [void](Control $f 109 'txtTitle' 6900 1500 6100 400)
    [void](Control $f 100 'lblMember' 400 2250 1600 350 'Display name')
    [void](Control $f 109 'txtMember' 2100 2200 10900 400)
    [void](Control $f 100 'lblStart' 400 2950 1600 350 'Period start')
    Control $f 109 'txtStart' 2100 2900 2900 400; $c=$script:swiftControl; $c.Format='yyyy-mm-dd'
    [void](Control $f 100 'lblEnd' 5500 2950 1300 350 'Period end')
    Control $f 109 'txtEnd' 6900 2900 2900 400; $c=$script:swiftControl; $c.Format='yyyy-mm-dd'
    [void](Control $f 100 'lblNotes' 400 3600 12400 350 'Draft notes (free text, no official narrative-fit guarantee)')
    Control $f 109 'txtNotes' 400 4000 12600 2400; $c=$script:swiftControl; $c.EnterKeyBehavior=$true; $c.ScrollBars=2
    [void](Control $f 104 'cmdSave' 400 6750 2800 550 'Save draft' '=SaveEditor()')
    [void](Control $f 104 'cmdClose' 3500 6750 2800 550 'Close editor' '=CloseEditor()')
    [void](Control $f 104 'cmdSaveHelp' 5000 9950 3000 550 'Saving and revision help' '=ShowHelp("Saving")')
    [void](Control $f 100 'lblSaved' 6600 6800 6300 400 'Not saved')
    [void](Control $f 100 'lblRevisions' 400 7500 12400 350 'Committed revisions - select an earlier revision to restore its values')
    Control $f 110 'lstRevisions' 400 7900 12600 1750; $c=$script:swiftControl
    $c.ColumnCount=4; $c.BoundColumn=1; $c.ColumnWidths='0;1200;3000;8000'
    [void](Control $f 104 'cmdRestore' 400 9950 4200 550 'Restore as new revision' '=RestoreSelected()')
    [void](Control $f 104 'cmdCompare' 8400 9950 4500 550 'Compare with latest saved' '=CompareSelectedRevision()')
    [void](Control $f 100 'lblOfficialComments' 400 11000 12600 500 'Official narrative')
    Control $f 109 'txtOfficialComments' 400 11600 12600 2500; $c=$script:swiftControl
    $c.EnterKeyBehavior=$true; $c.ScrollBars=2; $c.AfterUpdate='=RefreshNarrativeLabel()'
    [void](Control $f 100 'lblNarrativeCount' 400 14300 5000 350 '0 characters')
    [void](Control $f 104 'cmdSaveNarrative' 9300 14300 3700 550 'Save draft and narrative' '=SaveEditor()')
    [void](Control $f 100 'lblImportOrigin' 400 15000 12600 350 '')
    [void](Control $f 100 'lblRecovery' 400 15500 12600 350 'Pending edits checkpoint automatically every 30 seconds.')
    [void](Control $f 104 'cmdCheckpoint' 400 16000 5000 450 'Checkpoint pending edits now' '=CheckpointEditor()')
    [void](Control $f 104 'cmdChecklist' 7000 16000 5900 450 'Check latest saved draft' '=ChecklistFromEditor()')
    [void](Control $f 104 'cmdSnippets' 400 16600 5000 450 'Browse narrative snippets' '=OpenScreen("frmSnippets")')
    [void](Control $f 104 'cmdComments' 7000 16600 5900 450 'Review comments on saved draft' '=CommentsFromEditor()')
    [void](Control $f 104 'cmdRoute' 400 17200 12600 550 'Assign saved draft for review' '=RoutingFromEditor()')
    foreach($navControl in $f.Controls){if($navControl.Top -ge 2900){$navControl.Top=$navControl.Top+3300}}
    [void](Control $f 100 'lblAdministration' 400 2850 12600 650 'Administrative draft fields - entered separately from the local display name. Incomplete values can be saved. Length limits reflect the inspected NAVFIT schema; codes and format are not yet policy-validated.')
    [void](Control $f 100 'lblOfficialName' 400 3650 2300 350 'Official full name (27)')
    [void](Control $f 109 'txtOfficialName' 2800 3600 10200 400)
    [void](Control $f 100 'lblRankRate' 400 4300 2300 350 'Grade / rate (5)')
    [void](Control $f 109 'txtRankRate' 2800 4250 2500 400)
    [void](Control $f 100 'lblDesignator' 6000 4300 2500 350 'Desig / designation (12)')
    [void](Control $f 109 'txtDesignator' 8700 4250 4300 400)
    [void](Control $f 100 'lblUnitUIC' 400 4950 2300 350 'Unit UIC (5)')
    [void](Control $f 109 'txtUnitUIC' 2800 4900 2500 400)
    [void](Control $f 100 'lblShipStation' 6000 4950 2500 350 'Ship / station (18)')
    [void](Control $f 109 'txtShipStation' 8700 4900 4300 400)
    [void](Control $f 100 'lblPromotionStatus' 400 5400 2900 350 'Promotion status')
    Control $f 111 'cboPromotionStatus' 3400 5350 2300 350; $c=$script:swiftControl
    $c.RowSourceType='Value List'; $c.RowSource='REGULAR;FROCKED;SELECTED;SPOT'; $c.LimitToList=$false
    [void](Control $f 104 'cmdReportingSenior' 6000 5350 3300 350 'Reporting senior details' '=OpenReportingSenior()')
    [void](Control $f 100 'lblObservationStatus' 400 5800 2900 350 'Whole-report observation')
    Control $f 111 'cboObservationStatus' 3400 5750 2300 400; $c=$script:swiftControl
    $c.RowSourceType='Value List'; $c.RowSource='Observed;NOB'; $c.LimitToList=$false
    [void](Control $f 100 'lblPromotionRecommendation' 6000 5800 2900 350 'Promotion recommendation')
    Control $f 111 'cboPromotionRecommendation' 9100 5750 3900 400; $c=$script:swiftControl
    $c.RowSourceType='Value List'; $c.RowSource='NOB;Significant Problems;Progressing;Promotable;Must Promote;Early Promote'; $c.LimitToList=$false
    foreach($navControl in $f.Controls){if($navControl.Top -ge 6900){$navControl.Top=$navControl.Top+7800}}
    [void](Control $f 100 'lblTraits' 400 6800 12600 600 'Form-specific trait marks. Blank differs from NOB. Reference descriptions come from the inspected NAVFIT catalogue; whole-report observation status and policy are not assessed.')
    foreach($navBlock in 33..39){
        $navY=7550+($navBlock-33)*600
        [void](Control $f 100 ('lblTrait'+$navBlock) 400 ($navY+50) 9600 450 '')
        Control $f 111 ('cboTrait'+$navBlock) 10400 $navY 2600 400; $c=$script:swiftControl
        $c.RowSourceType='Value List'; $c.RowSource='NOB;1.0;2.0;3.0;4.0;5.0'; $c.LimitToList=$false
        $c.AfterUpdate='=ShowTraitGuidance('+$navBlock+')'; $c.OnGotFocus='=ShowTraitGuidance('+$navBlock+')'
    }
    foreach($navControl in $f.Controls){if($navControl.Top -ge 12100){$navControl.Top=$navControl.Top+1600}}
    Control $f 109 'txtPendingTraitSummary' 400 12100 12600 450; $script:swiftControl.Locked=$true
    Control $f 109 'txtSavedTraitSummary' 400 12600 12600 450; $script:swiftControl.Locked=$true
    [void](Control $f 100 'lblTraitSummaryScope' 400 13100 12600 450 'Arithmetic only: NOB and blanks are excluded. No recommendation, group, or Navy policy assessment.')
    Control $f 109 'txtTraitGuidance' 400 13700 12600 2100; $c=$script:swiftControl; $c.Locked=$true; $c.ScrollBars=2
    foreach($seniorKey in @('ReportingSeniorName','ReportingSeniorGrade','ReportingSeniorDesignator','ReportingSeniorTitle','ReportingSeniorUIC')){
        Control $f 109 ('txt'+$seniorKey) 10 10 100 100; $script:swiftControl.Visible=$false
    }
    foreach($contextKey in @('OccasionPeriodic','OccasionDetachment','OccasionFrocking','OccasionSpecial','ReportRegular','ReportConcurrent','ReportOperationalCommander','PrimaryDuty','DutyDescription')){
        Control $f 109 ("txt"+$contextKey) 10 10 100 100; $script:swiftControl.Visible=$false
    }
    [void](Control $f 104 'cmdMemberDetails' 9500 5350 3500 350 'Member details' '=OpenMemberDetails()')
    foreach($memberKey in @('MemberSSN','MemberDoDID','ComponentActive','ComponentTAR','ComponentInactive','ComponentATADSW','DateReported')){
        Control $f 109 ("txt"+$memberKey) 10 10 100 100; $script:swiftControl.Visible=$false
    }
    [void](Control $f 104 'cmdServiceDetails' 10000 6250 3000 400 'More report details' '=OpenServiceDetails()')
    foreach($serviceKey in @('CounselingDate','CounselorName','ReadinessSummary','ReadinessDetail','ReadinessDate','BilletSubcategory')){
        Control $f 109 ("txt"+$serviceKey) 10 10 100 100; $script:swiftControl.Visible=$false
    }
    foreach($decisionControl in $f.Controls){if($decisionControl.Top -ge 6800){$decisionControl.Top=$decisionControl.Top+550}}
    [void](Control $f 104 'cmdDecisionDetails' 400 6750 4000 400 'Recommendations / statement' '=OpenDecisionDetails()')
    foreach($decisionKey in @('AssignmentRecommendationA','AssignmentRecommendationB','RetentionYes','RetentionNo','StatementYes','StatementNo')){
        Control $f 109 ('txt'+$decisionKey) 10 10 100 100; $script:swiftControl.Visible=$false
    }
    [void](Control $f 104 'cmdReportPeople' 4600 6750 4000 400 'Senior contact and raters' '=OpenReportPeople()')
    foreach($peopleKey in @('ReportingSeniorSSN','ReportingSeniorDoDID','ReportingSeniorAddress1','ReportingSeniorAddress2','ReportingSeniorCity','ReportingSeniorState','ReportingSeniorZip','ReportingSeniorPhone','ReportingSeniorDSN','RaterName','RaterDate','SeniorRaterName','SeniorRaterDate')){
        Control $f 109 ('txt'+$peopleKey) 10 10 100 100; $script:swiftControl.Visible=$false
    }
    [void](Control $f 104 'cmdPartyNames' 8800 6750 4200 400 'Structured names' '=OpenPartyNames()')
    foreach($nameKey in @('MemberFirstName','MemberMiddleInitial','MemberLastName','MemberSuffix','ReportingSeniorLastName','ReportingSeniorFirstInitial','ReportingSeniorMiddleInitial','CounselorLastName','CounselorFirstInitial','CounselorMiddleInitial')){
        Control $f 109 ('txt'+$nameKey) 10 10 100 100; $script:swiftControl.Visible=$false
    }
    # Appended catalogue fields use the same unbound editor/recovery buffer.
    foreach($field in $canonicalFields | Where-Object {$_.definitionVersion -ge 2}){
        Control $f 109 $field.control 10 10 100 100
        $script:swiftControl.Visible=$false
        $script:swiftControl.EnterKeyBehavior=[bool]$field.multiline
    }
    $f.TimerInterval=30000; $f.OnTimer='=CheckpointEditor()'
    $f.OnLoad='=LoadEditor()'; $f.ScrollBars=2; SaveScreen $f 'frmEditor'
    NewScreen 'frmReportingSenior' 'Reporting senior' 6500; $f=$script:swiftForm
    $f.PopUp=$true; $f.Modal=$true
    [void](Control $f 100 'lblScope' 400 950 12600 550 'Apply transfers pending values and checkpoints them. Save draft in the editor commits them. Cancel discards panel edits. Storage limits do not establish reporting authority or Navy compliance.')
    $seniorKeys=@('ReportingSeniorName','ReportingSeniorGrade','ReportingSeniorDesignator','ReportingSeniorTitle','ReportingSeniorUIC')
    $seniorLabels=@('Name (18)','Grade (5)','Designator (5)','Title (14)','UIC (5)')
    foreach($seniorIndex in 0..4){
        $seniorY=1750+$seniorIndex*650
        [void](Control $f 100 ('lbl'+$seniorKeys[$seniorIndex]) 400 ($seniorY+50) 2300 350 $seniorLabels[$seniorIndex])
        [void](Control $f 109 ('txt'+$seniorKeys[$seniorIndex]) 2800 $seniorY 10200 400)
    }
    [void](Control $f 104 'cmdApply' 400 5350 5800 550 'Apply pending values and return' '=ApplyReportingSenior()')
    [void](Control $f 104 'cmdCancel' 7000 5350 6000 550 'Cancel panel changes' '=CancelReportingSenior()')
    SaveScreen $f 'frmReportingSenior'
    $contextPanels=@(
        @{Name='Occasions';Title='Report occasions';Keys=@('OccasionPeriodic','OccasionDetachment','OccasionFrocking','OccasionSpecial');Labels=@('Periodic','Detachment of individual','Frocking','Special')},
        @{Name='Categories';Title='Report categories';Keys=@('ReportRegular','ReportConcurrent','ReportOperationalCommander');Labels=@('Regular report','Concurrent report','Operational commander report')},
        @{Name='Duties';Title='Duty descriptions';Keys=@('PrimaryDuty','DutyDescription');Labels=@('Primary duty (14)','Duty description')}
    )
    foreach($contextPanel in $contextPanels){
        $screen='frmContext'+$contextPanel.Name
        NewScreen $screen $contextPanel.Title 8000; $f=$script:swiftForm; $f.PopUp=$true; $f.Modal=$true
        [void](Control $f 100 'lblScope' 400 950 12600 650 'Apply checkpoints pending values; Save draft commits them. Cancel discards panel changes. Blank means unentered. Combinations, applicability and Navy policy are not assessed. Imports remain narrative-only.')
        foreach($contextIndex in 0..($contextPanel.Keys.Count-1)){
            $contextY=1850+$contextIndex*850
            [void](Control $f 100 ('lbl'+$contextPanel.Keys[$contextIndex]) 400 ($contextY+50) 5000 400 $contextPanel.Labels[$contextIndex])
            if($contextPanel.Name -eq 'Duties'){
                if($contextIndex -eq 0){[void](Control $f 109 ('txt'+$contextPanel.Keys[$contextIndex]) 5500 $contextY 7500 450)}
                else{Control $f 109 'txtDutyDescription' 400 3250 12600 3000; $c=$script:swiftControl; $c.EnterKeyBehavior=$true; $c.ScrollBars=2}
            }else{
                Control $f 111 ('txt'+$contextPanel.Keys[$contextIndex]) 5500 $contextY 7500 450; $c=$script:swiftControl
                $c.RowSourceType='Value List'; $c.RowSource='Yes;No'; $c.LimitToList=$false
            }
        }
        [void](Control $f 104 'cmdApply' 400 6800 5800 550 'Apply pending values and return' ('=ApplyReportContext("'+$contextPanel.Name+'")'))
        [void](Control $f 104 'cmdCancel' 7000 6800 6000 550 'Cancel panel changes' ('=CancelReportContext("'+$contextPanel.Name+'")'))
        SaveScreen $f $screen
    }
    NewScreen 'frmMemberDetails' 'Member details' 9600; $f=$script:swiftForm; $f.PopUp=$true; $f.Modal=$true
    [void](Control $f 100 'lblScope' 400 950 12600 650 'Apply checkpoints pending values; Save draft commits them. Blank is unentered. Identifiers, component eligibility and reported-date applicability are not Navy-validated.')
    $memberKeys=@('MemberSSN','MemberDoDID','ComponentActive','ComponentTAR','ComponentInactive','ComponentATADSW','DateReported')
    $memberLabels=@('SSN (up to 9 digits)','DoD ID (up to 10 digits)','Active','TAR (reference field)','Inactive','AT / ADSW','Date reported (yyyy-mm-dd)')
    foreach($memberIndex in 0..6){
        $memberY=1850+$memberIndex*850
        [void](Control $f 100 ('lbl'+$memberKeys[$memberIndex]) 400 ($memberY+50) 5000 400 $memberLabels[$memberIndex])
        if($memberIndex -ge 2 -and $memberIndex -le 5){
            Control $f 111 ('txt'+$memberKeys[$memberIndex]) 5500 $memberY 7500 450; $c=$script:swiftControl
            $c.RowSourceType='Value List'; $c.RowSource='Yes;No'; $c.LimitToList=$false
        }else{[void](Control $f 109 ('txt'+$memberKeys[$memberIndex]) 5500 $memberY 7500 450)}
    }
    [void](Control $f 104 'cmdApply' 400 8300 5800 550 'Apply pending values and return' '=ApplyMemberDetails()')
    [void](Control $f 104 'cmdCancel' 7000 8300 6000 550 'Cancel panel changes' '=CancelMemberDetails()')
    SaveScreen $f 'frmMemberDetails'
    $serviceReference=Get-Content (Join-Path $root 'src/service-reference.json') -Raw | ConvertFrom-Json
    if($serviceReference.Readiness.Count -ne 10 -or $serviceReference.Billet.Count -ne 37){throw 'Unexpected service reference inventory'}
    NewScreen 'frmServiceDetails' 'Counseling, readiness and billet' 9000; $f=$script:swiftForm; $f.PopUp=$true; $f.Modal=$true
    [void](Control $f 100 'lblScope' 400 950 12600 650 'Apply checkpoints pending values; Save draft commits them. Blank remains saveable. Suggestions are reference vocabulary; applicability, eligibility and Navy rules are not assessed.')
    $serviceKeys=@('CounselingDate','CounselorName','ReadinessSummary','ReadinessDetail','ReadinessDate','BilletSubcategory')
    $serviceLabels=@('Counseling date (yyyy-mm-dd)','Counselor name (20)','Readiness summary (4)','Readiness detail (15)','Readiness date (yyyy-mm-dd)','Billet subcategory (10)')
    foreach($serviceIndex in 0..5){
        $serviceY=1850+$serviceIndex*850
        [void](Control $f 100 ('lbl'+$serviceKeys[$serviceIndex]) 400 ($serviceY+50) 5000 400 $serviceLabels[$serviceIndex])
        if($serviceIndex -eq 3 -or $serviceIndex -eq 5){
            Control $f 111 ('txt'+$serviceKeys[$serviceIndex]) 5500 $serviceY 7500 450; $c=$script:swiftControl
            $c.RowSourceType='Value List'; $c.LimitToList=$false
            if($serviceIndex -eq 3){$c.RowSource=$serviceReference.Readiness -join ';'}
            else{$c.RowSource=$serviceReference.Billet -join ';'}
        }else{[void](Control $f 109 ('txt'+$serviceKeys[$serviceIndex]) 5500 $serviceY 7500 450)}
    }
    [void](Control $f 104 'cmdApply' 400 7500 5800 550 'Apply pending values and return' '=ApplyServiceDetails()')
    [void](Control $f 104 'cmdCancel' 7000 7500 6000 550 'Cancel panel changes' '=CancelServiceDetails()')
    SaveScreen $f 'frmServiceDetails'
    NewScreen 'frmDecisionDetails' 'Recommendations and member statement' 9000; $f=$script:swiftForm; $f.PopUp=$true; $f.Modal=$true
    [void](Control $f 100 'lblScope' 400 950 12600 650 'Apply checkpoints pending values; Save draft commits them. For each box, Yes means marked, No means unmarked, blank means unentered. Signatures, official acceptance and Navy eligibility are not assessed.')
    $decisionKeys=@('AssignmentRecommendationA','AssignmentRecommendationB','RetentionYes','RetentionNo','StatementYes','StatementNo')
    $decisionLabels=@('Assignment recommendation A (20)','Assignment recommendation B (20)','Retention Yes marked','Retention No marked','Member statement Yes marked','Member statement No marked')
    foreach($decisionIndex in 0..5){
        $decisionY=1850+$decisionIndex*850
        [void](Control $f 100 ('lbl'+$decisionKeys[$decisionIndex]) 400 ($decisionY+50) 5000 400 $decisionLabels[$decisionIndex])
        if($decisionIndex -ge 2){
            Control $f 111 ('txt'+$decisionKeys[$decisionIndex]) 5500 $decisionY 7500 450; $c=$script:swiftControl
            $c.RowSourceType='Value List'; $c.RowSource='Yes;No'; $c.LimitToList=$false
        }else{[void](Control $f 109 ('txt'+$decisionKeys[$decisionIndex]) 5500 $decisionY 7500 450)}
    }
    [void](Control $f 104 'cmdApply' 400 7500 5800 550 'Apply pending values and return' '=ApplyDecisionDetails()')
    [void](Control $f 104 'cmdCancel' 7000 7500 6000 550 'Cancel panel changes' '=CancelDecisionDetails()')
    SaveScreen $f 'frmDecisionDetails'
    NewScreen 'frmReportPeople' 'Senior contact and raters' 15000; $f=$script:swiftForm; $f.PopUp=$true; $f.Modal=$true; $f.ScrollBars=2
    [void](Control $f 100 'lblScope' 400 950 12600 650 'Apply checkpoints pending entries; Save draft commits them. Incomplete entries remain saveable. Dates are yyyy-mm-dd. These details are separate from local reviewer profiles.')
    [void](Control $f 100 'lblIdentifiers' 400 1650 12600 350 'Reporting senior identifiers - digits only; leading zeros retained')
    [void](Control $f 100 'lblContact' 400 3700 12600 350 'Reporting senior contact details')
    [void](Control $f 100 'lblRaters' 400 10000 12600 350 'Rater and senior rater - dates do not establish signatures or approval')
    $peopleKeys=@('ReportingSeniorSSN','ReportingSeniorDoDID','ReportingSeniorAddress1','ReportingSeniorAddress2','ReportingSeniorCity','ReportingSeniorState','ReportingSeniorZip','ReportingSeniorPhone','ReportingSeniorDSN','RaterName','RaterDate','SeniorRaterName','SeniorRaterDate')
    $peopleLabels=@('Reporting senior SSN (9)','Reporting senior DoD ID (10)','Address line 1 (30)','Address line 2 (30)','City (15)','State / region (2)','ZIP / postal code (9)','Phone (10)','DSN (7)','Rater name (28)','Rater date (10)','Senior rater name (28)','Senior rater date (10)')
    foreach($peopleIndex in 0..12){
        $peopleY=2050+$peopleIndex*850
        if($peopleIndex -ge 2){$peopleY+=350}
        if($peopleIndex -ge 9){$peopleY+=350}
        [void](Control $f 100 ('lbl'+$peopleKeys[$peopleIndex]) 400 ($peopleY+50) 5000 400 $peopleLabels[$peopleIndex])
        [void](Control $f 109 ('txt'+$peopleKeys[$peopleIndex]) 5500 $peopleY 7500 450)
    }
    [void](Control $f 104 'cmdApply' 400 13700 5800 550 'Apply pending values and return' '=ApplyReportPeople()')
    [void](Control $f 104 'cmdCancel' 7000 13700 6000 550 'Cancel panel changes' '=CancelReportPeople()')
    SaveScreen $f 'frmReportPeople'
    NewScreen 'frmPartyNames' 'Structured names' 12500; $f=$script:swiftForm; $f.PopUp=$true; $f.Modal=$true; $f.ScrollBars=2
    [void](Control $f 100 'lblScope' 400 950 12600 650 'Apply checkpoints pending entries; Save draft commits them. These parts remain separate from full-name entries. No automatic splitting or replacement occurs. Blank remains saveable.')
    [void](Control $f 100 'lblMemberParts' 400 1650 12600 350 'Member name parts')
    [void](Control $f 100 'lblSeniorParts' 400 5400 12600 350 'Reporting senior name parts')
    [void](Control $f 100 'lblCounselorParts' 400 8300 12600 350 'Counselor name parts')
    $nameKeys=@('MemberFirstName','MemberMiddleInitial','MemberLastName','MemberSuffix','ReportingSeniorLastName','ReportingSeniorFirstInitial','ReportingSeniorMiddleInitial','CounselorLastName','CounselorFirstInitial','CounselorMiddleInitial')
    $nameLabels=@('Member first name (27)','Member middle initial(s) (3)','Member last name (27)','Member suffix (10)','Reporting senior last name (18)','Reporting senior first initial (1)','Reporting senior middle initial(s) (3)','Counselor last name (20)','Counselor first initial (1)','Counselor middle initial(s) (3)')
    foreach($nameIndex in 0..9){
        $nameY=2050+$nameIndex*850
        if($nameIndex -ge 4){$nameY+=350}
        if($nameIndex -ge 7){$nameY+=350}
        [void](Control $f 100 ('lbl'+$nameKeys[$nameIndex]) 400 ($nameY+50) 5000 400 $nameLabels[$nameIndex])
        [void](Control $f 109 ('txt'+$nameKeys[$nameIndex]) 5500 $nameY 7500 450)
    }
    [void](Control $f 104 'cmdApply' 400 11200 5800 550 'Apply pending values and return' '=ApplyPartyNames()')
    [void](Control $f 104 'cmdCancel' 7000 11200 6000 550 'Cancel panel changes' '=CancelPartyNames()')
    SaveScreen $f 'frmPartyNames'
    NewScreen 'frmSupport' 'Support and limitations' 7400; $f=$script:swiftForm
    Control $f 109 'txtDiagnostics' 400 1100 12600 4500; $c=$script:swiftControl; $c.Locked=$true; $c.ScrollBars=2
    [void](Control $f 104 'cmdHome' 400 6150 2800 550 'Home' '=OpenScreen("frmHome")')
    $f.OnLoad='=LoadSupport()'; SaveScreen $f 'frmSupport'
    NewScreen 'frmImportReview' 'NAVFIT source review' 12500; $f=$script:swiftForm
    [void](Control $f 100 'lblScope' 400 950 12600 700 'Partial narrative import only. Create a new draft, or compare with a saved draft before accepting changes. Other official fields remain unmapped.')
    Control $f 109 'txtSource' 400 1750 12600 450; $c=$script:swiftControl; $c.Locked=$true
    Control $f 110 'lstImportReports' 400 2400 4700 4400; $c=$script:swiftControl
    $c.ColumnCount=3; $c.ColumnWidths='600;1300;2700'; $c.BoundColumn=1; $c.AfterUpdate='=ReviewReportFields()'
    Control $f 110 'lstImportFields' 5500 2400 7500 4400; $c=$script:swiftControl
    $c.ColumnCount=4; $c.ColumnWidths='0;4000;1300;1300'; $c.BoundColumn=1; $c.AfterUpdate='=ReviewFieldValue()'
    [void](Control $f 100 'lblFieldColumns' 5500 2000 7500 350 'Field name / DAO type / Null or Value')
    Control $f 109 'txtFieldValue' 400 7100 12600 1800; $c=$script:swiftControl
    $c.Locked=$true; $c.ScrollBars=2
    [void](Control $f 104 'cmdLeave' 400 9250 3000 500 'Close review' '=LeaveImportReview()')
    [void](Control $f 100 'lblImportMember' 400 10000 3500 350 'Local member display name (optional)')
    [void](Control $f 109 'txtImportMember' 4000 9950 6000 450)
    [void](Control $f 104 'cmdImport' 400 10800 4200 550 'Create draft from narrative' '=ImportSelectedNarrativeUI()')
    [void](Control $f 100 'lblImportResult' 5000 10800 7900 650 '')
    Control $f 111 'cboDestinationReport' 400 11700 7000 500; $c=$script:swiftControl
    $c.ColumnCount=3; $c.ColumnWidths='0;5000;1800'; $c.BoundColumn=1; $c.LimitToList=$true
    $c.RowSource='SELECT ReportID,Title,FormKind FROM Report WHERE IsArchived=False ORDER BY Title'
    [void](Control $f 104 'cmdCompareIncoming' 8000 11700 4900 500 'Compare with selected saved draft' '=PreviewReconcileUI()')
    $f.OnClose='=CloseReviewFile()'; SaveScreen $f 'frmImportReview'
    NewScreen 'frmRevisionCompare' 'Compare saved revisions' 10000; $f=$script:swiftForm
    Control $f 109 'txtOlder' 10 10 100 100; $script:swiftControl.Visible=$false
    Control $f 109 'txtNewer' 120 10 100 100; $script:swiftControl.Visible=$false
    [void](Control $f 100 'lblSummary' 400 1050 12600 650 '')
    Control $f 106 'chkChangedOnly' 400 1850 350 350; $c=$script:swiftControl; $c.AfterUpdate='=RefreshComparisonFields()'
    [void](Control $f 100 'lblChangedOnly' 850 1850 4500 350 'Changed fields only')
    Control $f 110 'lstCompareFields' 400 2400 12600 1800; $c=$script:swiftControl
    $c.ColumnCount=3; $c.ColumnWidths='0;7000;5000'; $c.BoundColumn=1; $c.AfterUpdate='=LoadComparisonValues()'
    [void](Control $f 100 'lblOlder' 400 4500 6000 450 '')
    [void](Control $f 100 'lblNewer' 6950 4500 6000 450 '')
    Control $f 109 'txtOldValue' 400 5100 6000 3600; $c=$script:swiftControl; $c.Locked=$true; $c.ScrollBars=2
    Control $f 109 'txtNewValue' 6950 5100 6000 3600; $c=$script:swiftControl; $c.Locked=$true; $c.ScrollBars=2
    [void](Control $f 104 'cmdClose' 400 9150 3300 550 'Return to editor' '=CloseRevisionComparison()')
    SaveScreen $f 'frmRevisionCompare'
    NewScreen 'frmRecovery' 'Recover pending edits' 7600; $f=$script:swiftForm
    [void](Control $f 100 'lblScope' 400 1000 12600 800 'Checkpoints are pending edits, not committed revisions. Recover to inspect, then Save draft. If the saved report changed, recover as a new draft.')
    Control $f 110 'lstRecoveries' 400 2100 12600 3400; $c=$script:swiftControl
    $c.ColumnCount=4; $c.ColumnWidths='0;6000;1800;4200'; $c.BoundColumn=1
    [void](Control $f 104 'cmdRecover' 400 5900 3600 600 'Recover pending edits' '=RecoverSelectedUI()')
    [void](Control $f 104 'cmdRecoverNew' 4300 5900 3600 600 'Recover as new draft' '=RecoverSelectedUI(True)')
    [void](Control $f 104 'cmdDiscard' 8200 5900 3600 600 'Discard checkpoint' '=DeleteRecoveryUI()')
    [void](Control $f 104 'cmdHome' 400 6800 3000 500 'Home' '=OpenScreen("frmHome")')
    $f.OnLoad='=LoadRecoveries()'; $f.OnActivate='=LoadRecoveries()'; SaveScreen $f 'frmRecovery'
    NewScreen 'frmReconcile' 'Review incoming official narrative' 10100; $f=$script:swiftForm
    [void](Control $f 100 'lblContext' 400 1000 12600 450 '')
    [void](Control $f 100 'lblScope' 400 1600 12600 700 'Two-way narrative comparison only. Accept replaces this one saved field. Other fields stay unchanged. Close the editor before accepting.')
    foreach($name in @('txtExternalID','txtReportID','txtExpectedRevision','txtSourceHash')){Control $f 109 $name 10 10 100 100; $script:swiftControl.Visible=$false}
    [void](Control $f 100 'lblCurrent' 400 2450 6000 350 'Current Swift narrative')
    [void](Control $f 100 'lblIncoming' 6950 2450 6000 350 'Incoming NAVFIT narrative')
    Control $f 109 'txtCurrent' 400 2900 6000 5000; $c=$script:swiftControl; $c.Locked=$true; $c.ScrollBars=2
    Control $f 109 'txtIncoming' 6950 2900 6000 5000; $c=$script:swiftControl; $c.Locked=$true; $c.ScrollBars=2
    [void](Control $f 104 'cmdAccept' 400 8500 5000 600 'Accept incoming narrative' '=AcceptIncomingUI()')
    [void](Control $f 104 'cmdClose' 6000 8500 5000 600 'Keep Swift value / close' '=CloseReconcileUI()')
    SaveScreen $f 'frmReconcile'
    NewScreen 'frmSnippets' 'Narrative snippets' 13100; $f=$script:swiftForm
    [void](Control $f 100 'lblScope' 400 950 12600 650 'Locally curated snippets with attribution. Deliberate insertion appends to the open draft. Review for the member before saving.')
    Control $f 109 'txtSearch' 400 1800 8000 450; $c=$script:swiftControl; $c.AfterUpdate='=LoadSnippets()'
    [void](Control $f 104 'cmdSearch' 8800 1800 4000 450 'Search snippets' '=LoadSnippets()')
    Control $f 110 'lstSnippets' 400 2500 12600 2200; $c=$script:swiftControl
    $c.ColumnCount=4; $c.ColumnWidths='0;5000;1800;5500'; $c.BoundColumn=1; $c.AfterUpdate='=SelectSnippet()'
    Control $f 109 'txtSnippetID' 10 10 100 100; $script:swiftControl.Visible=$false
    [void](Control $f 100 'lblTitle' 400 4900 1500 350 'Title')
    [void](Control $f 109 'txtTitle' 2100 4900 10700 450)
    [void](Control $f 100 'lblAttribution' 400 5550 1500 350 'Attribution')
    [void](Control $f 109 'txtAttribution' 2100 5550 10700 450)
    Control $f 111 'cboKind' 2100 6200 4000 450; $c=$script:swiftControl; $c.RowSourceType='Value List'; $c.RowSource='ALL;EVAL;CHIEFEVAL;FITREP'; $c.LimitToList=$true
    [void](Control $f 100 'lblKind' 400 6200 1500 350 'Report type')
    Control $f 109 'txtBody' 400 7000 12600 3800; $c=$script:swiftControl; $c.ScrollBars=2; $c.EnterKeyBehavior=$true
    [void](Control $f 104 'cmdNew' 400 11200 2300 550 'New snippet' '=NewSnippetUI()')
    [void](Control $f 104 'cmdSave' 3000 11200 2300 550 'Save snippet' '=SaveSnippetUI()')
    [void](Control $f 104 'cmdInsert' 5600 11200 3300 550 'Append to open draft' '=InsertSnippetUI()')
    [void](Control $f 104 'cmdClose' 9500 11200 3300 550 'Close snippets' '=CloseSnippetsUI()')
    $f.OnLoad='=LoadSnippets()'; SaveScreen $f 'frmSnippets'
    . (Join-Path $PSScriptRoot 'invoke-build-phase.ps1') -BuildPhase 'base-resources'
    NewScreen 'frmOperations' 'Operations and import undo' 12300; $f=$script:swiftForm
    [void](Control $f 100 'lblScope' 400 950 12600 750 'Started entries may represent interrupted work. Inspect receipts before retrying. Undo preserves history and refuses later edits.')
    Control $f 110 'lstOperations' 400 2000 12600 2100; $c=$script:swiftControl
    $c.ColumnCount=4; $c.ColumnWidths='0;4200;3000;4800'; $c.BoundColumn=1; $c.AfterUpdate='=OperationDetails()'
    Control $f 109 'txtOutcome' 400 4400 12600 1900; $c=$script:swiftControl; $c.Locked=$true; $c.ScrollBars=2
    [void](Control $f 100 'lblReceipts' 400 6600 12600 450 'Import receipts - Swift report / source report / time / undone')
    Control $f 110 'lstReceipts' 400 7200 12600 2300; $c=$script:swiftControl
    $c.ColumnCount=5; $c.ColumnWidths='0;2500;2500;5000;1600'; $c.BoundColumn=1
    [void](Control $f 104 'cmdUndo' 400 10000 4500 600 'Undo selected import' '=UndoImportUI()')
    [void](Control $f 104 'cmdReview' 5400 10000 7400 600 'Acknowledge selected incomplete operation' '=AcknowledgeOperationUI()')
    [void](Control $f 104 'cmdProvenance' 400 11700 12600 500 'Inspect selected import provenance' '=ProvenanceFromOperations()')
    [void](Control $f 104 'cmdHome' 400 11000 3500 550 'Home' '=OpenScreen("frmHome")')
    [void](Control $f 104 'cmdUndoBatch' 4500 11000 8000 550 'Undo selected folder operation' '=UndoFolderBatchUI()')
    $f.OnLoad='=LoadOperations()'; $f.OnActivate='=LoadOperations()'; SaveScreen $f 'frmOperations'
    NewScreen 'frmProvenance' 'Retained import provenance' 13100; $f=$script:swiftForm
    [void](Control $f 100 'lblScope' 400 950 12600 650 'Read-only retained source values, including unmapped fields. A receipt captures all three source tables. This is evidence, not current Swift data or export readiness.')
    Control $f 109 'txtReceiptID' 10 10 100 100; $script:swiftControl.Visible=$false
    Control $f 109 'txtContext' 400 1750 12600 2200; $c=$script:swiftControl; $c.Locked=$true; $c.ScrollBars=2
    Control $f 111 'cboTable' 400 4200 2500 450; $c=$script:swiftControl; $c.RowSourceType='Value List'; $c.RowSource='All;Reports;Folders;Summary'; $c.LimitToList=$true
    [void](Control $f 100 'lblRow' 3100 4200 1800 350 'Source row (optional)')
    [void](Control $f 109 'txtRow' 5000 4200 1300 450)
    [void](Control $f 100 'lblSearch' 6500 4200 1500 350 'Field name')
    [void](Control $f 109 'txtSearch' 8100 4200 2600 450)
    [void](Control $f 104 'cmdFilter' 11000 4200 2000 450 'Apply filters' '=FilterProvenanceUI()')
    [void](Control $f 100 'lblCount' 400 4800 12600 400 '')
    Control $f 110 'lstFields' 400 5300 12600 2500; $c=$script:swiftControl
    $c.ColumnCount=7; $c.ColumnWidths='0;2000;1400;1400;5200;1300;1300'; $c.BoundColumn=1; $c.AfterUpdate='=ShowProvenanceValue()'
    [void](Control $f 100 'lblValue' 400 8000 12600 600 '')
    Control $f 109 'txtValue' 400 8750 12600 3100; $c=$script:swiftControl; $c.Locked=$true; $c.ScrollBars=2
    [void](Control $f 104 'cmdReturn' 400 12200 12600 550 'Return to operations' '=OpenScreen("frmOperations")')
    SaveScreen $f 'frmProvenance'
    NewScreen 'frmMyWork' 'My Work - local review assignments' 17500; $f=$script:swiftForm
    [void](Control $f 100 'lblProfile' 400 950 12600 650 '')
    [void](Control $f 100 'lblScope' 400 1700 12600 650 'Review outcomes apply to a saved revision. Drafts remain editable; later saves make older reviews stale. Accepted is an internal review outcome, not official approval or export readiness.')
    Control $f 109 'txtReportID' 10 10 100 100; $script:swiftControl.Visible=$false
    Control $f 109 'txtExpectedRevision' 10 10 100 100; $script:swiftControl.Visible=$false
    Control $f 111 'cboContributor' 400 2550 3700 450; $c=$script:swiftControl; $c.ColumnCount=2; $c.ColumnWidths='0;3500'; $c.BoundColumn=1; $c.LimitToList=$true
    [void](Control $f 104 'cmdSwitch' 4400 2550 2600 450 'Switch contributor' '=ContributorUI("Switch")')
    [void](Control $f 109 'txtContributor' 7300 2550 3300 450)
    [void](Control $f 104 'cmdAdd' 10900 2550 2100 450 'Add contributor' '=ContributorUI("Add")')
    Control $f 106 'chkAll' 400 3300 300 300; $c=$script:swiftControl; $c.DefaultValue='False'; $c.AfterUpdate='=LoadMyWork()'
    [void](Control $f 100 'lblAll' 850 3250 9500 450 'Show all review history (otherwise my pending assignments)')
    [void](Control $f 104 'cmdRefresh' 11000 3250 2000 450 'Refresh' '=LoadMyWork()')
    Control $f 110 'lstTasks' 400 3900 12600 2400; $c=$script:swiftControl
    $c.ColumnCount=7; $c.ColumnWidths='0;3500;2800;1700;2100;1500;1000'; $c.BoundColumn=1; $c.AfterUpdate='=TaskDetails()'
    Control $f 109 'txtHistory' 400 6600 12600 2800; $c=$script:swiftControl; $c.Locked=$true; $c.ScrollBars=2
    [void](Control $f 100 'lblAssignment' 400 9700 12600 700 '')
    [void](Control $f 100 'lblReviewer' 400 10600 2000 350 'Assign to reviewer')
    Control $f 111 'cboReviewer' 2500 10600 4800 450; $c=$script:swiftControl; $c.ColumnCount=2; $c.ColumnWidths='0;4500'; $c.BoundColumn=1; $c.LimitToList=$true
    [void](Control $f 100 'lblDue' 7600 10600 2000 350 'Due date (optional)')
    Control $f 109 'txtDue' 9900 10600 3100 450; $script:swiftControl.Format='Short Date'
    [void](Control $f 100 'lblNotes' 400 11300 12600 400 'Assignment instructions or outcome notes (return requires a reason)')
    Control $f 109 'txtNotes' 400 11900 12600 2100; $c=$script:swiftControl; $c.EnterKeyBehavior=$true; $c.ScrollBars=2
    [void](Control $f 104 'cmdAssign' 400 14400 3500 550 'Assign saved revision' '=AssignReviewUI()')
    [void](Control $f 104 'cmdAccept' 4400 14400 3800 550 'Accept selected review' '=FinishReviewUI("Accepted")')
    [void](Control $f 104 'cmdReturn' 8700 14400 4300 550 'Return selected review' '=FinishReviewUI("Returned")')
    [void](Control $f 104 'cmdCancel' 400 15300 3500 550 'Cancel selected review' '=FinishReviewUI("Cancelled")')
    [void](Control $f 104 'cmdDraft' 4400 15300 3800 550 'Open selected draft' '=OpenReviewTaskDraft()')
    [void](Control $f 104 'cmdHome' 8700 15300 4300 550 'Home' '=OpenScreen("frmHome")')
    $f.OnLoad='=LoadMyWork()'; SaveScreen $f 'frmMyWork'
    NewScreen 'frmBulkFolders' 'Preview bulk folder moves' 10400; $f=$script:swiftForm
    [void](Control $f 100 'lblScope' 400 950 12600 700 'Select multiple active drafts, choose a destination, then preview. Folder moves are separate from official report groups and revisions.')
    Control $f 110 'lstReports' 400 2000 12600 2900; $c=$script:swiftControl
    $c.ColumnCount=4; $c.ColumnWidths='0;4000;3500;4500'; $c.BoundColumn=1; $c.MultiSelect=1
    Control $f 111 'cboDestination' 400 5200 12600 450; $c=$script:swiftControl
    $c.ColumnCount=2; $c.ColumnWidths='0;12000'; $c.BoundColumn=1; $c.LimitToList=$true
    foreach($name in @('txtSelection','txtDestination','txtStamp')){Control $f 109 $name 10 10 100 100; $script:swiftControl.Visible=$false}
    Control $f 109 'txtPreview' 400 6000 12600 2300; $c=$script:swiftControl; $c.Locked=$true; $c.ScrollBars=2
    [void](Control $f 104 'cmdPreview' 400 8800 3500 600 'Preview selected moves' '=PreviewFolderBatchUI()')
    [void](Control $f 104 'cmdApply' 4500 8800 3500 600 'Apply previewed batch' '=ApplyFolderBatchUI()')
    [void](Control $f 104 'cmdLibrary' 8500 8800 4300 600 'Return to library' '=OpenScreen("frmLibrary")')
    $f.OnLoad='=LoadBulkFolders()'; SaveScreen $f 'frmBulkFolders'
    NewScreen 'frmMajorWorkbench' 'Groups, reviewed changes and meetings' 22500; $f=$script:swiftForm
    $f.ScrollBars=2
    [void](Control $f 100 'lblInfo' 400 950 12600 650 'Select reports, then create a group, preview a field change, route a saved revision or propose a meeting change. Profiles record attribution; they do not authenticate reviewers.')
    [void](Control $f 110 'lstReports' 400 1700 12600 2300); $c=$script:swiftControl; $c.ColumnCount=4; $c.ColumnWidths='0;4200;4800;2000'; $c.MultiSelect=2
    [void](Control $f 100 'lblTitle' 400 4100 1800 500 'Group / route / meeting title')
    [void](Control $f 109 'txtTitle' 2200 4100 6200 450)
    [void](Control $f 111 'cboKind' 8600 4100 1900 450); $c=$script:swiftControl; $c.RowSourceType='Value List'; $c.RowSource='EVAL;CHIEFEVAL;FITREP'; $c.LimitToList=$true
    [void](Control $f 111 'cboPaygrade' 10700 4100 2300 450); $c=$script:swiftControl; $c.RowSourceType='Value List'; $c.RowSource='E1;E2;E3;E4;E5;E6;E7;E8;E9;W1;W2;W3;W4;W5;O1;O2;O3;O4;O5;O6'; $c.LimitToList=$true
    $selectors=@(@('cboGroup','Group',4750),@('cboSnapshot','Frozen roster',5350),@('cboBatch','Change batch',5950),@('cboTemplate','Route template',6550),@('cboRun','Review route',7150),@('cboMeeting','Meeting',7750),@('cboProposal','Proposal',8350),@('cboFolder','Default folder',8950))
    foreach($item in $selectors){[void](Control $f 100 ('lbl'+$item[0]) 400 ($item[2]+50) 1800 350 $item[1]);[void](Control $f 111 $item[0] 2200 $item[2] 10600 450);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;10000';$c.LimitToList=$true}
    [void](Control $f 100 'lblField' 400 9600 1600 350 'Field / default scope')
    [void](Control $f 111 'cboField' 2200 9550 6500 450);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.LimitToList=$true
    [void](Control $f 111 'cboScope' 9000 9550 3800 450);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='Command;Folder;Group;Report';$c.LimitToList=$true
    [void](Control $f 100 'lblValue' 400 10200 1600 350 'Proposed value')
    [void](Control $f 109 'txtValue' 2200 10150 9000 700);$script:swiftControl.ScrollBars=2
    [void](Control $f 106 'chkNull' 11500 10200 300 350);$script:swiftControl.DefaultValue='False'
    [void](Control $f 100 'lblNull' 11900 10200 900 350 'Clear to null')
    [void](Control $f 100 'lblNote' 400 11000 1600 650 'Reason / attestation / decision')
    [void](Control $f 109 'txtNote' 2200 10950 10600 750);$script:swiftControl.ScrollBars=2
    [void](Control $f 100 'lblReviewers' 400 11850 4000 450 'Choose route reviewers in order')
    [void](Control $f 111 'cboReviewer' 4700 11800 3500 450);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;3300';$c.LimitToList=$true
    [void](Control $f 104 'cmdAddReviewer' 8400 11800 2100 450 'Add reviewer step' '=MajorAction("AddReviewer")')
    [void](Control $f 104 'cmdClearReviewers' 10700 11800 2100 450 'Clear reviewer steps' '=MajorAction("ClearReviewers")')
    [void](Control $f 109 'txtReviewers' 10 10 100 100);$script:swiftControl.Visible=$false
    [void](Control $f 109 'txtReviewerNames' 400 12450 7000 450);$script:swiftControl.Locked=$true
    [void](Control $f 100 'lblOrder' 7700 12450 2000 450 'Meeting order (report IDs)')
    [void](Control $f 109 'txtMeetingOrder' 10000 12450 2800 450)
    $actions=@(@('CreateGroup','Create group'),@('AddMembers','Add selected reports'),@('RemoveMember','Remove selected report'),@('FreezeGroup','Attest and freeze roster'),@('PreviewBatch','Preview field change'),@('ShowBatch','Show batch values'),@('ApplyBatch','Apply reviewed batch'),@('PreviewUndo','Preview batch undo'),@('SaveDefault','Save scoped default'),@('SuggestDefault','Suggest inherited default'),@('CreateRoute','Create route template'),@('StartRoute','Start selected route'),@('ShowRoute','Show route steps'),@('AcceptStep','Accept current step'),@('ReturnStep','Return current step'),@('CancelRoute','Cancel route'),@('CreateMeeting','Create frozen meeting'),@('ShowMeeting','Show frozen meeting'),@('ProposeField','Propose field change'),@('ApplyProposal','Apply selected proposal'),@('RunChecks','Run limited Navy draft checks'),@('ApproveRoute','Approve saved revision locally'),@('ReopenApproval','Reopen local approval'),@('OrderMeeting','Set meeting display order'),@('ShowProposals','Show meeting proposals'),@('ApplySelectedProposals','Apply selected proposals'))
    for($i=0;$i -lt $actions.Count;$i++){[void](Control $f 104 ('cmd'+$actions[$i][0]) (400+($i%4)*3200) (13100+[math]::Floor($i/4)*600) 3000 450 $actions[$i][1] ('=MajorAction("'+$actions[$i][0]+'")'))}
    [void](Control $f 110 'lstDetails' 400 17500 12600 1100);$c=$script:swiftControl;$c.ColumnCount=6;$c.ColumnWidths='1000;1800;2200;3500;2300;1600';$c.MultiSelect=2;$c.OnClick='=ShowMajorDetails()'
    [void](Control $f 109 'txtDetail' 400 18750 12600 1800);$c=$script:swiftControl;$c.Locked=$true;$c.ScrollBars=2
    [void](Control $f 100 'lblStatus' 400 20800 12600 600 'Ready. Select a detail row to review its complete values before applying changes.')
    [void](Control $f 104 'cmdHome' 400 21700 12600 450 'Return home' '=OpenScreen("frmHome")')
    $f.OnLoad='=LoadMajorWorkbench()'; SaveScreen $f 'frmMajorWorkbench'
    NewScreen 'frmPDFImport' 'Import structured NAVPERS PDF' 9600; $f=$script:swiftForm
    [void](Control $f 100 'lblInfo' 400 950 12600 850 'Extracts supported form data locally without Acrobat. Supports the supplied May 2025 EVAL, CHIEF EVAL and FITREP form revisions. Previewed fields: name, grade/rate, designation, identifier, ship/station, report dates and official comments. Other PDF values are retained as XML for review. Flattened PDFs require manual entry.')
    [void](Control $f 104 'cmdChoose' 400 1950 12600 550 'Choose PDF and preview its values' '=ChoosePDFPreview()')
    [void](Control $f 109 'txtPreview' 400 2700 12600 3800);$c=$script:swiftControl;$c.Locked=$true;$c.ScrollBars=2
    [void](Control $f 100 'lblCentury' 400 6800 5500 400 'Century for six-digit report dates (explicit choice)')
    [void](Control $f 111 'cboCentury' 6200 6750 6600 450);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='19;20;21';$c.LimitToList=$true
    [void](Control $f 100 'lblIdentifier' 400 7450 5500 400 'Combined identifier field means')
    [void](Control $f 111 'cboIdentifier' 6200 7400 6600 450);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='SSN;DoDID;Retain only';$c.LimitToList=$true
    [void](Control $f 104 'cmdImport' 400 8100 12600 550 'Create a draft from reviewed values' '=ImportPDFUI()')
    [void](Control $f 100 'lblStatus' 400 8800 12600 400 'No PDF selected. Source PDFs are read only.')
    [void](Control $f 104 'cmdHome' 400 9250 12600 350 'Return home' '=OpenScreen("frmHome")')
    SaveScreen $f 'frmPDFImport'
    NewScreen 'frmBackup' 'Verified workspace backup and restore' 10300; $f=$script:swiftForm
    [void](Control $f 100 'lblInfo' 400 950 12600 1250 'Backup checkpoints pending editor values and closes this Access session before copying the whole workspace. Restore copies a closed backup to a new file. Existing files are never overwritten. The worker verifies the SHA256 byte hash and opens the copied database read only. A verification file is written beside the new database; reopen Swift NAVFIT to read the result.')
    [void](Control $f 104 'cmdBackup' 400 2450 6000 550 'Checkpoint, close Access and back up' '=GuidedWorkspaceCopy("Backup")')
    [void](Control $f 104 'cmdRestore' 6800 2450 6000 550 'Close Access and restore to a new file' '=GuidedWorkspaceCopy("Restore")')
    [void](Control $f 104 'cmdStatus' 400 3250 12600 550 'Read a copy verification result' '=ShowCopyVerification()')
    [void](Control $f 109 'txtResult' 400 4150 12600 4400);$c=$script:swiftControl;$c.Locked=$true;$c.ScrollBars=2
    [void](Control $f 100 'lblStatus' 400 8800 12600 650 'Requires Windows PowerShell and Access DAO. Other Access sessions must also release the source. Pending checkpoint values still require deliberate recovery after reopening. Runtime and other workstations remain untested.')
    [void](Control $f 104 'cmdHome' 400 9700 12600 450 'Return home' '=OpenScreen("frmHome")')
    SaveScreen $f 'frmBackup'
    . (Join-Path $PSScriptRoot 'invoke-build-phase.ps1') -BuildPhase 'presentation'
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
    . (Join-Path $PSScriptRoot 'invoke-build-phase.ps1') -BuildPhase 'finalize'
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
