param([string]$OutputPath = (Join-Path $PSScriptRoot '..\output\Swift_NAVFIT_0.12.0-assessment.accdb'))
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
if (Test-Path -LiteralPath $OutputPath) { throw "Refusing to overwrite existing database: $OutputPath" }
[void](New-Item -ItemType Directory -Force -Path (Split-Path $OutputPath -Parent))
$app = $null
try {
    $app = New-Object -ComObject Access.Application
    $app.NewCurrentDatabase($OutputPath, 12)
    $db = $app.CurrentDb()
    $definitions = @(
        'CREATE TABLE AppMetadata (MetadataID LONG CONSTRAINT PK_AppMetadata PRIMARY KEY, AppVersion TEXT(40), SchemaVersion LONG, CompatibilityStatus TEXT(120))',
        'CREATE TABLE UserProfile (UserProfileID COUNTER CONSTRAINT PK_UserProfile PRIMARY KEY, ProfileGuid TEXT(38) NOT NULL, DisplayName TEXT(120) NOT NULL)',
        'CREATE TABLE Workspace (WorkspaceID LONG CONSTRAINT PK_Workspace PRIMARY KEY, WorkspaceGuid TEXT(38) NOT NULL, CommandName TEXT(120) NOT NULL, ActiveProfileID LONG NOT NULL)',
        'CREATE TABLE Folder (FolderID COUNTER CONSTRAINT PK_Folder PRIMARY KEY, FolderGuid TEXT(38) NOT NULL, FolderName TEXT(80) NOT NULL, ParentFolderID LONG)',
        'CREATE TABLE Report (ReportID COUNTER CONSTRAINT PK_Report PRIMARY KEY, ReportGuid TEXT(38) NOT NULL, FormKind TEXT(12) NOT NULL, Title TEXT(160), MemberName TEXT(120), PeriodStart DATETIME, PeriodEnd DATETIME, DraftNotes LONGTEXT, OfficialComments LONGTEXT, WorkflowState TEXT(40), IsArchived BIT, CreatedAt DATETIME, ModifiedAt DATETIME, CurrentRevisionID LONG, FolderID LONG, OfficialName TEXT(27), RankRate TEXT(5), Designator TEXT(12), UnitUIC TEXT(5), ShipStation TEXT(18), Trait33 TEXT(3), Trait34 TEXT(3), Trait35 TEXT(3), Trait36 TEXT(3), Trait37 TEXT(3), Trait38 TEXT(3), Trait39 TEXT(3), ObservationStatus TEXT(8), PromotionRecommendation TEXT(20))',
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
        'CREATE TABLE RecoveryDraft (RecoveryDraftID COUNTER CONSTRAINT PK_RecoveryDraft PRIMARY KEY, SessionGuid TEXT(38) NOT NULL, AuthorProfileID LONG NOT NULL, ReportID LONG, BaseRevisionID LONG, CheckpointAt DATETIME, FormKind TEXT(12), DraftTitle LONGTEXT, MemberDisplay LONGTEXT, StartEntry LONGTEXT, EndEntry LONGTEXT, DraftNotes LONGTEXT, OfficialComments LONGTEXT, OfficialName LONGTEXT, RankRate LONGTEXT, Designator LONGTEXT, UnitUIC LONGTEXT, ShipStation LONGTEXT, Trait33 LONGTEXT, Trait34 LONGTEXT, Trait35 LONGTEXT, Trait36 LONGTEXT, Trait37 LONGTEXT, Trait38 LONGTEXT, Trait39 LONGTEXT, ObservationStatus LONGTEXT, PromotionRecommendation LONGTEXT)',
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
        'ALTER TABLE Report ADD CONSTRAINT FK_ReportCurrentRevision FOREIGN KEY (CurrentRevisionID) REFERENCES ReportRevision (RevisionID)',
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
        "INSERT INTO AppMetadata VALUES (1,'0.12.0-prototype',11,'Unverified - NAVFIT exchange disabled')"
    )
    foreach ($sql in $definitions) { $db.Execute($sql,128) }
    for($ti=0;$ti -lt $db.TableDefs.Count;$ti++) {
        $table=$db.TableDefs.Item($ti)
        if($table.Name.StartsWith('MSys')) {continue}
        for($fi=0;$fi -lt $table.Fields.Count;$fi++) {
            $field=$table.Fields.Item($fi)
            if(($field.Type -eq 10 -or $field.Type -eq 12) -and -not $field.Required) {$field.AllowZeroLength=$true}
        }
    }
    $app.LoadFromText(5,'modBulkFolders',(Join-Path $root 'src\modBulkFolders.bas'))
    $app.LoadFromText(5,'modOperations',(Join-Path $root 'src\modOperations.bas'))
    $app.LoadFromText(5,'modProvenance',(Join-Path $root 'src\modProvenance.bas'))
    $app.LoadFromText(5,'modReviewComments',(Join-Path $root 'src\modReviewComments.bas'))
    $app.LoadFromText(5,'modTraits',(Join-Path $root 'src\modTraits.bas'))
    [void](New-Item -ItemType Directory -Force -Path (Join-Path $root 'build'))
    $navTraitModule=Join-Path $root ('build\trait-catalogue-'+[guid]::NewGuid()+'.bas')
    & (Join-Path $PSScriptRoot 'write-trait-catalogue.ps1') -SourcePath (Join-Path $root 'src\trait-catalogue.json') -OutputPath $navTraitModule
    $app.LoadFromText(5,'modTraitCatalogue',$navTraitModule)
    $app.LoadFromText(5,'modAssessment',(Join-Path $root 'src\modAssessment.bas'))
    $app.LoadFromText(5,'modAdministration',(Join-Path $root 'src\modAdministration.bas'))
    $app.LoadFromText(5,'modRouting',(Join-Path $root 'src\modRouting.bas'))
    $app.LoadFromText(5,'modSnippets',(Join-Path $root 'src\modSnippets.bas'))
    $app.LoadFromText(5,'modLibraryViews',(Join-Path $root 'src\modLibraryViews.bas'))
    $app.LoadFromText(5,'modChecklist',(Join-Path $root 'src\modChecklist.bas'))
    $app.LoadFromText(5,'modRecovery',(Join-Path $root 'src\modRecovery.bas'))
    $app.LoadFromText(5,'modRevisionCompare',(Join-Path $root 'src\modRevisionCompare.bas'))
    $app.LoadFromText(5,'modFileHash',(Join-Path $root 'src\modFileHash.bas'))
    $app.LoadFromText(5,'modSwift',(Join-Path $root 'src\modSwift.bas'))
    $app.LoadFromText(5,'modWorkspace',(Join-Path $root 'src\modWorkspace.bas'))
    $app.LoadFromText(5,'modImportReview',(Join-Path $root 'src\modImportReview.bas'))
    function Control($f,$type,$name,$x,$y,$w,$h,$caption='',$event='') {
        $c = $app.CreateControl($f.Name,$type,0,'','',$x,$y,$w,$h)
        $c.Name=$name
        if ($type -eq 100 -or $type -eq 104) { $c.Caption=$caption }
        if ($type -ne 106) { $c.FontName='Segoe UI'; $c.FontSize=10 }
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
        $app.DoCmd.Save(2,$name); $app.DoCmd.Close(2,$name,2)
    }
    NewScreen 'frmSetup' 'Set up your Swift NAVFIT workspace' 5300; $f=$script:swiftForm
    [void](Control $f 100 'lblInfo' 400 1000 12600 650 'One authoritative local file. Profiles identify contributors; they do not authenticate users. Separate copies do not synchronize.')
    [void](Control $f 100 'lblCommand' 400 1950 2800 350 'Command name')
    [void](Control $f 109 'txtCommand' 3300 1900 7000 420)
    [void](Control $f 100 'lblUser' 400 2750 2800 350 'Contributor name')
    [void](Control $f 109 'txtProfile' 3300 2700 7000 420)
    [void](Control $f 104 'cmdSetup' 3300 3550 4000 550 'Create workspace' '=SaveSetup()')
    SaveScreen $f 'frmSetup'
    NewScreen 'frmHome' 'Swift NAVFIT' 7800; $f=$script:swiftForm
    [void](Control $f 100 'lblCommand' 400 1000 12400 450 '')
    [void](Control $f 100 'lblProfile' 400 1550 12400 450 '')
    [void](Control $f 100 'lblPrototype' 400 2200 12400 800 'Draft prototype. Administrative block editing, policy validation, PDF import and NAVFIT exchange are not available in this build.')
    [void](Control $f 104 'cmdLibrary' 400 3300 4000 650 'Open draft library' '=OpenScreen("frmLibrary")')
    [void](Control $f 104 'cmdSupport' 4900 3300 4000 650 'Support and limitations' '=OpenScreen("frmSupport")')
    [void](Control $f 100 'lblUse' 400 4400 12400 700 'Use one person at a time. Open this file exclusively. Close Access before making a backup copy. This prototype does not enforce exclusive use.')
    [void](Control $f 104 'cmdImportReview' 9400 3300 3500 650 'Review NAVFIT database' '=ChooseReviewFile()')
    [void](Control $f 104 'cmdRecovery' 400 5250 5500 550 'Recover pending edits' '=OpenScreen("frmRecovery")')
    [void](Control $f 104 'cmdSnippets' 400 6100 5500 550 'Narrative snippets' '=OpenScreen("frmSnippets")')
    [void](Control $f 104 'cmdOperations' 6500 6100 6300 550 'Operations and import undo' '=OpenScreen("frmOperations")')
    [void](Control $f 104 'cmdMyWork' 400 6800 12600 550 'My Work and local contributors' '=OpenMyWork()')
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
    NewScreen 'frmEditor' 'Draft editor' 31000; $f=$script:swiftForm
    $f.Width=13800
    [void](Control $f 100 'lblScope' 400 900 12400 550 'Draft administration and form-specific narrative. Official validation, full NAVFIT mapping and printed fit remain pending.')
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
    [void](Control $f 100 'lblAssessmentScope' 400 5350 12600 350 'Explicit selections; no inference from trait marks. Eligibility and Navy rules remain unassessed.')
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
    $f.TimerInterval=30000; $f.OnTimer='=CheckpointEditor()'
    $f.OnLoad='=LoadEditor()'; $f.ScrollBars=2; SaveScreen $f 'frmEditor'
    NewScreen 'frmSupport' 'Support and limitations' 7400; $f=$script:swiftForm
    Control $f 109 'txtDiagnostics' 400 1100 12600 4500; $c=$script:swiftControl; $c.Locked=$true; $c.ScrollBars=2
    [void](Control $f 104 'cmdHome' 400 6150 2800 550 'Home' '=OpenScreen("frmHome")')
    $f.OnLoad='=LoadSupport()'; SaveScreen $f 'frmSupport'
    NewScreen 'frmHelp' 'Contextual help' 6300; $f=$script:swiftForm
    [void](Control $f 100 'lblTopic' 400 1050 12400 450 '')
    Control $f 109 'txtHelp' 400 1700 12400 3200; $c=$script:swiftControl; $c.Locked=$true; $c.ScrollBars=2
    [void](Control $f 104 'cmdClose' 400 5300 2800 550 'Return to work' '=CloseHelp()')
    SaveScreen $f 'frmHelp'
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
    NewScreen 'frmChecklist' 'Saved draft completeness checklist' 9900; $f=$script:swiftForm
    [void](Control $f 100 'lblContext' 400 1000 12600 450 '')
    [void](Control $f 100 'lblSummary' 400 1550 12600 650 '')
    Control $f 109 'txtReportID' 10 10 100 100; $script:swiftControl.Visible=$false
    Control $f 110 'lstFindings' 400 2400 12600 3300; $c=$script:swiftControl
    $c.ColumnCount=3; $c.ColumnWidths='0;2200;10000'; $c.BoundColumn=1; $c.AfterUpdate='=ChecklistGuidance()'
    Control $f 109 'txtGuidance' 400 6100 12600 2200; $c=$script:swiftControl; $c.Locked=$true; $c.ScrollBars=2
    [void](Control $f 104 'cmdRefresh' 400 8850 3500 550 'Refresh saved values' '=RefreshDraftChecklist()')
    [void](Control $f 104 'cmdClose' 4300 8850 3500 550 'Return to workspace' '=CloseDraftChecklist()')
    SaveScreen $f 'frmChecklist'
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
    NewScreen 'frmReviewComments' 'Saved draft review comments' 13900; $f=$script:swiftForm
    [void](Control $f 100 'lblScope' 400 950 12600 700 'Review comments stay separate from official narratives. Threads reference saved revisions; replies preserve that reference. Local contributor profiles identify activity without authentication.')
    Control $f 109 'txtReportID' 10 10 100 100; $script:swiftControl.Visible=$false
    [void](Control $f 100 'lblContext' 400 1800 12600 600 '')
    Control $f 110 'lstThreads' 400 2600 12600 2000; $c=$script:swiftControl
    $c.ColumnCount=4; $c.ColumnWidths='0;2000;5600;5000'; $c.BoundColumn=1; $c.AfterUpdate='=ShowReviewThread()'
    [void](Control $f 100 'lblHistory' 400 4800 12600 350 'Selected thread - comments and resolution history')
    Control $f 109 'txtHistory' 400 5300 12600 3800; $c=$script:swiftControl; $c.Locked=$true; $c.ScrollBars=2
    [void](Control $f 100 'lblCompose' 400 9300 12600 400 'Write a new thread or reply to the selected thread')
    Control $f 109 'txtCompose' 400 9800 12600 1700; $c=$script:swiftControl; $c.ScrollBars=2; $c.EnterKeyBehavior=$true
    [void](Control $f 104 'cmdPost' 400 11800 3800 550 'Post new thread' '=PostReviewUI()')
    [void](Control $f 104 'cmdReply' 4500 11800 3800 550 'Post reply' '=PostReviewUI(True)')
    [void](Control $f 104 'cmdResolve' 8600 11800 4400 550 'Resolve thread' '=ReviewStateUI(True)')
    [void](Control $f 104 'cmdReopen' 400 12600 3800 550 'Reopen thread' '=ReviewStateUI(False)')
    [void](Control $f 104 'cmdEditor' 4500 12600 8500 550 'Return to open editor' '=ReturnFromReviewComments()')
    SaveScreen $f 'frmReviewComments'
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
    $macro=Join-Path $root 'src\AutoExec.txt'
    $app.LoadFromText(4,'AutoExec',$macro)
    $app.DoCmd.RunCommand(126)
    $db.Close(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($db)
    $app.CloseCurrentDatabase()
    [pscustomobject]@{Output=$OutputPath;SHA256=(Get-FileHash -LiteralPath $OutputPath -Algorithm SHA256).Hash;Status='Built; smoke test required'} | ConvertTo-Json
} finally {
    if ($null -ne $app) { $app.Quit(2); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
}
