-- schema: 24
-- kind: captured
-- source: captured from Swift_NAVFIT_UI_Blank_2026-10-07-r89.accdb (0.25.0-dev)
-- appVersion: 0.25.0-dev
-- note: Same structure, ignoring generated index names, in: Swift_NAVFIT_UI_Blank_2026-10-06-r86.accdb, Swift_NAVFIT_UI_Blank_2026-10-06-r87.accdb, Swift_NAVFIT_UI_Blank_2026-10-06-r88.accdb, Swift_NAVFIT_UI_Blank_2026-10-07-r89.accdb
CREATE TABLE [AppMetadata] ([MetadataID] LONG, [AppVersion] TEXT(40), [SchemaVersion] LONG, [CompatibilityStatus] TEXT(120))
CREATE TABLE [ChangeBatch] ([BatchID] COUNTER, [BatchGuid] TEXT(38) NOT NULL, [BatchKind] TEXT(30), [BatchState] TEXT(20), [AuthorID] LONG NOT NULL, [CreatedAt] DATETIME, [BatchNote] LONGTEXT)
CREATE TABLE [ChangeItem] ([ItemID] COUNTER, [BatchID] LONG NOT NULL, [ReportID] LONG NOT NULL, [ExpectedRevision] LONG NOT NULL, [FieldKey] TEXT(64) NOT NULL, [NewValue] LONGTEXT, [NewIsNull] BIT, [BeforeValue] LONGTEXT, [BeforeIsNull] BIT, [AppliedRevision] LONG)
CREATE TABLE [CommandMember] ([MemberID] COUNTER, [MemberGuid] TEXT(38) NOT NULL, [FullName] TEXT(120) NOT NULL, [ReportName] TEXT(27), [RankRate] TEXT(5), [MemberSSN] TEXT(9), [SourceKind] TEXT(20), [IsArchived] BIT, [VersionCounter] LONG NOT NULL, [CreatedBy] LONG NOT NULL, [CreatedAt] DATETIME, [ModifiedBy] LONG NOT NULL, [ModifiedAt] DATETIME)
CREATE TABLE [ComparisonMember] ([ComparisonMemberID] COUNTER, [SetID] LONG NOT NULL, [ReportID] LONG NOT NULL, [RevisionID] LONG NOT NULL, [SortOrdinal] LONG)
CREATE TABLE [ComparisonSet] ([SetID] COUNTER, [GroupSnapshotID] LONG NOT NULL, [Title] TEXT(120), [CreatedBy] LONG NOT NULL, [CreatedAt] DATETIME)
CREATE TABLE [FieldDefault] ([DefaultID] COUNTER, [ScopeKind] TEXT(12) NOT NULL, [ScopeID] LONG NOT NULL, [FieldKey] TEXT(64) NOT NULL, [ValueText] LONGTEXT, [IsNullValue] BIT, [VersionCounter] LONG)
CREATE TABLE [Folder] ([FolderID] COUNTER, [FolderGuid] TEXT(38) NOT NULL, [FolderName] TEXT(80) NOT NULL, [ParentFolderID] LONG)
CREATE TABLE [FolderBatchChange] ([FolderBatchChangeID] COUNTER, [OperationID] LONG NOT NULL, [ReportID] LONG NOT NULL, [OldFolderID] LONG, [NewFolderID] LONG, [AppliedMoveID] LONG NOT NULL, [BaseRevisionID] LONG NOT NULL, [IsUndone] BIT)
CREATE TABLE [FolderMove] ([FolderMoveID] COUNTER, [ReportID] LONG NOT NULL, [OldFolderID] LONG, [NewFolderID] LONG, [AuthorProfileID] LONG NOT NULL, [MovedAt] DATETIME)
CREATE TABLE [GroupMember] ([MemberID] COUNTER, [GroupID] LONG NOT NULL, [ReportID] LONG NOT NULL, [Paygrade] TEXT(5) NOT NULL, [ExceptionReason] LONGTEXT)
CREATE TABLE [GroupRSCABaseline] ([GroupID] LONG, [SeniorID] LONG NOT NULL, [BaselineID] LONG, [ExcludesGroup] BIT)
CREATE TABLE [GroupSnapshot] ([SnapshotID] COUNTER, [GroupID] LONG NOT NULL, [Stamp] LONGTEXT, [AttestedBy] LONG NOT NULL, [AttestedAt] DATETIME, [CompletenessNote] LONGTEXT)
CREATE TABLE [GroupSnapshotMember] ([SnapshotMemberID] COUNTER, [SnapshotID] LONG NOT NULL, [ReportID] LONG NOT NULL, [RevisionID] LONG NOT NULL, [Promotion] TEXT(20), [TraitAverage] DOUBLE)
CREATE TABLE [ImportRawValue] ([ImportRawValueID] COUNTER, [ImportReceiptID] LONG NOT NULL, [SourceTable] TEXT(40) NOT NULL, [SourceRow] LONG NOT NULL, [FieldOrdinal] LONG NOT NULL, [FieldKey] TEXT(64) NOT NULL, [DAOType] LONG NOT NULL, [RawValue] LONGTEXT, [IsNullValue] BIT)
CREATE TABLE [ImportReceipt] ([ImportReceiptID] COUNTER, [ReceiptGuid] TEXT(38) NOT NULL, [SourceHash] TEXT(64) NOT NULL, [SourcePath] LONGTEXT, [ExternalReportID] LONG NOT NULL, [ReportID] LONG NOT NULL, [AuthorProfileID] LONG NOT NULL, [ImportedAt] DATETIME, [AdapterVersion] TEXT(40), [ValueEncoding] TEXT(80), [BeforeRevisionID] LONG, [AppliedRevisionID] LONG, [OperationID] LONG, [IsUndone] BIT)
CREATE TABLE [LibraryPreference] ([LibraryPreferenceID] COUNTER, [AuthorProfileID] LONG NOT NULL, [ViewPreset] TEXT(12))
CREATE TABLE [MeetingProposal] ([ProposalID] COUNTER, [SetID] LONG NOT NULL, [ReportID] LONG NOT NULL, [BaseRevision] LONG NOT NULL, [FieldKey] TEXT(64), [ValueText] LONGTEXT, [IsNullValue] BIT, [ProposalState] TEXT(20), [AppliedRevision] LONG)
CREATE TABLE [MemberReport] ([ReportID] LONG, [MemberID] LONG NOT NULL, [MatchState] TEXT(20), [MatchSource] TEXT(30), [UpdatedBy] LONG NOT NULL, [UpdatedAt] DATETIME)
CREATE TABLE [NarrativeSnippet] ([SnippetID] COUNTER, [SnippetGuid] TEXT(38) NOT NULL, [CreatedBy] LONG NOT NULL, [SnippetTitle] TEXT(120), [SnippetBody] LONGTEXT, [Attribution] TEXT(255), [FormKind] TEXT(12), [ModifiedAt] DATETIME)
CREATE TABLE [OperationJournal] ([OperationID] COUNTER, [OperationGuid] TEXT(38) NOT NULL, [AuthorProfileID] LONG NOT NULL, [OperationKind] TEXT(40), [OperationState] TEXT(20), [SourceDescription] LONGTEXT, [StartedAt] DATETIME, [FinishedAt] DATETIME, [Outcome] LONGTEXT)
CREATE TABLE [PriorEvaluation] ([PriorID] COUNTER, [PriorGuid] TEXT(38) NOT NULL, [MemberID] LONG NOT NULL, [FormKind] TEXT(12) NOT NULL, [RankRate] TEXT(5), [PeriodStart] DATETIME NOT NULL, [PeriodEnd] DATETIME NOT NULL, [PromotionRecommendation] TEXT(30), [ObservationStatus] TEXT(16), [SourceNote] TEXT(255), [IsArchived] BIT, [VersionCounter] LONG NOT NULL, [CreatedBy] LONG NOT NULL, [CreatedAt] DATETIME, [ModifiedBy] LONG NOT NULL, [ModifiedAt] DATETIME)
CREATE TABLE [PriorEvaluationTrait] ([PriorID] LONG NOT NULL, [BlockNumber] SHORT NOT NULL, [MarkValue] TEXT(3))
CREATE TABLE [RecoveryDraft] ([RecoveryDraftID] COUNTER, [SessionGuid] TEXT(38) NOT NULL, [AuthorProfileID] LONG NOT NULL, [ReportID] LONG, [BaseRevisionID] LONG, [CheckpointAt] DATETIME, [FormKind] TEXT(12), [DraftTitle] LONGTEXT, [MemberDisplay] LONGTEXT, [StartEntry] LONGTEXT, [EndEntry] LONGTEXT, [DraftNotes] LONGTEXT, [OfficialComments] LONGTEXT, [OfficialName] LONGTEXT, [RankRate] LONGTEXT, [Designator] LONGTEXT, [UnitUIC] LONGTEXT, [ShipStation] LONGTEXT, [Trait33] LONGTEXT, [Trait34] LONGTEXT, [Trait35] LONGTEXT, [Trait36] LONGTEXT, [Trait37] LONGTEXT, [Trait38] LONGTEXT, [Trait39] LONGTEXT, [ObservationStatus] LONGTEXT, [PromotionRecommendation] LONGTEXT, [PromotionStatus] LONGTEXT, [ReportingSeniorName] LONGTEXT, [ReportingSeniorGrade] LONGTEXT, [ReportingSeniorDesignator] LONGTEXT, [ReportingSeniorTitle] LONGTEXT, [ReportingSeniorUIC] LONGTEXT, [OccasionPeriodic] LONGTEXT, [OccasionDetachment] LONGTEXT, [OccasionFrocking] LONGTEXT, [OccasionSpecial] LONGTEXT, [ReportRegular] LONGTEXT, [ReportConcurrent] LONGTEXT, [ReportOperationalCommander] LONGTEXT, [PrimaryDuty] LONGTEXT, [DutyDescription] LONGTEXT, [MemberSSN] LONGTEXT, [MemberDoDID] LONGTEXT, [ComponentActive] LONGTEXT, [ComponentTAR] LONGTEXT, [ComponentInactive] LONGTEXT, [ComponentATADSW] LONGTEXT, [DateReported] LONGTEXT, [CounselingDate] LONGTEXT, [CounselorName] LONGTEXT, [ReadinessSummary] LONGTEXT, [ReadinessDetail] LONGTEXT, [ReadinessDate] LONGTEXT, [BilletSubcategory] LONGTEXT, [AssignmentRecommendationA] LONGTEXT, [AssignmentRecommendationB] LONGTEXT, [RetentionYes] LONGTEXT, [RetentionNo] LONGTEXT, [StatementYes] LONGTEXT, [StatementNo] LONGTEXT, [ReportingSeniorSSN] LONGTEXT, [ReportingSeniorDoDID] LONGTEXT, [ReportingSeniorAddress1] LONGTEXT, [ReportingSeniorAddress2] LONGTEXT, [ReportingSeniorCity] LONGTEXT, [ReportingSeniorState] LONGTEXT, [ReportingSeniorZip] LONGTEXT, [ReportingSeniorPhone] LONGTEXT, [ReportingSeniorDSN] LONGTEXT, [RaterName] LONGTEXT, [RaterDate] LONGTEXT, [SeniorRaterName] LONGTEXT, [SeniorRaterDate] LONGTEXT, [MemberFirstName] LONGTEXT, [MemberMiddleInitial] LONGTEXT, [MemberLastName] LONGTEXT, [MemberSuffix] LONGTEXT, [ReportingSeniorLastName] LONGTEXT, [ReportingSeniorFirstInitial] LONGTEXT, [ReportingSeniorMiddleInitial] LONGTEXT, [CounselorLastName] LONGTEXT, [CounselorFirstInitial] LONGTEXT, [CounselorMiddleInitial] LONGTEXT, [CommandEmployment] LONGTEXT)
CREATE TABLE [Report] ([ReportID] COUNTER, [ReportGuid] TEXT(38) NOT NULL, [FormKind] TEXT(12) NOT NULL, [Title] TEXT(160), [MemberName] TEXT(120), [PeriodStart] DATETIME, [PeriodEnd] DATETIME, [DraftNotes] LONGTEXT, [OfficialComments] LONGTEXT, [WorkflowState] TEXT(40), [IsArchived] BIT, [CreatedAt] DATETIME, [ModifiedAt] DATETIME, [CurrentRevisionID] LONG, [FolderID] LONG, [OfficialName] TEXT(27), [RankRate] TEXT(5), [Designator] TEXT(12), [UnitUIC] TEXT(5), [ShipStation] TEXT(18), [Trait33] TEXT(3), [Trait34] TEXT(3), [Trait35] TEXT(3), [Trait36] TEXT(3), [Trait37] TEXT(3), [Trait38] TEXT(3), [Trait39] TEXT(3), [ObservationStatus] TEXT(8), [PromotionRecommendation] TEXT(20), [PromotionStatus] TEXT(8), [ReportingSeniorName] TEXT(18), [ReportingSeniorGrade] TEXT(5), [ReportingSeniorDesignator] TEXT(5), [ReportingSeniorTitle] TEXT(14), [ReportingSeniorUIC] TEXT(5), [OccasionPeriodic] TEXT(3), [OccasionDetachment] TEXT(3), [OccasionFrocking] TEXT(3), [OccasionSpecial] TEXT(3), [ReportRegular] TEXT(3), [ReportConcurrent] TEXT(3), [ReportOperationalCommander] TEXT(3), [PrimaryDuty] TEXT(14), [DutyDescription] LONGTEXT, [MemberSSN] TEXT(9), [MemberDoDID] TEXT(10), [ComponentActive] TEXT(3), [ComponentTAR] TEXT(3), [ComponentInactive] TEXT(3), [ComponentATADSW] TEXT(3), [DateReported] TEXT(10), [CounselingDate] TEXT(10), [CounselorName] TEXT(20), [ReadinessSummary] TEXT(4), [ReadinessDetail] TEXT(15), [ReadinessDate] TEXT(10), [BilletSubcategory] TEXT(10), [AssignmentRecommendationA] TEXT(20), [AssignmentRecommendationB] TEXT(20), [RetentionYes] TEXT(3), [RetentionNo] TEXT(3), [StatementYes] TEXT(3), [StatementNo] TEXT(3), [ReportingSeniorSSN] TEXT(9), [ReportingSeniorDoDID] TEXT(10), [ReportingSeniorAddress1] TEXT(30), [ReportingSeniorAddress2] TEXT(30), [ReportingSeniorCity] TEXT(15), [ReportingSeniorState] TEXT(2), [ReportingSeniorZip] TEXT(9), [ReportingSeniorPhone] TEXT(10), [ReportingSeniorDSN] TEXT(7), [RaterName] TEXT(28), [RaterDate] TEXT(10), [SeniorRaterName] TEXT(28), [SeniorRaterDate] TEXT(10), [MemberFirstName] TEXT(27), [MemberMiddleInitial] TEXT(3), [MemberLastName] TEXT(27), [MemberSuffix] TEXT(10), [ReportingSeniorLastName] TEXT(18), [ReportingSeniorFirstInitial] TEXT(1), [ReportingSeniorMiddleInitial] TEXT(3), [CounselorLastName] TEXT(20), [CounselorFirstInitial] TEXT(1), [CounselorMiddleInitial] TEXT(3), [CommandEmployment] LONGTEXT)
CREATE TABLE [ReportEvent] ([ReportEventID] COUNTER, [ReportID] LONG NOT NULL, [AuthorProfileID] LONG NOT NULL, [EventAt] DATETIME, [EventKind] TEXT(40))
CREATE TABLE [ReportingSeniorRecord] ([SeniorID] COUNTER, [SeniorGuid] TEXT(38) NOT NULL, [Title] TEXT(120) NOT NULL, [CreatedBy] LONG NOT NULL, [ModifiedAt] DATETIME, [VersionCounter] LONG)
CREATE TABLE [ReportingSeniorValue] ([ValueID] COUNTER, [SeniorID] LONG NOT NULL, [FieldKey] TEXT(64) NOT NULL, [ValueText] TEXT(255), [IsNullValue] BIT)
CREATE TABLE [ReportNote] ([NoteID] COUNTER, [NoteGuid] TEXT(38) NOT NULL, [ReportID] LONG NOT NULL, [AuthorProfileID] LONG, [ContributorName] TEXT(160) NOT NULL, [NoteText] LONGTEXT, [CreatedAt] DATETIME, [SourceKind] TEXT(20), [LegacyReportID] LONG)
CREATE TABLE [ReportNoteDraft] ([DraftID] COUNTER, [DraftGuid] TEXT(38) NOT NULL, [ReportID] LONG NOT NULL, [AuthorProfileID] LONG NOT NULL, [NoteText] LONGTEXT, [VersionCounter] LONG NOT NULL, [CheckpointAt] DATETIME)
CREATE TABLE [ReportRevision] ([RevisionID] COUNTER, [ReportID] LONG NOT NULL, [RevisionNumber] LONG NOT NULL, [ParentRevisionID] LONG, [AuthorProfileID] LONG NOT NULL, [CommittedAt] DATETIME, [Reason] TEXT(255), [SnapshotSchema] LONG)
CREATE TABLE [ReviewComment] ([CommentID] COUNTER, [CommentGuid] TEXT(38) NOT NULL, [ReportID] LONG NOT NULL, [RevisionID] LONG NOT NULL, [ThreadID] LONG, [AuthorProfileID] LONG NOT NULL, [CommentAt] DATETIME, [CommentBody] LONGTEXT)
CREATE TABLE [ReviewCommentEvent] ([CommentEventID] COUNTER, [ThreadID] LONG NOT NULL, [AuthorProfileID] LONG NOT NULL, [EventAt] DATETIME, [EventKind] TEXT(20))
CREATE TABLE [ReviewTask] ([TaskID] COUNTER, [TaskGuid] TEXT(38) NOT NULL, [ReportID] LONG NOT NULL, [RevisionID] LONG NOT NULL, [RequestedBy] LONG NOT NULL, [AssignedTo] LONG NOT NULL, [RequestedAt] DATETIME, [DueDate] DATETIME, [Instructions] LONGTEXT, [TaskState] TEXT(20), [FinishedAt] DATETIME)
CREATE TABLE [ReviewTaskEvent] ([TaskEventID] COUNTER, [TaskID] LONG NOT NULL, [AuthorProfileID] LONG NOT NULL, [EventAt] DATETIME, [EventKind] TEXT(20), [Notes] LONGTEXT)
CREATE TABLE [RevisionChange] ([RevisionChangeID] COUNTER, [RevisionID] LONG NOT NULL, [FieldKey] TEXT(64) NOT NULL, [OldValue] LONGTEXT, [NewValue] LONGTEXT, [OldIsNull] BIT, [NewIsNull] BIT)
CREATE TABLE [RevisionValue] ([RevisionValueID] COUNTER, [RevisionID] LONG NOT NULL, [FieldKey] TEXT(64) NOT NULL, [ValueText] LONGTEXT, [IsNullValue] BIT)
CREATE TABLE [RouteRun] ([RunID] COUNTER, [ReportID] LONG NOT NULL, [RevisionID] LONG NOT NULL, [TemplateID] LONG NOT NULL, [CurrentStep] LONG, [RunState] TEXT(20), [GroupID] LONG, [GroupStamp] LONGTEXT, [RequestedBy] LONG NOT NULL, [RequestedAt] DATETIME, [ApprovedBy] LONG, [ApprovedAt] DATETIME)
CREATE TABLE [RouteRunStep] ([RunStepID] COUNTER, [RunID] LONG NOT NULL, [StepOrdinal] LONG NOT NULL, [AssigneeID] LONG NOT NULL, [Instructions] LONGTEXT, [StepState] TEXT(20), [DecisionNote] LONGTEXT, [DecisionAt] DATETIME)
CREATE TABLE [RouteTemplate] ([TemplateID] COUNTER, [Title] TEXT(120), [IsArchived] BIT)
CREATE TABLE [RouteTemplateStep] ([TemplateStepID] COUNTER, [TemplateID] LONG NOT NULL, [StepOrdinal] LONG NOT NULL, [AssigneeID] LONG NOT NULL, [Instructions] LONGTEXT)
CREATE TABLE [RSCABaseline] ([BaselineID] COUNTER, [SeniorID] LONG NOT NULL, [Paygrade] TEXT(2) NOT NULL, [AsOfDate] DATETIME NOT NULL, [CurrentRSCA] DOUBLE, [TraitCount] LONG NOT NULL, [TraitSum] DOUBLE, [SourceNote] TEXT(255), [CreatedBy] LONG NOT NULL, [CreatedAt] DATETIME)
CREATE TABLE [RuleDefinition] ([RuleID] COUNTER, [PackID] LONG NOT NULL, [RuleCode] TEXT(64), [FieldKey] TEXT(64), [CheckKind] TEXT(30), [ExpectedValue] LONGTEXT, [FormKind] TEXT(12), [Severity] TEXT(20), [SourceLocator] TEXT(120), [Description] LONGTEXT)
CREATE TABLE [RulePack] ([PackID] COUNTER, [VersionName] TEXT(120), [SourceTitle] TEXT(200), [SourceURL] LONGTEXT, [SourceSHA256] TEXT(64), [EffectiveFrom] DATETIME, [EffectiveTo] DATETIME)
CREATE TABLE [SummaryGroup] ([GroupID] COUNTER, [GroupGuid] TEXT(38) NOT NULL, [Title] TEXT(120) NOT NULL, [FormKind] TEXT(12) NOT NULL, [Paygrade] TEXT(5) NOT NULL, [KeySpec] LONGTEXT, [VersionCounter] LONG, [IsArchived] BIT)
CREATE TABLE [UserProfile] ([UserProfileID] COUNTER, [ProfileGuid] TEXT(38) NOT NULL, [DisplayName] TEXT(120) NOT NULL)
CREATE TABLE [ValidationFinding] ([FindingID] COUNTER, [ValidationRunID] LONG NOT NULL, [RuleID] LONG, [FindingState] TEXT(20), [Message] LONGTEXT)
CREATE TABLE [ValidationRun] ([ValidationRunID] COUNTER, [ReportID] LONG NOT NULL, [RevisionID] LONG NOT NULL, [PackID] LONG NOT NULL, [RunAt] DATETIME, [GroupSnapshotID] LONG)
CREATE TABLE [Workspace] ([WorkspaceID] LONG, [WorkspaceGuid] TEXT(38) NOT NULL, [CommandName] TEXT(120) NOT NULL, [ActiveProfileID] LONG NOT NULL)
CREATE UNIQUE INDEX [PK_AppMetadata] ON [AppMetadata] ([MetadataID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_ChangeBatch] ON [ChangeBatch] ([BatchID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_ChangeItem] ON [ChangeItem] ([ItemID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_MajorChangeItem] ON [ChangeItem] ([BatchID], [ReportID], [FieldKey])
CREATE UNIQUE INDEX [PK_CommandMember] ON [CommandMember] ([MemberID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_CommandMemberGuid] ON [CommandMember] ([MemberGuid])
CREATE UNIQUE INDEX [UX_CommandMemberSSN] ON [CommandMember] ([MemberSSN]) WITH IGNORE NULL
CREATE UNIQUE INDEX [PK_ComparisonMember] ON [ComparisonMember] ([ComparisonMemberID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_MajorComparison] ON [ComparisonMember] ([SetID], [ReportID])
CREATE UNIQUE INDEX [PK_ComparisonSet] ON [ComparisonSet] ([SetID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_FieldDefault] ON [FieldDefault] ([DefaultID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_MajorDefault] ON [FieldDefault] ([ScopeKind], [ScopeID], [FieldKey])
CREATE UNIQUE INDEX [PK_Folder] ON [Folder] ([FolderID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_FolderGuid] ON [Folder] ([FolderGuid])
CREATE UNIQUE INDEX [PK_FolderBatchChange] ON [FolderBatchChange] ([FolderBatchChangeID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_FolderMove] ON [FolderMove] ([FolderMoveID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_GroupMember] ON [GroupMember] ([MemberID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_MajorGroupMember] ON [GroupMember] ([GroupID], [ReportID])
CREATE UNIQUE INDEX [PK_GroupRSCABaseline] ON [GroupRSCABaseline] ([GroupID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_GroupSnapshot] ON [GroupSnapshot] ([SnapshotID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_GroupSnapshotMember] ON [GroupSnapshotMember] ([SnapshotMemberID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_MajorSnapshotMember] ON [GroupSnapshotMember] ([SnapshotID], [ReportID])
CREATE UNIQUE INDEX [PK_ImportRawValue] ON [ImportRawValue] ([ImportRawValueID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_ImportRaw] ON [ImportRawValue] ([ImportReceiptID], [SourceTable], [SourceRow], [FieldOrdinal])
CREATE UNIQUE INDEX [PK_ImportReceipt] ON [ImportReceipt] ([ImportReceiptID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_ImportSource] ON [ImportReceipt] ([SourceHash], [ExternalReportID])
CREATE UNIQUE INDEX [PK_LibraryPreference] ON [LibraryPreference] ([LibraryPreferenceID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_LibraryProfile] ON [LibraryPreference] ([AuthorProfileID])
CREATE UNIQUE INDEX [PK_MeetingProposal] ON [MeetingProposal] ([ProposalID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_MemberReport] ON [MemberReport] ([ReportID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_NarrativeSnippet] ON [NarrativeSnippet] ([SnippetID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_OperationJournal] ON [OperationJournal] ([OperationID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_PriorEvaluation] ON [PriorEvaluation] ([PriorID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_PriorGuid] ON [PriorEvaluation] ([PriorGuid])
CREATE UNIQUE INDEX [PK_PriorTrait] ON [PriorEvaluationTrait] ([PriorID], [BlockNumber]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_RecoveryDraft] ON [RecoveryDraft] ([RecoveryDraftID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_RecoverySession] ON [RecoveryDraft] ([SessionGuid])
CREATE UNIQUE INDEX [PK_Report] ON [Report] ([ReportID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_ReportGuid] ON [Report] ([ReportGuid])
CREATE UNIQUE INDEX [PK_ReportEvent] ON [ReportEvent] ([ReportEventID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_ReportingSeniorRecord] ON [ReportingSeniorRecord] ([SeniorID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_RSGuid] ON [ReportingSeniorRecord] ([SeniorGuid])
CREATE UNIQUE INDEX [PK_ReportingSeniorValue] ON [ReportingSeniorValue] ([ValueID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_RSValue] ON [ReportingSeniorValue] ([SeniorID], [FieldKey])
CREATE UNIQUE INDEX [PK_ReportNote] ON [ReportNote] ([NoteID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_LegacyNote] ON [ReportNote] ([LegacyReportID]) WITH IGNORE NULL
CREATE UNIQUE INDEX [UX_ReportNoteGuid] ON [ReportNote] ([NoteGuid])
CREATE UNIQUE INDEX [PK_ReportNoteDraft] ON [ReportNoteDraft] ([DraftID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_NoteDraftGuid] ON [ReportNoteDraft] ([DraftGuid])
CREATE UNIQUE INDEX [PK_ReportRevision] ON [ReportRevision] ([RevisionID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_ReportRevision] ON [ReportRevision] ([ReportID], [RevisionNumber])
CREATE UNIQUE INDEX [PK_ReviewComment] ON [ReviewComment] ([CommentID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_CommentGuid] ON [ReviewComment] ([CommentGuid])
CREATE UNIQUE INDEX [PK_ReviewCommentEvent] ON [ReviewCommentEvent] ([CommentEventID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_ReviewTask] ON [ReviewTask] ([TaskID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_TaskGuid] ON [ReviewTask] ([TaskGuid])
CREATE UNIQUE INDEX [PK_ReviewTaskEvent] ON [ReviewTaskEvent] ([TaskEventID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_RevisionChange] ON [RevisionChange] ([RevisionChangeID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_RevisionValue] ON [RevisionValue] ([RevisionValueID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_RevisionValue] ON [RevisionValue] ([RevisionID], [FieldKey])
CREATE UNIQUE INDEX [PK_RouteRun] ON [RouteRun] ([RunID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_RouteRunStep] ON [RouteRunStep] ([RunStepID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_MajorRunStep] ON [RouteRunStep] ([RunID], [StepOrdinal])
CREATE UNIQUE INDEX [PK_RouteTemplate] ON [RouteTemplate] ([TemplateID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_RouteTemplateStep] ON [RouteTemplateStep] ([TemplateStepID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_MajorRouteStep] ON [RouteTemplateStep] ([TemplateID], [StepOrdinal])
CREATE UNIQUE INDEX [PK_RSCABaseline] ON [RSCABaseline] ([BaselineID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_RuleDefinition] ON [RuleDefinition] ([RuleID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_RulePack] ON [RulePack] ([PackID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_SummaryGroup] ON [SummaryGroup] ([GroupID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_UserProfile] ON [UserProfile] ([UserProfileID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_ProfileGuid] ON [UserProfile] ([ProfileGuid])
CREATE UNIQUE INDEX [PK_ValidationFinding] ON [ValidationFinding] ([FindingID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_ValidationRun] ON [ValidationRun] ([ValidationRunID]) WITH PRIMARY
CREATE UNIQUE INDEX [PK_Workspace] ON [Workspace] ([WorkspaceID]) WITH PRIMARY
CREATE UNIQUE INDEX [UX_WorkspaceGuid] ON [Workspace] ([WorkspaceGuid])
ALTER TABLE [FolderBatchChange] ADD CONSTRAINT [FK_BatchMove] FOREIGN KEY ([AppliedMoveID]) REFERENCES [FolderMove] ([FolderMoveID])
ALTER TABLE [FolderBatchChange] ADD CONSTRAINT [FK_BatchOperation] FOREIGN KEY ([OperationID]) REFERENCES [OperationJournal] ([OperationID])
ALTER TABLE [FolderBatchChange] ADD CONSTRAINT [FK_BatchReport] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [FolderBatchChange] ADD CONSTRAINT [FK_BatchRevision] FOREIGN KEY ([BaseRevisionID]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [RevisionChange] ADD CONSTRAINT [FK_ChangeRevision] FOREIGN KEY ([RevisionID]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [CommandMember] ADD CONSTRAINT [FK_CommandMemberCreator] FOREIGN KEY ([CreatedBy]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [CommandMember] ADD CONSTRAINT [FK_CommandMemberEditor] FOREIGN KEY ([ModifiedBy]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [ReviewComment] ADD CONSTRAINT [FK_CommentAuthor] FOREIGN KEY ([AuthorProfileID]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [ReviewCommentEvent] ADD CONSTRAINT [FK_CommentEventAuthor] FOREIGN KEY ([AuthorProfileID]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [ReviewCommentEvent] ADD CONSTRAINT [FK_CommentEventThread] FOREIGN KEY ([ThreadID]) REFERENCES [ReviewComment] ([CommentID])
ALTER TABLE [ReviewComment] ADD CONSTRAINT [FK_CommentReport] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [ReviewComment] ADD CONSTRAINT [FK_CommentRevision] FOREIGN KEY ([RevisionID]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [ReviewComment] ADD CONSTRAINT [FK_CommentThread] FOREIGN KEY ([ThreadID]) REFERENCES [ReviewComment] ([CommentID])
ALTER TABLE [ReportEvent] ADD CONSTRAINT [FK_EventAuthor] FOREIGN KEY ([AuthorProfileID]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [ReportEvent] ADD CONSTRAINT [FK_EventReport] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [Folder] ADD CONSTRAINT [FK_FolderParent] FOREIGN KEY ([ParentFolderID]) REFERENCES [Folder] ([FolderID])
ALTER TABLE [ImportReceipt] ADD CONSTRAINT [FK_ImportApplied] FOREIGN KEY ([AppliedRevisionID]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [ImportReceipt] ADD CONSTRAINT [FK_ImportAuthor] FOREIGN KEY ([AuthorProfileID]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [ImportReceipt] ADD CONSTRAINT [FK_ImportBefore] FOREIGN KEY ([BeforeRevisionID]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [ImportReceipt] ADD CONSTRAINT [FK_ImportOperation] FOREIGN KEY ([OperationID]) REFERENCES [OperationJournal] ([OperationID])
ALTER TABLE [ImportRawValue] ADD CONSTRAINT [FK_ImportRawReceipt] FOREIGN KEY ([ImportReceiptID]) REFERENCES [ImportReceipt] ([ImportReceiptID])
ALTER TABLE [ImportReceipt] ADD CONSTRAINT [FK_ImportReport] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [LibraryPreference] ADD CONSTRAINT [FK_LibraryProfile] FOREIGN KEY ([AuthorProfileID]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [GroupMember] ADD CONSTRAINT [FK_Major0] FOREIGN KEY ([GroupID]) REFERENCES [SummaryGroup] ([GroupID])
ALTER TABLE [GroupMember] ADD CONSTRAINT [FK_Major1] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [ChangeItem] ADD CONSTRAINT [FK_Major10] FOREIGN KEY ([ExpectedRevision]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [ChangeItem] ADD CONSTRAINT [FK_Major11] FOREIGN KEY ([AppliedRevision]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [RouteTemplateStep] ADD CONSTRAINT [FK_Major12] FOREIGN KEY ([TemplateID]) REFERENCES [RouteTemplate] ([TemplateID])
ALTER TABLE [RouteTemplateStep] ADD CONSTRAINT [FK_Major13] FOREIGN KEY ([AssigneeID]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [RouteRun] ADD CONSTRAINT [FK_Major14] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [RouteRun] ADD CONSTRAINT [FK_Major15] FOREIGN KEY ([RevisionID]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [RouteRun] ADD CONSTRAINT [FK_Major16] FOREIGN KEY ([TemplateID]) REFERENCES [RouteTemplate] ([TemplateID])
ALTER TABLE [RouteRun] ADD CONSTRAINT [FK_Major17] FOREIGN KEY ([GroupID]) REFERENCES [SummaryGroup] ([GroupID])
ALTER TABLE [RouteRun] ADD CONSTRAINT [FK_Major18] FOREIGN KEY ([RequestedBy]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [RouteRunStep] ADD CONSTRAINT [FK_Major19] FOREIGN KEY ([RunID]) REFERENCES [RouteRun] ([RunID])
ALTER TABLE [GroupSnapshot] ADD CONSTRAINT [FK_Major2] FOREIGN KEY ([GroupID]) REFERENCES [SummaryGroup] ([GroupID])
ALTER TABLE [RouteRunStep] ADD CONSTRAINT [FK_Major20] FOREIGN KEY ([AssigneeID]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [ComparisonSet] ADD CONSTRAINT [FK_Major21] FOREIGN KEY ([GroupSnapshotID]) REFERENCES [GroupSnapshot] ([SnapshotID])
ALTER TABLE [ComparisonMember] ADD CONSTRAINT [FK_Major22] FOREIGN KEY ([SetID]) REFERENCES [ComparisonSet] ([SetID])
ALTER TABLE [ComparisonMember] ADD CONSTRAINT [FK_Major23] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [ComparisonMember] ADD CONSTRAINT [FK_Major24] FOREIGN KEY ([RevisionID]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [MeetingProposal] ADD CONSTRAINT [FK_Major25] FOREIGN KEY ([SetID]) REFERENCES [ComparisonSet] ([SetID])
ALTER TABLE [MeetingProposal] ADD CONSTRAINT [FK_Major26] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [MeetingProposal] ADD CONSTRAINT [FK_Major27] FOREIGN KEY ([BaseRevision]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [RuleDefinition] ADD CONSTRAINT [FK_Major28] FOREIGN KEY ([PackID]) REFERENCES [RulePack] ([PackID])
ALTER TABLE [ValidationRun] ADD CONSTRAINT [FK_Major29] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [GroupSnapshot] ADD CONSTRAINT [FK_Major3] FOREIGN KEY ([AttestedBy]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [ValidationRun] ADD CONSTRAINT [FK_Major30] FOREIGN KEY ([RevisionID]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [ValidationRun] ADD CONSTRAINT [FK_Major31] FOREIGN KEY ([PackID]) REFERENCES [RulePack] ([PackID])
ALTER TABLE [ValidationFinding] ADD CONSTRAINT [FK_Major32] FOREIGN KEY ([ValidationRunID]) REFERENCES [ValidationRun] ([ValidationRunID])
ALTER TABLE [ValidationFinding] ADD CONSTRAINT [FK_Major33] FOREIGN KEY ([RuleID]) REFERENCES [RuleDefinition] ([RuleID])
ALTER TABLE [GroupSnapshotMember] ADD CONSTRAINT [FK_Major4] FOREIGN KEY ([SnapshotID]) REFERENCES [GroupSnapshot] ([SnapshotID])
ALTER TABLE [GroupSnapshotMember] ADD CONSTRAINT [FK_Major5] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [GroupSnapshotMember] ADD CONSTRAINT [FK_Major6] FOREIGN KEY ([RevisionID]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [ChangeBatch] ADD CONSTRAINT [FK_Major7] FOREIGN KEY ([AuthorID]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [ChangeItem] ADD CONSTRAINT [FK_Major8] FOREIGN KEY ([BatchID]) REFERENCES [ChangeBatch] ([BatchID])
ALTER TABLE [ChangeItem] ADD CONSTRAINT [FK_Major9] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [MemberReport] ADD CONSTRAINT [FK_MemberReport] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [MemberReport] ADD CONSTRAINT [FK_MemberReportAuthor] FOREIGN KEY ([UpdatedBy]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [FolderMove] ADD CONSTRAINT [FK_MoveAuthor] FOREIGN KEY ([AuthorProfileID]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [FolderMove] ADD CONSTRAINT [FK_MoveNewFolder] FOREIGN KEY ([NewFolderID]) REFERENCES [Folder] ([FolderID])
ALTER TABLE [FolderMove] ADD CONSTRAINT [FK_MoveOldFolder] FOREIGN KEY ([OldFolderID]) REFERENCES [Folder] ([FolderID])
ALTER TABLE [FolderMove] ADD CONSTRAINT [FK_MoveReport] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [ReportNote] ADD CONSTRAINT [FK_NoteAuthor] FOREIGN KEY ([AuthorProfileID]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [ReportNoteDraft] ADD CONSTRAINT [FK_NoteDraftAuthor] FOREIGN KEY ([AuthorProfileID]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [ReportNoteDraft] ADD CONSTRAINT [FK_NoteDraftReport] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [ReportNote] ADD CONSTRAINT [FK_NoteReport] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [OperationJournal] ADD CONSTRAINT [FK_OperationProfile] FOREIGN KEY ([AuthorProfileID]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [PriorEvaluation] ADD CONSTRAINT [FK_PriorCreator] FOREIGN KEY ([CreatedBy]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [PriorEvaluation] ADD CONSTRAINT [FK_PriorEditor] FOREIGN KEY ([ModifiedBy]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [PriorEvaluation] ADD CONSTRAINT [FK_PriorMember] FOREIGN KEY ([MemberID]) REFERENCES [CommandMember] ([MemberID])
ALTER TABLE [PriorEvaluationTrait] ADD CONSTRAINT [FK_PriorTrait] FOREIGN KEY ([PriorID]) REFERENCES [PriorEvaluation] ([PriorID])
ALTER TABLE [RecoveryDraft] ADD CONSTRAINT [FK_RecoveryAuthor] FOREIGN KEY ([AuthorProfileID]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [RecoveryDraft] ADD CONSTRAINT [FK_RecoveryReport] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [RecoveryDraft] ADD CONSTRAINT [FK_RecoveryRevision] FOREIGN KEY ([BaseRevisionID]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [Report] ADD CONSTRAINT [FK_ReportCurrentRevision] FOREIGN KEY ([CurrentRevisionID]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [MemberReport] ADD CONSTRAINT [FK_ReportDirectoryMember] FOREIGN KEY ([MemberID]) REFERENCES [CommandMember] ([MemberID])
ALTER TABLE [Report] ADD CONSTRAINT [FK_ReportFolder] FOREIGN KEY ([FolderID]) REFERENCES [Folder] ([FolderID])
ALTER TABLE [ReportRevision] ADD CONSTRAINT [FK_RevisionAuthor] FOREIGN KEY ([AuthorProfileID]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [ReportRevision] ADD CONSTRAINT [FK_RevisionParent] FOREIGN KEY ([ParentRevisionID]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [ReportRevision] ADD CONSTRAINT [FK_RevisionReport] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [ReportingSeniorRecord] ADD CONSTRAINT [FK_RSAuthor] FOREIGN KEY ([CreatedBy]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [RSCABaseline] ADD CONSTRAINT [FK_RSBaseline] FOREIGN KEY ([SeniorID]) REFERENCES [ReportingSeniorRecord] ([SeniorID])
ALTER TABLE [RSCABaseline] ADD CONSTRAINT [FK_RSBaselineAuthor] FOREIGN KEY ([CreatedBy]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [GroupRSCABaseline] ADD CONSTRAINT [FK_RSGroup] FOREIGN KEY ([GroupID]) REFERENCES [SummaryGroup] ([GroupID])
ALTER TABLE [GroupRSCABaseline] ADD CONSTRAINT [FK_RSGroupBaseline] FOREIGN KEY ([BaselineID]) REFERENCES [RSCABaseline] ([BaselineID])
ALTER TABLE [GroupRSCABaseline] ADD CONSTRAINT [FK_RSGroupSenior] FOREIGN KEY ([SeniorID]) REFERENCES [ReportingSeniorRecord] ([SeniorID])
ALTER TABLE [ReportingSeniorValue] ADD CONSTRAINT [FK_RSValue] FOREIGN KEY ([SeniorID]) REFERENCES [ReportingSeniorRecord] ([SeniorID])
ALTER TABLE [NarrativeSnippet] ADD CONSTRAINT [FK_SnippetProfile] FOREIGN KEY ([CreatedBy]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [ReviewTask] ADD CONSTRAINT [FK_TaskAssignee] FOREIGN KEY ([AssignedTo]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [ReviewTaskEvent] ADD CONSTRAINT [FK_TaskEventAuthor] FOREIGN KEY ([AuthorProfileID]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [ReviewTaskEvent] ADD CONSTRAINT [FK_TaskEventTask] FOREIGN KEY ([TaskID]) REFERENCES [ReviewTask] ([TaskID])
ALTER TABLE [ReviewTask] ADD CONSTRAINT [FK_TaskReport] FOREIGN KEY ([ReportID]) REFERENCES [Report] ([ReportID])
ALTER TABLE [ReviewTask] ADD CONSTRAINT [FK_TaskRequester] FOREIGN KEY ([RequestedBy]) REFERENCES [UserProfile] ([UserProfileID])
ALTER TABLE [ReviewTask] ADD CONSTRAINT [FK_TaskRevision] FOREIGN KEY ([RevisionID]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [RevisionValue] ADD CONSTRAINT [FK_ValueRevision] FOREIGN KEY ([RevisionID]) REFERENCES [ReportRevision] ([RevisionID])
ALTER TABLE [Workspace] ADD CONSTRAINT [FK_WorkspaceProfile] FOREIGN KEY ([ActiveProfileID]) REFERENCES [UserProfile] ([UserProfileID])
