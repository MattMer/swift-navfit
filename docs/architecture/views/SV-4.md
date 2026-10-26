# SV-4 Systems Functionality

Functions F1–F11 support the OV-5b activities A1–A10. The modules are the VBA
source under `src/`, and the screens are forms under `src/forms/`.

| Function | Main modules | Main screens |
| --- | --- | --- |
| F1 Workspace and contributors | `modSwift`, `modWorkspace`, `modContributorStartup`, `modAdministration` | Setup, contributor start, Settings & Help |
| F2 Command directory | `modCommandDirectory`, `modUXDirectory`, `modDirectoryBatch`, `modDirectoryMerge`, `modRUADImport`, `modRUADProfiles` | Command member directory, batch creation, duplicate review, RUAD import |
| F3 Counseling | `modCounseling` | Member history, counseling editor, My Work → Counseling |
| F4 Reports and groups | `modSummaryGroups`, `modUXGroupOrganization`, `modUXGroupPages`, `modGroupPolicy`, `modBulkFolders`, `modLibraryViews` | Reports, Summary groups, bulk defaults |
| F5 Report editor | `modUXSurface`, `modReportPages`, `modReportPageMap`, `modTraits`, `modRecovery`, `modReportState` | Report editor (fixed NAVPERS pages) |
| F6 Validation | `modValidation`, `modLiveValidation`, `modValidationExpansion`, `modValidationReferences`, `modSavedReadiness` | Validation checks, validation pop-out |
| F7 Review and messages | `modRouting`, `modRouteTemplates`, `modReviewComments`, `modReportMessages`, `modUXMyWork`, `modUXReview` | My Work, Report tools |
| F8 Group calculations | `modUXTraitGrid`, `modRSAccounting`, `modRSCompletion`, `modReportingSenior`, `modUXSeniorTable` | Trait table, RSCA table, reporting seniors |
| F9 Interchange | `modNAVFITAdapter`, `modNAVFITExchange`, `modNAVFITMapping`, `modNAVFITUpgrade`, `modPDFOutput`, `modPDFImport`, `modPDFMapping`, `modPDFLocal` | NAVFIT98A exchange, PDF output, PDF import |
| F10 Lifecycle and history | `modSignoff`, `modRevisionCompare`, `modAudit`, `modProvenance`, `modOperations` | Mark signed, History, revision compare, provenance |
| F11 Continuity | `modBackupUI`, `modMigration`, `modWorkspaceImport`, `modFileHash` | Backup and restore, workspace import |
| Shared UI shell | `modUX`, `modUXViewport`, `modUXVisual`, `modUXHelp`, `modTableSort` | Navigation rail, layout, color palette, contextual help, sortable lists |
