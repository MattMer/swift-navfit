# C01/C02: folder protection and reviewed directory merging

21 October 2026. Application 0.27.0-dev, schema 27, r127.

## C01 decision and inventory

The user explicitly declined an Access password/encryption design: “No, protection is covered by file folder permissions.” Windows file/folder controls remain the storage boundary. There is no new password, embedded secret, DPAPI dependency, network service or protection setup. SSNs remain readable inside Access; masking does not provide storage encryption. Contributor names identify attribution, not authentication.

The implementation preserves these retained locations rather than scrubbing historical data:

| Location | Purpose and handling |
| --- | --- |
| `Report.MemberSSN`, `Report.ReportingSeniorSSN` | Current official report fields; existing authorized edit/import/export paths retained. |
| `CommandMember.MemberSSN` | Directory identity signal; stable `MemberReport.MemberID` links survive subsequent report edits. |
| `RecoveryDraft` identifier fields | Pending editor recovery remains intact. |
| `RevisionValue` and `RevisionChange` | Original snapshots and old/new field values remain intact. |
| `ImportRawValue`, retained import/source XML | Original source/provenance can contain identifiers; source values remain preserved. |
| `ReportingSeniorValue` | Reusable reporting-senior SSN field values remain intact. |
| Exports, backups and workspace copies | Can retain the same identifiers; no app-specific encryption or remote protection service added. Folder/file permissions apply independently to each destination. |

This is acceptance of the existing protection boundary, not a claim that plaintext identifiers were removed. README and C01 backlog now say so explicitly.

## C02 behavior

See [duplicate-review procedure](../docs/Duplicate_Sailor_Review.md). The new modal review form is launched from the existing directory, preserving its navigation rail. It offers normalized-surname candidates and deliberate changed-name selection, masked SSN/DoD ID summaries, and affected report/prior lists. Apply requires a fresh SHA-256 state fingerprint and a review reason. The fingerprint covers both profiles, links/current revisions, prior fields/traits and linked report identifiers. No plaintext identifier is copied into the new mapping tables or generated operation messages.

`DirectoryMerge` retains original/surviving IDs, operation/actor, review reason, state, original profile modification metadata and the post-merge fingerprint. `DirectoryMergeReport` and `DirectoryMergePrior` retain original link metadata keyed by the merge and existing report/prior IDs. New relationships use descriptive names. Existing `OperationJournal` supplies the author relationship; redundant new UserProfile foreign keys were avoided because the existing Access table is at its index limit. Original actor IDs in snapshots remain retained metadata.

Within one DAO transaction, Apply archives the original profile, moves report/prior links, advances affected version counters, records mappings and completes the operation. It preserves all report fields, report revisions, notes/reviews, prior fields and trait values. Provisional/Needs review states are not upgraded. Undo checks the post-merge fingerprint before restoring original links and metadata, keeps the audit and advances versions. Open editors block both operations. Later affected changes block undo. A retained active merge blocks nested merging involving either member.

Identity checks include both profiles and all linked reports, including archived or unconfirmed reports. Conflicting nonblank SSNs/DoD IDs block the merge. If only the source profile has an SSN, users must choose it as survivor; profile fields are never silently copied or cleared to evade uniqueness constraints. The original archived profile remains queryable through the retained mapping. Other unreferenced field differences are retained in that profile rather than reconciled automatically.

Schema 27 is additive. Production table-level transfer includes the three new tables after their referenced members, reports, priors and operations. Schemas 1–26 create empty merge history; schema 27 preserves committed and undone mappings exactly. Source databases remain read-only and their code is never executed.

## Build and verification

The release uses `build-directory-merge-update.ps1` on the verified schema-26 r126 blank, imports current changed modules, applies the additive schema/form stage, compiles and reopens, and checks all 99 static modules and event bindings across 66 forms. It retains the known full-generation limitation documented for r126; this is not a claim that the unrelated full fresh-generation failure is fixed.

Native merge checks cover conflicting identities, changed names, missing identifiers, stale profile/link previews, open-editor guards, injected failure rollback without audit leakage, report/prior/trait preservation, retained committed/undone mappings, later revision/prior/link undo rejection, exact restored metadata and monotonic versions. Existing directory/history, batch creation and native directory UI regressions are also exercised. Final release gates passed: 47 merge, 72 directory/history, 20 batch, 21 directory UI, 57 context-help, 27 cleanup, 48 smoke, 39 Demo and 35 report-tools checks; all 27 retained-schema transfers and rejection/rollback guards; populated schema-20/current transfers; and exact committed/undone merge transfer with native undo and next AutoNumber. Oldest retained fixtures are blank distributions. Final release receipts, actual fixture paths/counts and Drive name/size/parent readbacks are recorded in `DIRECTORY_MERGE_R127_VERIFICATION.json`.

Blank contains no user rows; Demo contains 13 synthetic reports, five groups and 13 directory members. No scratch merge rows are distributed. Both files are under `output` and uploaded to the designated Drive folder. The original r126 Blank and populated r116 workspace match recorded hashes. The local r126 Demo differs from its older published receipt; its current contents passed exact read-only transfer and before/after hash checks and were preserved as found. Its prior hash difference was not attributed to an unverified cause. Access is closed.

Physical display/accessibility acceptance remains deferred. No official print-layout claim is made. C03 startup contributor selection and C04 validation-policy applicability are the next critical items.
