# Assignment recommendations, retention and member statement

> Historical service/verification reference. Current procedures are in the
> [user guide](../docs/Swift_NAVFIT_User_Guide.md); current architecture and
> decisions are in [PROJECT_STATE.md](../PROJECT_STATE.md). Retained field/source
> details remain evidence, but old screens and test commands may be superseded.

Swift NAVFIT 0.18.0, schema 17 adds three features for EVAL, CHIEFEVAL and FITREP. Open **Recommendations and statement** in the editor.

| Feature | Swift fields | Observed NAVFIT source |
| --- | --- | --- |
| Assignment recommendations | AssignmentRecommendationA/B: 20-character single-line text | RecommendA/B: text 20 |
| Retention selections | RetentionYes/No: Yes, No or blank | RetentionYes/No: Boolean |
| Member statement selections | StatementYes/No: Yes, No or blank | StatementYes/No: Boolean |

Assignment recommendations are separate from promotion recommendation. No rank, duty or promotion selection fills them automatically.

For each source checkbox, Yes means marked, No means unmarked and blank means unentered. Each box has its own entry; combinations remain saveable for drafting. No exclusivity or form-applicability rule is enforced. These are draft selections, not signatures, evidence of member acknowledgment or official approval. They are separate from local review acceptance. Swift preserves unentered versus explicit unmarked values; conversion to NAVFIT Boolean fields remains pending.

Apply transfers all six pending values and checkpoints them. Save draft commits them with all 55 fields. Cancel preserves the editor. Unapplied panel edits are not checkpointed. Archived reports cannot apply edits. Long or multiline recommendations, nontext values and unsupported flags reject the whole save without truncation. Incomplete entries remain saveable.

History, duplication, comparison, restore and recovery include every new field; saved changes stale earlier review assignments. Older snapshots show Not recorded and restore these fields to null. Schema-17 snapshots require all six rows. Null and empty remain distinct, including in raw pending-edit recovery.

The source field names, types and capacities were inspected read-only in the supplied NAVFIT98A/Data/NF98A_empty.accdb. This verifies storage structure only. Recommendation eligibility, statement handling requirements, retention rules, signatures, full import/export and printed layout remain pending. The local checklist checks presence only.

Use scripts/upgrade-workspace.ps1 with the blank 0.18.0 template and a fresh destination to migrate populated workspaces. The source is preserved and copied rows/fields are compared exactly.

## Verification

198 focused behavioral assertions and three layout checks passed, with pending recovery in a new Access process. The complete release passed 1356 behavioral assertions and 218 reference checks, plus exact narrative-import/provenance checks of 2472 values. Schema 16-to-17, 17-to-17 and 5-to-17 migrations preserved original data and source hashes and reopened. All 20 user-data tables in the distribution are empty. Runtime/x86 and comprehensive visual verification remain pending.
