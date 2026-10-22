# Counseling, readiness and billet details

> Historical service/verification reference. Current procedures are in the
> [user guide](../docs/Swift_NAVFIT_User_Guide.md); current architecture and
> decisions are in [PROJECT_STATE.md](../PROJECT_STATE.md). Retained field/source
> details remain evidence, but old screens and test commands may be superseded.

Version 0.17.0, schema 16 adds three features for EVAL, CHIEFEVAL and FITREP. Open **More report details** in the editor.

| Feature | Entry fields | Storage limits |
| --- | --- | --- |
| Counseling | Counseling date, counselor name | ISO date YYYY-MM-DD or blank; name 20 characters |
| Physical readiness | Summary, detail, date | Summary 4 characters; detail 15; ISO date or blank |
| Billet | Billet subcategory | 10 characters |

Dates must be real calendar dates. Text fields reject line breaks and tabs. Incomplete entries remain saveable. Readiness detail and billet dropdowns provide suggestions and allow custom entries within the limits. They do not enforce Navy codes or eligibility rules.

Apply transfers all six pending values to the editor and creates a recovery checkpoint. Save draft commits a revision. Cancel preserves the editor. Unapplied panel edits are not checkpointed. Archived reports cannot apply changes.

All 49 draft fields participate in history, duplication, comparison, restore and recovery. Saved changes stale earlier review assignments. Older snapshots show Not recorded for these fields and restore them to null. Schema-16 snapshots require all six rows. Null and empty values remain distinct.

## Reference and interchange scope

Suggestions were extracted read-only from NAVFIT98A 2.2.0.332's Data/NF98A_Dte.accdb, SHA256 ABB8CA070891CB7BF76D3CB93599F9BBEC40FAAB7BFAF0124B594E2A3E1577A8. The source JSON embeds ten readiness and 37 billet suggestions during the build. No runtime source database is required.

Observed source fields are DateCounseled (text 8), Counseler (text 20), PhysicalReadiness (text 4), PhysicalReadiness2 (text 15), PhysicalReadinessDt (DateTime) and BilletSubcat (text 10). Swift uses ISO text dates internally; conversion to the external date formats remains pending. Separate counselor name parts are also pending.

The checklist checks presence only. Counseling frequency, readiness interpretation, waiver eligibility and billet applicability rules remain pending, along with full NAVFIT import/export.

## Verification

152 focused behavioral assertions and 49 reference checks passed, including all three report types, dates and limits, rollback, stale reviews, legacy snapshots, panel actions and recovery in a new Access process. The complete release passed 1158 behavioral assertions and 218 reference checks. Schema 15-to-16, 16-to-16 and 5-to-16 migrations preserved original rows and fields, source hashes and reopening. Runtime/x86 and comprehensive visual testing remain pending.
