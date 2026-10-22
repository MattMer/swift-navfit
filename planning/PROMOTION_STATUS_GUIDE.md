# Promotion status

> Historical service/verification reference. Current procedures are in the
> [user guide](../docs/Swift_NAVFIT_User_Guide.md); current architecture and
> decisions are in [PROJECT_STATE.md](../PROJECT_STATE.md). Retained field/source
> details remain evidence, but old screens and test commands may be superseded.

Swift NAVFIT 0.13.0 adds manual promotion status near the top of the editor for EVAL, CHIEFEVAL and FITREP. Choose REGULAR, FROCKED, SELECTED or SPOT, or leave it blank. These four labels were read from PromotionStatus in the supplied NAVFIT98A 2.2.0.332 lookup database. Local storage is text with an eight-character capacity matching the inspected Reports.PromotionStatus field.

Promotion status is separate from promotion recommendation, observation status, grade/rate and trait marks. None is inferred or overwritten when another changes. Blank drafts remain saveable. Unsupported selections, case/spelling variants and nontext codes reject the save without truncation or partial writes.

The field participates in snapshots, comparison, restore, duplication, pending-editor recovery and stale-review checks. Null and empty text remain distinct in history. A restored revision before schema 12 clears unrecorded status to null; comparison displays Not recorded. New snapshots require the status row, including explicit nulls. Recovery preserves raw invalid input for deliberate correction without auto-saving it. The checklist adds a local presence prompt.

This is entry and reference vocabulary, not Navy validation. Eligibility, grade/status consistency, occasion requirements and effective dates remain unassessed. NAVFIT import remains narrative-only; promotion status mapping/export and target-app round-trip proof are pending. Printed layout remains deferred.

Use scripts/upgrade-workspace.ps1 with the new blank template and a fresh destination for populated workspaces. Verification: 66 new assertions and 539 prior behavioral assertions passed, plus 169 catalogue checks and all 2472 retained import values. Schema 11-to-12, 12-to-12 and 5-to-12 migrations preserve original fields, source hashes and reopen. All 20 distribution user-data tables are empty. Runtime/x86 and comprehensive visual checks remain pending.
