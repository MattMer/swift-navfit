# Observation status and promotion recommendation

> Historical service/verification reference. Current procedures are in the
> [user guide](../docs/Swift_NAVFIT_User_Guide.md); current architecture and
> decisions are in [PROJECT_STATE.md](../PROJECT_STATE.md). Retained field/source
> details remain evidence, but old screens and test commands may be superseded.

Swift NAVFIT 0.12.0 adds two manual selections near the top of the editor for all three form types. Whole-report observation status accepts Observed, NOB or blank. Promotion recommendation accepts NOB, Significant Problems, Progressing, Promotable, Must Promote, Early Promote or blank. The six recommendation labels were read from PromotionRecom in the supplied NAVFIT98A 2.2.0.332 lookup database; local storage uses text rather than external numeric codes.

These selections are independent of individual trait NOB marks and the arithmetic average. No value is inferred, and a trait change never overwrites the selections. Blank and null remain distinct in saved history. Incomplete drafts remain saveable; unsupported codes, spelling/case variants and nontext inputs reject the entire save. Potential policy conflicts remain saveable pending the validation-rule implementation.

The two fields participate in snapshots, comparison, duplicate, restore, checkpoints and stale-review checks. Restoring a revision before schema 11 clears unrecorded selections to null. Comparison displays Not recorded. New schema-11 snapshots require both rows, including explicit nulls. Raw invalid pending input survives checkpoints and recovery without being committed.

The checklist adds presence prompts. These are local prompts, not official requirements. Whole-report eligibility, recommendation restrictions, quota limits, trait/observation consistency and Navy policy are not assessed. NAVFIT import remains narrative-only; external recommendation/observation codes and export round trips remain unverified. No block-number mapping is claimed.

Existing populated workspaces require scripts/upgrade-workspace.ps1 with the new blank template and a fresh destination. Schema 11 retains original revisions without manufacturing new snapshot fields.

Verification: 70 assessment assertions plus the existing 469 behavioral assertions, 169 catalogue checks and all 2472 retained import values passed. Populated schema 10-to-11, 11-to-11 and 5-to-11 migrations compare every original row/field, preserve source hashes and reopen. Old reports/checkpoints receive null assessment fields. All 20 user-data tables in the distribution are empty. Runtime/x86 and comprehensive visual verification remain pending.
