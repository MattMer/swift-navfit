# Reporting-senior identifiers, contact details and raters

Swift NAVFIT 0.19.0, schema 18 adds three features for EVAL, CHIEFEVAL and FITREP. Open **Senior contact and raters** in the editor. The panel has three labeled sections and scrolls to the Apply and Cancel controls.

| Feature | Entry fields and limits | Observed NAVFIT source |
| --- | --- | --- |
| Senior identifiers | SSN up to 9 digits; DoD ID up to 10 digits | RSSSN text 9; RSDoDID text 10 |
| Senior contact details | Address lines 1/2: 30 each; city: 15; state/region: 2; ZIP/postal code: 9; phone: 10; DSN: 7 | RSAddress1/2, RSCity, RSState, RSZipCd, RSPhone, RSDSN: text with matching capacities |
| Rater and senior rater | Two names: 28 each; two dates: yyyy-mm-dd or blank | Rater/SeniorRater: text 28; RaterDate/SeniorRaterDate: DateTime |

Identifiers are text, retain leading zeros and permit incomplete digit entries. They remain separate from member identifiers, the existing reporting-senior name and local contributor profiles. Identity authenticity and authority are not assessed.

Contact fields accept single-line text within the observed capacities. Postal, phone, DSN and state formats are not inferred or validated as official codes. The source's combined RSAddress memo remains unmapped; structured address entries do not automatically compose it.

Rater dates must be real calendar dates, entered explicitly. Names and dates do not create signatures or official approval. No automatic dates, chronological rule, form applicability or reporting-authority inference is applied. Date conversion to NAVFIT DateTime remains pending.

Apply transfers all 13 pending values and checkpoints them. Save draft commits them with all 68 fields. Cancel preserves the editor. Unapplied panel edits are not checkpointed. Archived reports cannot apply changes. Nontext entries, line breaks/tabs, invalid dates/identifiers and overlength text reject the whole save without truncation. Blanks remain saveable.

All fields participate in history, comparison, duplication, restore and raw pending recovery. Saved changes stale earlier review assignments. Older snapshots show Not recorded and restore these fields to null; schema-18 snapshots require all 13 rows. Null and empty remain distinct.

The supplied NAVFIT98A/Data/NF98A_empty.accdb was inspected read-only for names, types and capacities. Its SHA256 remains 43348B36198AF6D8E3D34A5F40090BB06AE1D0ADC01295066BAA8799A9078252. Source structure is not a Navy policy rule. The local checklist checks presence only; full NAVFIT import/export, identity/authority rules, signatures and printed layout remain pending.

Use scripts/upgrade-workspace.ps1 with the blank 0.19.0 template and a fresh destination for populated workspaces. The script preserves the source and compares copied rows and fields exactly.

## Verification

364 focused assertions and three new layout checks cover all fields, all report types, leading zeros, partial identifiers, boundary lengths, dates, null/empty values, failed-write rollback, stale reviews, older/corrupt snapshots and recovery in a new Access process. The complete release verification comprises 1720 behavioral assertions, 218 reference checks and six layout checks, plus exact import/provenance checks of 2472 values. Migration paths are schema 17-to-18, 18-to-18 and 5-to-18. Runtime/x86 and comprehensive visual verification remain pending.
