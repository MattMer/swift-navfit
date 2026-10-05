# Structured member, reporting-senior and counselor names

> Current native navigation: use the [workspace guide](UX_WORKSPACE_GUIDE.md) and
> [current state](CURRENT_STATE.md). This guide retains service details and
> earlier verification history; legacy screen instructions may differ.

Swift NAVFIT 0.20.0, schema 19 adds three features for EVAL, CHIEFEVAL and FITREP. Open **Structured names** in the editor. The panel groups member, reporting-senior and counselor parts separately.

| Feature | Entry fields and capacities | Observed NAVFIT fields |
| --- | --- | --- |
| Member name parts | First name 27; middle initial(s) 3; last name 27; suffix 10 | FirstName, MI, LastName, Suffix |
| Reporting-senior name parts | Last name 18; first initial 1; middle initial(s) 3 | RSLastName, RSFI, RSMI |
| Counselor name parts | Last name 20; first initial 1; middle initial(s) 3 | CounselerLN, CounselerFI, CounselerMI |

All source fields are text with the listed capacities. Values accept single-line text within those limits, including punctuation and accented characters. Case is preserved. No alphabet-only rule, uppercase conversion, trimming, name splitting or composition is applied. Initials are entered explicitly.

These entries remain separate from Official full name, Reporting senior name, Counselor name and the local member display name. Editing parts leaves those existing entries intact; editing a full name does not populate the parts. This preserves deliberate entries and avoids guessing how a name should be represented. Identity and Navy name-format validation remain pending.

Apply transfers all ten pending values and checkpoints them. Save draft commits them with all 78 fields. Cancel preserves the editor. Unapplied panel edits are not checkpointed. Archived reports cannot apply changes. Overlength, nontext or multiline/tab entries reject the whole save without truncation. Blanks remain saveable.

All parts participate in history, comparison, duplication, restore and raw pending recovery. Saved changes stale earlier review assignments; case changes are visible in history. Older snapshots show Not recorded and restore these fields to null. Schema-19 snapshots require all ten rows. Null and empty remain distinct.

The source structure was inspected read-only in NAVFIT98A/Data/NF98A_empty.accdb. Source SHA256 remains 43348B36198AF6D8E3D34A5F40090BB06AE1D0ADC01295066BAA8799A9078252. This establishes field structure only. The checklist assesses presence, not official name completeness or validity. Full NAVFIT interchange and printed layout remain pending.

Use scripts/upgrade-workspace.ps1 with the blank 0.20.0 template and a fresh destination for populated workspaces. The script preserves the source and compares copied rows and fields exactly; no names are automatically derived during migration.

## Verification

281 focused assertions cover all ten fields and all report types, preservation of full names, punctuation, accents, case changes, boundary lengths, null/empty entries, failed-write rollback, stale reviews, older/corrupt snapshots and recovery in a new Access process. The release verification comprises 2001 behavioral assertions, 218 reference checks and nine layout checks, plus all 2472 retained import values. Migration paths are schema 18-to-19, 19-to-19 and 5-to-19. Runtime/x86 and comprehensive visual verification remain pending.
