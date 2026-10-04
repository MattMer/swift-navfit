# Report occasions, report categories and duty descriptions

Swift NAVFIT 0.15.0 adds three panels at the top of the draft editor for EVAL, CHIEFEVAL and FITREP.

## Report occasions

Periodic, detachment of individual, frocking and special each accept Yes, No or blank. They are independent selections; multiple Yes values and all-No values remain saveable pending official combination/applicability rules. Promotion status does not set the frocking occasion. The inspected NAVFIT source columns are Periodic, DetInd, Frocking and Special (Boolean); Swift stores explicit text selections so blank remains distinguishable from No. External Boolean mapping is not implemented.

## Report categories

Regular, concurrent and operational commander each accept Yes, No or blank. Category selections remain separate from occasions, whole-report observation status, promotion recommendation and organizational folders. No exclusivity, eligibility or group rules are inferred. The inspected source columns are Regular, Concurrent and OpsCdr (Boolean); external mapping remains pending.

## Duty descriptions

Primary duty supports 14 characters, single-line without tabs, matching the inspected Reports.PrimaryDuty capacity. Duty description is multiline long text, corresponding to the inspected Duties memo field. It stays separate from local draft notes and official comments. Long descriptions are retained without a 255-character cut-off. No printed fit, block-number mapping or policy-length guarantee is claimed.

## Editing and history

Open a panel, edit and Apply pending values and return. Apply transfers only that panel's values to the editor and checkpoints them; it does not commit a revision or change another panel. Cancel discards panel edits and preserves applied editor values. Unapplied panel edits are not checkpointed. Save draft commits all report fields atomically. Unsupported flags, nontext entries and invalid primary duty reject the entire save; raw pending invalid text survives recovery for correction. Archived drafts permit viewing, but cannot apply edits.

All nine new fields participate in snapshots, comparison, duplicate/restore, checkpoint recovery and stale-review checks. Old snapshots display Not recorded; restoring them clears unrecorded fields to null. New schema-14 snapshots require all nine rows, including explicit nulls. The checklist adds local presence prompts; a No selection counts as entered, not missing. These prompts are not official requirements. NAVFIT import remains narrative-only; full mapping/export and Navy validation remain pending.

Use scripts/upgrade-workspace.ps1 with the new blank template and a fresh destination for populated workspaces. Verification: 195 new assertions plus 686 prior behavioral assertions, 169 catalogue checks and all 2472 retained import values passed. Schema 13-to-14, 14-to-14 and 5-to-14 migrations compare every original row/field, preserve source hashes and reopen. All 20 distribution user-data tables are empty. Runtime/x86 and comprehensive visual verification remain pending; printed layout remains deferred.
