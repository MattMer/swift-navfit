# Calculation context and saved-draft checklist

Implemented in the r20 native Access preview; schema remains 21.

## Calculation context

- Report > Review > Calculation details is keyboard reachable. Double-clicking ITA or Summary Group Average also opens the details.
- Current editor traits and latest saved revision are shown separately. Explicit selected group context supplies the saved group calculation; pending comparison edits are excluded.
- Arithmetic displays trait sums, graded counts, whole-report NOB exclusions, individual NOB/blank exclusions, archived report exclusions and unavailable reasons. Existing arithmetic is unchanged.
- Summary group > Members > Reporting senior and PSG RSCA shows the saved group formula, manually entered historical baseline paygrade/date/count/source, exact versus rounded-RSCA estimate, and partial/unavailable reasons. Initial load and selection-save actually refresh the display.
- Zero historical traits do not imply an estimated historical sum. Baseline data remains manual; provenance text does not verify the source.

## Saved-draft checklist

- Report > Review > Draft checklist opens the saved checklist rather than duplicating Review routes.
- Findings use form-specific official captions and block numbers. Legacy DoDID fields and inapplicable fields are omitted; retained data is untouched.
- Select a field finding, then Open report section (or double-click). The matching section opens without saving a revision or replacing pending editor values.
- Nonfield items explain their destination and disable the field-navigation action. Archived/local-approved reports remain read-only with their existing explanation.
- Incomplete local inputs, invalid date order and limited source-linked Navy checks remain distinct. The checklist is not certification and does not silently extend any rule pack's effective dates.
- Guidance points to current editor sections, Overview recovery and Reports import/export workflows; obsolete detail-dialog/scrolling instructions were removed.

## Verification limits

Native COM harnesses check calculations, official targets, pending-value preservation, initial calculator feedback and keyboard-reachable action identities. Existing responsive/focus, editor, lifecycle, bulk/comparison, startup, tools and administration checks cover regression behavior. Human screen-reader testing and physical display/scaling checks remain outstanding; this is a review preview, not a release certification.
