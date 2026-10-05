# NAVFIT group export — r28, 5 October 2026

## Implemented

Summary group > Export group to NAVFIT98A opens the native exchange workspace
with that group selected. Reports > Import reports > NAVFIT export also offers
individual saved-report export and group export. Choose a verified blank native
template and a new destination. Preview the group, then export.

The preview lists every saved member, promotion totals and RSCA handling. Export
creates exactly one native folder with exactly those members, using NAVFIT's
`a 1` report parent convention. Promotion totals are recalculated on every report
and the folder; AutoSummary is enabled. NOB recommendations are listed separately
and excluded from the five promotion totals. Native-only report values from each
member's own provenance are preserved. Folder defaults and summary-letter contact
profiles are not inferred or merged from unrelated source files.

Native report IDs are reassigned deterministically within the package. A frozen
`.swift-baseline.accdb` companion records original Swift identities, revisions,
source receipts, counseling centuries and all 78 saved values per member, plus
the exported native records. Keep it locally. Send/open only the native ACCDB.
The operation journal seals the companion hash and records prepared/completed or
failed outcomes. Existing destinations are rejected; failures do not report
completion and leave diagnostic pending files. No export creates report revisions.

Group, revision and RSCA changes invalidate the preview. Archive, recovery and
editor guards apply to every member. Type, reporting senior, period end, report
type flags and component selections must agree. Group names must fit the native
50-character folder limit and promotion counts the three-character native fields.
These compatibility checks are not complete Navy eligibility certification.

By default native RSCA is preserved per imported report and remains unset for new
drafts. "Write exact PSG RSCA" deliberately writes the group's calculated value,
rounded to two decimal places, to every member's native RSCA field. It requires
an exact historical trait sum, confirmed exclusion of this group, matching senior
SSNs and complete graded traits. Estimated or partial calculations are rejected.

Single-report export/three-way returns remain available. A new-draft export now
uses the correct native folder-parent encoding. Group-return reconciliation is
still pending; the single-report return screen rejects group companions explicitly.

## Verification and limits

- 867 group assertions: all three report kinds, imported/new-draft members,
  75 decoded values per member, 120 unchanged native columns per imported member,
  roster/folder/counts, 78-value baselines, overwrite/stale/mismatch rejection,
  exact PSG export and membership-error cleanup.
- Existing service suites: 497 adapter, 410 single-report exchange and 98 guards.
- Native Access exchange UI: 35 assertions, including group preview,
  invalidation, rejected unpreviewed export and successful reviewed export.
- The membership error handler no longer closes an already-closed recordset and
  masks a kind/paygrade mismatch with DAO error 3420.
- Final blank/demo compile and match all 57 source modules. The seeded demo
  passes 36 assertions and normal/enabled-content repeat startup checks. A clean
  source rebuild passed; the final recordset-cleanup-only refresh was compiled
  and reverified separately. Artifact hashes are in
  [the r28 manifest](NAVFIT_R28_VERIFICATION.json).

Actual group opening/printing in NAVFIT98A is not yet counted as verified. The
computer-use helper returned stale file-dialog focus during the attempted native
check; it was canceled. Database and Access tests do not establish printed fit,
native policy validation, summary-letter completeness or accessibility/scaling.

See [print-engine inspection](NAVFIT_PRINT_ENGINE_FINDINGS.md). Printing through
NAVFIT98A uses its actual engine. Direct printing from Swift is not implemented.
No physical print job was submitted. Earlier r27 per-report actual native round
trips remain the individual-report evidence.
