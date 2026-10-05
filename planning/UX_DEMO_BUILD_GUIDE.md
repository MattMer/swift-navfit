# Native UX demo previews

For future user-review builds, run:

```powershell
powershell.exe -NoProfile -File scripts/build-preview.ps1
```

The default creates a uniquely named `output/ux/Swift_NAVFIT_UI_Demo_*.accdb`.
It builds a blank source template, copies it, then seeds only the copy. An
explicit `-TemplatePath` can reuse a verified blank template. `-OutputPath`
can choose a different new destination. Existing destinations are never
overwritten, and seeding refuses a populated database. The seed module is
removed from the finished demo; users receive normal app services and data.
The builder executes shared-block/default/comparison, responsive identity,
native surface/workspace and workflow checks before seeding and
the populated-demo journey checks afterward. A verification failure stops the
build from reporting success.

Open the demo in desktop Access. If Access blocks active content, startup now
opens an instruction screen instead of running blocked VBA. Click **Enable
Content**, then close and reopen the file. Alternatively, add only the preview
folder in Access Trust Center; the build scripts do not change Office security
settings. `test-preview-launch.ps1` checks both the normal-security startup
path and enabled-content startup on disposable copies.

Setup is already completed under
`SWIFT NAVFIT DEMO - SYNTHETIC DATA`, with `Alex Morgan (demo)` active.
Browser prototype changes are separate and are not imported.

## Try these journeys

1. Overview: inspect active reports, group/work counts, recent reports and the
   recovery checkpoint. Double-click a recent report to open it.
2. Reports: search names; filter report type, folder and local review status.
   Clear filters to restore the active report list. Use New report to choose
   E1-E6 EVAL, Chief EVAL or FITREP.
3. Jamie Taylor: edit complete official sections; pending values survive
   section changes. Enter an incomplete SSN and select Save draft to inspect
   the inline error and clickable summary. Correct it to save a revision.
4. Summary groups: inspect the Reports count, then open the E5 group; inspect group averages and reporting
   senior/PSG information. Comparison captures its baseline automatically.
   Select a member, stage changed marks, review proposals and Apply.
5. Jordan Reed: local approval locks editing. Review actions operate on saved
   revisions; local approval is not an official signature.
6. My Work: Casey Morgan has current review work; Avery Patel has intentionally
   stale review work. Contributors and assigning review have separate views.
7. Reporting seniors: Resources contains the fictional senior and E5/E7/O3
   historical baselines. Exact sums/counts are synthetic and exclude the group.
8. Recovery: Jamie Taylor has pending local notes that are not in the saved
   revision. Recover or discard the checkpoint to explore that workflow.
9. Reports: switch Show to Archived for Quinn Park; Sam Rivera demonstrates
   whole-report NOB. Drew Blake is an incomplete draft to fill in yourself.

The rebuild preserves all 78 report fields, schema 21 and existing transaction,
revision and recovery services. Native section fields are an unbound scrollable
surface bridged to the original editor buffer. The report editor fills the
window and reflows content-sized report block rows; related checkboxes share
compact rows with adjacent labels, and save actions stay near the bottom. Active text is preserved when resizing. Official captions vary by report
type. Name parts remain independent from full-name fields; legacy DoDID values
are retained in storage but hidden from normal editing.

This is a development preview, not a finished or published redesign. Remaining
work includes full current-template label review, responsive navigation/collapse,
native keyboard and screen-reader acceptance, scaling at 100/125/150 percent,
remaining review/import/bulk screen polish and native close-window dirty guards.
Automated behavior and bounds checks do not prove visual equivalence.

Use `scripts/build.ps1` for blank templates and release verification. Migrate
existing workspaces only into a blank template, never into a seeded demo.

The preview builder also runs `test-ux-responsive-identity.ps1` before seeding.
This gate covers compact report navigation, active-field/caret preservation,
shared viewport geometry and official-name identity. Native DPI, visual and
screen-reader acceptance remain separate manual checks.

## r20 calculation and checklist journeys

Open a report, choose Review, then Calculation details or Draft checklist.
In the checklist select a field and choose Open report section. For PSG RSCA,
open a summary group, Members, Reporting senior and PSG RSCA. The scrolling
context field shows counts, exclusions and manually entered baseline provenance.

## Current help and import journeys

Settings & Help > Workspace guide and quick start opens the native guide.
Reports > Import reports opens the shared source chooser. Preview full mapping
values before Apply; Results opens the created report. NAVFIT mapping remains
narrative-only and PDF mapping remains limited to supported fields. Source
inspection and existing-draft reconciliation remain separate contextual tools.

Native preview gates do not require Acrobat by default. Add
`-IncludePDFIntegrationChecks` for the complete PDF-import flow verification.
Candidate packaging includes those checks on the verified build host.

The current r23 demo is `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r23.accdb`. Builds now run the focused saved-report review and native Reports return suite before seeding. `scripts/verify-ux-review-build.ps1` checks installed standard modules against source and native review/return event hooks on a disposable copy. See `planning/UX_REVIEW_RETURN_2026-10-04.md` for behavior and evidence.
