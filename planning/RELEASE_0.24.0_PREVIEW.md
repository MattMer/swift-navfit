# Swift NAVFIT 0.24.0-preview

Deployment prerelease, schema 23. Includes a blank workspace and a synthetic demo with nine reports, three summary groups, nine directory members, eight prior evaluations, attributed notes, review work and recovery examples.

## Changes since 0.23.0-preview

- Summary-group correctness and policy context: editable member comparison with both career recommendations, per-row report opening, current and prior ITA/promotion, explicit group status and membership moves, global/group/report starting defaults and inheritance resets.
- Command directory and prior evaluation history: stable identities, confirmed/provisional matching, explicit corrections, manual prior periods/traits/recommendations and previewed atomic report creation for selected members. RUAD file mapping awaits a synthetic sample.
- Local approval no longer becomes stale solely from its own or a sibling report’s edit lock; membership/revision changes retain stale guards.
- Verified backup/restore now accepts schema-23 workspaces; packaged embedded worker is checked against the standalone tool.
- Attributed, append-only report notes with full-text viewing, separate unposted-note checkpoints, recovery and shutdown safeguards.
- Overview contributor switching, hidden Access Objects pane, contextual report review buttons, snippet return to the originating draft and failed-save undo/reload.
- Selected-summary-group NAVFIT import, full supported report values and provenance, default export filenames, existing guarded report/group exports and single-report returned-file reconciliation.
- Expanded NavyWriter catalog: 72 rate/source collections and 20 short excerpts; rate/search filters in the source catalog and saved snippets, plus manual snippet rate assignment. Full source collections are linked, not redistributed.
- Expanded live report and group-wide policy checks, 37 verified instruction excerpts/page locators, official NOT REQ/NOT PERF counseling support through save, PDF import and NAVFIT conversion. Narrative adequacy, referral and approval exceptions remain human review items.

## Use and migration

Extract the ZIP and open **Swift_NAVFIT_0.24.0-preview-DEMO.accdb** to explore. Use **Swift_NAVFIT_0.24.0-preview-blank.accdb** for a new workspace or as the template for a new migration destination. Never transfer operational records into the seeded demo.

The included upgrade script preserves originals and verifies original rows/fields. Schema-20 through schema-23 transfer acceptance is documented in the package evidence. Do not overwrite an operational database or use an older template to downgrade schema-23 data. Read START_HERE.md and guides/UX_WORKSPACE_GUIDE.md.

## Verification and limitations

Exact packaged databases are tested on disposable copies, audited against all 72 source modules and compiled. MANIFEST.json records package files, SHA-256 hashes, empty blank tables and successful suite logs. The demo intentionally has synthetic data and a recovery checkpoint.

Automated policy/catalog checks passed 102/29 assertions; service checks passed 170 plus 49 reference checks. Native adapter/exchange guards and structured PDF status import were verified in development. Automated display geometry is not physical display or accessibility acceptance.

Outstanding: actual 1366Ãƒâ€”768/1920Ãƒâ€”1080 displays at 100/125/150% scaling, keyboard/screen-reader acceptance, Access Runtime/x86, RUAD sample mapping, full instruction coverage, exact narrative print fit and native NAVFIT UI/printing round-trip acceptance. Returned-group reconciliation remains incomplete. Local approval is not an official signature or Navy acceptance. Contributor profiles provide attribution, not authentication; separate copies do not synchronize.

See NARRATIVE_POLICY_REVIEW_2026-10-06.md and MANUAL_FEEDBACK_PROGRESS_2026-10-06.md for coverage and research-backed next priorities.
