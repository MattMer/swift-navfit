# Manual-test change plan - 6 October 2026

Planned work against r32, not an implementation claim. This supersedes earlier UX decisions where requested. Preserve existing reports, revisions and published artifacts.

## Working assumptions and clarifications pending

- Defaults: global supplies starting values, group overrides global, report overrides group; or literal global-first winning precedence? Also determine whether parent changes update untouched inherited fields in existing drafts.
- Membership: moving removes the previous group, or reports may belong to multiple groups?
- Workflow: report-header buttons use the same four statuses as groups or retain separate review states? Does setting group status change reports automatically?
- History: use the latest prior evaluation ending before the current report starts? RUAD refresh updates one directory or maintains separate roster snapshots?
- Export naming: [RANK] is group paygrade or member rank/rate? What represents multiple ranks or ending dates?

Implementation proceeds with reversible working assumptions: Move replaces existing membership; global supplies starting values with group/report overrides; group workflow does not propagate to reports; report review retains local approval and ordered routes. Export uses report rank/rate or group paygrade and ISO date, with MIXED for differing group end dates. User corrections may change these choices. Parent-default edits do not silently rewrite saved drafts.

RUAD column mapping awaits the user's dummy sample. Narrative font, dimensions and wrapping require installed NAVFIT98A control/print inspection; do not invent measurements or claim verified printed fit.

## Phase 1 - Navigation and save-failure recovery

1. Put the active-contributor dropdown on Overview; update headers and subsequent attribution without rewriting old records.
2. Hide All Access Objects at startup; retain deliberate diagnostics/developer access.
3. Remember the originating report and section when opening snippets. Close returns to that report with pending values intact; Resources-originated browsing returns to Resources.
4. Failed saves must not trap navigation or application shutdown. Offer Undo pending changes / Reload last saved report and Retry. Discard clears only this session's pending values/checkpoint and preserves committed revisions and other contributors' work. Keep normal save-on-leave behavior, but avoid failure/retry loops. Test invalid fields, stale baselines, locked approval, database failures and closing Access afterward.

## Phase 2 - Summary groups become report organization

5. Remove destination-folder controls. Move report opens a summary-group chooser with type/paygrade compatibility and change preview; membership semantics await clarification.
6. Replace Reports folder filter with summary-group filter, including ungrouped reports. Preserve existing folder data/default/provenance dependencies during migration. Native NAVFIT folders remain necessary interchange structures.
7. Make the editable comparison table the default group member list. Add Open report to the left of each member name and preserve table state on return.
8. Add both editable career-recommendation columns with form-specific official captions, atomic revisions and existing lock/stale-review guards.
9. Show ITA and promotion recommendation in Reports and group tables. Reports also shows group membership. Label unavailable/NOB values and distinguish ITA from group average.
10. Show prior-evaluation ITA and promotion recommendation in group rows when history exists, with source period available. Depends on Phase 4.
11. Add Edit group defaults directly to the summary-group list.
12. Provide explicit group status action and display: Initialized, In-Work, Ready for Signature, Completed. Journal changes; distinguish local workflow from official signatures. Map old statuses safely and clarify report propagation.

## Phase 3 - Defaults and collaboration

13. Add global field-default editing using report-section layouts. Show inherited source, explicit override and Reset to inherited value. Fields remain editable. Confirm precedence and whether parent updates propagate before modifying the resolver; preview saved-draft changes and retain undo/revision safeguards.
14. Put review-status actions at the report header so the report remains visible. Preserve saved-revision semantics, pending-edit handling, contributor permissions and stale approvals/routes. State mapping awaits clarification.
15. Replace single draft notes with an entry list showing contributor and note. Store timestamps internally and preserve legacy notes with identifiable provenance. Final approvers read discussion alongside the report. Define edit/delete rules without changing others' attribution. Keep notes separate from NAVPERS comments.

## Phase 4 - Command directory, RUAD and evaluation history

16. Add RUAD import under Resources: Choose -> Preview -> Apply -> Results. Derive file formats and row/column mapping from the supplied dummy sample.
17. Maintain stable member identity, duplicate/conflict handling and safe refresh. Do not match by name alone, silently delete absent members or discard history. Distinguish manual entries from imported RUAD records.
18. Create reports for selected directory members into a group, previewing period/type/defaults first. Creating a report for a person absent from the directory adds a manual member record; link existing/imported reports without duplicates.
19. Member history lists evaluations written in Swift. Permit external prior-eval records with date range, promotion recommendation and individual traits. Record report kind/provenance and corrections; calculate ITA from applicable numeric marks with NOB handled separately. Prior-command history does not require creating a current draft.
20. Directory, prior evaluations and attributed notes require a separately justified additive schema migration and existing-workspace transfer tests.

## Phase 5 - Interchange, export naming and narrative fidelity

21. Generate Swift_NAVFIT_Export_[RANK]_[REPORT_END_DATE].accdb automatically. Choose destination directory without requiring a filename. Use consistent date formatting, sanitized tokens and safe collision suffixes; keep companion return-baseline files paired. Rank/mixed naming awaits clarification.
22. Replace stale narrative-only import wording. r27 already maps full supported report values. Extend NAVFIT import to a popup selecting native summary groups/folders, preview their details and contained reports, then apply selected groups/reports together. Preserve unmapped data, identities and rollback; handle duplicates/returned conflicts and prove group round trips in NAVFIT98A.
23. Explicitly size and font-lock official comments-on-performance using measured installed NAVFIT98A control/print behavior for each report type. Verify wrapping, line limits, paste/newlines and overflow against actual output before claiming fit. This brings narrative-fit verification into scope without redesigning every printed page.
24. Add a browsable NavyWriter example catalog with previewed snippet imports, deduplication and source URL/title. Examples are not policy. Verify available content and reuse terms before redistributing a catalog; use links/user-selected imports where appropriate. A starter catalog of eight short excerpts with www.navywriter.com source links has been implemented; this does not include the full website or establish a bulk redistribution license.

## Acceptance and rollout

Follow phases above, with history ahead of prior-eval columns and clarified semantics ahead of migrations. Build a new populated demo and separate blank artifact; preserve the published release.

Verify contributor switching; snippet/report return and pending buffers; successful and failed save/close; membership/filters/table editing; ITA/history values; defaults inheritance/override/reset; notes attribution; RUAD matching/refresh; selected-group native import; export collisions and paired baselines; narrative fit. Run existing recovery, approval, group and interchange suites plus targeted tests. Manually inspect keyboard behavior and native layouts at laptop/desktop sizes and scaling. Publish after exact-artifact and schema-transfer verification.
