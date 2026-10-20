# RSCA completion and report review

Development candidate: `build/rsca-review-v6.accdb`; synthetic demo: `build/rsca-review-demo-v6.accdb`. These are local candidates, not a published release or replacements for existing workspaces.

## Why the demo said unavailable

The two new E6 groups had no GroupRSCABaseline selection and no E6 historical baseline. Their displayed senior name came from member report values. A senior name or another paygrade's RSCA is insufficient: PSG needs the same senior, evaluated paygrade, previous graded trait count, and confirmation that history excludes this group. The demo now has a synthetic E6 baseline of 420 points / 100 traits and both E6 groups select it. For the clean group, (420 + 56) / (100 + 14) = 4.1754386, displayed as 4.18. Group-only average remains 4.00.

## Completion accounting

Marking a group Completed posts its saved, graded traits to a new immutable baseline for that senior/paygrade. Original historical observations remain intact. Count means numeric traits, not reports. Whole-report and individual NOB traits are excluded; blank or invalid observed traits stop completion. Pending editor changes must first be saved/discarded. Senior/grade mismatch, archived members and duplicate report contributions stop completion.

The group policy records the contribution and baseline identity. The status, appended baseline, provenance and audit operations share one transaction. Repeating completion with unchanged contributions creates no new baseline. Reopening/archiving retains the prior historical contribution; completing corrected saved traits applies the difference against the current local baseline chain. A report already counted through another group cannot be counted again. Projections follow subsequent local entries descended from the selected baseline, subtract this group's earlier contribution, then add its current saved traits once.

Local entries are explicitly marked as local completion, not BOL verification. Exact historical sums remain exact; rounded historical averages retain estimated provenance. A newer independent BOL observation is not silently merged with earlier local contributions. Groups with previously posted contributions stop with a reconciliation message if the selected history no longer contains that contribution. There is no automatic BOL reconciliation in this candidate.

No tables or fields were added: completion markers use the existing versioned group-policy XML, and baselines use existing history rows. This does not constitute release migration acceptance.

## Draft to locally complete

1. Create/open a report, enter Front Page and Back Page values, and Save draft. Incomplete drafts can be saved. Draft notes hold local contributor writing notes and recovery checkpoints.
2. Open Validation issues. Correct invalid/incomplete blocks and review policy findings. Validation does not itself approve a report.
3. Open Review setup. Choose a route template containing the contributors in review order and Start review. Templates are maintained through Resources. The route refers to saved values; save pending edits first.
4. The current assigned contributor opens the report from My Work or Review setup. Review comments are revision-linked feedback threads, separate from draft notes. Accept step advances to the next contributor. Return report records the correction reason and ends that route.
5. For corrections, edit/save the draft and start another route after a return. If a pending/accepted route became stale after a saved report or group-context change, its requester restarts it against current values. Earlier steps and decisions remain in history.
6. After every ordered step is accepted, the final reviewer records Local approval with a note. The report becomes Approved and editing is locked. Reopen approval with a reason returns it to Draft for correction and another review cycle.
7. Group statuses describe separate closeout work: Initialized, In-Work, Ready for Signature, Completed. Complete the group when closeout is finished; this now posts local RSCA totals. This action does not approve member reports or supply official signatures/submission. Export, official validation, printing and submission remain separate Navy/NAVFIT98A steps.

Legacy one-off assignments also exist in My Work. Accepting one accepts that assignment only; it does not complete an ordered route or approve a report.

## Distinct report tabs

Draft notes contains contributor notes and checkpoints. Review contains routing, revision-linked comments, validation and calculation actions. History contains saved revisions, compare and restore. The page transition hid the shared subform while parent painting was suspended, leaving its native child content over Review. The selector now clears and hides that child before suspending painting, uses explicit page branches, and restores/repaints the parent once after a transition. Fixed official field geometry and pending buffers are retained.

## Verification

Desktop Access 16.0 build 20430: 285 completion/child-content assertions, 29 page-switch preservation checks, 264 fixed-page checks, 40 lifecycle checks plus forced shutdown/recovery, 39 populated demo checks, 39 directory checks, 204 legacy/shared-cursor equivalence checks and 48 smoke checks passed on the final candidate. Visible synthetic group and Notes-to-Review checks passed. Candidate hashes and preservation evidence are recorded in RSCA_COMPLETION_REVIEW_VERIFICATION.json. Free Access Runtime is unsupported. Full release migration/print/accessibility acceptance is outside this local candidate.
