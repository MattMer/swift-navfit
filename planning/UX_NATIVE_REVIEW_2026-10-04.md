# Swift NAVFIT r23 native UX review

Reviewed October 4, 2026 in the user's running Microsoft Access installation through the computer-use plugin. Target: `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r23.accdb`. This is a native exploratory review, not a complete acceptance or regression certification. No application source fixes were made during this review.

## Display and tooling evidence

- The regenerated `@oai/sky` runtime loaded, listed windows and captured Access successfully.
- Captured the maximized Access window at 1920 x 1080, then resized the same native window to exactly 1366 x 768 (returned screenshot dimensions).
- Windows display scaling was not changed or verified. Access's status bar showed 100% zoom; this is not evidence of Windows DPI scaling.
- Screenshot-backed coordinate input worked. Screenshot IDs and accessibility element indexes were rejected as unavailable across calls. Focus metadata repeatedly reported the ribbon search box even when the visible caret was in an Access field. These are automation limitations; do not treat them as a screen-reader verdict.
- Tab moved visibly from the first comparison trait to the next trait. A full keyboard and screen-reader test remains outstanding.

## Highest priorities

### P1: Report reopening crashes when the list is at its top

Reproduction: Overview > New report > FITREP; enter `TESTER, UX` in the official full-name field; switch to Reporting senior; Back to reports; search TESTER; Find / refresh; select the matching report; Open report.

Observed: a report appeared in the search results, but opening raised VBA run-time error 13, Type mismatch. Debug highlighted `UXCaptureReportList` in `modUXReportReturn`:

```vb
If topPosition < f!lstReports.ListCount Then topReport = Nz(f!lstReports.Column(0, topPosition), 0)
```

The destination is a Long. The query list includes column headings; at top position zero, reading the header rather than a numeric report identity is the likely cause. This hypothesis needs a focused reproduction and fix. The report library's unrelated scroll bookkeeping must never prevent opening a selected report. Verify zero/one/many results, headings, top and scrolled positions, missing scrollbars, and return after save.

Creation, text entry and preservation through a section change were visible. Leaving without a confirmation returned to Reports and the new row was searchable. Successful saved-revision creation and reopening are **not** certified by this test. A later Overview capture still showed 8 reports and two recovery checkpoints, while Reports had shown 9 active rows; investigate refresh/persistence as part of this defect rather than assuming the save completed normally.

The debugger was reset and returned to Access after recording the failure. The synthetic test report/checkpoint was left in the demo for investigation. No data was deleted.

### P1: My Work does not represent the review workload

The r23 inbox lists only two simple assignments. Ordered route steps and unresolved comment threads are absent. A stale Avery Patel assignment shows its stale status and overdue due date in details, but Accept/Return remain visually available. The detail text exposes `Report 6 / saved revision ID 12`.

Integrate route steps and unresolved threads with names, report period, actual saved revision number, current responsible contributor and explicit next action. Disable actions that cannot succeed for the selected state, while retaining backend guards. No comment-assignee field or route-step due dates should be invented.

### P2: Make shared layout and workflow state consistent

- Reports heading is visibly cut off by the filter row at 1920 x 1080.
- Summary group helper text overlaps the Group status label/filter.
- Comparison title is cut off by the Trait marks / Preview staged changes buttons.
- Read-only editor explanation is clipped after `reopen local`; the user cannot read the complete next action.
- Save draft appeared as a blank button on the editable Casey report and new FITREP. It was labeled on the approved Jordan report. Investigate editable-state caption/rendering; keep the label visible in every state.
- Review comments, snippets, recovery and import chooser retain older fixed-width layouts without the application navigation/header. Closing snippets returned to Overview rather than its Resources origin.
- Report list exposes `FormKind`, `IsArchived` and numeric archive values. My Work, group members, comments, history and RSCA list regions lack readable column headings in the captured views.
- Multiple persistent Access document tabs accumulate, alongside the application's own navigation. The All Access Objects pane exposes backend tables and VBA modules and consumes about 210 pixels. A user-facing build needs a deliberate shell; technical objects belong behind diagnostics.
- Blank checkbox values render as blue indeterminate squares. New-report users can mistake them for selected choices. Make unset/unchecked semantics understandable without changing existing saved values silently.
- Final report details on CHIEFEVAL show controls explicitly labeled `Not used on this report type; stored value retained`, with some messages clipped. Preserve the data, but move irrelevant legacy fields into a clearly named secondary area.

## Screen coverage and observations

| Area | Native checks | Outcome / limitation |
| --- | --- | --- |
| Overview | Counts, recent report identities/status, recovery entry point, create report | Opens; counts/refresh after test creation need investigation |
| Reports | Selection, Open report, TESTER search, import chooser, return selection | Search works; normal Casey opening worked initially; single-result opening reproduced error 13 |
| EVAL editor | Member details, performance traits, recommendations/comments, review entry points | Compact fields and all seven marks visible at 1920 x 1080; narrative count and snippet access visible |
| FITREP editor | Create, official name, section switch, Reporting senior | Name survives section change; report creation/reopen workflow fails as described above |
| CHIEFEVAL editor | Approved report, member details, command/duties, counseling, final details, history, traits | Read-only wording and disabled Save visible; no mutation of approval attempted |
| Laptop editor | Member details and traits at 1366 x 768 | Reflow and vertical field scrolling work; fixed actions visible, no horizontal form scroll seen; trait guidance clipped in its panel |
| Report review | Saved report heading, empty route selector/steps, disabled actions and explanations | Opens; disabled-action explanation is a long generic paragraph; no route decisions made |
| Comments | Open Casey root, read thread, return to same editor | Opens and return works; raw report/revision IDs and missing shared shell remain |
| Groups | All three groups, report counts 3/2/2, select/open E5 workspace | Opens; focused group shows members, trait average 3.86 and PSG RSCA 4.16 |
| Comparison | Open directly from group; compact trait columns; Tab movement; Preview staged changes | Opens without manual freeze; no-change preview is empty with Apply apparently available and no explicit empty-state message; no proposals applied |
| My Work | Pending assignments, stale/overdue details, navigation | Opens; missing routes/comments and action-state problems above |
| Resources | Landing page, snippet screen, reporting senior directory/record, RSCA history tab | Opens; senior entry still uses oversized equal-width fields; RSCA history region appeared empty despite demo baselines and needs verification |
| Settings | Landing page, recovery screen and return | Opens; visible title `Settings _Help`; recovery data not applied/discarded |
| Imports | Choose stage and Back to Reports | Opens and return retains TESTER filter; no file import or Apply tested |

## Recommended sequence

1. Fix report-list capture/reopening and prove create > save > search > reopen, including empty and single-result lists.
2. Complete My Work routes/comments integration and selection-specific action guards.
3. Fix shared title/filter spacing, visible Save/locked-state feedback, readable table headers and consistent origin-aware returns.
4. Finish converting supporting screens to the shared workspace shell; simplify tab accumulation and default object-pane exposure for the user-facing build.
5. Run native mutation journeys on a fresh demo: bulk preview/apply/undo, route advance/stale restart/local approval, comment resolve/reopen, import preview/apply/results/undo, recovery, and failure-preserved edits.
6. Repeat keyboard and screen-reader checks and the full 1366 x 768 / 1920 x 1080 matrix at actual Windows scaling 100%, 125% and 150%.

Not tested to completion: all field/block mappings, all three report types across every section, bulk Apply/undo, import mapping/apply/undo, backup/restore, group membership mutation, reporting-senior baseline writes, full route lifecycle, approval permissions, validation failures, crash recovery, full keyboard order, screen-reader behavior and Windows DPI matrix. This review does not replace existing automated behavioral suites or authorize publication of a release.
