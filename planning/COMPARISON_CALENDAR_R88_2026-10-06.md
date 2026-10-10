# r88 comparison retry and native dates — 6 October 2026

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r88.accdb`; matching blank is `Swift_NAVFIT_UI_Blank_2026-10-06-r88.accdb`. Demo copied from the current r87 workspace without reseeding; existing practice records preserved. Version 0.25.0-dev/schema 24; GitHub release unchanged.

Comparison edits now isolate the temporary row database in its own DAO workspace, capture the expected revision before the approval prompt, guard against reentrant edit callbacks, retain owned database references in dependent operations, and restore the row after failed saves. Errors include the failing save stage and error number. The misplaced `stage` diagnostic in UXGridCurrent was removed before this build.

Native calendar selectors are enabled for report dates and applicable shared/default, prior-evaluation, review-due and reporting-senior date fields. Calendar dates normalize to yyyy-mm-dd; invalid typed entries remain recoverable and counseling status text remains supported.

Verification: compiled blank and demo; 128 calendar assertions (all three report types, raw invalid recovery and counseling statuses); 26 approved-trait rollback/retry/repeated-save/reopen assertions; 142 admin/comparison assertions; 46 major workflow, 39 summary-group correctness and 26 global-default assertions. Both packaged copies open their normal startup forms. Visual check on a disposable r88 demo confirmed the Date Reported calendar icon and calendar popup in native Access.

Limit: the original object-invalid error was not reproduced in clean r87 native approved-report tests. Even the older build passed the corrected synthetic failure/retry test. These changes harden transaction and object lifetimes; they do not establish a reproduced root cause. If it recurs, the new stage/error diagnostic is needed. No new physical scaling, screen-reader, NAVFIT printing or GitHub publication claim.
