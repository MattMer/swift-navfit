# r96 report revision cursor fix — 9 October 2026

The user's r95 save trace proves CommitDraft reached revision linking and failed with DAO 3218; CommitSavedFields then surfaced 3420. Navigation completed before this save stage.

CommitDraft now retains its report write recordset from the field update through revision linking, and sets CurrentRevisionID through that same cursor instead of issuing a second SQL UPDATE after closing it. The cursor is closed and cleared only after linking. The revision recordset is also cleared immediately after closing so failure cleanup does not operate on a closed revision cursor. Existing transaction, snapshot, history, approval-reopen and stale-baseline services remain in place. No schema change.

Candidate: output/ux/Swift_NAVFIT_UI_TestWorkspace_2026-10-09-r96.accdb. Contains the user's r95 data and local save-stage diagnostics. Originals preserved. Normal source fix is src/modSwift.bas; instrumented candidate source is build/tests/r94-diagnostic-sources/modSwift.bas.

Validation: native compile; all 49 tables identical to r95; 26 approval, rollback, retry, repeated-save and reopen checks; real dropdown/timer events across EVAL, CHIEFEVAL and FITREP; queued E6 1.0, multi-row identity, coalescing, failure reset and retry; existing user-created report native-event save on a scratch copy. All passed. Automated tests had also passed earlier builds; user's exact interactive failure is not declared resolved until verified in r96. Diagnostic trace retained for that verification. No computer use, GitHub push or release publication.

## User confirmation

On 9 October the user reported that r96 worked in the previously failing interactive journey. The revision save issue is resolved; subsequent regression checks retain the cursor fix.
