# Comparison event save changes — r93, 9 October 2026

The user still reports 3420 on the first trait change in r92; a prior debugger screenshot showed 3218 at the revision-pointer update. Retaining DAO owners and sharing the report write handle did not establish resolution. The exact user failure still has not reproduced in automation.

r93 changes the native event path. All ten editable comparison controls now queue immutable report ID, field and value instead of committing inside AfterUpdate. A 150 ms form timer finishes the temporary bound-cache edit before starting the guarded report transaction. Repeated changes to the same cell coalesce, edits to different member rows retain their identity, and native unload flushes queued edits. Report saves remain automatic. An invalid queued edit reloads the saved comparison values so subsequent cells do not appear saved when they were not applied; retry remains available. Existing local-approval prompts, stale-revision guards, atomic revision services and schema 25 remain.

Source build helpers install the queued control events and timer for full and incremental builds. This is an event-path change, not a claim that the original lock condition was reproduced.

Verification: native compile for all three delivered artifacts; actual control Text/SetFocus events with real timer dispatch in new EVAL, CHIEFEVAL and FITREP groups/reports; pending edits flushed on close; queued E6 1.0 marks; cross-member identity; repeated-cell coalescing; invalid career recommendation cleanup and retry; 26 approval rollback/retry assertions; Avery Patel/Riley Brooks repeated saves/reopen; 142 admin assertions; 129 group/directory/Overview assertions. Every one of the 49 user tables is identical between r92 and r93 in both the supplied test workspace and demo. The blank opens Setup. All ten bindings and the form timer were inspected in the final artifacts.

No computer use was used. During implementation, a temporary timer-property typo and an automation build stall were corrected before final verification; their failed runs are not counted as passes. Optional external-lock injection from r92 remains unverified. User retest is necessary before calling the reported production journey resolved.

Open Swift_NAVFIT_UI_TestWorkspace_2026-10-09-r93.accdb to keep the existing test group/report, or Swift_NAVFIT_UI_Blank_2026-10-09-r93.accdb for an empty workspace. The r92 sources and the author's edited baseline remain preserved. No GitHub publication.
