# Navigation updates — explicit defaults context and comparison views

Implemented 2026-10-19. Candidate: build/navigation-context-verified.accdb, based on the previous navigation/help candidate; existing populated workspaces are preserved.

## Changes

- Group Defaults opens using the group selected in its originating list or member page. The defaults route no longer creates or reads a hidden workbench to decide its target. Direct calls accept an explicit group ID; older no-argument bindings retain their workbench compatibility behavior.
- The list's Defaults action opens the selected group directly. Members-to-Defaults transitions pass the resolved group ID, including empty groups and report-type switches. Global defaults remain explicitly scope zero.
- Opening defaults drains existing queued member cells first. If a save fails, navigation stops and retains the pending cell and member window; correction/retry keeps the original contributor and authoritative revision checks. Confirmed defaults keep their existing per-field transaction/rollback behavior.
- Marks/Proposal view switching uses one visibility pass for retired controls, removing the intermediate show-then-hide pass. Grid/proposal visibility and actions are unchanged.
- Existing contextual help buttons and the fixed report editor are retained. Review/History and empty Members still use the workbench adapter; this change removes that dependency specifically from Defaults.

## Verification

Focused native tests cover direct defaults without a workbench, stale workbench selection, global/group/type switching, invalid page rejection, a blocked failed-cell transition, correction/retry, return to Members and Marks/Proposal control visibility.

Final results, installed-source alignment and artifact hash are recorded in NAVIGATION_UPDATES_VERIFICATION.json. Native tests use disposable copies; the distribution candidate remains blank. The standard build applies these modules after the contextual-help step. Full generator, physical display/scaling, keyboard/screen-reader, Runtime/x86 and release-migration acceptance remain outside this development-candidate check.

## Remaining navigation backlog

Unified review entry/action services, retirement of obsolete folder UI after a full binding audit, and an explicit full-build dependency manifest remain separate work. Assignment and ordered-route history have not been merged or rewritten.

Final native verification: 30 explicit-context/view assertions, 200 defaults assertions, 61 live-member assertions, 42 pending-cell assertions, 39 group assertions and 57 contextual-help assertions passed. Application smoke Status=PASS and its artifact hash matches the candidate. All five installed modules align with source; blank inventory and prior/original hash checks passed.
