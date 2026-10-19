# Priority cleanup batch 6 — U01, U02, U03

Implemented 2026-10-19 on a copy of priority-batch5-final.accdb.

- U01: Editor feedback distinguishes pending edits, recovery checkpoints, failed saves and saved revision numbers. Grid feedback distinguishes queued cells from completed automatic saves. Existing buffer, retry, discard, transaction and recovery mechanics remain intact.
- U02: Shared-field editing shows scope and potential local-approval impact before a field changes. Counts exclude archived items and filter field applicability on entry. Exact changes and approval reopening still use the existing confirmation and atomic transaction.
- U03: Editor/review actions highlight the next available action and explain disabled actions. Status text separates local approval, group status and official submission. Availability and authority rules remain unchanged.

The reported End If compile failure was caused by a single-line If incorrectly expanded into a block; it is corrected. Native testing also exposed an overlength tooltip during approval reopening; dynamic tooltips now respect Access's 255-character limit. Full text remains visible in status labels.

## Verification

Native Access tests ran on disposable copies of build/priority-batch6-final.accdb:

- 200 shared-default assertions, including the pre-edit scope banner.
- 24 header review assertions, including primary accept/approval, reopen, return and stale-route rejection.
- 42 pending-cell preservation, failure, retry and discard assertions.
- 12 recovery preparation and 23 recovery assertions, plus automatic timer/active-text checkpoint checks after reopening Access.
- 264 fixed report-page, pending buffer, local workflow and closeout assertions.
- Application smoke suite: PASS (see build/batch6-smoke-final.log).
- All nine installed modules match source after normalizing Access export formatting while preserving strings.
- Blank artifact inventory completed; prior batch 5 and populated r116 hashes unchanged.

The legacy header test expected every action to shrink into a 9000-twip window. That assertion also failed on batch 5. It now verifies fixed-page controls remain within the scrollable form when fixed pages are active, while retaining the old compact-window bound otherwise. The separate fixed-page suite verifies field geometry does not reflow and all seven traits remain on Front.

## Artifacts and limits

Candidate: build/priority-batch6-final.accdb
Manifest: build/priority-batch6-final.manifest.json
Evidence: planning/PRIORITY_CLEANUP_BATCH6_VERIFICATION.json and build/batch6-*-final.log

The incremental builder imports the nine modules and installs the scope label/field-entry hooks. The same apply step is wired into the standard build after theme application. A full standard form-generation build was not repeated for this batch. Visual review at display scaling and Access Runtime/x86 validation remain pending. No production workspace was replaced and no release was published.
