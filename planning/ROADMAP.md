# Swift NAVFIT roadmap — 4 October 2026

Current native development preview: **0.22.0-ux-preview, schema 21, r23**.
See [complete project-plan status](PROJECT_PLAN_STATUS_2026-10-04.md) for every
original and redesign workstream, evidence limits and outstanding acceptance.
The prior distributed checkpoint is 0.21.0/schema 20; old milestones below are
historical and do not certify current full NAVFIT interchange or Navy validation.

The core drafting/revision/recovery library, explicit groups, defaults/bulk/undo,
compact editable comparison, senior/RSCA directory, local review/approval,
partial imports and backup/migration services are implemented in their documented
scope. r23 completes focused saved-report review/stale-route restart and native
Reports scroll/selection return. Physical native acceptance remains outstanding.

Next focus:
1. Native end-to-end visual/keyboard/scaling acceptance and integrate ordered
   route steps/unresolved work into My Work.
2. Complete verified NAVFIT field mappings, full staged import/export and
   returned-file reconciliation/round-trip proof.
3. Complete policy-defined grouping, exception/completeness and versioned rules.
4. Exact-artifact broad regression, migration and Runtime/x86/workstation
   verification before replacing the distributed prerelease.
5. Later rollover/targeted comments/richer comparison/full PDF mapping. Printed
   layout stays deferred; no network/authentication/synchronization/signatures.

## Verified milestone

The 0.20.0 release passed 2001 assertions plus 218 reference checks across administrative, routing, comments, provenance, workflow, narrative, comparison, recovery, direct review and smoke suites. Narrative import verified all 2472 retained values. Schema-18-to-19, schema-19-to-19 and schema-5-to-19 migrations preserved every original row and field, source hashes and reopen behavior. All 20 user-data tables in the distribution are empty. See STRUCTURED_NAMES_GUIDE.md and BUILD_STATUS.md for exact scope and limits.
