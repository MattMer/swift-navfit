# Progress and roadmap

Current prototype: **0.18.0**, Swift schema **17**. This is a local Microsoft Access app under development. Printed-layout work is deferred. The app does not yet provide full NAVFIT interchange or Navy policy validation.

| Area | Current progress |
| --- | --- |
| Workspace | Single-file blank distribution, contributor profiles, organizational folders, search and saved library presets |
| Drafts | Three forms; administrative/reporting-senior fields, traits, summaries, observation/recommendation/status selections, occasions, categories, duty descriptions, member identifiers/components and date reported, counseling, physical readiness and billet subcategory, assignment recommendations, retention and statement selections; incomplete saves |
| History | Committed revisions, restore as new revision, duplication, archive/recover and field-level comparison |
| Recovery | Timed/manual pending-editor checkpoints and deliberate recovery |
| Imports | Read-only source review, confirmed narrative-only mapping, retained source provenance and explicit two-way narrative reconciliation |
| Review | Revision-linked comments/replies, resolve/reopen history, local assignments, due dates and My Work |
| Bulk actions | Folder move preview/apply, guarded undo and operation history |
| Readiness | Local draft completeness checklist; no official policy/export-readiness claim |

## Next implementation priorities

1. Remaining administrative fields (reporting-senior identifiers/contact details, rater/senior-rater details and dates, concurrent reporting-senior details and structured name parts), recommendation/observation context and eligibility rules for EVAL, CHIEFEVAL and FITREP, with verified field mappings.
2. Summary groups independent of organizational folders; eligibility, exceptions and complete-group context.
3. Source-linked, versioned Navy validation rules with boundary and exception tests.
4. Reviewed defaults/inheritance and official-field bulk changes.
5. Complete NAVFIT import/export, cross-file identity matching and three-way returned-file reconciliation, followed by target-app round-trip proof.
6. Multi-step route templates, official approval tracking and stale-approval handling.
7. Meeting comparison/rack-and-stack snapshots and proposals.
8. Structured PDF import and guided backup/restore.
9. Access Runtime/x86, workstation compatibility and comprehensive visual testing; printed layout later.

These are priorities, not delivery-date promises. Open issues can refine scope and sequence.

## Verified milestone

The 0.18.0 release passed 1356 assertions plus 218 reference checks across administrative, routing, comments, provenance, workflow, narrative, comparison, recovery, direct review and smoke suites. Narrative import verified all 2472 retained values. Schema-16-to-17, schema-17-to-17 and schema-5-to-17 migrations preserved every original row and field, source hashes and reopen behavior. All 20 user-data tables in the distribution are empty. See DECISION_DETAILS_GUIDE.md and BUILD_STATUS.md for exact scope and limits.
