# 0.21.0 foundations checkpoint

Summary-group snapshots, defaults and atomic official-field batches, multi-step local approval routes, meeting snapshots/proposals, limited source-linked validation, eight-field interactive PDF import and guided backup/restore are implemented. See [scope and limits](MAJOR_FOUNDATIONS_GUIDE.md). Earlier priorities below describe the broader intended scope; foundations do not complete full grouping, validation or interchange.

# Progress and roadmap

Current prototype: **0.20.0**, Swift schema **19**. This is a local Microsoft Access app under development. Printed-layout work is deferred. The app does not yet provide full NAVFIT interchange or Navy policy validation.

| Area | Current progress |
| --- | --- |
| Workspace | Single-file blank distribution, contributor profiles, organizational folders, search and saved library presets |
| Drafts | Three forms; administrative/reporting-senior fields, traits, summaries, observation/recommendation/status selections, occasions, categories, duty descriptions, member identifiers/components and date reported, counseling, physical readiness and billet subcategory, assignment recommendations, retention and statement selections, senior identifiers/contact and raters/dates and structured member/senior/counselor names; incomplete saves |
| History | Committed revisions, restore as new revision, duplication, archive/recover and field-level comparison |
| Recovery | Timed/manual pending-editor checkpoints and deliberate recovery |
| Imports | Read-only source review, confirmed narrative-only mapping, retained source provenance and explicit two-way narrative reconciliation |
| Review | Revision-linked comments/replies, resolve/reopen history, local assignments, due dates and My Work |
| Bulk actions | Folder move preview/apply, guarded undo and operation history |
| Readiness | Local draft completeness checklist; no official policy/export-readiness claim |

## Next implementation priorities

1. Remaining administrative fields (additional RRS reporting-senior fields (role/mapping to verify), qualifications/achievements and remaining narrative/context fields), recommendation/observation context and eligibility rules for EVAL, CHIEFEVAL and FITREP, with verified field mappings.
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

The 0.20.0 release passed 2001 assertions plus 218 reference checks across administrative, routing, comments, provenance, workflow, narrative, comparison, recovery, direct review and smoke suites. Narrative import verified all 2472 retained values. Schema-18-to-19, schema-19-to-19 and schema-5-to-19 migrations preserved every original row and field, source hashes and reopen behavior. All 20 user-data tables in the distribution are empty. See STRUCTURED_NAMES_GUIDE.md and BUILD_STATUS.md for exact scope and limits.
