# Swift NAVFIT

A local Microsoft Access app for drafting and reviewing Navy evaluations and fitness reports. One operational `.accdb` contains the workspace, VBA, forms, draft data, history and local review tools.

**Latest distributed prototype: 0.21.0 · schema 20.** Development is ongoing. Full NAVFIT import/export and Navy policy validation are not implemented. Printed-layout work is deferred.

See [major foundations](planning/MAJOR_FOUNDATIONS_GUIDE.md) for summary groups, reviewed defaults/batches, multi-step routes, meeting proposals, limited Navy checks, structured PDF import and guided backup/restore.

## UI redesign under review

The [r24 milestone notes](planning/GITHUB_R24_NOTES_2026-10-04.md) summarize the current native redesign, tested fixes and remaining acceptance work. The [native verification record](planning/UX_MY_WORK_2026-10-04.md) identifies the exact local demo and blank files, checks and limitations.

Start with the [native workspace guide](planning/UX_WORKSPACE_GUIDE.md) for current workflows and the [current state](planning/CURRENT_STATE.md) for exact native artifacts and verification limits. The [candidate guide](planning/UX_CANDIDATE_GUIDE.md) explains blank versus demo files and source-preserving migration.

The working source is 0.22.0-ux-preview, schema 21. The [clickable design prototype](prototype/index.html) runs locally through `scripts/build-ux-prototype.ps1`; see [review journeys and scope](planning/UX_REDESIGN_REVIEW.md). It adds section navigation, automatic comparison baselines, reusable reporting seniors and rank-specific historical RSCA/PSG calculations. The native rebuild is under verification; it has not replaced the distributed 0.21.0 file. Visual, keyboard, screen-reader and scaling acceptance remain pending.

Build reviewable native previews with `scripts/build-preview.ps1`. They include synthetic reports, groups, review work, reporting seniors/RSCA and recovery examples. Each preview has a unique `UI_Demo` filename and a visibly synthetic command name. Demo seeding refuses any populated workspace. Use `scripts/build.ps1` for separate blank templates and release verification; never migrate operational data into a seeded demo.

Transfer an existing workspace into a separate **blank** template using `scripts/upgrade-workspace.ps1`; the source remains unchanged. Schema 21 adds four directory/baseline tables while retaining all 78 report fields. The native redesign remains under review and has not been established as visually equivalent to the browser prototype.

See [native demo build guide](planning/UX_DEMO_BUILD_GUIDE.md) for the synthetic examples and journeys to try.

## Current features

- EVAL, CHIEFEVAL and FITREP drafts with administrative fields, reporting-senior details, traits/reference guidance, arithmetic summaries and official narratives; incomplete saves.
- Explicit promotion/observation selections, report occasions and categories, primary duty and separate multiline duty descriptions.
- Member identifiers, checkbox service-component selections and date reported in Member & report details.
- Counseling date/name, physical-readiness details/date and billet subcategory in their official report sections, with reference suggestions.
- Assignment recommendations A/B, retention selections and member-statement selections, distinct from promotion and review acceptance.
- Reporting-senior identifiers/contact details and rater/senior-rater names and dates.
- Separate structured name parts for members, reporting seniors and counselors, with explicit initials and member suffix.
- Organizational folders, search, saved library views and bulk folder moves with guarded undo.
- Committed revisions, restore as a new revision, duplicate, archive/recover and field-level comparison.
- Automatic/manual pending-editor checkpoints and recovery.
- Read-only NAVFIT review, narrative-only import/reconciliation and retained source provenance.
- Attributed narrative snippets and a local draft completeness checklist.
- Revision-linked review comments, replies and resolve/reopen history.
- Local reviewer assignments, due dates, My Work, contributor switching and accept/return/cancel history.

Internal review acceptance is separate from official approval. Local contributor profiles provide attribution, not authentication. Use one contributor at a time in the same authoritative file; separate copies do not synchronize.

## Try the app

For the redesigned native UI, use the latest demo linked in
[CURRENT_STATE.md](planning/CURRENT_STATE.md), or the separate local candidate
package described in [UX_CANDIDATE_GUIDE.md](planning/UX_CANDIDATE_GUIDE.md).
The candidate is for review; it has not replaced the distributed checkpoint.
Settings & Help includes an in-app workspace guide. For existing data, provide
an explicit blank candidate `-TemplatePath` to the migration script.


Download the blank `.accdb` from the [0.21.0 prerelease](https://github.com/MattMer/swift-navfit/releases/tag/v0.21.0) and make a working copy. Open it in trusted desktop Microsoft Access, complete workspace setup, and use Home to open the draft library. The Access Runtime and x86 environments have not been verified.

Existing populated workspaces need migration into the new app rather than replacement with its blank distribution:

```powershell
powershell -NoProfile -File scripts/upgrade-workspace.ps1 -SourcePath 'C:\path\existing.accdb' -DestinationPath 'C:\path\upgraded.accdb'
```

The script defaults to the current release template in `output/`; download that release asset into `output/` or provide `-TemplatePath`. It preserves the source, refuses destination overwrites and verifies copied records.

## Build and collaborate

Windows and desktop Microsoft Access are required for building and integration testing. Build from text source into a fresh file:

```powershell
powershell -NoProfile -File scripts/build.ps1 -OutputPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/smoke-test.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
```

See [CONTRIBUTING.md](CONTRIBUTING.md) for setup, focused checks, migration requirements and pull requests. Binary databases, generated evidence, local test fixtures and supplied reference PDFs are excluded from Git. Source changes belong in `src/`, `scripts/` and `tests/`. Verified blank binaries are release assets.

## Progress and guides

- [Roadmap and current progress](planning/ROADMAP.md)
- [Structured member, senior and counselor names](planning/STRUCTURED_NAMES_GUIDE.md)
- [Reporting senior contact details and raters](planning/REPORT_PEOPLE_GUIDE.md)
- [Assignment recommendations, retention and statement](planning/DECISION_DETAILS_GUIDE.md)
- [Counseling, readiness and billet details](planning/SERVICE_DETAILS_GUIDE.md)
- [Identifiers, components and date reported](planning/MEMBER_DETAILS_GUIDE.md)
- [Occasions, categories and duty descriptions](planning/REPORT_CONTEXT_GUIDE.md)
- [Reporting senior details](planning/REPORTING_SENIOR_GUIDE.md)
- [Promotion status](planning/PROMOTION_STATUS_GUIDE.md)
- [Observation and recommendation selections](planning/ASSESSMENT_GUIDE.md)
- [Trait summaries](planning/TRAIT_SUMMARY_GUIDE.md)
- [Trait editor guide](planning/TRAITS_GUIDE.md)
- [Administrative draft fields](planning/ADMINISTRATION_GUIDE.md)
- [NAVFIT package findings](planning/NAVFIT_PACKAGE_FINDINGS.md)
- [Review assignments and My Work](planning/REVIEW_ROUTING_GUIDE.md)
- [Review comments](planning/REVIEW_COMMENTS_GUIDE.md)
- [Import provenance](planning/PROVENANCE_GUIDE.md)
- [Workflow tools](planning/WORKFLOW_TOOLS_GUIDE.md)
- [Detailed build and verification history](planning/BUILD_STATUS.md)
- [Architecture and implementation plan](planning/IMPLEMENTATION_PLAN.md)
- [Current release manifest](output/MANIFEST.json)

The 0.21.0 build passed 2001 behavioral assertions and 218 reference checks, plus exact narrative-import/provenance checks. Schema-18-to-19, schema-19-to-19 and schema-5-to-19 migrations preserved every original row/field and reopened. All 20 user-data tables in the blank distribution are empty. Full visual inspection and Runtime/x86 testing remain pending.

This private repository has no selected open-source license. Coordinate use and redistribution with its owner.
