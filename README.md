# Swift NAVFIT

A local Microsoft Access app for drafting and reviewing Navy evaluations and fitness reports. One operational `.accdb` contains the workspace, VBA, forms, draft data, history and local review tools.

**Previous versioned prerelease: 0.24.0-preview · schema 23.** Download the blank workspace or synthetic demo from the [GitHub release](https://github.com/MattMer/swift-navfit/releases/tag/v0.24.0-preview). Includes the native workspace redesign, summary-group tools, command directory and prior evaluations, contributor notes, expanded rate-filtered narratives and cited live validation. Supported NAVFIT report/group exports and selected-group imports retain provenance; single-report returned files use three-way reconciliation. Full policy coverage, returned-group reconciliation and exact printed fit remain pending.

See [major foundations](planning/MAJOR_FOUNDATIONS_GUIDE.md) for summary groups, reviewed defaults/batches, multi-step routes, meeting proposals, limited Navy checks, structured PDF import and guided backup/restore.

See [0.24.0 release notes](planning/RELEASE_0.24.0_PREVIEW.md), [current state](planning/CURRENT_STATE.md) and [policy review/recommendations](planning/NARRATIVE_POLICY_REVIEW_2026-10-06.md) for verified behavior and remaining work.

## Current development baseline

**r113 — 15 October 2026 · 0.25.0-dev · schema 25.** Current tested downloads: [blank workspace](https://drive.google.com/file/d/1oVLR6_3xvLSnIydOmRxwbXCWXhRdkG6s/view) and [synthetic demo](https://drive.google.com/file/d/10Bqvj97rj4VyzMPIDPdYuy5LQULWuDx7/view). The latest GitHub-hosted deployment asset remains [r104](https://github.com/MattMer/swift-navfit/releases/tag/deployment-2026-10-10-r104); source and current-state notes now cover r105–r113.

Recent changes include automatic older NAVFIT schema conversion; immediate confirmed recursive defaults; reporting-senior RSCA tables and clearer group tools; shared live/saved/export readiness with policy references; visible import groups with unresolved paygrades; automatic PSG selection saving and group-page PSG refresh; SSN/unique-name matching and report-driven RUAD profile enrichment; automatic 20xx import dates; and paygrade/specific-rate filters for single and batch member report creation.

The r113 candidate passed **668 relevant native assertions**, including real Acrobat extraction, import preview/apply safeguards, atomic group import, directory creation and identity/history/PSG behavior. Both builds compile/start, the blank has 48 checked user-data tables empty, and the demo has nine synthetic reports. See [r113 scope and verification](planning/IMPORT_RANK_R113_2026-10-15.md), [r112 member/PSG behavior](planning/MEMBER_PSG_R112_2026-10-15.md), and [r110 readiness coverage](planning/READINESS_R110_2026-10-14.md). Manual display, keyboard/screen-reader and printed-layout acceptance remain outstanding.

It retains [r103's single-user performance and save-recovery improvements](planning/DEPLOYMENT_R103_2026-10-10.md). Larger synthetic groups were measured locally at 25/50/100 members per report type; see [scale results](planning/SCALE_R102_2026-10-10.md). Report editor startup is the next performance target. Exact printed wrapping, native round-trip and physical accessibility acceptance remain outstanding.

Build scripts and individual change/verification notes for r105–r113 are in `scripts/` and `planning/`. Access databases, local test data and supplied NAVFIT program files remain excluded from Git; use the explicit download links above for the app.

## UI redesign under review

The [r26 PDF-import notes](planning/PDF_IMPORT_R26.md) describe expanded mappings, existing-draft merges and the target-computer extraction check. The [r25 live-validation notes](planning/UX_LIVE_VALIDATION_2026-10-04.md) describe live report/group diagnostics and policy citations. The [r24 milestone notes](planning/GITHUB_R24_NOTES_2026-10-04.md) retain the earlier native redesign checkpoint and acceptance limitations.

Start with the [native workspace guide](planning/UX_WORKSPACE_GUIDE.md) for current workflows and the [current state](planning/CURRENT_STATE.md) for exact native artifacts and verification limits. The [candidate guide](planning/UX_CANDIDATE_GUIDE.md) explains blank versus demo files and source-preserving migration.

The working source is 0.25.0-dev, schema 25. The [clickable design prototype](prototype/index.html) runs locally through `scripts/build-ux-prototype.ps1`; see [review journeys and scope](planning/UX_REDESIGN_REVIEW.md). It adds section navigation, automatic comparison baselines, reusable reporting seniors and rank-specific historical RSCA/PSG calculations. The native rebuild is under verification; the prior 0.21.0 checkpoint is retained. Visual, keyboard, screen-reader and scaling acceptance remain pending.

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
- Familiar full-name fields for members, reporting seniors and counselors; structured export fields remain in the background.
- Summary-group filtering and report organization, search and saved library views; legacy folder storage remains for compatibility.
- Committed revisions, restore as a new revision, duplicate, archive/recover and field-level comparison.
- Automatic/manual pending-editor checkpoints and recovery.
- Full supported-value NAVFIT import, single-report export, frozen baselines, three-way returned-file conflicts and retained native provenance. See [r27 scope and proof](planning/NAVFIT_EXCHANGE_R27.md).
- Attributed narrative snippets and a local draft completeness checklist.
- Revision-linked review comments, replies and resolve/reopen history.
- Local reviewer assignments, due dates, My Work, contributor switching and accept/return/cancel history.

Internal review acceptance is separate from official approval. Local contributor profiles provide attribution, not authentication. Use one contributor at a time in the same authoritative file; separate copies do not synchronize.

## Try the app

Download the **blank r113 workspace** from the [current Drive copy](https://drive.google.com/file/d/1oVLR6_3xvLSnIydOmRxwbXCWXhRdkG6s/view), make a working copy, and open it in trusted desktop Microsoft Access. Workspace setup opens automatically. Enter your command and active contributor, then create summary groups and reports. Settings & Help contains the workspace guide.

The release includes a separate blank workspace and synthetic demo. Use the demo for practice and the blank for a new workspace. Preserve existing workspaces and migrate their data into a separate blank file with an explicit `-TemplatePath`. Access Runtime and x86 environments have not been verified.

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
