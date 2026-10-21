# Contributing to Swift NAVFIT

Start with README.md, planning/ROADMAP.md and the current usage guides. planning/BUILD_STATUS.md retains the development and verification history; older sections describe superseded builds.

## Development environment

Use Windows with desktop Microsoft Access installed. Current verification uses Access 16.0, build 20430, through Access COM and DAO. Free Access Runtime is not a supported target; x86 has not been verified. GitHub-hosted runners do not have the required Access installation, so Access integration tests run locally.

The editable source of truth is src/, scripts/build-manifest.json and its generator scripts. See planning/BUILD_PIPELINE.md for build ownership, ordering, input fingerprints and final verification. Do not merge binary Access databases as source. Keep test fixtures synthetic and local. TEST_DB.accdb and NAVFIT_BLANK.accdb are intentionally excluded from the repository, along with supplied Navy reference PDFs, the NAVFIT98A application bundle and third-party converter files. Arrange those fixtures with the maintainer when working on import compatibility. Missing fixtures are not a passing compatibility test.

## Build and verify

Run from the repository root in PowerShell. Use a fresh output name; the builder refuses overwrites:

```powershell
powershell -NoProfile -File scripts/test-build-plan.ps1
powershell -NoProfile -File scripts/test-vba-source-normalization.ps1
powershell -NoProfile -File scripts/build.ps1 -OutputPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-build-cleanup.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/smoke-test.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-context-help.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-ux-organization.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-compat-pages.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-ux-workflows.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/build-ux-admin-demo.ps1 -TemplatePath output/Swift_NAVFIT_review.accdb -OutputPath output/Swift_NAVFIT_demo_review.accdb
powershell -NoProfile -File scripts/test-ux-demo.ps1 -BuildPath output/Swift_NAVFIT_demo_review.accdb
powershell -NoProfile -File scripts/test-report-tools.ps1 -BuildPath output/Swift_NAVFIT_demo_review.accdb
powershell -NoProfile -File scripts/test-ux-my-work.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-trait-catalogue.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-swift-data-import.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-workspace-upgrade.ps1 -TemplatePath output/Swift_NAVFIT_review.accdb -SourceWorkspacePath output/Swift_NAVFIT_demo_review.accdb
powershell -NoProfile -File scripts/test-report-message-import.ps1 -BlankPath output/Swift_NAVFIT_review.accdb -DemoPath output/Swift_NAVFIT_demo_review.accdb
powershell -NoProfile -File scripts/test-engine-verification.ps1 -BuildPath output/Swift_NAVFIT_review.accdb -SourcePath output/Swift_NAVFIT_demo_review.accdb
```

The older focused field/provenance harnesses below retain historical fixtures and can include superseded UI or snapshot-count assumptions. They are not current release acceptance gates. For historical provenance investigation with the authorized synthetic TEST_DB.accdb at the repository root:

```powershell
powershell -NoProfile -File scripts/test-provenance.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
```

That last script runs the workflow suite before the provenance suite. Other focused scripts cover narratives, revision comparison, recovery and import compatibility. Run checks appropriate to the change. Test scripts use disposable copies and may enable VBA for those copies; run only trusted source. Keep human workspaces closed during automation.

For schema changes, update the build DDL, SCHEMA_VERSION, src/modMigration.bas and the upgrade script's table inventory. Verify populated old-to-new and new-to-new migrations with every original row/field compared, source hashes unchanged and successful reopen. Never test upgrades against the sole copy of a real workspace.

## Collaboration workflow

For native report lifecycle, bulk-preview or review-navigation changes, run
`scripts/test-ux-workflows.ps1 -BuildPath <fresh blank build>`. The harness uses a
scratch copy and scripted dialog choices to exercise the actual Unload event,
failed saves, preview invalidation and report-specific review context. Normal
non-lifecycle persistence fixtures deliberately disable Unload only in their
disposable editor session to retain their original abrupt-close semantics and
checkpoint assertions. The workflow harness exercises real user exits.

Create an issue for a feature or defect, then a short feature branch. Submit a pull request describing the behavior, scope and validation. Keep one feature per pull request where practical. Review text source changes; share a fresh blank build separately if visual review is needed.

Report defects with app/schema version, Access build/bitness, steps, expected/actual behavior and sanitized screenshots. Do not attach personnel data, SSNs, populated databases, credentials or sensitive report text to issues.

Preserve committed revisions and audit history. Official narrative, independent report notes/replies and review assignments have distinct purposes. Historical review-comment and integrated-tools harnesses include superseded standalone UI and are not current acceptance suites; current report-tools tests cover the unified note system, and build-cleanup tests cover retained folder batch history and undo. Draft saves can be incomplete. Internal acceptance is not official approval. Do not claim Navy validation or NAVFIT compatibility without the specific verified evidence.

## Releases

Publish only a verified Blank distribution and a clearly marked synthetic Demo. Check all Blank user-data tables are empty and all Demo data is synthetic; record artifact hashes and verification evidence in output manifests. Preserve earlier files and publish the current pair to the configured Drive folder after migration gates pass. Never upload a human working workspace or an unverified candidate as a release. GitHub source publication and a formal GitHub release are separate actions. Keep future source changes in Git even when a release binary is available.

No open-source license has been selected. The repository is private; coordinate permitted use and redistribution with the owner.

The curated src/trait-catalogue.json is embedded by scripts/write-trait-catalogue.ps1. Preserve its source version/hash and exact text; test-trait-catalogue.ps1 verifies 169 catalogue values. The older combined TraitsHarness has historical snapshot-count and standalone-checklist expectations; current trait behavior is covered by fixed-page, report-tools and workflow suites. Descriptions are reference guidance; they do not implement Navy validation.

## Counseling, readiness and billet checks

Run scripts/test-service-details.ps1 with -BuildPath pointing to the freshly built database. src/service-reference.json contains curated suggestions from the local NAVFIT reference database; changes require source attribution and reference fidelity checks. Suggestions are not policy rules.

## Assignment recommendations, retention and member statement

Run scripts/test-decision-details.ps1 with -BuildPath pointing to the fresh database. It checks behavioral integration, raw recovery in a new Access process, the new editor button for overlap and both forms for control bounds. Each source checkbox preserves Yes, No or blank independently. Applicable Navy rules and signatures require separate implementation.

## Reporting senior identifiers, contact and raters

Run scripts/test-report-people.ps1 with -BuildPath pointing to the fresh database. It checks all 13 fields through persistence, history, duplication/restore, rollback and raw recovery in a new Access process, plus button overlap and form bounds. Source capacities are documented in planning/REPORT_PEOPLE_GUIDE.md; signatures, authority, policy and external date conversion require separate work.

## Structured member, senior and counselor names

Run scripts/test-party-names.ps1 with -BuildPath pointing to the fresh database. It verifies ten name parts, preservation of full names, punctuation/accents/case, rollback, raw recovery in a new Access process, button overlap and form bounds. Capacities and source fields are documented in planning/STRUCTURED_NAMES_GUIDE.md. Do not silently split, normalize or compose names.

Responsive layout and report identity regressions run against a scratch copy:

```powershell
powershell.exe -NoProfile -File scripts/test-ux-responsive-identity.ps1 -BuildPath build/your-template.accdb
```

This verifies native control geometry and active-field/caret retention, not
physical DPI or screen-reader acceptance. The preview builder includes this gate.
