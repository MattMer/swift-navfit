# Contributing to Swift NAVFIT

Start with README.md, planning/ROADMAP.md and the current usage guides. planning/BUILD_STATUS.md retains the development and verification history; older sections describe superseded builds.

## Development environment

Use Windows with desktop Microsoft Access installed. Current verification uses Access 16.0, build 20430, through Access COM and DAO. Access Runtime and x86 have not been verified. GitHub-hosted runners do not have the required Access installation, so Access integration tests run locally.

The editable source of truth is src/*.bas and scripts/build.ps1. Do not merge binary Access databases as source. Keep test fixtures synthetic and local. TEST_DB.accdb and NAVFIT_BLANK.accdb are intentionally excluded from the repository, along with supplied Navy reference PDFs, the NAVFIT98A application bundle and third-party converter files. Arrange those fixtures with the maintainer when working on import compatibility. Missing fixtures are not a passing compatibility test.

## Build and verify

Run from the repository root in PowerShell. Use a fresh output name; the builder refuses overwrites:

```powershell
powershell -NoProfile -File scripts/build.ps1 -OutputPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/smoke-test.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-member-details.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-report-context.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-reporting-senior.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-promotion-status.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-assessment.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-trait-summary.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-traits.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-administration.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-review-comments.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/test-routing.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
```

With the authorized synthetic TEST_DB.accdb at the repository root:

```powershell
powershell -NoProfile -File scripts/test-provenance.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
```

That last script runs the workflow suite before the provenance suite. Other focused scripts cover narratives, revision comparison, recovery and import compatibility. Run checks appropriate to the change. Test scripts use disposable copies and may enable VBA for those copies; run only trusted source. Keep human workspaces closed during automation.

For schema changes, update the build DDL, SCHEMA_VERSION, tests/modMigration.bas and the upgrade script's table inventory. Verify populated old-to-new and new-to-new migrations with every original row/field compared, source hashes unchanged and successful reopen. Never test upgrades against the sole copy of a real workspace.

## Collaboration workflow

Create an issue for a feature or defect, then a short feature branch. Submit a pull request describing the behavior, scope and validation. Keep one feature per pull request where practical. Review text source changes; share a fresh blank build separately if visual review is needed.

Report defects with app/schema version, Access build/bitness, steps, expected/actual behavior and sanitized screenshots. Do not attach personnel data, SSNs, populated databases, credentials or sensitive report text to issues.

Preserve committed revisions and audit history. Official narrative, local notes, review comments and assignments have distinct purposes. Draft saves can be incomplete. Internal acceptance is not official approval. Do not claim Navy validation or NAVFIT compatibility without the specific verified evidence.

## Releases

Publish only a verified blank distribution. Check all user-data tables are empty, record its SHA256 in output/MANIFEST.json, tag the source commit and attach the blank .accdb plus manifest to a GitHub release. Never upload a candidate, working workspace or test database as a release. Keep future source changes in Git even when a release binary is available.

No open-source license has been selected. The repository is private; coordinate permitted use and redistribution with the owner.

The curated src/trait-catalogue.json is embedded by scripts/write-trait-catalogue.ps1. Preserve its source version/hash and exact text; test-traits.ps1 verifies 169 catalogue values. Descriptions are reference guidance; they do not implement Navy validation.

## Counseling, readiness and billet checks

Run scripts/test-service-details.ps1 with -BuildPath pointing to the freshly built database. src/service-reference.json contains curated suggestions from the local NAVFIT reference database; changes require source attribution and reference fidelity checks. Suggestions are not policy rules.

## Assignment recommendations, retention and member statement

Run scripts/test-decision-details.ps1 with -BuildPath pointing to the fresh database. It checks behavioral integration, raw recovery in a new Access process, the new editor button for overlap and both forms for control bounds. Each source checkbox preserves Yes, No or blank independently. Applicable Navy rules and signatures require separate implementation.

## Reporting senior identifiers, contact and raters

Run scripts/test-report-people.ps1 with -BuildPath pointing to the fresh database. It checks all 13 fields through persistence, history, duplication/restore, rollback and raw recovery in a new Access process, plus button overlap and form bounds. Source capacities are documented in planning/REPORT_PEOPLE_GUIDE.md; signatures, authority, policy and external date conversion require separate work.

## Structured member, senior and counselor names

Run scripts/test-party-names.ps1 with -BuildPath pointing to the fresh database. It verifies ten name parts, preservation of full names, punctuation/accents/case, rollback, raw recovery in a new Access process, button overlap and form bounds. Capacities and source fields are documented in planning/STRUCTURED_NAMES_GUIDE.md. Do not silently split, normalize or compose names.
