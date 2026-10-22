# Source-first development setup

21 October 2026. Application remains r127 / 0.27.0-dev / schema 27. This is
development tooling and verification work, not a new application release.

## Changes

- Installed Microsoft PowerShell 2025.4.0 and serkonda7 VBA 1.2.1. The user
  installed Access Explorer 1.2.0 during this work. Workspace recommendations
  include those extensions and the existing GitHub Pull Requests extension.
- Added VS Code tasks for Preflight, default fresh build plus focused tests,
  build-only, focused tests on a selected blank and release verification.
  Tasks invoke Windows PowerShell as a process with separate arguments.
- Added `scripts/develop.ps1`, `scripts/release-check.ps1` and shared logging.
  Commands stop on failed children, preserve diagnostic logs and write PASS
  receipts only after all steps and input/source-hash checks pass.
- Added read-only original-artifact inspection and disposable-copy compile,
  installed-source, binding and blank-data checks. Release verification also
  requires complete coverage from schema 1 through the canonical current schema.
- Removed stale app/schema expectations from build-plan tests. The report-tools
  suite now has an explicit synthetic-seed option for its own blank scratch copy;
  its existing Demo mode remains intact. Corrected the demo builder's stale
  reopen counts from 9 reports/3 groups to 13 reports/5 groups.
- Updated agent guidance, README, contributing and build documentation. See
  [Development workflow](../docs/Development_Workflow.md) for commands.

## Fresh generation and RG01

The full `scripts/build.ps1` created
`build/setup-fresh-r127-20261021-01.accdb` directly from source with no earlier
binary as an assembly template. All 71 stages passed, including report tools,
followed by compile/reopen, 99 static-module comparisons and bindings across
66 forms. The receipt's input hashes still match the final declared build inputs.

The earlier r126 report-tools closed-object error did not reproduce. Current
fresh-generation proof is now available; no specific historical root cause or
targeted bug fix is claimed. Keep the older failed attempt as historical
evidence, and capture stage/stack details if the error returns.

The development test command passed report tools, smoke, 264 fixed-page/buffer
assertions, 41 lifecycle assertions, forced-shutdown recovery and 47 merge
checks. Native release verification also exercises all 27 retained schemas,
populated schema-20/current transfers, merge history/undo, exact values and
rollback/rejection guards against the fresh blank. Exact receipts, fixture
paths/counts and final results are in
[verification](DEVELOPMENT_SETUP_2026-10-21_VERIFICATION.json).

Oldest retained fixtures are blank distributions. They do not establish
populated historical workflow coverage beyond the separately populated cases.
Physical display, accessibility and printed fit remain deferred.

## Scope and preservation

Current published Blank/Demo names, sizes and hashes are checked against their
existing receipt. Existing workspaces, earlier releases and unrelated FORMS
files are preserved. Generated development databases remain under `build`;
no new release or Drive upload is part of this setup change. Close owned Access
automation sessions after all checks.

Access Explorer's installed commands support interactive investigation on
scratch copies. Its MCP backend is not exposed to this agent session; installing
the extension is not proof of an operational agent connection. The standard
build/test commands remain the reproducible path.

C03 startup contributor selection and C04 validation applicability remain the
next critical implementation items.

Follow-up: the user subsequently requested the MCP connection. The backend is
now installed, registered and protocol-tested; see
[Access Explorer MCP setup](ACCESS_MCP_SETUP_2026-10-21.md). The earlier paragraph
records the state before that separate connection task.
