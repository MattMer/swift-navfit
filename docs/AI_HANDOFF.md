# Shared AI handoff

Updated 30 October 2026. Read this at startup, then verify Git status/diffs,
versions and artifacts. Requirements remain in [PROJECT_STATE](../PROJECT_STATE.md),
[backlog](../planning/PRIORITY_BACKLOG.md), [AGENTS](../AGENTS.md) and the
[development workflow](Development_Workflow.md). Private conversation history is
not needed. Commits need the user's approval; pushes have not been requested.

## State

- Branch main, ahead of origin/main (not pushed). APP_VERSION 0.49.0 /
  SCHEMA_VERSION 35 ([modSwift](../src/modSwift.bas)).
- Claude and Codex both work here. Claude Code has the Codex plugin
  (openai/codex-plugin-cc, ChatGPT sign-in). Claude delegates contained pieces
  (reviews, restructuring, one-module fixes) to Codex and runs native Access
  checks, stages and commits itself: Codex's sandbox cannot write .git, and it
  does not launch Access. `develop.ps1` aborts if a build input changes mid-build,
  so don't edit source while a build runs.

## Done in development

- M04 (error diagnostics): primary errors survive cleanup across the editor,
  pages, scope, signoff, trait grid and workspace import. Remaining suppression
  sites are optional probes; review one only after reproducing a failure.
  Open note: an undeletable trait cache file stays in TEMP after the next grid
  load (it now has a diagnostic).
- M08 (test suites): 33 registry entries, no area above five. Suites are top-level
  `scripts/test-*.ps1`; their case scripts live in `scripts/cases` and run through
  `scripts/suite-common.ps1`. Core suites take `-Tier extended` via registry
  `extendedArgs`. Cases needing other inputs print SKIP without them.

## Latest verification

Build build/dev/Swift_NAVFIT_0.49.0_20261030-101953_c461ec65.accdb: BuildAndTest
PASS; ExtendedTest PASS with one FLAKY runner, receipt
build/workflow/20261030-102446-2b9da61c-48c4-4952-8f4e-da39cc5e31e6/verification.json.
test-import-engine's first run hit error 3059 ("Operation canceled by user") in
its 16-minute old-VBA benchmark; the retry passed. Not yet diagnosed.
Not run: release gates, NAVFIT98A, physical printing. No Access sessions open.

## M03 revisit (30 October)

PDF packets: each page now resets only the slots the previous page showed
(f3a7e20; 250 members 65 s to 46 s, pages identical). Editor open profiled at about
1.1 s on the 4,700-report workspace: mostly per-control layout writes in
UXRenderFixedSheet (modReportPages). Scratch profilers (not in the repository)
loaded an instrumented modReportPages copy into a disposable build copy.

## Exact next actions

1. Editor page-layout cache is done (b5e7e54); editor open stays about 1.1 s by
   user decision. Open lead: the fourth consecutive editor open took 12-13 s in
   both profiles.
2. L03 workspace size headroom only when the user schedules it.
3. Optional: diagnose the test-import-engine 3059 flake.
4. Update this file at milestones and before stopping.
