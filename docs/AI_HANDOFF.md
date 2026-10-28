# Shared AI handoff

Updated 28 October 2026. Read this at startup, then verify Git status/diffs,
versions and artifacts. Requirements remain in [PROJECT_STATE](../PROJECT_STATE.md),
[backlog](../planning/PRIORITY_BACKLOG.md), [AGENTS](../AGENTS.md) and the
[development workflow](Development_Workflow.md). Private conversation history is
not needed. Preserve all existing changes; no commit/push/discard is authorized.

## Objective and acceptance

Continue M04: preserve primary error identity, source, stage and transaction
outcome through cleanup; narrow suppression; show redacted feedback; prove
failure/retry and saved/pending-data preservation in native Access. Completed
handlers below must not be repeated. Remaining suppression/code-identity review
is open; do not treat inventory counts as defects. M03's remaining editor-open/PDF
performance work is deferred by user decision; M11 remains planning only.

## Repository and authority

- Branch main; HEAD and local origin/main remain
  116e84ada483189df7f8596eaccb65e2c8f3a8ab. No remote fetch. Release commit
  9577903 is documented as 0.49.0; no release/publication in this work.
- Latest recovered Codex-associated checkpoint: 7c1655b (24 Oct), linked by its
  Git message to session af502e4e-7938-4860-a993-b7b027140c60. Its tree equals
  fbfd64f: 0.44.0-dev/schema 33. The registry returns no transcript; do not infer
  authorship from commit names. Subsequent history covers forms as source,
  counseling/signoff (schemas 34/35), accessibility/M09, architecture/security,
  0.49.0 release, XML hardening and M03 save/list/grid performance.
- Current APP_VERSION 0.49.0 / SCHEMA_VERSION 35:
  [modSwift](../src/modSwift.bas). Schema: [major-schema](../src/major-schema.json).
  Report catalogue version 2: [field definitions](../src/field-definitions.json).
- Authoritative source: [src](../src/), including [forms](../src/forms/), and
  resource scripts in the [manifest](../scripts/build-manifest.json). Use the
  existing [pipeline](../planning/BUILD_PIPELINE.md); databases are ignored
  artifacts, never mergeable source. Checkout: D:/Development/Swift NAVFIT.

## Completed changes and preserved working tree

The approved local Claude transcript established the M04 authorization and
usage-limit interruption. Preserve its seven unstaged source edits: modReportPages,
modSignoff, modUX, modUXFeedback, modUXGroupPages, modUXHeaderReview and
modUXResourceTools. They captured errors before cleanup/probes and exposed page,
review and committed-signoff refresh failures. The first Codex continuation
completed Boolean failure propagation and native failure/retry coverage.

The latest continuation reproduced a calculated surface-control assignment failure
on the previous verified blank, then fixed modUXSurface refresh/restore diagnostics
and their modUX/modReportPages callers. Failures retain native error/source/field
stage, keep saved/pending data, release the busy flag and allow retry. Optional
active-control lookup stays narrowly suppressed. DiagnosticsContractsHarness now
covers pending text, visible feedback, page propagation, painting and retry.

Claude continuation (28 Oct, afternoon): a Codex adversarial review (via the
Claude Code Codex plugin) of the diff above found two gaps, both fixed in modUX:
UXSaveEditor now reports a post-save surface refresh failure as Committed with
"Saved; display refresh failed" instead of the unsaved message, and UXOpenGroup
keeps the original source and records a failed Echo restore as cleanup. Also
fixed UXSharedScope (modUXRecursiveDefaults): it keeps the native failure/stage
in the scope tooltip, closes its recordset through SwiftClose, and cannot raise
into UXConfirmBlockChange after a commit (which would have shown "Not saved").
Tests: DiagnosticsContractsHarness CheckSaveRefreshFailure; RecursiveDefaultsHarness
scope failure/retry (an unknown name in the member list; Access accepts a
trailing comma in IN lists, so that input does not fail).

Other unstaged edits are AGENTS.md, CLAUDE.md, PROJECT_STATE.md, the backlog,
CONSOLIDATION_PLAN.md and the diagnostics harness; this handoff is untracked.
No files are staged, committed or discarded. No schema, field geometry or build
contract change. Never store credentials, personnel or report content here.

## Artifacts and controller state

- Current verified blank:
  build/dev/Swift_NAVFIT_0.49.0_20261028-094214_08ce0af7.accdb (includes the
  afternoon fixes). Runners used disposable copies. No authoritative user
  database was identified/modified.
- Earlier blanks: build/dev/Swift_NAVFIT_0.49.0_20261028-091340_2151d832.accdb
  (morning surface fixes, before the afternoon fixes) and
  build/dev/Swift_NAVFIT_0.49.0_20261028-085514_ba5428c8.accdb (before surface fixes).
  build/dev/grid-list2.accdb predates the seven original M04 edits.
  build/dev/editor-hidden.accdb also has reverted experimental form state; do
  not use it as a baseline. Existing artifacts were preserved.
- Desktop runners closed their owned Access sessions normally. No live owned
  controller/database to resume; sandbox process inspection sees no Access
  processes. Establish database/PID ownership before starting new native work.

## Latest checks and limitations

- VBA static checker: 222 files, zero errors, ten existing mixed-line-ending
  warnings. Editor Problems and git diff --check pass.
- Afternoon: VBA checker 0 errors; BuildAndTest PASS, receipt
  build/workflow/20261028-094214-2780d9a5-e3e9-4175-acdf-e4b3e79b2d6b/verification.json;
  diagnostics 109 PASS (build/tests/diagnostics-contracts-fa9f3552-4c2f-4ad3-8d69-496972ee57d5);
  recursive defaults 205 PASS; signoff 63 PASS plus fresh-session persistence.
- Morning BuildAndTest: all 12 steps PASS, no flaky result. Receipt:
  build/workflow/20261028-091340-fa362fc3-96a7-4284-b00b-f76026fe4914/verification.json.
  Includes compile/reopen/source/bindings/blank, report tools, contributor startup,
  smoke, fixed pages/buffers, workflow/recovery, directory merge and architecture.
- Focused diagnostics: 103 assertions PASS (includes the earlier postcommit
  signing/contributor tests). Receipt and source inventories:
  build/tests/diagnostics-contracts-d7eb045a-8c01-40fb-95a7-e6bc24be239f/verification.json.
  Current inventory: 51 suppression sites, 933 raise sites; not defect counts.
- Signoff: 63 assertions PASS plus fresh-session signed/history/disposal/value
  persistence; build/tests/signoff-results.json.
- Unified exec has failed before startup with helper_unknown_error. Node child
  processes work for source checks; sandbox Access COM fails with 0x80070520.
  Use normal desktop Windows PowerShell 5.1 / VS Code tasks for native runners.
  A reused VS Code task can retain its old resolved arguments: inspect the actual
  executed command. A newly named temporary task ran the current artifact. All
  temporary entries were removed by restoring the exact original tasks.json.
- No NAVFIT98A execution, full historical-schema rerun, release/extended tier,
  physical printing, screen-reader or managed-workstation acceptance in this work.
  Preserve PROJECT_STATE's separate limitations and M08's stale-runner findings.

## Exact next actions

1. Review the current diff and the suppression-inventory.csv/error-identities.csv
   beside the diagnostics receipt. UXSharedScope is done; continue with
   UXGridClose's derived-cache cleanup and importer display cleanup.
   Reproduce a specific failure before changing behavior; classify legitimate
   optional probes separately. See the [M04 audit](../planning/CONSOLIDATION_PLAN.md#m04m06-source-and-consumer-audit).
2. Extend existing feature harnesses; preserve primary/cleanup distinctions and
   do not renumber error codes. After relevant changes, run from the repo root:

   ~~~powershell
   powershell.exe -NoProfile -File scripts/check-vba-source.ps1
   powershell.exe -NoProfile -File scripts/develop.ps1 -Action BuildAndTest
   powershell.exe -NoProfile -File scripts/test-diagnostics-contracts.ps1 -BuildPath "<fresh blank>"
   powershell.exe -NoProfile -File scripts/test-signoff.ps1 -BuildPath "<fresh blank>"
   git diff --check
   ~~~

   Use the new build path, never a populated workspace. Do not use diagnostics
   -InstallSource on a current build: its legacy mode rewrites schema metadata.
3. Update this same handoff at milestones/before stopping. Capture primary errors
   before cleanup, allow one justified fresh-copy retry, and close only owned
   Access sessions normally. No release, publication or security change implied.
