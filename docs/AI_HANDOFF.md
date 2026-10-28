# Shared AI handoff

Updated 28 October 2026. Read this at startup, then verify Git status/diffs,
versions and artifacts. Requirements remain in [PROJECT_STATE](../PROJECT_STATE.md),
[backlog](../planning/PRIORITY_BACKLOG.md), [AGENTS](../AGENTS.md) and the
[development workflow](Development_Workflow.md). Private conversation history is
not needed. Commits need the user's approval; pushes have not been requested.

## State

- Branch main, ahead of origin/main by the M04 commits (not pushed).
  APP_VERSION 0.49.0 / SCHEMA_VERSION 35 ([modSwift](../src/modSwift.bas)).
- Claude and Codex both work here. Claude Code has the Codex plugin
  (openai/codex-plugin-cc, codex-cli 0.165.0, ChatGPT sign-in). Claude delegates
  contained pieces (reviews, one-module fixes) to Codex and runs native Access
  checks itself. Codex works without launching Access while Claude runs builds;
  `develop.ps1` aborts if a build input changes mid-build.

## M04: done in development

Primary errors keep number, description, source, stage and transaction outcome
through cleanup in report pages, surface refresh/restore, editor save, signoff
refresh, contributor switches, group open, shared-field scope, trait-grid
load/close and workspace-import preview/progress/display. See the
[M04 audit](../planning/CONSOLIDATION_PLAN.md#m04m06-source-and-consumer-audit).
The remaining suppression inventory (written by scripts/cases/test-diagnostics-contracts.ps1) is
optional probes and best-effort UI restoration; review another site only after
reproducing a specific failure.

Open note: when UXGridClose cannot delete the trait cache, the next UXGridLoad
creates a new cache path, so the old file stays in TEMP (now with a diagnostic).

## Latest verification

Build build/dev/Swift_NAVFIT_0.49.0_20261028-100430_af491202.accdb:
BuildAndTest PASS; pending-cells 45, wizard-state (34 migration/rollback/retry),
diagnostics 109, swift-data-import (schemas 1-35 plus 7 kept releases),
bulk-workspace-import and engine-import PASS. Earlier the same day: recursive
defaults 205 and signoff 63 PASS. Not run: extended tier, release gates,
NAVFIT98A, physical printing. No Access sessions are left open.

M08 baseline: ExtendedTest PASS, 100 steps, on the same build. Receipt
build/workflow/20261028-105117-667ab856-6f5e-457c-bb0e-81e4fa4b09e7/verification.json.
M08 consolidation and runner deletion are user-approved (28 Oct). Done, one
commit per area: traits and narratives (test-traits, test-narratives); backup and
recovery (test-recovery `-BuildPath [-InstallSource] [-CompiledPath]`). Suites use
scripts/suite-common.ps1: every case runs on its own copy and any failure fails
the suite. Codex can't write .git (index.lock) from its sandbox; Claude stages
and commits.

## M08 continuation (28 October)

Workspace/settings: 19 runners -> four workspace suites. Report editor/pages:
22 runners -> four report suites; test-contributor-context stays a top-level
helper. Case scripts moved to scripts/cases with only script-root substitutions.
Registry, develop/release/directory/demo callers and current docs updated.
ExtendedTest now passes registry extendedArgs; review-history includes -SeedDemo.
Startup core needs a blank; extended demo coverage needs optional -DemoPath
(populated synthetic Demo), otherwise prints SKIP. Stale batched-reads remains
unchanged, manual only with -IncludeStaleBatchedReads and -BaselinePath.

Checks: static test-source-contracts PASS (all three cases), git diff --check
PASS; all 41 moved bodies match HEAD apart from script-root paths; 11 stubbed
suite dispatch checks PASS. No native verification performed. Codex launched no
Access, owns no database/PID, and has nothing to close. git mv was denied by the
read-only index; plain moves used. Nothing staged, committed or pushed.

## Exact next actions

1. User runs native suites on disposable artifacts, including startup with
   `powershell.exe -NoProfile -File scripts/test-workspace-startup.ps1 -BuildPath <blank> -Tier extended -DemoPath <demo>`.
   Run `powershell.exe -NoProfile -File scripts/develop.ps1 -Action ExtendedTest -BuildPath <blank>`
   for scheduled coverage; the optional demo and stale cases are explicitly skipped
   without their inputs. Do not interpret static checks as native coverage.
2. Review/commit workspace and report areas separately (shared support files need
   hunk staging); then continue remaining M08 areas in backlog order.
