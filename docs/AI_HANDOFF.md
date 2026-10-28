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
The remaining suppression inventory (written by test-diagnostics-contracts) is
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

## Exact next actions

1. M08: consolidate runners per [tests/runners.json](../tests/runners.json) and the
   backlog entry, starting with the stale `pending` runners it names.
2. Update this file at milestones and before stopping.
