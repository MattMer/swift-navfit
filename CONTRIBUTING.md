# Contributing to Swift NAVFIT

Start with [PROJECT_STATE.md](PROJECT_STATE.md) for current architecture and
support, [AGENTS.md](AGENTS.md) for task/data/release constraints, and
[Development_Workflow.md](docs/Development_Workflow.md) for environment and commands.
[BUILD_PIPELINE.md](planning/BUILD_PIPELINE.md) owns build order and the form workflow.
Historical release notes are evidence, not additional daily task requirements.

## Working on a change

Declare behavior, expected file scope, targeted checks and exclusions before
implementation. Edit tracked source/catalogues/generators; use disposable native
copies for diagnosis and bring useful edits back into source. Keep unrelated edits.
Do not merge binary databases as source or introduce another assembly path.

Use the standard development command and separate release command documented in
the workflow guide. Select checks below for the actual change; this table is not
an instruction to run all suites. Current native tests operate on disposable
copies. Missing external fixtures are an unverified case, not a passing test.

## Change-specific checks

All scripts below live in [scripts/](scripts/). Native tests generally take
`-BuildPath <fresh build>`; inspect parameters before running.

| Changed behavior | Focused checks / source contract |
| --- | --- |
| Any VBA module (`src/`, `tests/`) | check-vba-source.ps1 (a few seconds; no Access). It checks the module name, Option Explicit, ASCII-only text, the 1023-character line limit, the 24-continuation limit and duplicate procedures |
| Build graph/source import | test-source-contracts.ps1 (build plan, runner inventory, VBA normalization, field definitions); final installed-source/binding verification |
| Retired routes/folders | verify-build-artifact.ps1 -Tier extended (retired-route cleanup case); retain folder batch history and undo |
| Report navigation, pending values and lifecycle | test-work-review.ps1; test-report-pages.ps1; real Unload, failed saves, recovery and stale previews |
| Startup contributor attribution | test-workspace-startup.ps1 |
| Conversations/report tools | test-report-review-history.ps1; test-workspace-migration.ps1 -Cases test-report-message-import -BuildPath <blank> -DemoPath <demo> for transfer; no resolve/reopen requirement |
| Audit/conversation consolidation and routes | [test-report-review-history.ps1](scripts/test-report-review-history.ps1) -Tier extended: identity/correlation, outcome actor, undo/history, canonical adapters, private ownership, ancestry rejection, rollback, routes and full native form/macro binding inventory |
| Retained-schema audit/conversation import | [test-workspace-migration.ps1](scripts/test-workspace-migration.ps1) -BuildPath <fresh blank> -FixtureMapPath <JSON schema-to-path map>; actual retained schema files on disposable copies, Settings import, exact original fields, late rollback, private/reply preservation, source hashes and reopen. Do not relabel fixtures or execute historical source VBA. |
| My Work/help | test-work-review.ps1 -Tier extended; test-report-navigation.ps1 |
| Form controls, labels or the visual palette | test-workspace-presentation.ps1: 4.5:1 text contrast, a linked label on each input control, a name on each button. Every new input needs an accessible name: link its visible label (LabelName). Where there is none, such as a grid cell or a list under a heading, attach a 15x15-twip label named `an<Control>` whose caption is the spoken name. tests/accessibility-gaps.json is empty; adding to it needs a stated reason |
| Schema, field catalogue, validation rule packs or states, runner inventory | test-source-contracts.ps1 -BuildPath: the generated DoDAF views in docs/architecture/views must match the build. After an intended change, run generate-architecture-views.ps1 -BuildPath <fresh build> and commit the views |
| File importers (PDF, RUAD CSV, NAVFIT98A) | test-external-import.ps1: malformed and control inputs must be rejected cleanly or load, with no rows written, the source unchanged, state cleared and no temporary files. Add a case for any new parser limit or file format |
| Directory merging/history | test-directory-management.ps1 -Tier extended; applicable directory transfer/release checks |
| Traits and narratives | test-traits.ps1 (persistence, rollback, legacy restore, fresh-session recovery, exact catalogue text/source/hash and transfer, grid events, queue, summary, retry); test-narratives.ps1 (narrative snapshots, layout, NavyWriter); fixed-page/report-tools checks cover current UI |
| Service/counseling/readiness fields | test-report-field-persistence.ps1; src/service-reference.json suggestions require source fidelity, not policy-rule claims |
| Assignment/retention/member statement | test-report-field-persistence.ps1; persistence, raw recovery and native bounds |
| Mark signed / draft disposal | test-report-review-history.ps1 -Tier extended; lock and save refusal on every route, reopen prompt (editor and grid), bulk skip, disposal removes only drafts (signed revisions, conversations, groups kept), restore/compare refusal and fresh-session persistence |
| Counseling | test-member-profiles.ps1; schedule by paygrade and date, status, record rules, carry-over to blocks 30–31, My Work and history screens, disposition, member merge/undo and fresh-session persistence |
| Report people | test-member-profiles.ps1; all 13 fields, revisions, restore, rollback and fresh-process recovery |
| Structured names | test-member-profiles.ps1; ten parts/full names, punctuation/accents/case; no silent name normalization |
| Viewing area/identity | test-workspace-presentation.ps1; control geometry, active field/caret; not physical DPI acceptance |
| Schema/import engine | [test-workspace-migration.ps1](scripts/test-workspace-migration.ps1) -BuildPath <fresh blank> imports a seeded synthetic source for every schema from 1 to the build's, built from the per-schema files in [tests/historical-schemas](tests/historical-schemas/README.md), plus populated kept releases in `output/`; -Schemas '1-7,20' narrows a focused run. Use test-import-engine.ps1 -Tier release for engine verification as applicable; exact values, IDs/relationships, source hashes, rollback/retry and reopen |
| Canonical report definitions/revisions and typed persistence | test-source-contracts.ps1 (static; -BuildPath -Tier extended adds native catalogue recovery); test-recovery.ps1; test-workspace-migration.ps1 -BuildPath <fresh blank> -Cases test-persistence-transfer -LegacyTemplate <schema-27 blank> (manual case); exact raw/null/status/choice preservation, catalogue field extension, revision/cache ownership, constraints and rollback |
| Summary-group completion/RSCA | test-reporting-senior-accounting.ps1; reconciliation, correction, duplicate/deletion/senior guards, cancelled legacy state and idempotent XML backfill. The default full suite also retains historical UI assertions; report any separate preexisting UI failure explicitly |
| Provenance viewer | test-workspace-resources.ps1 on a current blank; self-contained synthetic receipts/raw values, filter isolation, Null/empty/long text and reopen; no external fixture required |
| NAVFIT/PDF exchange | Applicable adapter/group export/return/local-PDF suites and authorized synthetic/reference fixtures |

[tests/runners.json](tests/runners.json) lists every runner with its feature area
and tier. Runners in the `extended` tier take only `-BuildPath` (plus any listed
`extraArgs`); optional manual inputs are passed through and their cases print
`SKIP` when those inputs are absent. Core suites with `extendedArgs` also run
in their extended mode. See the [suite reference](docs/Development_Workflow.md#m08-groups-and-workflow-suites)
for parameters and cases. Run the suites for the area you changed, or the whole tier with
`scripts/develop.ps1 -Action ExtendedTest -BuildPath <fresh blank>` (weekly, and
from release-check). A new runner must be added there, or `test-source-contracts.ps1`
fails; a runner whose feature is removed is deleted with it.

`ExtendedTest` closes VBA error dialogs in the Access sessions a runner starts
(OK for compile errors, End for run-time errors), so a dialog fails the step
instead of hanging the tier; each runner is stopped after 45 minutes. A runner
that fails is retried once on a fresh copy. Passing on the retry is reported as
**FLAKY**, by name, with the first run's error, dialog text and Access processes,
in the output and the receipt; it is never a clean pass, so investigate it.
Failing twice fails the tier.

The trait suite was restored and modernized: it preserves current service,
rollback and recovery checks without retired checklist UI or a duplicated
catalogue loop. Current report-tools/narrative/cleanup harnesses cover long
messages, snippet pending/save behavior, failed folder batches and journal state.
The standalone review-comment suite, historical foundations runner and old tools
bundle were retired; git history keeps them. Their retired UI, resolve/reopen and
source-fixture assumptions are not current acceptance gates. Current adapter/engine suites retain
native import-capture coverage; this consolidation does not claim every legacy
API/fixture assertion was rerun. The provenance viewer owns a synthetic fixture.
Legacy benchmark fixtures may still be active test inputs; trace callers before removal.
Normal persistence fixtures may disable Unload in their disposable session for
abrupt-close assertions; lifecycle tests must exercise actual user exits.
Snapshot completeness uses the canonical field list. New Access test sessions
must explicitly select their synthetic contributor using the existing test
context helper. The report-field-persistence and member-profiles suites offer an explicit
`-PersistenceOnly` mode that skips their detail-panel bounds checks; the extended tier
runs them in full.

For schema changes, assess persistent-data impact explicitly and update build DDL,
version constants, metadata, migration and relevant table inventories together.
First capture the outgoing schema with `scripts/capture-schema-definition.ps1
-BuildPath <build of that schema>`. Otherwise the import runner stops with "no
definition for schema N".
Never test against the only copy of a workspace. Historical import support and
the r128 old-output retest waiver are distinguished in PROJECT_STATE.md; do not
silently narrow compatibility or recreate deleted fixtures.

## Review and delivery

Review text diffs and relevant check results. Use a focused branch/PR when requested
by the collaboration workflow; source publication, binary delivery and formal
GitHub releases are separate actions. An artifact receipt is not physical acceptance.
Follow AGENTS.md for preservation, Blank/Demo inventories, Drive verification
and normal Access cleanup.

Report defects with app/schema version, Access build/bitness, sanitized steps and
expected/actual behavior. Never attach personnel data, SSNs, populated user
databases, credentials or sensitive report text to issues.

Local approval is not an official signature. Draft saves may be incomplete.
Do not claim Navy validation, interchange coverage or printing beyond verified
evidence. The private repository has no selected open-source license; coordinate
use and redistribution with its owner.
