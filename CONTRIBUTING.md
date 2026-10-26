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
| Build graph/source import | test-build-plan.ps1; test-vba-source-normalization.ps1; final installed-source/binding verification |
| Retired routes/folders | test-build-cleanup.ps1; retain folder batch history and undo |
| Report navigation, pending values and lifecycle | test-ux-workflows.ps1; test-compat-pages.ps1; real Unload, failed saves, recovery and stale previews |
| Startup contributor attribution | test-contributor-startup.ps1 |
| Conversations/report tools | test-report-tools.ps1; test-report-message-import.ps1 for transfer; no resolve/reopen requirement |
| Audit/conversation consolidation and routes | [test-audit-conversations-navigation.ps1](scripts/test-audit-conversations-navigation.ps1): identity/correlation, outcome actor, undo/history, canonical adapters, private ownership, ancestry rejection, rollback, routes and full native form/macro binding inventory |
| Retained-schema audit/conversation import | [test-audit-message-transfer.ps1](scripts/test-audit-message-transfer.ps1) -BuildPath <fresh blank> -FixtureMapPath <JSON schema-to-path map>; actual retained schema files on disposable copies, Settings import, exact original fields, late rollback, private/reply preservation, source hashes and reopen. Do not relabel fixtures or execute historical source VBA. |
| My Work/help | test-ux-my-work.ps1; test-context-help.ps1 |
| Directory merging/history | test-directory-merge.ps1; applicable directory transfer/release checks |
| Traits | test-traits.ps1 for persistence, rollback, legacy restore and fresh-session recovery; test-trait-catalogue.ps1 for exact text/source version/hash; fixed-page/report-tools checks cover current UI |
| Service/counseling/readiness fields | test-service-details.ps1; src/service-reference.json suggestions require source fidelity, not policy-rule claims |
| Assignment/retention/member statement | test-decision-details.ps1; persistence, raw recovery and native bounds |
| Mark signed / draft disposal | test-signoff.ps1; lock and save refusal on every route, reopen prompt (editor and grid), bulk skip, disposal removes only drafts (signed revisions, conversations, groups kept), restore/compare refusal and fresh-session persistence |
| Counseling | test-counseling.ps1; schedule by paygrade and date, status, record rules, carry-over to blocks 30–31, My Work and history screens, disposition, member merge/undo and fresh-session persistence |
| Report people | test-report-people.ps1; all 13 fields, revisions, restore, rollback and fresh-process recovery |
| Structured names | test-party-names.ps1; ten parts/full names, punctuation/accents/case; no silent name normalization |
| Viewing area/identity | test-ux-responsive-identity.ps1; control geometry, active field/caret; not physical DPI acceptance |
| Schema/import engine | [test-swift-data-import.ps1](scripts/test-swift-data-import.ps1) -BuildPath <fresh blank> imports a seeded synthetic source for every schema from 1 to the build's, built from the per-schema files in [tests/historical-schemas](tests/historical-schemas/README.md), plus populated kept releases in `output/`; -Schemas '1-7,20' narrows a focused run. Add test-workspace-upgrade.ps1 and test-engine-verification.ps1 as applicable; exact values, IDs/relationships, source hashes, rollback/retry and reopen |
| Canonical report definitions/revisions and typed persistence | test-field-definitions.ps1 (also with -BuildPath); test-persistence.ps1; test-persistence-transfer.ps1 -BuildPath <fresh blank> -LegacyTemplate <real schema-27 blank>; exact raw/null/status/choice preservation, catalogue field extension, revision/cache ownership, constraints and rollback |
| Summary-group completion/RSCA | test-rs-completion.ps1 -AccountingOnly; reconciliation, correction, duplicate/deletion/senior guards, cancelled legacy state and idempotent XML backfill. The default full suite also retains historical UI assertions; report any separate preexisting UI failure explicitly |
| Provenance viewer | test-provenance.ps1 on a current blank; self-contained synthetic receipts/raw values, filter isolation, Null/empty/long text and reopen; no external fixture required |
| NAVFIT/PDF exchange | Applicable adapter/group export/return/local-PDF suites and authorized synthetic/reference fixtures |

[tests/runners.json](tests/runners.json) lists every runner with its feature area
and tier. Runners in the `extended` tier take only `-BuildPath` (plus any listed
`extraArgs`); run the ones for the area you changed, or the whole tier with
`scripts/develop.ps1 -Action ExtendedTest -BuildPath <fresh blank>` (weekly, and
from release-check). A new runner must be added there, or `test-build-plan.ps1`
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
bundle are held in [PROPOSED_PURGE](PROPOSED_PURGE/README.md). Their retired UI,
resolve/reopen and source-fixture assumptions are not current acceptance gates;
the originals remain available as evidence. Current adapter/engine suites retain
native import-capture coverage; this consolidation does not claim every legacy
API/fixture assertion was rerun. The provenance viewer owns a synthetic fixture.
Legacy benchmark fixtures may still be active test inputs; trace callers before removal.
Normal persistence fixtures may disable Unload in their disposable session for
abrupt-close assertions; lifecycle tests must exercise actual user exits.
Snapshot completeness uses the canonical field list. New Access test sessions
must explicitly select their synthetic contributor using the existing test
context helper. The decision/report-people/party-name runners offer an explicit
`-PersistenceOnly` mode for data and recovery checks without their retained legacy
button-layout assertions; current fixed-page acceptance remains a separate gate.

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
