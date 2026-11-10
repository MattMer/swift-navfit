# Contributing to Swift NAVFIT

Start with [PROJECT_STATE.md](PROJECT_STATE.md) for current architecture and
support, [Development_Workflow.md](docs/Development_Workflow.md) for environment
and commands, and the [project rules](#sources-of-truth) below for task, data and
release constraints. The rules apply to human contributors and coding agents alike.
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
| Schema, field catalogue, validation rule packs or states, runner inventory | test-source-contracts.ps1 -BuildPath: the generated DoDAF views in docs/architecture/views and the generated facts in TECHNICAL_README.md must match the build, and every table must be classified in TECHNICAL_README.md section 5.2. After an intended change, run generate-architecture-views.ps1 -BuildPath <fresh build> and commit the views |
| Data held, imported or exported; security behaviour (access, attribution, audit, integrity, input handling, encryption, signing, dependencies); platform, interfaces, testing, release or known limitations | Update [TECHNICAL_README.md](TECHNICAL_README.md), the security and system description, in the same commit and add a line to its document history (section 17 lists what each section covers) |
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
| Schema/import engine | [test-workspace-migration.ps1](scripts/test-workspace-migration.ps1) -BuildPath <fresh blank> imports a seeded synthetic source for every schema that shipped in a release (listed in [release-schemas.json](tests/historical-schemas/release-schemas.json)) and the build's own, built from the per-schema files in [tests/historical-schemas](tests/historical-schemas/README.md), plus populated kept releases of those schemas in `output/`; -Schemas '35-41,44' narrows a focused run. Use test-import-engine.ps1 -Tier release for engine verification as applicable; exact values, IDs/relationships, source hashes, rollback/retry and reopen |
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

`develop.ps1` and `release-check.ps1` print a timestamped line as each step,
runner and suite case starts or finishes, with elapsed times, and an extended
runner that is quiet prints "Still running" every 5 minutes. Redirect the output
to a file and read it to see where a long run is.

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
Follow the project rules below for preservation, Blank/Demo inventories, Drive
verification and normal Access cleanup.

Report defects with app/schema version, Access build/bitness, sanitized steps and
expected/actual behavior. Never attach personnel data, SSNs, populated user
databases, credentials or sensitive report text to issues.

Local approval is not an official signature. Draft saves may be incomplete.
Do not claim Navy validation, interchange coverage or printing beyond verified
evidence. Swift NAVFIT is licensed under the Apache License 2.0 (LICENSE and
NOTICE); contributions are accepted under the same license.

## Sources of truth

- [PROJECT_STATE.md](PROJECT_STATE.md): current architecture, supported behavior,
  decisions and open gaps.
- [planning/PRIORITY_BACKLOG.md](planning/PRIORITY_BACKLOG.md): planned work.
- [docs/Development_Workflow.md](docs/Development_Workflow.md): commands and focused
  test runners. The table above maps each kind of change to its checks.
- [planning/BUILD_PIPELINE.md](planning/BUILD_PIPELINE.md): build order and the form workflow.
- Other dated `planning/` files, including `PROJECT_STATE_HISTORY.md`, are
  historical evidence, not current requirements.

If documents conflict in a way that affects the task, point out the conflict
and ask. The user's instructions for the current task take precedence over
plans, but they never silently override the product invariants below.

## Product invariants (non-negotiable)

- **Historical import:**
  - Every release must import data from every supported schema: schema 35
    (release 0.49.0, the first public release) through the current one.
    Release 0.55.0 is the last to import schemas 1–34; older workspaces go
    through a 0.55.0 Blank first.
  - Preserve rows, fields, IDs and relationships, revision/review history,
    provenance, recovery, command members, notes, defaults and reporting-senior
    history.
  - Use the production migration engine and the Settings & Help import. Read
    sources through read-only DAO without executing their code, and verify copied
    values before commit.
  - Never silently replace a populated workspace or merge records by ID.
- **Schema policy:**
  - Changes are additive and versioned.
  - The importer accepts source schemas 35 through the current `SCHEMA_VERSION`.
    When the schema changes, raise the hard-coded upper bound in `modMigration`
    alongside it, and add the outgoing release's schema to
    `tests/historical-schemas/release-schemas.json`. Raise the floor only on the
    owner's decision, keeping one release that bridges the dropped schemas.
  - The release gate tests every schema that shipped in a release plus the
    current one. Development-only schemas between releases stay accepted but
    are not retested.
  - Do not claim retesting you did not perform. PROJECT_STATE.md records which
    schemas were tested when, and the r128 old-output waiver.
- **Report editor:**
  - Keep the fixed Front/Back geometry in NAVPERS block order with the official
    captions.
  - All seven traits stay together on the editing Front Page; printed page
    assignments do not change.
  - Content stays to the right of the navigation rail.
  - Resizing changes the viewport, never the field layout.
  - Preserve the unbound buffer, pending values, saves, revisions and recovery
    across page changes.
  - Never claim printed fit without physical verification.
- **Protected state:**
  - Never modify user workspaces, earlier releases, supplied reference material
  (`FORMS_REFERENCES_TEMPLATES/`, `NAVFIT98A*`), or data/recovery resources.
    Work on copies.
  - Blank builds stay blank, and Demos contain only synthetic data. Never put
    personnel data in tracked fixtures or send it to external services.
  - Retained report history is never deleted. New-policy working states may be
    replaced atomically after a save, only when no review, signature, message,
    preview or undo reference needs them. Existing/imported history, narrative
    saves and workflow checkpoints are retained. A separate exception is draft disposal that the
    user confirms after Mark signed (BUPERSINST 1610.10H, enclosure (2)
    Introduction paragraph 7a). It keeps every signed revision, all
    conversations, reviews, group history and accounting, and records what it
    removed. Final reports are kept; the workspace is the command's reporting
    senior binder.

## Permissions

- **Allowed while implementing a requested change:**
  - the edits the task needs
  - small cleanup in files you are already changing
  - focused tests and fresh development builds
  - scratch files outside the repository
- **Ask first:**
  - refactors outside the task's files, or a materially larger file set than you
    described
  - deleting tracked files
  - moving or deleting untracked, ignored, binary or user files
  - new runtime dependencies or external services
  - security or Trust Center changes
  - commits and pushes, unless the user asked for them
- **Explicit request only:** releases, publishing Blank/Demo files, Drive
  uploads, signing.

Before a nontrivial change, state the intended behavior, the files involved and
the checks you will run. Keep it brief.

## How to work

- Edit authoritative tracked source: `src/` (modules, catalogues and the forms
  in `src/forms`) and the resource scripts listed in the manifest. Forms are
  source: edit their text, or change them in Access on a disposable build and
  re-export with `scripts/export-form-text.ps1`. A catalogue change includes the
  matching form edit. Change generators rather than generated output. Native Access experiments happen only on disposable copies, and useful
  results are ported back to source.
- Check `git status` before significant changes. Do not overwrite someone
  else's uncommitted work.
- Use whatever tool fits the job best. Don't add permanent scripts for one-off
  tasks, and don't add a second build or assembly path. Propose changes to the
  build contract (`develop.ps1`, `build-manifest.json`) before making them.
- Access code can be reached dynamically: control expressions (`=Fn()`), macros,
  `Eval`, form events and generators. Before calling code unused, check those as
  well as text references.

## Tests

- Removing a feature removes its tests and their harnesses in the same commit.
- Extend the existing test for a feature rather than adding a runner named
  after a change, milestone or release.
- Fixtures are generated from tracked definitions, or kept small and synthetic in
  the repository. Never depend on files outside it or on old outputs.
- [tests/runners.json](tests/runners.json) lists every runner with its area and
  tier. `scripts/cases/test-build-plan.ps1` (run by `test-source-contracts`) fails if a runner or harness is missing from it or
  orphaned. Run `scripts/develop.ps1 -Action ExtendedTest -BuildPath <blank>`
  weekly; release-check runs it too.

## Verification and reporting

- Run checks in proportion to the change:
  - the focused runner for the behavior you touched (see CONTRIBUTING.md)
  - before committing persistence, migration, generated-form or cross-module
    changes: `scripts/develop.ps1 -Action BuildAndTest`
  - release gates only for releases.
- Never claim coverage you did not run, and never weaken an assertion to make it
  pass. Diagnose a failure before adding defensive code, and report preexisting
  failures separately from regressions.
- At the end, report the changed files, behavior, tests run with their results,
  and anything left unverified.
- Update PROJECT_STATE.md only when behavior, architecture or a decision
  changes, and keep it current-state only. Verification detail (assertion counts,
  receipts) goes in the commit message, not the docs. Don't create per-session
  narrative documents.

## Access automation safety

- Allow one native controller (COM, MCP or person) per database at a time.
  Know which database you own and its PID. Before handing a file to another
  controller, release it.
- Close only Access sessions you started, through normal save and close. Never
  force-kill a session with unsaved work, and never pick processes by application
  name alone.
- Capture the primary error before cleanup. Allow one justified retry on a fresh
  disposable copy; after that, diagnose instead of retrying again.

## Cleanup

Deleting confirmed-obsolete tracked files is the normal path, and git history is
how you recover them. For uncertain, untracked or binary items, list each item's
path, reason and hash for the user and ask before deleting. Don't keep a holding
folder in the repository.

## Releases (explicit request only)

1. Set `APP_VERSION` to the release number and `APP_RELEASE_DATE` to the release
   date (yyyy-mm-dd) in `src/modSwift.bas`; after the release, return to the next
   `-dev` version with an empty `APP_RELEASE_DATE`. Then build a new Blank and a
   separate synthetic Demo.
2. Run `scripts/release-check.ps1 -BlankPath <blank> -DemoPath <demo>` plus the
   feature-specific, data-transfer, rollback and unsupported-source checks that
   apply.
3. Document which schema fixtures were tested.
4. Put dated, versioned files in `output/` and keep the prior outputs.
5. Upload to the owner's Drive folder, then verify the
   names, sizes and folder membership.
