# Swift NAVFIT: agent and contributor rules

These rules apply to every coding agent (Claude Code, Codex or others) and to
human contributors. Tool-specific notes belong in that tool's own file, such as
[CLAUDE.md](CLAUDE.md).

## Shared assistant handoff

At the start of every continuation, read [docs/AI_HANDOFF.md](docs/AI_HANDOFF.md).
Verify it against Git status/diffs, current versions and database artifacts before
acting; it is a resumption record, not a replacement for product requirements.
Preserve all existing work, regardless of assistant, and avoid repeating completed
tasks. Update the same handoff after meaningful milestones and before stopping:
record unfinished work, exact next actions/commands, checks and failures, and
owned Access database/PID and safe closure state. Keep it concise, reference the
backlog, and omit credentials, personnel data and sensitive report content.
Manual assistant switching requires no private conversation history.

## Sources of truth

- [PROJECT_STATE.md](PROJECT_STATE.md): current architecture, supported behavior,
  decisions and open gaps.
- [planning/PRIORITY_BACKLOG.md](planning/PRIORITY_BACKLOG.md): planned work.
- [docs/Development_Workflow.md](docs/Development_Workflow.md): commands and focused
  test runners. [CONTRIBUTING.md](CONTRIBUTING.md) maps each kind of change to its checks.
- [planning/BUILD_PIPELINE.md](planning/BUILD_PIPELINE.md): build order and the form workflow.
- Other dated `planning/` files, including `PROJECT_STATE_HISTORY.md`, are
  historical evidence, not current requirements.

If documents conflict in a way that affects the task, point out the conflict
and ask. The user's instructions for the current task take precedence over
plans, but they never silently override the product invariants below.

## Product invariants (non-negotiable)

- **Historical import:**
  - Every release must import data from every previously supported schema.
  - Preserve rows, fields, IDs and relationships, revision/review history,
    provenance, recovery, command members, notes, defaults and reporting-senior
    history.
  - Use the production migration engine and the Settings & Help import. Read
    sources through read-only DAO without executing their code, and verify copied
    values before commit.
  - Never silently replace a populated workspace or merge records by ID.
- **Schema policy:**
  - Changes are additive and versioned.
  - The importer accepts source schemas 1 through the current `SCHEMA_VERSION`.
    When the schema changes, raise the hard-coded bound in `modMigration`
    alongside it. Never narrow the range.
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
  - Report history is never deleted, with one exception: draft disposal that the
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
  tier. `test-build-plan.ps1` fails if a runner or harness is missing from it or
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

1. Build a new Blank and a separate synthetic Demo.
2. Run `scripts/release-check.ps1 -BlankPath <blank> -DemoPath <demo>` plus the
   feature-specific, data-transfer, rollback and unsupported-source checks that
   apply.
3. Document which schema fixtures were tested.
4. Put dated, versioned files in `output/` and keep the prior outputs.
5. Upload to the Drive folder named in Development_Workflow.md, then verify the
   names, sizes and folder membership.
