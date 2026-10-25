# Development in VS Code

This document owns development commands and workstation use. Read
[PROJECT_STATE.md](../PROJECT_STATE.md) for architecture/support decisions,
[AGENTS.md](../AGENTS.md) for scope/data/release requirements and
[CONTRIBUTING.md](../CONTRIBUTING.md) for change-specific checks.

Maintain tracked VBA, JSON catalogues and generators as authoritative source.
Choose the most suitable available tool for each task, then validate at a level
proportionate to the change and its risks. The existing Access COM build is the
verified reproducible-build route, not a requirement that every inspection or
edit must be executed through PowerShell. Routine feature work must not introduce
another release-specific assembly script.

## Tool selection and development boundaries

Before implementing a task, identify its affected objects, expected files,
preferred editing/inspection method, and validation level. Discover available
native, embedded, MCP and project tools when their capabilities are relevant;
do not assume previously used PowerShell scripts are the best interface.
Choose based on correctness, safety, reproducibility, operational cost, and
maintenance burden, not novelty or incumbency.

- **Inspect/search:** Prefer direct repository readers, editor facilities, and
  available MCP tools for bounded source inspections and dependency tracing.
  Do not manufacture a new script just to list, read, search, or compare files.
- **Access object investigation:** Prefer Access Explorer MCP for native object
  discovery, procedure-level diagnostics, and cross-object usage searches when
  it offers a clear advantage. Its 66 tools are callable; select the relevant
  operation and respect the verified scope below. Connect to a disposable copy,
  not a user workspace or authoritative template.
- **Source modification:** Prefer direct, precise edits to tracked source.
  Experimental native modifications belong only on disposable copies; verify
  the relevant MCP mutation individually before adopting it into routine use.
  Mirror useful native results into authoritative tracked source before rebuilding.
- **Build/verification:** Use the established reproducible build and native
  regression gates where applicable. A more direct tool may validate a localized
  change earlier, but cannot substitute for release acceptance.
- **Tooling improvement:** If existing tooling is inefficient, propose a specific
  alternative and a bounded comparison. Do not silently replace the build pipeline
  or add a competing maintenance workflow during unrelated feature development.
- **One-off tasks:** Prefer existing tools and ephemeral commands over permanent
  helper scripts. Persist a new script only if it solves an ongoing, documented
  operational need and does not duplicate an existing capability.

### Validation levels

- **Focused:** For isolated source-level edits, perform appropriate source
  checks, affected-object compilation where available, and targeted functional
  tests. Document what was and was not validated. Source compilation on a scratch
  copy is not equivalent to verification of a fresh distributable build.
- **Integration:** For cross-object, generated-form, persistence, report,
  recovery, migration, or build-pipeline changes, create a fresh build and run
  the relevant native integration checks. Escalate when dependencies or risk
  justify broader testing.
- **Release:** For explicitly authorized publication, execute applicable release
  gates, artifact verification, compatibility/rollback checks, and any additional
  physical or feature acceptance needed. This level is not automatically
  authorized by a normal feature request.

Do not weaken a safety invariant to save execution time. If a required native
check cannot run, report the limitation instead of claiming verification.
Avoid needless repeated fresh builds of identical inputs; use the documented
fingerprinted verification receipts only under their existing strict validity
conditions.


## Environment

Use desktop Windows with Microsoft Access, Windows PowerShell 5.1 and Git.
Tasks explicitly use `powershell.exe`, preserving the tested COM host even when
the interactive terminal uses PowerShell 7. Do not run Access builds through WSL.
Python and other available tools may be used where they are a better fit; do not introduce dependencies or parallel toolchains without justification.

Workspace recommendations are in [extensions.json](../.vscode/extensions.json):

- Microsoft PowerShell: script navigation, diagnostics and script debugging.
- VBA by serkonda7: syntax highlighting and snippets, not Access compilation
  or an Access debugger.
- GitHub Pull Requests: repository review integration.
- Access Explorer: object browsing, SQL, VBA editing/compilation and cross-object
  usage searches through its separately managed MCP-Access backend.

Use Access's VBA editor for native breakpoints and runtime investigation. Bring
native edits back to tracked source before rebuilding. Automatic VBA formatting
on save is disabled to avoid unrelated source changes.

### Agent tool registrations

Agent tooling is per-client. On this workstation the read-only
`@modelcontextprotocol/server-filesystem` reader is registered in Codex as
`swift_files` (write tools excluded by its allow-list), and Access Explorer is
registered for VS Code and Codex. Claude Code has neither; see
[CLAUDE.md](../CLAUDE.md). Any agent may read source directly. Use the existing
PowerShell build and native-check commands when their verification scope is
required. [Reader verification](../planning/FILESYSTEM_MCP_SETUP_2026-10-21.json)
is historical evidence.

### Access Explorer

Access Explorer 1.2.0 and its MCP backend are configured on this workstation.
The runtime is under the extension's user-level managed storage, with its own
Python 3.12 virtual environment. Access Explorer user settings select that
interpreter and backend script. VS Code's user MCP configuration registers
`access-explorer-local`; Codex registers `access_explorer`.
Existing settings were backed up before registration. No machine-specific
runtime paths are committed in this workspace's configuration.

The [capability audit](../planning/ACCESS_EXPLORER_CAPABILITY_AUDIT.md) confirms
66 callable tools and records successful disposable-copy probes for discovery,
SQL reads, table/relationship inspection, procedure reads, control inventory,
and static lint. It owns the full tool inventory; availability is not evidence
that mutation operations work reliably. Mutation, compilation, maintenance,
UI actions, and data transfer were not exercised by that audit.

For daily investigation, copy a development build to a new scratch path and
pass that path to MCP tools, or use **Access: Add Connection** on the copy.
Read only the relevant procedures/objects; combine cross-object usage searches
with control event properties, queries, macros, generators, and dynamic callers
before deciding something is unused. Prefer tracked-source edits for the fix.
If a native experiment is necessary, verify the specific operation and its
result on the disposable copy before routine adoption; port useful changes
back to source and run the applicable existing checks. Direct **Save Code to
Access** changes the database, not tracked source.

Safety boundaries:

- **Events and writes:** The disabled arbitrary VBA/macro execution gate is
  not a read-only mode. Opening forms/reports, screenshot/report-output tools,
  and UI actions can execute existing events. Use trusted disposable copies;
  do not change execution permissions as part of ordinary development.
- **Compilation:** `access_compile_vba` automatically decompiles first when
  needed in its session and can discard unsaved VBA. Save/export experimental
  edits before using it on a scratch copy. `access_vbe_check_syntax` checks
  structure only; it does not prove native compilation or functional correctness.
- **SQL:** Prefer bounded SELECT diagnostics. SQL batches are not atomic:
  stopping on error does not roll back earlier writes. Confirmation flags do
  not make a write safe or substitute for production transactions.
- **Results and lint:** Inspect returned text/JSON for errors even when
  `isError=false`. Static lint findings are investigation leads, not verified
  visible defects, accessibility compliance, or printed fit; interpret them
  against actual page/visibility state and fixed report geometry.
- **Verification and sessions:** Explorer imports/transfers do not replace
  production migration, recovery, reproducible builds, or release gates. Apply
  the [Access-session ownership and handoff rule](../AGENTS.md#access-automation-safety)
  before switching between Explorer and another controller. The backend close path
  has a force-kill fallback; do not use it on a session with unsaved user work
  or use restart-processes on user sessions.

## Standard reproducible build and test command

Use this command from the repository root for a full routine build-and-test cycle,
not automatically for every inspection or focused change:

```powershell
powershell.exe -NoProfile -File scripts/develop.ps1
```

This runs preflight, creates a uniquely named fresh blank under `build/dev`,
then runs current-source/blank verification, smoke, embedded report tools,
fixed-page/buffer, workflow/recovery and directory-merge checks. Report-tools
tests explicitly seed their own scratch copy with synthetic approved reports;
the supplied blank remains empty. Suites run
serially in child processes on disposable copies. Testing must not change the
supplied artifact's hash. Failures stop the command. Successful receipts record
the artifact, input fingerprints, steps and logs. Failed runs retain logs
without a PASS receipt.

Press **Ctrl+Shift+B** for the same default build-and-test task. **Tasks: Run
Task** also exposes Preflight, Fresh build only, Focused tests on a blank build,
and Release verification. Paths are process arguments, preserving spaces.

```powershell
powershell.exe -NoProfile -File scripts/develop.ps1 -Action Preflight
powershell.exe -NoProfile -File scripts/develop.ps1 -Action Build
powershell.exe -NoProfile -File scripts/develop.ps1 -Action Test -BuildPath "D:\path\new-blank.accdb"
```

`-BuildPath` can choose a new output for Build/BuildAndTest. Existing files are
refused. The lower-level [build.ps1](../scripts/build.ps1) and focused suites
remain available for diagnosis. A fresh build must not use a retained binary
as its source. Preflight follows canonical version/schema sources.

### Preview ownership and recovery (M02)

For M02 preview changes, use the focused native runner before the final integrated
build, then repeat against the unchanged final artifact:

```powershell
powershell.exe -NoProfile -File scripts/test-wizard-state.ps1 -BuildPath "D:\path\blank.accdb" -LegacyPath "D:\path\verified-schema29-blank.accdb" -PDFPath "D:\path\synthetic-filled-eval.pdf"
```

`-InstallSource` installs changed services on an existing schema-30-or-newer
disposable copy for early checks; it preserves staging rows and metadata. Use
production migration for older sources. It does not verify generated form layouts.
The final run omits it. The runner invokes an actual VBE Reset and an intentional
unhandled VBA error, starts a separate Access session, and tests source/target/
contributor changes, cancel, staged-state migration, late rollback/retry and source
hashes. The classic VBE error dialog has no UI Automation Invoke pattern; a bounded
supervisor recognizes only the literal synthetic error on the owned PID/start time
and sends one End shortcut. No arbitrary dialogs, user sessions, C#/Win32 helper,
new framework or persistent background service is used. Missing/unexpected UI
must not be counted as unhandled-error coverage. The optional legacy proof uses
an actual schema-29 source and the production Settings import; it does not imply
that all historical schemas were retested. Inspect Resume/Cancel placement and
binding on all three wizard forms. Saved staging lives inside the protected
workspace; Cancel targets only the current contributor's preview, not source files.
The optional `-PDFPath` is a supported synthetic filled EVAL input (such as an
existing PDF fixture-generator output); it adds positive PDF Resume-callback and
resumed-Apply checks. Without it, PDF source/choice recovery uses the supplied blank
template and positive filled-PDF UI/apply coverage is not claimed. The ignored local
`verification.json` records which optional proofs ran, assertion counts and hashes.

Use `-CrossWorkflowOnly -PDFPath <synthetic-filled-eval.pdf>` for deterministic
contributor/task-order cases without the reset/error-dialog phase. It covers
stale/superseded cancellation, owner-private drafts, archive-between-preview-and-
Apply and invalid pending editor values. It does not claim reset/restart coverage.
Cancel with a live context must match its saved token; Resume is required before
cancelling a different or newer preview. Workspace-import cases additionally test
contributor/superseded cancellation and close/reopen without automatic Apply.
All three cancel adapters must return explicit failure guidance and preserve choices
when a live token is stale, rather than exposing an unhandled event-expression error.

Use the existing wizard runner with `-ApplyFaultsOnly -PDFPath <synthetic-filled-eval.pdf>`
for PDF/return application outcomes. It owns a separate disposable database for each
case: native precommit rejection and retry, an explicitly instrumented postcommit
fault, a display fault after successful service return, and a simulated uncertain
outcome. Check saved rows/revisions/receipts, consumed versus retained descriptors,
disabled Apply, honest recovery feedback and source hashes. This mode does not run
or claim VBA reset/restart. `-InstallSource` is for early focused copies; omit it on
the final fresh artifact. Fault injection remains local to disposable modules.
`-SkipUnhandledFailure` omits only the intentional visible VBA-error dialog from
the full wizard run while retaining real VBA reset, restart and migration checks.
Use it when another attended Access session is open; its receipt explicitly records
the skipped unhandled-error proof. The default still requires the supervised error.

### Batched read verification (M03)

Use [test-batched-reads.ps1](../scripts/test-batched-reads.ps1) for native output,
guard and performance proof before the final fresh build, then repeat without
`-InstallSource` against the exact final artifact:

```powershell
powershell.exe -NoProfile -File scripts/test-batched-reads.ps1 -BuildPath "D:\path\candidate-blank.accdb" -BaselinePath "D:\path\verified-before-change-blank.accdb" -Size 500 -Trials 5
```

It owns sequential disposable copies, seeds synthetic data, copies it through the
production read-only-source importer, compares complete cell values/order, and
measures medians with `Stopwatch`. Guard/fault tests run on a separate disposable
copy after measurements. CSV timings survive a later assertion failure; only a
fully passing run gets `verification.json`. Both supplied artifacts remain hashed
and unchanged. `-InstallSource` is only for focused development on disposable
copies; it cannot establish a fresh integrated artifact. Optional `-DataSourcePath`
reuses an unchanged synthetic benchmark fixture through production import, avoiding
repeated seeding; its command marker/row counts and source hash are checked. Do
not use a populated operational workspace as benchmark data.

Run [test-ux-read-rows.ps1](../scripts/test-ux-read-rows.ps1) for Reports
prepared-row equivalence, sorting/filtering and null/NOB/missing-field cases,
[test-ux-my-work.ps1](../scripts/test-ux-my-work.ps1) for My Work behavior, and
[test-ux-performance.ps1](../scripts/test-ux-performance.ps1) for refresh/state
invariants. Deliberate cache divergence must reject writes; restore the exact
synthetic values before testing unrelated mutations. The benchmark accepts up to
1,000 reports plus 1,000 prior evaluations; same-version comparisons measure
regression/equivalence, not a new optimization's speedup.

[measure-ux-read-rows.ps1](../scripts/measure-ux-read-rows.ps1) with
`-BuildPath <blank.accdb> -Size 1000` compares prepared/unprepared Reports and
My Work reads at 25, 100 and the selected maximum (100-1000), alternating sample
order over five trials. Traits are seeded through normal saves. Exact values/order
and unchanged history are required; per-run CSV and verification files sit beside
the unique disposable fixture. Timings cover query/data preparation, not screen
paint, cold startup, tail latency or a production performance guarantee.

Also run the existing persistence, reporting-senior/RSCA, directory/history,
folder cleanup/rollback/undo and performance/recovery suites for changes to the
shared readers. A preexisting fixture that intentionally corrupts a C06 projection
must not justify disabling that persistence guard. Profile before widening scope;
keep display snapshots separate from live write-time validation.

### Relationship names and revision-ID aliases (L01–L02)

Use the existing major-workflow runner before the integrated build, then repeat
against the exact final blank without `-InstallSource`:

```powershell
powershell.exe -NoProfile -File scripts/test-major.ps1 -BuildPath "D:\path\blank.accdb" -LegacyPath "D:\path\schema31-blank.accdb" -OlderPath "D:\path\schema30-blank.accdb"
powershell.exe -NoProfile -File scripts/test-ux-trait-retry.ps1 -BuildPath "D:\path\blank.accdb"
```

The major runner owns sequential disposable sessions and checks existing batch,
meeting, routing and undo behavior, writable revision-ID aliases and FK rejection.
It also exercises the current comparison's before/after/detail rows and the
retained hidden workbench adapter; these checks do not reopen a retired route.
It seeds synthetic populated sources, imports them through Settings & Help, and
verifies original fields/IDs, complete relationship signatures (including field
order and attributes), late rollback, source hashes, populated-target rejection
and reopen. Future schemas and unknown source tables remain rejected. Optional
legacy paths must be actual retained templates, not relabeled current schemas.
This does not change historical-output waivers or establish release acceptance.

`-InstallSource` prepares an early blank disposable candidate, installing the two
schema-owned alias queries and transactionally replacing only opaque relationship
names with the same definitions. It verifies signature equality and imports the
changed source modules. It is not the production migration path or a fresh build.
Production import reads the source without executing its code, leaving source
relationships, physical column names and original IDs untouched; the fresh target
already contains the reviewed schema-32 names/queries. No in-place rename of a
populated workspace is needed.

During alias transaction checks, use the owned DAO reader and roll back before
closing the query cursor. The initial harness sequence raised native error 3246;
that test orchestration was corrected without adding a production workaround.

### NAVFIT98A export (H02)

The supported export screen is summary-group scoped: select a group from Summary
groups or open its Members page, choose **Export group to NAVFIT98A**, and choose
an output directory. The build embeds the tracked, hash-reviewed native blank;
users do not supply a template or baseline. Readiness findings remain informative.
Legacy returned-report services and provenance remain tested backend behavior,
but their compatibility controls are hidden from this export screen.

Run the existing native suites sequentially on disposable copies before the final
integrated build, then repeat against the exact final blank:

```powershell
powershell.exe -NoProfile -File scripts/test-navfit-group-export.ps1 -BuildPath "D:\path\blank.accdb"
powershell.exe -NoProfile -File scripts/test-ux-navfit.ps1 -BuildPath "D:\path\blank.accdb"
powershell.exe -NoProfile -File scripts/test-navfit-adapter.ps1 -BuildPath "D:\path\blank.accdb"
```

These runners default to installed application code. Use `-InstallSource` only for
early focused changes on their disposable copies; that is not a fresh integrated
build. Source-only installers require compatible generated forms and resource
modules; use a fresh candidate when those inputs changed. They check source-build
hashes and preserve primary errors during owned
session cleanup. The adapter suite retains mapped/native field and return/undo
checks; group/UI suites cover incomplete/mixed/archived/approved exports, every
selected member, informational findings, exact/unavailable PSG, pending buffers,
recovery, source/history preservation, stale previews and operational failures.
These are Access/DAO format and workflow checks, not a new NAVFIT98A desktop
acceptance or printed-fit claim. Actual serialization/capacity limits still fail
without truncation/coercion; readiness findings do not require acknowledgement.

For editor navigation/trait changes after a VBA reset, run the existing focused
runner with `scripts/test-compat-pages.ps1 -BuildPath <blank> -EditorReset`.
It resets the real VBA project between preparation and verification, then checks
authored Chief comments, original revision guards, raw invalid-value recovery,
and return to the same summary group. A successful recovery checkpoint permits
navigation after a save failure; a failed checkpoint must retain the open buffer.

The adapter guard suite also injects a missing template together with a rejected
failure-journal write, and separately rejects completion-journal writes after files
have been published, for report and group output. Require the original error
number/identity, separate cleanup details, redacted diagnostics and truthful
publication state. Existing published/pending files are recovery evidence, not
permission to overwrite them; a retry must preserve their bytes and saved revisions.

### Diagnostics, view contracts and trait data (M04–M06)

Directory-batch preview/failure checks use the existing
`scripts/test-directory-batch.ps1 -BuildPath <blank>` and its
`-AdversarialPreviews` mode. The latter isolates each case in a new owned Access
controller, injects postcommit/display faults only into disposable copies, tests
an explicitly simulated uncertain outcome, and executes a real VBE reset. Require
truthful feedback, disabled/stamp-free repeat Apply and persisted report/link
counts after reopening. No fault hooks are shipped in application source.

For the summary-list M03 read optimization, use the existing lifecycle runner:

```powershell
powershell.exe -NoProfile -File scripts/test-ux-group-lifecycle.ps1 -BuildPath "D:\path\candidate.accdb" -BaselinePath "D:\path\baseline.accdb" -GroupListBenchmark -Groups 40 -Members 10 -Trials 5
```

It compares all eight displayed cells and order for All/Active/Archived, covers
eight empty/default/missing-lifecycle/duplicate-title boundaries, checks no data/
history row changes, and alternates five samples per artifact. A lookup counter
is installed only in test copies. `-InstallSource` is an early compatible-copy
option, not integrated build proof. `-DataSourcePath <synthetic-summary-list.accdb>`
reuses the matching fixture through production migration with row/field verification
and source hashes. The retained Completed row is a read-only legacy-state case,
not an RSCA completion or release fixture. Timings are local measurements.

For M02/M04/M05 preview ownership and group-import outcomes, use the existing
native runners sequentially on disposable copies:

```powershell
powershell.exe -NoProfile -File scripts/test-ux-native-groups.ps1 -BuildPath "D:\path\blank.accdb"
powershell.exe -NoProfile -File scripts/test-ux-native-groups.ps1 -BuildPath "D:\path\blank.accdb" -AdversarialPreviews
powershell.exe -NoProfile -File scripts/test-ux-organization.ps1 -BuildPath "D:\path\blank.accdb"
powershell.exe -NoProfile -File scripts/test-ux-organization.ps1 -BuildPath "D:\path\blank.accdb" -DisplayFaultOnly
powershell.exe -NoProfile -File scripts/test-directory-merge.ps1 -BuildPath "D:\path\blank.accdb"
powershell.exe -NoProfile -File scripts/test-directory-merge.ps1 -BuildPath "D:\path\blank.accdb" -DisplayFaultOnly
```

These check actual contributor switching, fresh review after failure/reopen,
atomic native group import and retained history. Adversarial cases inject faults
only into owned test copies at the real commit/display boundaries; the uncertain
outcome is explicitly simulated. The reset case executes native VBE Reset with
all three previews open and verifies that Apply writes nothing. Reopen checks
verify report/group/receipt counts. Display-fault merge tests retain guarded Undo.
The normal installed application contains none of this fault instrumentation.
After a reset, return to import choices and reselect the native group source
before previewing; its source reader is in-session. Move/merge selections remain
visible for a fresh preview. No automatic Resume is added to these bounded flows.

For workspace-import outcomes, run the existing
[test-engine-verification.ps1](../scripts/test-engine-verification.ps1) with
`-BuildPath <blank.accdb> -SourcePath <retained-synthetic-workspace.accdb> -CommitFaults`.
It verifies adversarial field comparisons, a precommit mismatch with confirmed
rollback/retry, and a fault injected immediately after the real production commit
on a separate disposable copy. The injection is explicit test instrumentation,
not an application hook. Require retained committed data, the original error,
disabled re-Apply, usable close/unlocked UI and unchanged source hashes. Use only
synthetic source workspaces. `-InstallSource` loads changed diagnostic/import/view
services for early focused checks; final runs omit it. A dispatched import or an
error message alone does not prove whether the transaction committed.

The diagnostics/contracts runner also corrupts status/progress controls and action
bindings on all four import/return/PDF screens. `-InstallViewContracts` loads only
the changed view-contract module on a disposable current-schema copy, preserving
metadata and schema. Final checks omit it. The existing installed verifier reads
the same literal contracts; no second contract registry is introduced.

Run the focused checks before the final integrated build, then repeat without
source installation against the exact final blank:

```powershell
powershell.exe -NoProfile -File scripts/test-diagnostics-contracts.ps1 -BuildPath "D:\path\blank.accdb" -VerifyViews
powershell.exe -NoProfile -File scripts/test-trait-catalogue.ps1 -BuildPath "D:\path\blank.accdb" -LegacyPath "D:\path\verified-schema30-blank.accdb"
```

The diagnostics runner owns disposable copies and tests original/cleanup error
separation, redaction, rollback/retry, and renamed or incorrectly bound controls.
It writes suppression/error-identity inventories and a verification receipt.
`-InstallSource` is for early disposable-copy checks only; `-PrepareOnly` prepares
a candidate without behavioral assertions. Neither establishes a fresh build.
`-VerifyViews` reuses installed-source/form/report/macro verification and proves
the installed checker rejects renamed controls and wrong existing callbacks.
Build verification reads literal contract data from verified VBA and inspects
native properties without Eval or Trust Center changes. Unsupported contract
syntax fails explicitly. Expected failures are caught in their VBA harness or
native-property checker so no unhandled modal dialog is expected.

The catalogue runner compares every trait/mark/title/description/source field,
checks uniqueness/domains and missing-definition failure, and uses production
migration to test populated transfer, retained catalogue versions, malformed
source-column rejection and atomic conflict rejection. Optional `-LegacyPath` must be a real schema-30 blank; its
copy receives synthetic data. Both originals remain hash-checked and unchanged.
This does not claim all historical fixtures were retested. Run affected native
save/recovery, trait, group, migration and navigation suites as well.

Schema 31 includes seeded `TraitDefinition` reference data. Blank verification
requires its exact reviewed contents and still rejects all user/domain rows.
Never delete the reference rows to make a blank check pass, reinterpret them as
user content, or overwrite conflicting source definitions during import.

### Remote regression and accessibility preparation

For backup/recovery changes, run [test-backup.ps1](../scripts/test-backup.ps1)
with `-BuildPath <blank.accdb>` and optional `-CompiledPath <candidate.accde>`,
and [test-recovery.ps1](../scripts/test-recovery.ps1) with `-BuildPath <blank.accdb>`.
The backup runner covers a backup taken while the editor holds pending values
(it is checkpointed and closed, and Access stays open), literal filenames, the
verification record and file hash, an independent row-by-row comparison, restore
through the Settings & Help import with raw recovery fields, and rejection of
existing output or evidence, missing folders, wrong extensions and linked tables.
With `-CompiledPath` it also backs up an ACCDE workspace. `-InstallSource`
prepares only its disposable copy; final verification omits it. Backup starts no
other process.

Use the existing focused runners for task-order and view regressions:

```powershell
powershell.exe -NoProfile -File scripts/test-ux-validation.ps1 -BuildPath "D:\path\blank.accdb"
powershell.exe -NoProfile -File scripts/test-major-ui.ps1 -BuildPath "D:\path\blank.accdb"
powershell.exe -NoProfile -File scripts/test-live-members.ps1 -BuildPath "D:\path\blank.accdb"
powershell.exe -NoProfile -File scripts/test-group-exit.ps1 -BuildPath "D:\path\blank.accdb" -TaskOrders
powershell.exe -NoProfile -File scripts/test-compat-pages.ps1 -BuildPath "D:\path\blank.accdb"
```

The six task orders alternate PDF preview, saved NAVFIT export and read-only
findings with invalid pending grid/editor values. They assert unchanged saved
revisions, retained raw values, explicit discard and guarded editing navigation.
Members checks include live/historical views, stale revisions/membership and a
canonical non-grid proposal. The retained workbench is tested as a hidden adapter,
not presented as a second Members UI. List headers are not data rows, and native
VBA checks real control bounds rather than nullable external COM properties.

Validation tests set native `InsideWidth`/`InsideHeight` before asserting resized
geometry; an initial zero client width is not a rendered layout. Repeated sizes
check evidence/action clearance, Help, paging, selection and focus. Fixed-page
tests check all 21 Chief comment labels, ordered tabs and native subform focus.
The entry-label test checks ten associated labels/tab stops in setup, contributor
selection and PDF preview. `-EntryLabelsOnly` is a bounded subset. The validation,
Members, task-order and entry-label runners offer `-InstallSource` for supported
preparation of their disposable copies; final checks use the fresh artifact
without that switch. This does not replace full generation or visual acceptance.

M07's development targets follow the [Access Board software provisions](https://www.access-board.gov/ict/#chapter-5-software)
and Microsoft's [Access tab-order guidance](https://support.microsoft.com/en-us/access/set-the-tab-order-for-controls).
Treat this as a traceable test plan, not a compliance determination:

| Area | Development evidence | Remaining attended acceptance |
| --- | --- | --- |
| Names and state | Native attached-label and control-property checks | Screen reader announces correct label, role, value, state and findings |
| Keyboard and focus | Tab indices, native focus and pending-value guards | Tab/Shift+Tab, activation, Escape, visible focus and no trapped dialogs across complete tasks |
| Display and contrast | Bounds at explicit viewports; selected static contrast inspection | Target DPI/scaling, high contrast, zoom and physical readability |
| Dialogs and output | Context, error/retry, source preservation and PDF checks where actually run | Screen-reader notifications, PDF reading order/accessibility and physical print fit |

Retain attended results for contributor startup, all report kinds, Members,
review, findings, import/export, dialogs and output. Do not infer keyboard-key
delivery, screen-reader behavior or Windows DPI acceptance from `SetFocus`,
geometry or a static lint verdict. The inspected screenshot tool refused normal
PopUp opening before capture; do not count a design image as runtime evidence or
cycle through alternate capture mechanisms to bypass that boundary.

### Compiled desktop candidate (M01)

After a fresh ACCDB build and its applicable focused/integrated tests pass, use
the same entry point to create a local, unsigned ACCDE candidate:

```powershell
powershell.exe -NoProfile -File scripts/develop.ps1 -Action Compile -BuildPath "D:\path\verified-blank.accdb"
```

Optional `-CompiledPath "D:\path\new-name.accde"` chooses a new destination;
existing files are refused. The default is a unique versioned, architecture-labelled
file under `build/dev`. Compile checks the source identity/source bindings and
blank inventory, converts a disposable copy, then runs
[test-compiled-distribution.ps1](../scripts/test-compiled-distribution.ps1) on
copies of the actual ACCDE before accepting it. No VBA test harness or replacement
application source is imported into the compiled binary. Source/candidate hashes
must remain unchanged. The receipt records compiler file version, executable
architecture, source and output SHA-256, build inputs and checker fingerprints.
These are trusted development checks (`AutomationSecurity=1` on owned copies),
not evidence of normal-policy launch or signing acceptance. The observed host's
normal Office policy displayed an unsigned-file security prompt; cancelling it
produced a security/signature refusal. This does not establish what an explicit
user trust decision or an approved trusted location would permit.
`-CheckNormalSecurity` on the focused runner is a separate **attended**
acceptance option: security prompts require user interaction and must not be
bypassed. Compiled PDF generation is still unverified because the external COM
compiled-form/control interface failed during the attempted test. Do not infer
  compiled PDF support from the passing ACCDB PDF suite. The import picker lists
  both ACCDB and ACCDE; production data-import handlers read either through DAO
  without executing source code.
Keep the source ACCDB and its receipt: compiled VBA source cannot be edited or
recovered through the ordinary Access development workflow.

The conversion uses Access's **undocumented SysCmd 603** action, isolated in
[develop.ps1](../scripts/develop.ps1). Access Explorer has no conversion operation;
its compile-VBA tool does not create an ACCDE. This route is accepted only with
the native behavioral proof on the recorded desktop host, not assumed reliable
on another Access version or architecture. If it fails, preserve the primary
error and candidate, close the owned controller, and stop. Do not cycle through
UI/Win32 conversion techniques. Microsoft's documented manual route is
**File > Save As > Make ACCDE** on a disposable verified source; manually created
candidates still require the focused compiled tests and separate release gates.

ACCDE support is limited to the tested desktop Access version and bitness;
32-bit and 64-bit binaries are separate distributions. Other workstation versions
need their own compilation and acceptance. Free Access Runtime remains out of
scope. ACCDE compilation is code/design hardening, not encryption or access
control. Keep publishing the existing Blank and synthetic Demo ACCDBs when a
release is explicitly authorized. Compile does not publish or replace either.

**Signing and integrity:** candidates are explicitly unsigned. Do not generate
a self-signed production identity, choose a certificate automatically, change
Trust Center policy or publish an unsigned candidate as signed. When an approved
certificate is available, use Access **Package and Sign** for a final ACCDC package
containing the verified ACCDE. Record the approved publisher, thumbprint, expiry,
package hash and contained-file hash; independently verify the signature/publisher,
extract and compare the binary hash, and test normal startup/import on the target
desktop before publication. Retain the editable source and final verification
receipts. Repackage/reverify whenever distributed bytes change. A package signature
attests the delivered package; it does not promise ongoing integrity of a mutable
workspace. Prefer this package route over assuming an embedded VBA signature
survives migrations that create queries. Do not apply PowerShell Authenticode
signing to an Access binary. Certificate availability and actual signing remain
separate acceptance gates.

Microsoft documents [ACCDE creation and restrictions](https://support.microsoft.com/en-us/access/hide-vba-code-from-users),
[deployment/version considerations](https://support.microsoft.com/en-us/access/deploy-an-access-application)
and [Access signing/package behavior](https://support.microsoft.com/en-us/access/show-trust-by-adding-a-digital-signature-to-an-access-database).
Retain all existing compatibility, migration, rollback, release and physical
acceptance gates; current-schema compiled tests do not prove historical coverage.

For H01 PDF changes, run the [focused PDF suite](../scripts/test-pdf-output.ps1)
against a fresh blank. It owns a disposable database and synthetic output root:

```powershell
powershell.exe -NoProfile -File scripts/test-pdf-output.ps1 -BuildPath "D:\path\new-blank.accdb"
```

During refinement, `-InstallSource` applies the tracked PDF catalogue, service,
forms/report and affected layout modules on the disposable copy. Add
`-ReusePresentation` for service/asset iterations against an already installed
PDF fixture; it refreshes embedded backgrounds while retaining form/report controls,
and does not verify structural generator changes. These modes do not replace
the final fresh build. The suite checks draft/grid preservation, nonblocking invalid
data, ordering, output inventory, repeat runs, failure state, all report types and
native page preview. Also reopen the produced PDFs with an independent renderer:
check every mapped block, all seven traits, exact long-text readback, absence of
`#Error`/blank pages, combined order, summary continuation and Letter page size.
Review actual filled pages before the final integrated build; a native export
return value or file length alone is insufficient.

For the schema-33 Chief correction, the same runner also checks all 21 authored
comments, mark-change isolation, native mapping/return conflicts/clear, history
restore, editor page changes and overflow. Its print matrix includes blank,
short, two-line and overflowing comments, all 35 block/grade combinations, a
summary sheet and alphabetical combined/individual/Both output. During focused
work, `-MetricsOnly`, `-ChiefFieldsOnly`, `-ChiefProofOnly` and `-ChiefCasesOnly`
select bounded subsets. `-InstallSource` additionally prepares nullable report/
recovery columns, regenerated metadata and new editor controls on its disposable
copy; production upgrade remains fresh-build plus Settings & Help import.

Retain the related focused gates:

```powershell
powershell.exe -NoProfile -File scripts/test-major.ps1 -BuildPath "D:\path\blank.accdb" -LegacyPath "D:\path\schema32-blank.accdb" -OlderPath "D:\path\schema31-blank.accdb" -ChiefComments
powershell.exe -NoProfile -File scripts/test-ux-workflows.ps1 -BuildPath "D:\path\blank.accdb" -ChiefComments
powershell.exe -NoProfile -File scripts/test-pdf-expanded.ps1 -BuildPath "D:\path\blank.accdb" -FixturePath "D:\path\synthetic-xfa-fixtures"
```

The transfer suite verifies blank new values in old snapshots, actual authored
values in current snapshots, old sealed native baselines, untouched native-only
comments after unrelated saves, late rollback and source hashes. The workflow
suite verifies all 21 values after normal shutdown and restoration in a new
Access session. Structured input fixtures must carry distinct synthetic markers
in all 21 fields; never use personnel data in tracked test fixtures.

The font-metrics choice is deliberate: native report font properties are limited
to the print context ([Microsoft's report-font example](https://learn.microsoft.com/en-us/office/vba/api/access.textbox.fontsize))
and returned Access 2455 during planning. The existing renderer now uses built-in
Windows [font creation](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-createfontw)
and [text metrics](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-gettextmetricsw),
with face verification and owned-handle release. This is application text
measurement, not a development UI-control/capture helper. No C#, PDF.js, Python or
Acrobat is introduced into the application runtime. Plain and rich Access print
slots both rounded the source's 6.5-point trait-34 grade-4/5 wording; the recorded
development fallback is 6 points for those cells, with the other grades at 7.
That difference, actual printed fit and an Acrobat-absent workstation remain
explicit acceptance items. The unsuccessful rich-text route was removed.

The [print catalogue](../src/pdf-print-layout.json) records public Navy source and
300-dpi image hashes. Development asset preparation used approved temporary
PDF.js 6.4.299 and pikepdf 10.17.0; neither is an application dependency. For XFA
blank assets, render the print layer at 300 dpi, hide widget chrome/values, retain
official text/geometry, and top-align static draw content to avoid PDF.js vertically
centering tall single-line containers into subsequent lines (EVAL/FITREP only).
CHIEFEVAL retains original static alignment and explicit centered 10-point checkbox
outlines, matching the check-button regions in the print metadata. Inspect both pages of
each template and update reviewed asset hashes together. CGTimes fallback and actual
printer margins/fit require physical acceptance; do not infer it from screenshots.
Keep synthetic readback evidence under `build/tests`. Never send personnel PDFs to
external services. Verification on a machine with Acrobat absent remains distinct
from proving that this Access-only path makes no Acrobat calls.

## Release verification command (explicit release work only)

After building a new Blank and creating a separate synthetic Demo:

```powershell
powershell.exe -NoProfile -File scripts/release-check.ps1 -BlankPath "D:\path\new-blank.accdb" -DemoPath "D:\path\synthetic-demo.accdb"
```

This checks both artifacts against current source, requires all Blank tables
except AppMetadata to be empty, and runs the existing directory release gate:
merge/history/UI, retained-schema transfers, populated schema-20/current
transfers and rejection/rollback. It adds fixed pages, workflow/recovery, group
exit, engine exactness and canonical message/private-draft transfer. Add
feature-specific suites as needed; this baseline does not replace local PDF
fixtures, benchmarks or other applicable acceptance.

The command verifies supplied files; it neither creates a Demo nor publishes a
release. Create a Demo with [build-ux-admin-demo.ps1](../scripts/build-ux-admin-demo.ps1)
using a current blank template and new output path; see the
[Demo build guide](../planning/UX_DEMO_BUILD_GUIDE.md). This builder verifies the
blank, seeds synthetic data and runs focused Demo/startup checks. Other feature
suites remain change-specific; the duplicate build-preview entry point is held
in PROPOSED_PURGE. Record historical fixture limitations: blank old fixtures do not
prove populated historical workflows. Logs/receipts are in `build/workflow`.
Physical display, accessibility and printed fit require separate acceptance.

Artifact verification reuses the adjacent `.accdb.build.json` structural result
only when artifact path/hash, app/schema, every current build input (including
source, generators and the structural checker), module coverage and completed
verification flags match. Inputs are rehashed before accepting the receipt.
Missing, malformed, stale or incomplete receipts take the normal disposable-copy
native compile/source/binding path. This applies to both development and release
commands through the existing artifact verifier. It does not reuse workflow or
migration results; current metadata and requested blank-table checks always run.
To force a native structural recheck:

```powershell
powershell.exe -NoProfile -File scripts/verify-build-artifact.ps1 -BuildPath "D:\path\new-blank.accdb" -RequireBlank -ForceNative
```

The verifier reports whether it reused the build result. It does not create a
new cache or automatically copy receipts alongside seeded/renamed distributions.

Historical import coverage no longer depends on old output files.
`test-swift-data-import.ps1` builds a data-only source for every schema from 1 to
the build's from [tests/historical-schemas](../tests/historical-schemas/README.md),
seeds synthetic rows into every table, imports each through the production
Settings import and compares every source row and field itself. Each file
records its origin (`git-ddl`, `captured` or `reconstructed`). Before a schema
change, `scripts/capture-schema-definition.ps1 -BuildPath <build>` adds the
outgoing schema. A schema number that shipped with more than one structure has
one file per variant. Kept populated releases in `output/` are imported as well. `release-check.ps1`
requires the receipt to cover every schema. `-CurrentFixturesOnly` still records a
user-granted waiver, as for r128; it never ends support for older schemas. See
PROJECT_STATE.md.

Before publication, follow [AGENTS.md](../AGENTS.md): preserve earlier releases
and user workspaces, place dated/versioned Blank and Demo directly in `output`,
upload both to the designated Drive folder, verify names/sizes/parents, record
the exact final artifacts and close owned Access sessions normally.

## Review and commit

When importing tracked VBA through Access Explorer's `access_set_code`, normalize
the supplied text to LF first (`text.replace(/\r\n/g, "\n")`). Its current text
import path can double CRLF line endings and break VBA line continuations.
The audit did not verify `access_set_code` mutations; before routine use, verify
the specific import by readback and relevant compilation/functional checks on
a disposable copy. Save/export edits before any Explorer compile because of
its decompile side effect. The standard builder uses its own tested
source-normalization path; diagnostic native edits do not replace the required
integration/release checks. Expected-error VBA tests should call services
directly under their error trap rather than nesting Access `Eval` callbacks.

Review Git diffs and `git diff --check`, then prepare relevant source, generators,
tests and documentation for commit after applicable verification. Commit only
when requested or explicitly authorized by the established task workflow. Databases, scratch
builds and logs remain ignored. Preserve unrelated untracked files such as the
user's FORMS directory. No task commits, pushes, uploads or force-closes another
person's Access session automatically.

### Efficient Investigation and Execution

During development, minimize unnecessary tool calls, repeated inspection, and duplicated native automation while preserving required validation.

1. **Search once, inspect selectively.** Start with targeted source searches and read only the relevant ranges. Avoid repeatedly reading unchanged files or reopening previously inspected source without a specific reason.

2. **Reuse established test runners.** Use existing focused runners as the preferred mechanism for stateful Access behavior testing, including context setup, pending edits, saves, recovery and return paths. Select the relevant suite from [CONTRIBUTING.md](../CONTRIBUTING.md), such as [group-exit](../scripts/test-group-exit.ps1) or [group lifecycle](../scripts/test-ux-group-lifecycle.ps1). Prefer a narrowly scoped extension or parameter of that runner when coverage is missing, rather than recreating LoadFromText/compile/Eval/Quit wrappers or adding a parallel harness.

3. **Use MCP where it provides native value.** Prefer Access Explorer for suitable installed-object inspection, VBA references, control properties and native callback discovery when the operation fits its verified scope. Retain COM and existing runners where they provide capabilities or reliability that the enabled MCP does not. An advertised schema is not verification, static syntax checks are not compilation, and SQL must not be assumed to invoke Swift VBA or establish UI context. Do not use MCP merely to replace efficient source searches or enable additional execution permissions to bypass a capability gap.

4. **Monitor meaningful progress.** Wait on the running execution or task's completion channel, normally using 30–60 second intervals for long builds; provide concise user updates independently. Inspect logs on completion, failure, a meaningful milestone or a suspected stall. Additional log tails or process/CPU checks need a specific diagnostic question; do not repeatedly combine short polls with unchanged logs and process listings. Adjust intervals when evidence warrants closer diagnosis; no separate polling helper is needed.

5. **Sequence focused and integrated verification.** Settle focused behavior tests and, for UI changes, visual review of relevant states such as empty/populated rosters and pending-edit failures on disposable copies before freezing tracked inputs for the final authoritative integrated build. Run the required fresh build and affected integration checks on those final inputs. If later refinements change source, perform the necessary final verification; an earlier build or stale fingerprint receipt cannot cover changed inputs. There is no arbitrary build-count cap. Repeat equivalent tests only when a changed artifact, environment, failure or validation purpose justifies it; retain the separate release gates.

6. **Preserve errors and recover the owner safely.** Capture the primary error, stage, target database and controller identity before cleanup; record its owned PID when available and always before process recovery. Preserve pending edits and recovery evidence, and record cleanup errors separately so they do not replace the original failure. Establish ownership and database state before releasing a stale controller or handing off the file. After diagnosis, permit one justified retry on a fresh disposable copy; repeated failure requires further diagnosis rather than blind retries or controller switching. If an existing runner has already masked the primary exception, preserve its logs and report the uncertainty: documentation cannot recover that exception, and a runner fix requires separate scoped work. Do not alter tests merely to satisfy an unexplained failure.

7. **Preserve useful evidence.** Record meaningful test outcomes, baseline comparisons, and unresolved failures without generating unnecessary permanent logs, reports, or documentation.

8. **Bound unusual UI automation.** Prefer one suitable, verified operation over custom orchestration. Apply the [tool-selection justification rule](../AGENTS.md#how-to-work) before embedded C# or Win32 use. For an exceptional fallback, state the capability gap, required UI context, event/save effects, attempt budget and timeout or other stop condition before running it. A form-opening timeout is not a screenshot-capture timeout; the inspected Explorer screenshot backend also uses synchronous PrintWindow and has no OpenArgs parameter. Do not assume it avoids a custom capture's failure. If the bounded fallback fails or cannot establish the required state, stop and retain the evidence; use an existing runner/manual review or report the unverified case instead of cycling through capture/control APIs. New evidence may justify a revised diagnostic plan, not automatic additional attempts.

These working defaults allow justified debugging, native testing and broader validation. They do not relax mandatory session ownership, data protection, source/artifact verification or release requirements.
