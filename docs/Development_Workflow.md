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

### Direct project file reader

The official `@modelcontextprotocol/server-filesystem` 2.0.0 is installed in
user-local tool storage and registered in Codex as `swift_files`. Its allowed
directory is this project. Codex's tool allow-list exposes only text reads,
multiple-file reads, directory listings/tree, file search, metadata and the
allowed-directory listing; write/edit/move tools are excluded from Codex.
The upstream package itself also contains write tools, so retain that Codex
allow-list if changing the registration.

A protocol smoke test read the project instructions and confirmed that an
outside-project path was rejected. See
[reader verification](../planning/FILESYSTEM_MCP_SETUP_2026-10-21.json). New
servers require tool discovery in the client; this existing chat may need a
new session before the reader tools appear. Prefer direct reader tools for source inspection. Use the existing PowerShell
build and native-check commands when their verification scope is required;
do not route unrelated inspection or editing through them.

### Access Explorer

Access Explorer 1.2.0 and its MCP backend are configured on this workstation.
The runtime is under the extension's user-level managed storage, with its own
Python 3.12 virtual environment. Access Explorer user settings select that
interpreter and backend script. VS Code's user MCP configuration registers
`access-explorer-local`; Codex's user configuration registers `access_explorer`.
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
  production migration, recovery, reproducible builds, or release gates. Close
  only owned sessions normally after preserving edits. The backend close path
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

For r128, the user explicitly waived old-output fixture retesting. The existing
`-CurrentFixturesOnly` switch records that limited scope; it is not a declaration
that older-schema import support ended. Future historical test policy remains
unresolved. Do not recreate removed outputs or silently omit required checks;
record exactly which fixtures were verified. See PROJECT_STATE.md.

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

2. **Reuse established test runners.** Before constructing an ad hoc Access COM harness, inspect whether an existing focused test runner supports the required validation. Prefer extending or parameterizing an established runner over repeatedly recreating equivalent automation.

3. **Use MCP where it provides native value.** Prefer Access Explorer for installed-object inspection, VBA references, control properties, native callback discovery, and other operations that would otherwise require custom Access automation. Do not use MCP merely to replace efficient source searches.

4. **Avoid excessive process polling.** Prefer execution mechanisms that report completion. For long-running builds, use appropriately spaced status checks and inspect logs when execution completes, fails, stalls, or reaches a meaningful milestone. Do not repeatedly inspect unchanged logs without a diagnostic reason.

5. **Avoid redundant validation.** Test targeted changes before integration, then run the required fresh-build verification. Repeating equivalent tests is justified when the artifact, environment, or validation purpose differs; otherwise avoid duplicate runs.

6. **Diagnose failures before retrying.** When native automation or a test command fails, capture the actual error and determine its cause before rerunning or modifying the command. Do not repeatedly change test code merely to satisfy an unexplained failure.

7. **Preserve useful evidence.** Record meaningful test outcomes, baseline comparisons, and unresolved failures without generating unnecessary permanent logs, reports, or documentation.

8. **Prefer minimal orchestration.** A single well-scoped tool call is generally preferable to a multi-stage custom automation procedure when both offer equivalent correctness, safety, and diagnostics.

These are efficiency preferences, not restrictions that prevent necessary debugging, native testing, or broader validation when justified.