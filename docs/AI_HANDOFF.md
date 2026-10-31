# Shared AI handoff

Updated 30 October 2026. Read this at startup, then verify Git status/diffs,
versions and artifacts. Requirements remain in [PROJECT_STATE](../PROJECT_STATE.md),
[backlog](../planning/PRIORITY_BACKLOG.md), [AGENTS](../AGENTS.md) and the
[development workflow](Development_Workflow.md). Private conversation history is
not needed. Commits need the user's approval; pushes need it too (Matt approved the 31 October push).

## State

- Branch main, in sync with origin/main after the 31 October push. APP_VERSION 0.49.0 /
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

## Narrative line-fit meter (31 October, committed a72eed8)

Plan: project doc "Plan: narrative line-fit meter and spell check", section 1;
Matt then asked for NAVFIT98A field sizes, CPO separate from EVAL/FITREP.
Committed and pushed:
- modPDFOutput: `PDFFieldFit` (Swift PDF, shares FitText via `BoxLines`) and
  `NAVFITFieldFit` (GDI+ `GdipMeasureString`, NAVFIT98A 2.2.0.332 print
  rectangles at default pitch; limits comments 18 / CHIEFEVAL 6, blocks 28/29
  3/4, Chief performance comments 2). Swift has no `CommentsPitch`, so exports
  print at default pitch; the 12-pitch branches are not modelled.
- Editor meter `UXNarrativeStatus` (NAVFIT98A lines; Swift PDF in the tip),
  Information finding `APP-NARRATIVE-FIT` (not in saved readiness or the issue
  count), guide/help/PROJECT_STATE, SBOM/views regenerated.
- Tests in test-narratives: `TestNarrativeLineFit` and `navfit-line-fit` (99
  cases against a System.Drawing reference built in the suite script).
- Verified: BuildAndTest PASS build/workflow/20261030-184108-d9fdb6a2-1aec-4dc9-9fd6-badcd9c426ee;
  test-narratives PASS; test-pdf-output PASS, 28 PDFs / 390 pages identical in
  content, images and text to HEAD build 8b82e582. Codex reviewed twice; fixes applied.
- Not verified (Matt's attended print is backlog H03): a real NAVFIT98A print (CHIEF comments rectangle fits 7 lines vs
  its 6-line limit; block 28 third line and Chief comment second line are slightly
  clipped in the rectangles); the no-printer path; guide PDF visual check.

## Spell check (31 October, committed 289093e; word list in the next commit)

Plan section 2, committed and pushed:
- `UXSpellCheck` / `UXSpellKey` / `UXSpellApply` (modUXSurface): Check spelling
  button by the comments (`cmdSpelling`) and F7 (frmUXEditorFields class module,
  KeyPreview) run `acCmdSpelling` on the comments, blocks 28/29 and Chief
  performance comments; corrections go through `UXSurfaceSetValue`; signed
  reports ask to reopen.
- Navy word list, Matt's option (b): about 800 terms drafted by a fast model and
  reviewed, embedded in `NavyWordList` (modUXResourceTools). Resources > Save
  Navy word list for spell check writes a UTF-16 .dic; the user adds it under
  File > Options > Proofing > Custom Dictionaries. Swift never changes Office settings.
- Verified: BuildAndTest PASS build/workflow/20261030-202555-8fcc32bf-6941-41e1-a3e7-0409af38df5e;
  test-narratives PASS (1,686 spelling and word-list assertions); test-workspace-presentation PASS.
- Needs Matt attended: the Office dialog (corrections, cancel, Add to
  Dictionary), adding the saved .dic in Office, and a workstation without
  proofing tools.

## Visual walkthrough (31 October, committed locally, not pushed)

Matt asked for a full visual pass. 122 screens were captured on a synthetic demo
with runtime geometry, triaged by Sonnet agents and verified, then fixed: 22-pt
titles 500 -> 600 twips on 56 forms, editor page title/header/group box, two-line
block 22/49 captions (`modReportPages`), `&` escaped in field captions, Report
tools sort headings placed per view (`UXToolsRefresh`) and corrected Validation
headings (`modTableSort`), Mark signed button styling, overlapping status labels
(group Members, workspace import), RUAD mojibake and duplicate header, captions
following their field on wide windows (`UXElasticPlace`). Report geometry,
captions and PDF code are unchanged. Proposals awaiting Matt (header right column,
trait-grid headings, raw list headings, rail height, Overview, button rows,
contributor start) are in the project doc "Swift NAVFIT visual walkthrough:
fixes and proposals". Capture/triage scripts were scratch only, not in the repo.
Verified: BuildAndTest PASS build/workflow/20261030-221106-cf50613f-73b4-4795-96b1-e4be89dd9608;
test-workspace-presentation and test-report-navigation PASS on that build;
test-traits, test-summary-groups, test-pdf-output PASS on the prior build
(before the last viewport refinement). Not checked: other DPI/scaling, physical print.
Unrelated finding: docs/architecture/views/SV-7.md has mojibake (en dash, arrow).

## Exact next actions

1. Editor page-layout cache is done (b5e7e54); editor open stays about 1.1 s by
   user decision. The slow fourth open only happens back to back (see M03).
2. L03 steps 1-2 done (size warning; no first-save empty-to-empty change rows).
   Step 3 (sparse revision values, schema change) is next only if the user
   schedules it.
3. Optional: diagnose the test-import-engine 3059 flake.
4. Spell check and the Navy word list are committed; Matt's attended checks are listed above.
5. Update this file at milestones and before stopping.
