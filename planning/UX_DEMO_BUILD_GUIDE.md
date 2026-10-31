# Build and review the synthetic Demo

Use the existing [Demo builder](../scripts/build-ux-admin-demo.ps1) with a
verified current blank template and a new destination. If a fresh blank is needed,
use [develop.ps1](../scripts/develop.ps1) first:

```powershell
powershell.exe -NoProfile -File scripts/develop.ps1 -Action Build -BuildPath "build/dev/Swift_NAVFIT_demo_template.accdb"
powershell.exe -NoProfile -File scripts/build-ux-admin-demo.ps1 -TemplatePath "build/dev/Swift_NAVFIT_demo_template.accdb" -OutputPath "build/tests/Swift_NAVFIT_demo_review.accdb"
```

Both destinations must be new. The Demo builder verifies the blank through
read-only DAO before creating its copy, seeds synthetic examples, removes seed
modules, restores the editor's normal unload handler and startup macro, and
compiles. It then runs the populated Demo journey and normal/enabled-content
startup checks on disposable copies. The template hash must remain unchanged.
A failed check does not produce a success report or authorize publication.

The old build-preview entry point is in [PROPOSED_PURGE](../PROPOSED_PURGE/README.md)
for review. It is not an alternate supported build command. Its pre-seed suites
remain available as focused tests; creating a Demo no longer automatically runs
all historical feature suites.

## Check the changed behavior

| Changed area | Focused scripts retained under scripts/ |
| --- | --- |
| Review return / My Work | test-ux-review-return.ps1; test-ux-my-work.ps1 |
| Validation and citations | test-ux-validation.ps1; applicable readiness/policy checks |
| NAVFIT exchange | test-navfit-adapter.ps1; test-navfit-group-export.ps1; test-ux-navfit.ps1 |
| PDF / guided import | test-pdf-expanded.ps1; test-ux-import-flow.ps1; applicable local-PDF checks |
| Context / group defaults / shared blocks | test-ux-context.ps1; test-ux-group-lifecycle.ps1; test-ux-block-tools.ps1 |
| Native surface / viewing area | test-ux-surface.ps1; test-ux-responsive-identity.ps1 |
| Help | test-ux-help.ps1 |
| Pending edits and normal exits | test-ux-workflows.ps1 |

These checks have distinct assertions; they were not declared equivalent to
each other or deleted. Run those relevant to a change, and use
[release-check.ps1](../scripts/release-check.ps1) plus applicable feature gates
before publication. The former IncludePDFIntegrationChecks option belonged to
the retired builder; invoke PDF checks directly when relevant. Demo creation is
not a release-verification receipt.

## Try the Demo

Open it in full desktop Access. Enable Content if prompted, then close and reopen.
See [contributor startup](../docs/Contributor_Startup.md): setup is complete but
each initialized launch requires an explicit contributor choice; nothing is
preselected. Local contributors provide attribution, not authentication.

The synthetic workspace contains 13 reports, five groups, two contributors,
one reporting senior, four RSCA baselines and 12 prior evaluations. Practice:

1. Open Jamie Taylor, edit a field, switch Front/Back pages, correct an invalid
   SSN and save. Pending values should survive page changes.
2. Inspect the E5/E7/O3 groups and reporting-senior baselines; compare marks.
3. Inspect Jordan Reed's local approval, Quinn Park's archive and Sam Rivera's NOB.
4. Use My Work for current and intentionally stale review examples.
5. Recover or discard Jamie Taylor's uncommitted checkpoint.
6. Compare the clean E6 group with the group containing three deliberate report issues.
7. Inspect synthetic member history and shared report conversations.

Report fields use fixed geometry; window resizing changes the viewing area.
All seven traits remain together on Front Page. Browser prototype state is separate.
Native checks do not establish physical display, keyboard/screen-reader,
scaling or printed fit. NAVFIT98A remains the final native validation/printing
destination. Structured PDF extraction is local; it does not require Acrobat.
Never use a populated user workspace as the seed template or import real data
into a synthetic Demo. Follow [CONTRIBUTING.md](../CONTRIBUTING.md#releases-explicit-request-only) for any release delivery.
