# Unified Members: development orchestration review

22 October 2026. Review only: existing edits, source, infrastructure and tool
configuration are unchanged. No new native experiments were run.

**Greatest expected improvement:** finish focused functional and visual iteration
before the final integrated build, serialize Access controllers, and stop
repeated capture experiments once a capability or reliability limit is known.
Existing rules already cover much of this; tighten their execution criteria.

## Findings and preferred replacements

| Issue | Session evidence and capability comparison | Minimal change / expected benefit |
| --- | --- | --- |
| **1. Repeated embedded C# / Win32 commands** | Window enumeration, dialog dispatch, capture, positioning and cleanup were repeatedly recreated. Verified Explorer procedure/control reads and exports cover ordinary inspection, but not runtime debugging or every modal-dialog diagnostic. | Use Explorer inspection or the native VBA debugger first. Require a capability gap and stop condition before one exceptional Win32 fallback. Less fragile disposable code. |
| **2. Duplicated Access COM orchestration** | `access_export_text` and `access_close` succeeded in this session; the earlier capability audit verifies procedure reads, control inventory and relationships. Repeated raw `LoadFromText` / compile / Eval / Quit wrappers overlap the existing focused runners. Enabled MCP SQL returned “Undefined function” for both `SwitchContributor` and `UXShowGroupMembers`; arbitrary VBA execution tools are absent. | Use Explorer for installed-object inspection; use the existing group runners for stateful VBA behavior. Do not route UI setup through SQL or replace native compilation with static syntax checks. High reduction in wrapper duplication without losing coverage. |
| **3. Repeated full builds** | Three full-generation attempts occurred: one failed during theming, one passed BuildAndTest, and a final Build followed an empty-state wording change. Successful generation took **458.6 s** and **441.7 s**. The first build began before the empty-grid assumption was resolved; the last began after visual review found stale wording. | Extend the existing focused test first; settle empty/populated layouts and wording on a disposable copy, then freeze inputs and run the final integrated build. Avoiding the late wording rebuild could save about **7.4 minutes** in a comparable session. This is a counterfactual opportunity, not a measured optimization. Changed final inputs still require valid final verification; never reuse a stale receipt. |
| **4. Excessive polling** | Short execution polls repeatedly accompanied unchanged log tails and Access process/CPU checks. Existing workflow commands report stages, completion, failures and receipts. | Wait on the running execution/task at 30–60 second intervals. Inspect logs/processes for failure, meaningful milestones or suspected stalls. Fewer calls; no new monitor. |
| **5. Repeated screenshot techniques** | The first capture confused `hWndAccessApp()` with a property; another attempt stalled, and ordinary desktop captures initially included an obscuring Explorer window. Foreground captures eventually showed the intended layouts. MCP screenshot was not exercised: its schema lacks OpenArgs, while Members needs group context. Inspection of installed `ui.py` shows the same synchronous `PrintWindow(..., 2)` mechanism; its open timeout does not bound capture. | Define the required UI state before capture. Explorer screenshot is a candidate for directly openable forms, subject to a scratch-copy probe; it is **not yet a verified reliable substitute** here. Use the focused runner to establish context and a normal manual capture when necessary. After a capture failure, diagnose once rather than cycle through APIs. Very high reduction in exploratory work. |
| **6. COM failures and cleanup** | Explorer and direct COM used the same inspection copy before the Explorer session was closed; a fresh-copy compile then succeeded. Build and test automation also overlapped. One build reported RPC failure in `finally` at `Quit`, potentially masking the original error. A stalled synthetic preview needed owned-PID termination after normal close failed. Concurrency is a risk factor, **not an established cause** of that build failure. | Default to sequential native work and one controller per scratch database. Hand off explicitly with the owner's normal close. Explorer close is verified but is not risk-free: inspected backend tracks spawned/attached ownership and has a force-kill fallback for spawned sessions. Existing runners/build cleanup remain appropriate; switching to MCP does not eliminate COM. Preserve the primary error before cleanup and identify ownership before recovery. High reliability benefit; no quantified speed claim. |

## Recommended rule edits

Refine existing sections instead of duplicating the eight-item efficiency checklist.

In [AGENTS.md](../AGENTS.md), under tool selection / Access sessions:

> Default to sequential native Access automation and one identified controller
> per disposable database. Close the owning controller before handing that file
> to another tool. Recovery must identify the owned database and PID; never
> select processes for termination solely by application name.

> Before introducing embedded C# or Windows API automation, document the missing
> capability in verified existing tools. Permit one bounded diagnostic fallback
> with a stop condition.

In [Development_Workflow.md](../docs/Development_Workflow.md), refine
“Efficient Investigation and Execution”:

- **Build sequencing:** Settle focused functional/visual cases before freezing
  inputs for integration. Changed inputs or diagnosed failures still justify rebuilding.
- **Waiting:** Use execution completion with 30–60 second waits; additional
  log/process checks need a diagnostic purpose. Keep progress updates separate.
- **Capture:** Check context, event/save effects and timeout scope first. If MCP
  cannot reach the state, retain the runner/manual workflow; do not enable new
  execution permissions or cycle through capture implementations.
- **Recovery:** Record error, stage, database and ownership before cleanup;
  release the stale owner before one diagnosed retry. If runner cleanup masks
  the original error, retain logs and report uncertainty. Documentation cannot
  repair that implementation issue; any runner fix needs separate scoped work.

## Evidence and limits

- [Feature receipt and known failures](../build/workflow/unified-members-85a9f12d-ef8d-4c07-aa44-8f4191310679/verification.json), [failed build log](../build/workflow/20261022-090247-51ae901c-c1e3-4d8c-a448-b9cd7aca84e0/03-build.log), [BuildAndTest timing](../build/workflow/20261022-090741-282e69cf-4c5d-41d8-a066-d8bca63133d9/verification.json), [final Build timing](../build/workflow/20261022-092532-83f31e78-a40f-4c0d-9713-70c0bb660fb9/verification.json).
- [Verified Explorer scope and safety limits](ACCESS_EXPLORER_CAPABILITY_AUDIT.md); current callable schemas and installed screenshot/session implementation were inspected read-only. Availability is not reliability certification.
- Reuse [group-exit](../scripts/test-group-exit.ps1), [lifecycle](../scripts/test-ux-group-lifecycle.ps1), [development](../scripts/develop.ps1), and exact-fingerprint [artifact verification](../scripts/verify-build-artifact.ps1). Explorer compilation may decompile first; static syntax/lint cannot replace native or visual acceptance.
- No new framework, helper, server or configuration change is proposed. RPC-disconnect and capture-stall root causes remain unproven.
