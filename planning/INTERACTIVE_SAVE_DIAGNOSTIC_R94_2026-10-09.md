> Historical checkpoint and evidence. Current architecture/decisions are in
> [PROJECT_STATE.md](../PROJECT_STATE.md); current procedures are in the
> [user guide](../docs/Swift_NAVFIT_User_Guide.md). Old versions, preview builders
> and pending findings below describe that checkpoint, not current instructions.

# Interactive recording / r94 diagnostic — 9 October 2026

Read the supplied seven-second recording locally using a workspace-only PyAV decoder. The title bar identifies r93 TestWorkspace. Multiple trait dropdown choices revert; the same 3420 Save report revision message remains. This contradicts any inference that automated passes establish the user's journey fixed. No computer use or media upload was used.

Separate r94-Diagnostic workspace retains all 49 user tables unchanged. Local .save-trace.log records queue/drain/commit entry and stages; no names, SSNs, values or narratives are logged. Standard application source remains uninstrumented; copies of the instrumented modules are in build/tests/r94-diagnostic-sources and tests/modSaveDiagnostics.bas. Native compilation and 26 approval/save assertions passed. This is a diagnostic artifact, not a claimed fix or release. Next required evidence is one interactive failed edit and the generated log. Do not modify/overwrite r93 or the user's original edited database.

## r95 navigation candidate

Removed the visible legacy Members detour and duplicate LoadMajorWorkbench call when opening a populated group. ContextOnly opening skips the legacy Members rendering. Direct member opening uses one guarded screen-paint transition, retains the backend adapter hidden, and restores painting on success or failure. Group-tab navigation avoids rebuilding the legacy member list before opening the grid. Empty groups retain their existing member-management destination. Back explicitly reveals its legacy destination.

Candidate: output/ux/Swift_NAVFIT_UI_TestWorkspace_2026-10-09-r95-Diagnostic.accdb. Native compilation passed; scratch direct opening, matching group context, Members/Comparison tab switching with hidden adapter, and Back destination assertions passed. The initial test selected the list header instead of a data row; corrected harness passes. The 26 approval/save assertions passed. All 49 user tables compare identical to r93. Interactive flicker and the user's failing save still require native user verification; diagnostic trace remains enabled. No GitHub release/push made.
