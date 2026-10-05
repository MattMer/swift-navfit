# Workspace navigation performance - r30

Adaptive sizing is retained. UXResizeWorkspace caches the successfully applied
width and height for each form instance. Repeated requests at unchanged sizes
return without moving controls. Closing/reopening a form uses a new instance;
UXRestoreViewport invalidates the cache when switching contextual layouts.

UXLoadContext updates captions only when their text changes, and requests
geometry only for the active workspace. Hidden open forms no longer receive a
layout pass from context refresh. Existing activation data refreshes remain so
report, group and review changes are visible when returning. Redundant Requery
calls after RowSource assignment were removed from the My Work inbox, Overview
recent reports and summary-group list. No timed data cache was introduced.

The editor's section-specific layout remains intact: changing sections, errors
or lock status can require layout even with unchanged window dimensions.

Verification: compile and exact installed source audit (57 modules), native
editor/review/list event guards, 1801 responsive layout/focus/identity assertions
including repeated-size skip, changed-height layout, contextual invalidation and
hidden-tab isolation. Native wall-clock tab-switch timings are not measured.

Demo: output/ux/Swift_NAVFIT_UI_Demo_2026-10-05-r30.accdb
Blank: build/ux-r30-blank.accdb
Version/schema unchanged: 0.22.0-ux-preview / 21. Existing workspaces preserved.
