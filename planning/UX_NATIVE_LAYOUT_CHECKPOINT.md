# Native layout checkpoint - 2026-10-04

The user correctly identified that candidate9 did not resemble the reviewed
browser prototype. It was a partial functional adapter, not a faithful native UI.
Do not describe it as the implemented redesign.

Candidate13 adds a consistent left navigation rail to thirteen main workspaces,
selected destination styling, Segoe UI, content bounds, an Overview with four
journey cards, and a Reports folder filter instead of the legacy folder tree.
The report type selector is visible. Existing service-facing control identities
and editor buffer remain intact. Existing workspaces are not modified.

The new folder-filter callback was initially inserted ahead of module-level
declarations, breaking their scope. The corrected module places all declarations
before procedures. Access RunCommand(126) alone failed to establish executable
validity; behavioral execution must accompany build checks.

Verified on build/ux-candidate13.accdb:
- 38 reporting senior, RSCA and comparison assertions.
- 34 recovery assertions.
- 6 workbench UI assertions and 75 visible-control bounds checks.
- 46 smoke assertions, including setup, three report types, atomic saves,
  archive guards, folder filtering, editor actions and reopening.
- 560 navigation/content bounds checks.

This remains an intermediate development artifact. It has not been visually
verified against the prototype, checked at Windows scaling levels, or tested
with a screen reader. Native editor pagination, the Overview's missing recent
reports, long workflow screens, inline validation, and persistent navigation
while vertically scrolling still require redesign. Bounds checks do not prove
readable labels, absence of overlap, or visual fidelity. The rail is currently
within form detail, rather than a fixed shell hosting destination content.

The published 0.21.0 release and existing workspaces are unchanged.
