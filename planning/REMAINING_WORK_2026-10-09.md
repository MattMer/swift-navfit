# Remaining work after r97 — 9 October 2026

## Confirmed and newly completed

The user confirmed r96 fixes the interactive 3218/3420 comparison save failure. r97 adds reviewed atomic RUAD CSV import/refresh with all 35 fields and corrects archive-filter wording. Existing manual-feedback items are implemented apart from exact comments-on-performance fit verification. Research proposals are not counted as accepted implementations.

## In order

1. Narrative fidelity (manual feedback item 23): use the installed NAVFIT98A comments control and print behavior to establish exact fonts, dimensions, wrapping and overflow for all three report types. Existing character counts/source inspection are insufficient. Do not claim printed fit yet.
2. Native compatibility acceptance: open full-field report/group exports and returned-file scenarios in actual NAVFIT98A, proving EVAL/CHIEFEVAL/FITREP round trips. Automated adapters and exchange checks already exist; native UI/output proof remains outstanding.
3. Native usability acceptance: keyboard and screen reader, physical laptop/desktop displays at 100/125/150 percent; remaining report-column truncation and narrow-screen layouts. No computer use per user preference; COM behavioral checks do not establish physical accessibility.
4. Release readiness: verify schema-20 through schema-25 transfer, including RUAD XML provenance, exact blank/demo artifacts, installed-source audit and target-computer PDF extraction dependencies; update guides and publish a separately tested prerelease when requested. Current published release remains 0.24.0-preview.
5. Expand policy coverage with verified effective-dated rules and human exception facts: continuity/occasions, PFA/PRIMS context, substantive mandatory/prohibited statements, historical applicability and reproducible result snapshots. Existing live findings retain policy locators/excerpts; no complete-policy certification is claimed.

## Further directory refinements

RUAD missing-member reconciliation view, repeat-refresh match suggestions with explicit confirmation, and journaled undo are useful follow-ups. Current refresh does not archive absent people and preserves manually edited fields. Source matching must never infer identity from names/BIN alone.

## Research suggestions awaiting prioritization

Member report-continuity timeline, ready-for-signature checklist, counseling schedule, accomplishment/evidence log, submission/returned-report tracker, policy coverage dashboard, deliberate retention review and advisory narrative-quality tools. Source review and dependencies are in NARRATIVE_POLICY_REVIEW_2026-10-06.md.
