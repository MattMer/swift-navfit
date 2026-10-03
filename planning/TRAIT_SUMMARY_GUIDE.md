# Trait summaries

Version 0.11.0 adds two read-only summaries below the seven trait marks: Editor values and Saved revision. The first updates when a mark is changed or focused and after checkpoint recovery. The second reads the latest committed revision and refreshes after saving. Opening or refreshing these summaries creates no revisions or checkpoints.

The numeric average is the sum of numeric marks divided by their count, displayed to two decimal places using the workstation's number format. NOB and blank marks are excluded; their counts are shown separately. When blanks remain, the average is labelled partial. With no numeric marks the average is unavailable, rather than zero. Unsupported pending input suppresses the editor average while preserving the saved summary.

These are descriptive arithmetic summaries, not official Navy assessments. They do not determine whole-report observation status, promotion recommendations, summary-group eligibility or policy compliance. No derived average is stored or exported. Schema remains 10; use upgrade-workspace.ps1 to carry a populated workspace into the new application at a fresh path.

Verification: 44 focused summary assertions, the existing 425 behavioral assertions and 169 catalogue checks passed. All 2472 retained import values were verified. A populated schema-10 migration preserved every original row/field, unchanged source hash and successful reopen. Runtime/x86 and comprehensive visual testing remain pending.
