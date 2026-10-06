# Summary-group correctness - r31, first increment

PR #15 merged into main at 8c79cd7. Every existing local feature branch is already
an ancestor of main. New work is on codex/summary-group-correctness.

The live group findings now withhold provisional promotion capacity when observed
members have unresolved or conflicting grouping keys. They require a resolved
worn paygrade, form kind, ending date, promotion status, exactly one duty status,
report type and reporting-senior identity, plus supported officer category or
an enlisted UIC. ACT and TAR still group together for enlisted personnel.

NOB reports remain visible as findings in a local planning roster, but contribute
neither to the observed denominator nor to EP/MP usage. Their different keys do
not alter the observed roster's comparison. Over-capacity observed rosters get a
policy issue, subject to the existing historical-applicability downgrade.
All findings retain verified instruction/page/paragraph/excerpt references.

No drafts, memberships, revisions or approvals are mutated by diagnostics.
Incomplete drafts remain editable and can stay in local planning groups. Local
capacity remains provisional: it does not establish complete command membership.

Verification: nine targeted assertions cover persisted roster counts, ACT/TAR,
NOB with conflicting keys, over-capacity, mixed status, missing senior identity,
multiple components, wrong grade and unsupported officer category; diagnostics
are checked for no revision/operation mutation. Existing 59 live validation/UI
assertions and all 20 fingerprinted policy excerpts pass. The populated build
compiles and all 57 installed modules match source; native event guards pass.

Preview: output/ux/Swift_NAVFIT_UI_Demo_2026-10-05-r31.accdb
Blank: build/ux-r31-blank.accdb
Deployment 0.23.0-preview is preserved; these are later development artifacts.

Next: explicit enlisted multi-UIC and officer dual-hatted review, supported
LDO/CWO category resolution, completeness evidence, and before/after membership
and bulk-edit consequences. Capacity integration with comparison and export must
follow; their current compatibility safeguards are not full policy certification.
No schema change, printing change or automatic membership reassignment is added.
