# r25 - live report and summary-group validation

Version remains 0.22.0-ux-preview, schema 21. This development preview adds live,
read-only diagnostics and policy explanations; it is not a Navy-ready certification.
Existing workspaces, earlier artifacts and immutable historical rule definitions remain intact.

## Use

- In a report, select the action-area status/issue button, or Review > Validation issues.
  Text changes update the issue count without saving a revision. Checkbox and list
  changes update through their normal edit events. Recovery checkpoints remain separate.
- In any focused summary-group view, select All summary-group validation issues.
  The list includes every member's saved values and the open member's pending values.
  Pages contain 50 findings; Previous/Next expose the complete result set.
- Select a finding to read its complete explanation and evidence. Open affected report
  block navigates to the actual field. Navigating within the current editor preserves
  pending values; switching reports uses the existing save-on-leave guards.
- Invalid entry, Incomplete draft, Policy issue and Needs review are text labels.
  Incomplete drafts and policy findings remain saveable. Existing storage validation
  still prevents invalid saves and preserves raw recovery values.

## Policy evidence and implemented coverage

Each verified Navy finding names BUPERSINST 1610.10H CH-2, enclosure (2), the printed
page, paragraph 1-2 and relevant block, note or table, a short excerpt, and source
SHA-256. The compilation is dated 26 May 2026; unchanged pages retain their earlier
30 July 2025 or 16 December 2025 headers. Source locators refer to the instruction's
printed pages; report-field captions continue to use current form mappings.

The supplied PDF fingerprint is
`5E71CA59D66C509640B48845960D236670ABFB5D57AE76435A7C7983F0DC9058`.
Its recorded extraction fingerprint is
`318926467046A731A0D77A5F57B6FE2849A3C886C2CF9287579B55824FF11F1F`.
`scripts/verify-validation-policy.ps1` checks both fingerprints and all 20 excerpts,
normalizing whitespace and quotation typography. Local reference files are excluded
from Git. An unverified custom rule receives an explicit missing-source notice;
application storage rules receive no invented Navy citation.

Implemented checks cover required core identification/period information, component
selection, report type, report occasions (including Special exclusivity), wholly
versus partially graded NOB, 1.0/2.0 promotion restrictions, critical observed traits,
the two-2.0 limit, ENS/LTJG non-LDO restrictions, group paygrade/type/ending date,
component grouping, promotion status, billet subcategory and reporting-senior
consistency. Officer competitive categories include URL, selected restricted-line,
information-warfare and staff categories. Enlisted ACT/TAR group together. Different
ratings may share the same current paygrade. EP and combined EP/MP capacities show
provisional local-roster counts against the existing tested capacity calculator.

LDO/CWO competitive categories, UIC exceptions, source-date applicability, partial
NOB justification, unknown abbreviations and complete eligible membership require
review. The app does not infer complete eligible membership or require a manual
roster freeze. Earlier/future report dates downgrade definite policy conflicts to
Needs review pending confirmation of the applicable instruction. Historical saved
rule-pack runs retain their original effective-period scope; new runs include
expanded citation text. Live checks never write ValidationRun, ValidationFinding,
rule packs or revisions.

## Verification and limitations

- 59 focused assertions cover source details, valid Concurrent/Regular selection,
  NOB distinctions, promotion restrictions, grade mapping, live text before
  AfterUpdate, all 78 pending fields, group pending values, selection retention,
  direct block navigation, pagination and absence of database writes.
- The expanded UXHelpText initially exceeded VBA's source-line limit. It is now
  split into continuation strings; all nine complete topics execute and match
  the canonical JSON and written guide. This is included in native verification.
- Existing behavioral suites and exact final-artifact verification are recorded
  below after completion. Source-preserving copies are used for tests and seeding.
- Manual screen-reader, physical display/DPI, Runtime/x86 and full policy acceptance
  remain outstanding. Bulk previews and staged meeting proposals are not included
  as pending editor values in the live group list; saved results are rechecked.
  Printed-layout work remains deferred.

## Final local artifacts and evidence

- Synthetic demo: `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r25.accdb`.
  SHA256: `1A9A04986CD4694E4C7E7E5EBFDA32FFD3C94A6D2AA9B97F1FB721364304BE3D`.
  Nine reports, three groups, two contributors and recovery/review/approval/NOB examples.
- Blank template: `build/ux-r25-candidate.accdb`.
  SHA256: `66983E8556D6180FB2D210B6BC741A3F7080E2C210BEA7A79A82D8A57D4F07C7`.
  All 42 user-data tables empty; AppMetadata retained.
- Exact installed-source comparison: all 52 source standard modules match and compile
  in both artifacts. Native editor unload/return hooks remain intact. Verification
  used disposable copies and left the deliverable bytes unchanged.
- Final preview gates passed: 145 review/return, 59 live validation, 32 My Work,
  21 guided NAVFIT import, 30 help, 23 calculation/checklist, 41 group lifecycle,
  90 block/comparison, 1793 responsive, 108 editor surface, 49 close/workflow,
  forced shutdown/recovery, 36 populated demo and normal/enabled-content startup.
  The 97 source-table/historical-rule assertions also passed on a copy of the final demo.
  PDF integration was not requested in this focused preview run.
- Native computer-use review on a separate seeded probe confirmed the empty report
  issue state, readable group rows and policy/page/paragraph/excerpt details. A group
  finding opened Sam Rivera's actual Performance traits section and focused trait 33.
  No operational workspace or final demo was used for interactive mutations.
  This is a focused visual/navigation check, not full screen-reader or DPI acceptance.
- Exact-source evidence: `build/tests/review-build-77207908-bc8d-4d4a-8f24-c51b0e63c04e/verification.json`
  (blank) and `build/tests/review-build-5e6c36e4-170d-4301-8c95-31673d18d3a2/verification.json`
  (demo). Local tests/probes remain excluded from Git.

No GitHub push, release publication or operational workspace migration was performed
for this feature milestone. Earlier artifacts remain available.
