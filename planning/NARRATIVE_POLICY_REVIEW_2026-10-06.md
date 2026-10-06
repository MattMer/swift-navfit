# Narrative expansion, policy coverage and leadership tools â€” 6 October 2026

Development candidate: `output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r75.accdb`.
Blank candidate: `build/ux-r75-policy-blank.accdb`. Schema remains 23; published 0.23.0-preview/schema-21 release is unchanged. This is a development build, not Navy certification.

## Narrative sources implemented

Resources > NavyWriter examples catalog now contains **92 entries**: **72 rate/source collections and 20 short, attributable excerpts**, across 72 source rate/category filters plus General. Search and rate filter intersect. Collections open their full source and cannot be imported as invented narrative; excerpts can be adapted and added deliberately. Existing attribution and duplicate guards remain.

Resources > Narrative snippets also has a rate filter and an editable rate selector. Imported examples retain their source and rate; manual snippets can be assigned a rate. Existing untagged snippets appear under Unassigned / legacy. Rate metadata uses an explicit `[Rate: ...]` attribution tag, retaining schema compatibility. Attribution remains visible and the existing 255-character limit applies without truncation.

The catalog is an examples index, not an authoritative current rating roster. Legacy CTN, PC, SH, SK and AV source collections are explicitly identified; CWT and CTN remain separate. Source category codes such as FN/AW are retained as source categories, not asserted to be current distinct ratings. No mass download of complete narrative collections is distributed. All **78 distinct source URLs returned HTTP 200** during this audit; future availability is external.

Source of truth: `src/navywriter-catalog.json`; generator: `scripts/generate-narrative-catalog.py`. Discovery and link-health evidence: `build/research/narrative-sources/`. Recheck source text before replacing excerpts. Examples must reflect verified member accomplishments.

Catalog reference: https://www.navywriter.com/navy-eval-bullets-by-rating.htm

## Policy source and verification

MyNavyHR's instruction index lists **BUPERSINST 1610.10H CH-2, 26 May 2026**. The supplied `BUPERSINST 1610.10.pdf` is already that compilation, including retained pages bearing earlier change dates. Its SHA-256 is **5E71CA59D66C509640B48845960D236670ABFB5D57AE76435A7C7983F0DC9058**. Live web retrieval of some MyNavyHR pages returned 403; detailed rules were reviewed against the local fingerprinted instruction, not an assumed freshly downloaded PDF.

`scripts/verify-validation-policy.ps1` now verifies **37 exact source excerpts and their actual PDF-page locations**, in addition to PDF/extraction fingerprints. Findings show instruction, enclosure, printed page, paragraph/block, PDF page and excerpt. Application storage rules are labeled separately. Historical/future report dates prompt applicability review; this compilation is not automatically asserted to apply to every past report.

### Coverage added in this increment

| Check / behavior | Instruction locator | Treatment |
|---|---|---|
| Missing ship/station, promotion status, date reported | 1-2, blocks 7â€“9; printed 1-3/1-4 | Incomplete draft; save permitted |
| Missing reporting senior name, grade, title and UIC | 1-2, blocks 22, 23, 25, 26; printed 1-9â€“1-11 | Incomplete draft |
| UIC first four characters numeric; five-character representation | 1-2, block 6; printed 1-2 | Invalid entry; fifth character permits letter/digit |
| Counseling date within report period; NOT REQ / NOT PERF support | 1-2, block 30; printed 1-13 | Cited finding; official status text saves/imports/exports |
| Counselor missing after actual counseling; NOT PERF explanation; SEE COMMENT without comments | 1-2, block 31; printed 1-13 | Incomplete draft |
| First career recommendation missing | 1-2, career recommendations; printed 1-15 | Incomplete draft; NA/NONE valid; second optional; whole/partial NOB exception retained |
| Partial NOB has more than three numeric traits, or blank ungraded traits | 6-3; printed 6-1 | Policy/incomplete findings; remaining traits must be NOB |
| Partial NOB adverse marks | 6-3; printed 6-1 | Policy finding for 1.0, three 2.0s, or climate/character below 3.0; adverse narrative still manual |
| Adverse-trait justification | 13-4a/b; printed 13-3 | Review warning; app cannot judge adequacy |
| Significant Problems / recommendation against retention justification | 13-4c(1); printed 13-3 | Review warning |
| Adverse-report referral to member | Introduction 4aâ€“c; printed I-2 | Review warning; statement checkbox/local approval is not proof of referral |
| Regular report exceeds 15 months | 3-2; printed 3-1 | Review PERS-32 approval and extensions; exceptions are not hard rejected |
| Both Yes and No selected for statement or applicable retention response | Application consistency | Invalid entry; no fabricated policy citation |

These checks run against pending report values and **all summary-group members**, including open-member overrides. They do not create revisions, validation runs or journal entries. Counseling status support was corrected across service validation, native field validation, PDF import and NAVFIT conversion. Readiness dates still require real dates.

### Existing coverage retained

Storage lengths/options/digits/date validation; required member identity, rank, UIC, SSN, report period and reporting senior SSN; component exclusivity; supported Regular/Concurrent combination; Special occasion exclusivity; whole-NOB traits/career/promotion rules; 1.0/2.0 promotion restrictions; chief Accountability / climate critical traits; count of 2.0 traits; ENS/LTJG non-LDO restrictions; group type/paygrade/key consistency; officer competitive categories and UIC breakout authority; eligible roster and provisional EP/MP capacity limits; immutable saved-revision checklist/rule runs; RSCA arithmetic and provenance guards.

### Rules that still require data, judgment or separate work

- **Regular continuity and occasions:** reliable gap/overlap detection needs prior report type, Regular/Concurrent distinctions, extensions and departure/promotion event dates. Current prior history stores period, traits and recommendation but does not establish these facts.
- **Reporting senior authority / delegation:** chapter 2 waivers, assigned/dual-hatted seniors and unusual civilian/foreign-service designators need explicit context. Do not reject every senior using a Navy-officer-only designator rule.
- **Specialized reports:** chapters 4â€“12 cover Concurrent, Operational Commander, NOB, adverse/punishment contexts, education, Reserve/INACT, administrative corrections and PIM exceptions. General field checks cannot establish applicability of every exception.
- **Unmodeled official content:** command employment/block 28, EVAL qualifications/block 44, chief hard breakout/summary ranking, complete signature facts and some native-only summary values remain preserved where imported but are not fully editable/validated as Swift saved fields.
- **PFA/PRIMS:** instruction codes, cycle facts, actual PRIMS results and report-date applicability need a dedicated verified model. Do not infer results from free text or invent an online integration.
- **Narrative substance/prohibitions:** misconduct disposition, prohibited references, mandatory GRGB/SAPR/command-climate statements and required block-number justification need a contextual checklist plus human judgment. Blank-text tests cannot certify compliance.
- **Final processing:** signatures, flag endorsements, member statements, transmission, receipt, official record acceptance, correction/appeal deadlines and paper-copy requirements remain outside local approval. A completed local group is not an accepted Navy submission.
- **Retention/privacy:** instruction Introduction paragraph 7 limits retention of notes/drafts and distinguishes officer/enlisted records. Plan explicit retention review and deliberate disposal; no automatic deletion has been introduced.
- **Historical rule packs:** the older immutable persisted pack still has its recorded applicability window. New live checks do not silently alter past validation runs. A new effective-dated pack and result snapshot are needed for complete reproducible policy-history support.
- **Printed fit:** character counts and a locked font are insufficient to prove wrap/line behavior. Exact NAVFIT print-output acceptance remains outstanding.

## Research-backed feature recommendations

These are proposed tools, **not implemented by this increment**. Existing directory/RUAD work, reporting-senior history, group comparison and review routes are not counted as new features.

| Priority | Tool | Why / useful first scope | Dependency |
|---|---|---|---|
| 1 | Member report-continuity timeline | Show Regular coverage, gaps/overlaps, extensions and next due event; link each segment to its report. Prevent missing periods and distinguish Concurrent coverage. | Add explicit prior report type and extension provenance; preserve unknown historical classifications. |
| 2 | Ready-for-signature checklist | Combine deterministic findings with required human confirmations, exact policy references and unresolved review notes. Show what blocks readiness and who confirmed exceptions. | Versioned rule pack; contextual facts; no official-signature claims. |
| 3 | Counseling schedule and worksheet | Track due/completed counseling, goals and follow-up; prefill report block 30 only when within its period. Useful throughout the year rather than only at evaluation time. | Effective-dated periodic schedule and explicit NOT REQ/NOT PERF reasons. |
| 4 | Member accomplishment / evidence log | Record date, duty, action, measured result and source; filter to report period and insert reviewed facts. Reduces reliance on generic copied bullets. | Directory identity and contributor attribution already available. |
| 5 | Submission / returned-report tracker | Track ready, signed locally, sent, receipt, accepted/returned and correction task. Map official error codes to affected blocks. | Local manual evidence; no network integration or inference of acceptance. |
| 6 | Policy-version and coverage dashboard | Explain which checks applied on a date and which facts still need human review; retain the exact source/rule version with final readiness results. | New immutable packs; maintain old results unchanged. |
| 7 | Retention review and workspace access guidance | Review old draft notes/copies, document deliberate retention/disposal, make local attribution versus authentication explicit. | Current source policy and command authorization; no automatic purge. |
| 8 | Narrative quality aids | Detect unresolved placeholders, unusually reused text and missing user-entered evidence; maintain command-approved abbreviations. | Advisory only; no automated promotion/ranking or factual invention. |

Primary research sources:
- Current instruction: https://www.mynavyhr.navy.mil/References/BUPERS-Instructions/
- Mid-term counseling resources: https://www.mynavyhr.navy.mil/Career-Management/Performance-Evaluation/Mid-Term-Counseling/
- Official error-code/block mapping: https://www.mynavyhr.navy.mil/Career-Management/Performance-Evaluation/Error-Codes-Reasons/
- Continuity / reporting-senior submission reports: https://www.mynavyhr.navy.mil/Career-Management/Performance-Evaluation/FAQ/q/
- Current templates: https://www.mynavyhr.navy.mil/Career-Management/Performance-Evaluation/Software-Forms/
- Future compatibility watch: https://www.mynavyhr.navy.mil/Career-Management/Talent-Management/TIDE/ â€” monitor official direction before adopting a different scoring system. This is research, not a planned network integration.

## Acceptance evidence and limits

- 29 catalog/source preview, rate intersection, legacy handling, manual rate assignment and snippet import assertions.
- 102 live validation assertions across all three report types, positive/negative exceptions, summary-group inclusion, navigation and zero-write behavior.
- 170 service-detail assertions plus 49 readiness/billet checks, including status save/reopen/native editor acceptance.
- 501 NAVFIT adapter, 410 exchange and 98 stale/hash/identity/lock/undo assertions; new status round trips included. This is automated data acceptance, not actual NAVFIT UI/printing proof.
- 250 expanded PDF extraction/mapping/merge/undo assertions and 12 synthetic PDF counseling-status extraction/save assertions. Test fixtures preserve XFA encryption and do not modify original PDFs.
- 37 source excerpts and PDF locators verified against the fingerprinted instruction.
- Final r75: 72 installed source modules match; 36 populated-demo, 30 guide and 29 populated-catalog assertions pass. Schema-23 transfer preserves all original rows/fields and the source hash. r74 automated responsive checks pass 1,972 assertions; no physical-DPI claim. Evidence is recorded in CURRENT_STATE.md and the progress log.
- Physical display/DPI, screen-reader and native NAVFIT UI round-trip acceptance remain outstanding. No release, Git push or deployment occurred in this task.
