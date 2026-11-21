# PERS-32 reject reasons and Swift NAVFIT checks

PERS-32 returns reports it cannot accept and records a numerical error code for each
reason. NAVPERSCOM publishes the codes on its
[Error Codes & Reasons](https://www.mynavyhr.navy.mil/Career-Management/Performance-Evaluation/Error-Codes-Reasons/)
page (read 20 Nov 2026). Its
[Rejected Reports & Trends](https://www.mynavyhr.navy.mil/Career-Management/Performance-Evaluation/Rejected-Reports/)
page lists the most common reasons, in order:

1. enlisted summary block errors
2. officer summary block errors
3. member signature missing
4. period of report
5. physical readiness
6. reporting senior signature missing
7. member signature missing on an adverse report
8. billet subcategory
9. supplements
10. reporting senior information

This table shows which Swift NAVFIT check catches each reason. Checks with an `RJ-`
code are in `src/modValidationRejects.bas` and quote the built-in EVALMAN
(BUPERSINST 1610.10H CH-4). A check marked **Review** asks a person to decide.

**Signatures cannot be checked.** Missing signatures account for three of the seven
most common reasons, and they are on paper. Swift NAVFIT reminds you that an adverse
report needs the member's signature (`RJ-ADVERSE-SIGN`).

## Report blocks

| Code | Reason | Swift NAVFIT |
| --- | --- | --- |
| 1 | Name does not match SSN | Not checked: needs the Navy's records |
| 2, 3, 5B, 8, 10-13B, 14-15D, 21A, 22-27 | Summary group members differ in a block that must match | Group checks (`GROUP-ENLISTED`, `GROUP-OFFICER`, `GROUP-CATEGORY`, `NV-RS-GROUP`, `NV-GROUP`) |
| 2A | Block 2 is not the grade or rate | `NV-RATE` (review) |
| 4 | Member SSN missing | `SSN-PRESENT`, `APP-DIGITS` |
| 5A | Duty status missing | `COMPONENT` |
| 6, 7 | UIC or ship/station missing | `UIC-PRESENT`, `UIC-FORMAT`, `SHIP-PRESENT` |
| 8A | SELECTED for an ensign | `RJ-ENS-SELECTED` |
| 8B | APPROVED in block 8 | Block 8 accepts only REGULAR, FROCKED, SELECTED or SPOT |
| 8C | Promotion status missing | `STATUS-PRESENT` |
| 9 | Date reported missing | `REPORTED-PRESENT` |
| 10-13A | Occasion missing | `OCCASION` |
| 10-13C | Special report without its reason in the comments | `RJ-SPECIAL` |
| 10-13D | Officer Special report for superior performance | `RJ-SPECIAL` (review) |
| 10-13E, 10-13F | Special report reason not allowed; reservist annual training | Not checked: depends on the circumstances |
| 10-13G | Promotion/frocking report on an officer or CPO when not required | `RJ-FROCK-OFFICER` (review) |
| 14-15A, 14-15E | Period missing or invalid | `PERIOD`, `APP-DATE`, `APP-PERIOD` |
| 14-15E | End date in the future | `RJ-FUTURE-END` (information) |
| 14-15B | Periodic report on a non-periodic date | `VR-PERIODIC` (review) |
| 14-15C | Period longer than 15 months | `REGULAR-LENGTH` (review) |
| 14-15F, 14-15G | Detachment report in place of a required Periodic report | Not checked |
| 16A, 16C, 33-39D, 33-39E | Not observed and partly graded reports | `WHOLE-NOB`, `PARTIAL-NOB`, `PARTIAL-NOB-LIMIT`, `PARTIAL-NOB-MARKS`, `PARTIAL-NOB-ADVERSE` |
| 17-19A | Report type missing | `REPORT-TYPE` |
| 17-19B | Ops Cdr report by someone who is not a commander | `RJ-OPS-CDR-TITLE` (review) |
| 17-19C | Ops Cdr report on someone who is not a CO or OIC | `RJ-OPS-CDR-MEMBER` (review) |
| 20A, 20B | Physical readiness missing or invalid | `VR-PFA`, `NV-PFA` |
| 21B | Billet subcategory blank | `RJ-SUBCAT-BLANK` |
| 21C | Special billet subcategory not approved | `NV-SUBCAT` (review) |
| 21D | Billet subcategory the primary duty does not support | `RJ-SUBCAT-DUTY` (review) |
| 22-27 | Reporting senior SSN missing | `RS-SSN` |
| 33-39A | Trait grade missing | `APP-TRAIT` |
| 41C, 42B, 45B | Observed report without a promotion recommendation | `PROMOTION-PRESENT`, `PARTIAL-NOB` |
| 41D, 45D | Progressing while an advancement recommendation is in effect | `RJ-PROGRESSING` (review; uses the member's reports in the workspace) |
| 42C, 45C | Promotion recommendation on a not observed report | `WHOLE-NOB` |
| 43B, 46B, 48B | 2.0 trait above Promotable | `TRAIT-TWO` |
| 43D, 46D, 48D | Promotable with three or more 2.0 traits | `TRAIT-TWO-COUNT` |
| 43E, 46E, 48E | 1.0 trait with Promotable or higher | `TRAIT-RECOMMENDATION` |
| 43G, 46G, 48G | 2.0 or less in Equal Opportunity, or in CHIEFEVAL block 37 (Accountability), with Promotable or higher | `TRAIT-CRITICAL`, `TRAIT-TWO` |
| 43L, 46L | Third PFA failure in 4 years without the required marks | `RJ-PFA-FAILURES` (review: Swift sees only this report's failures) |
| 47-ENL | Retention not marked | `VR-RETENTION` |
| NOTE 2/H | Non-LDO ensign or LTJG above Promotable | `JUNIOR-OFFICER` |
| NOTE 2/I | Promotion/frocking report without SELECTED in block 8 | `RJ-FROCK-STATUS` |
| NOTE 2/J | Date reported after the end of the period | `RJ-DATE-REPORTED` |
| NOTE 2/C | A second Regular report for the same period | `RJ-DUPLICATE` (review; uses the member's reports in the workspace) |
| NOTE 2/G | Reporting senior's grade does not allow them to sign | Not checked |

## Summary block

Swift NAVFIT fills the summary block (FITREP 43, EVAL 46, CHIEFEVAL 48) from the
group's saved reports, so it always matches the members and counts. The group
checks give the promotion limits (`CAPACITY`) and keep not observed reports out of
the totals.

| Code | Reason | Swift NAVFIT |
| --- | --- | --- |
| 43A, 46A, 48A | Reports missing from the summary group | Not checked: Swift cannot see reports outside the workspace |
| 43C, 43I, 43K, 46C, 46I, 46K, 48C, 48I, 48K | Early or Must Promote over the limit | `CAPACITY` |
| 43H, 43J, 43N, 46H, 46J, 46N, 48H, 48N, NOTE 2/D | Summary entries wrong or different | Written by Swift from the group |
| 43M, 46M, 48M | Not observed report counted in the summary | Group check: not observed reports are left out |

## Not checkable from report fields

| Code | Reason |
| --- | --- |
| 40, 41, 43-ENL | Comments handwritten or centered |
| 45, 50, 47-OFF, 52 | Reporting senior or regular reporting senior signature missing |
| 46A, 49A, 51A | Member signature missing (an adverse report gets `RJ-ADVERSE-SIGN`) |
| NOTE 2/A, NOTE 2/B | Unreadable copy; wrong form version |
| NOTE 2/C.1 | Reports in two summary groups with the same end date, filed by different reporting seniors |
| NOTE 2/E, NOTE 2/F | Corrections to a summary group already accepted |
| NOTE 2/L | Digital or computer-generated signatures |
| 16B, 33-39B, 43F, 46F | Rules for reports ending before 1 May 2000 |
