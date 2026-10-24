# Swift NAVFIT User Guide

Version 1.2 | 24 October 2026

For Swift NAVFIT development source | 0.42.0-dev | schema 33

The accompanying previously published PDF guide predates these development updates.

## Draft. Coordinate. Review. Preserve.

A practical operating guide for EVAL, CHIEFEVAL, and FITREP workspaces in Microsoft Access.

This guide follows the training-manual approach of the supplied eNavFit User Guide: overview, responsibilities, setup, navigation, numbered procedures, workflow, and reference material. It describes Swift NAVFIT's own controls and local workflow. It does not describe an eNavFit connection or Navy submission service.

**Independent development application.** Local approval and group completion do not constitute official signatures, submission, or Navy approval. Use governing instructions and the command's authorized final processing procedure.

**Guide scope:** Current r126 navigation and behavior. Screen maps are orientation diagrams, not screenshots. Older releases can have different controls. Physical display, accessibility, and final print acceptance remain separate from automated application checks.

---

# Contents and quick start

## Find a procedure

| Chapter | Topic | Page |
| --- | --- | --- |
| 1 | Overview, responsibilities, and data protection | 3 |
| 2 | Open and set up a workspace | 4 |
| 3 | Navigate the application | 5 |
| 4 | Create, edit, and save reports | 6 |
| 5 | Add notes and replies | 7 |
| 6 | Check validation and calculations | 8 |
| 7 | Use revision history and recover work | 9 |
| 8 | Manage command members and prior evaluations | 10 |
| 9 | Create and manage summary groups | 11 |
| 10 | Edit traits and compare members | 12 |
| 11 | Manage reporting seniors and PSG RSCA | 13 |
| 12 | Run an ordered review route | 14 |
| 13 | Use My Work and individual assignments | 15 |
| 14 | Complete a group and prepare final output | 16 |
| 15 | Exchange data with NAVFIT98A | 17 |
| 16 | Import PDFs and use narrative resources | 18 |
| 17 | Upgrade, back up, and protect continuity | 19 |
| 18 | Troubleshooting and status reference | 20 |
| 19 | Practice exercise and glossary | 21 |
| 20 | Scope, references, and document control | 22 |

## First useful session

1. Open a **copy of the Demo** for practice, or a **Blank** for a new operational workspace.
2. Complete setup and confirm the active contributor's name.
3. Open **Reports** and create or open a report. Work through **Front Page** and **Back Page**, then **Save draft**.
4. Use **Report tools > Notes and replies** for communication; its other views provide review steps, validation checks, and trait calculations.
5. Put the report in the correct summary group. Check its saved traits, reporting senior, and group findings.
6. Run the command's local review process. Export for the separately authorized final validation, printing, signatures, and submission process.

---

# 1. Overview, responsibilities, and data protection

## 1.1 What the application does

Swift NAVFIT stores report drafts, summary groups, members, prior evaluations, reporting-senior history, local reviews, notes, revisions, and recovery checkpoints in one local Access database. Supported report types are EVAL, CHIEFEVAL, and FITREP.

Use one authoritative workspace file and one contributor at a time. Separate copies do not synchronize. Sending someone a copy does not create a live shared conversation or merge their edits back automatically.

## 1.2 Who does what

| Responsibility | Swift activity |
| --- | --- |
| Workspace coordinator | Maintains the authoritative file, backups, members, groups, and contributor list. |
| Drafter | Enters report information, saves revisions, addresses findings, and participates in notes. |
| Reviewer | Uses the correct contributor profile to accept or return eligible assigned work. |
| Route requester | Starts routes and handles stale routes, restarts, or cancellation. |
| Final local reviewer | Records local approval when all required route steps are accepted on current saved values. |
| Reporting-senior administrator | Maintains reusable senior information and verified historical RSCA inputs. |

These are operating responsibilities, not authenticated security roles. A local contributor profile records attribution; it does not prove identity or enforce folder access.

## 1.3 Local data protection

Normal report drafting and review do not send data to the internet, synchronize to a service, or send telemetry. Swift requires no internet account or application-specific network protection. Optional reference links open an external site only when selected; report contents are not included in those links.

Store the database in an approved protected folder. Under normal Windows inheritance, it receives that folder's permissions; explicit file permissions still apply. Device encryption, organizational handling rules, backups, and any configured folder synchronization remain relevant. Swift adds no separate authentication or encryption layer. Microsoft Access and Office operate under their own settings.

**Expected result:** Everyone knows which file is authoritative, whose contributor name is active, and who owns backup and final processing responsibilities.

---

# 2. Open and set up a workspace

## 2.1 Prerequisites

Use Windows with full desktop Microsoft Access. Free Access Runtime is not supported. The application is a development candidate; consult the current release notes for platform and interoperability boundaries.

Choose the file that matches your purpose:

| File | Purpose |
| --- | --- |
| Blank | New command workspace or empty destination for an older Swift workspace. |
| Demo | Practice using synthetic reports, groups, reviews, and reporting-senior history. |
| Existing populated workspace | Continue real work; preserve it when upgrading. |

## 2.2 Start a new workspace

1. Keep the downloaded Blank as a master and make a working copy in the approved location.
2. Open the working copy in desktop Access. Follow organizational Access trust policy; application VBA must be allowed to run for the buttons to work.
3. Enter the command/workspace information and your contributor name in first-run setup.
4. Open **Overview**. Confirm the workspace context and active contributor before entering report data.
5. Make a backup before substantial data entry or import.

**Expected result:** The workspace is configured and navigation opens normally. A new Blank has no existing reports or command history.

## 2.3 Continue an older Swift workspace

Do not overwrite the old file with a new Blank. Open a separate new Blank and use **Import data from a previous Swift NAVFIT version**. Chapter 17 describes the transfer. A populated Demo is not an upgrade destination.

## 2.4 Change the active contributor

1. Finish or save report work and close the report editor.
2. Open **My Work**, then the contributor controls.
3. Add the contributor if needed, or select the existing name and switch contributors.
4. Confirm the displayed active name before posting notes or recording review decisions.

Contributor switching is guarded while report work is open. Do not bypass the guard by editing tables or creating a duplicate profile to impersonate another person.

---

# 3. Navigate the application

## 3.1 Main destinations

| Left navigation | Use it for |
| --- | --- |
| Overview | Workspace status, work counts, and recovery entry points. |
| Reports | Search, filters, report creation, import, and selected-report actions. |
| Summary groups | Membership, shared defaults, traits, group status, and closeout. |
| My Work | Assigned reviews, current route work, contributor controls, and report conversations. |
| Resources | Reporting seniors, narrative snippets, and reference tools. |
| Settings & Help | Help, backups, transfer from older Swift versions, operation history, and diagnostics. |

The top-right **?** opens help for the current window. Use the application's navigation and return buttons to keep the report or group context correct. Leaving a report follows its normal save/recovery guards. Leaving a group closes its working view so another group can open with the correct context.

## 3.2 Report workspace map

| Layer | Controls and purpose |
| --- | --- |
| Report header | Member, report type, period, ITA, saved group average, and local review context. |
| Report pages | Front Page, Back Page, Report tools, History. |
| Report tools views | Notes and replies, Review steps, Validation checks, Trait calculations. |
| Action area | Back to reports, Undo pending changes, Validation issues, Save draft, and save/issue status. |

The four Report tools buttons switch content inside the editor. **Notes and replies** contains report conversations. Validation can also open in a separate popout.

## 3.3 Color and availability cues

The navigation rail shows the current destination. Stronger button emphasis indicates a current or primary action; warning and issue treatments draw attention to pending work or problems. Read the label and status text as well as the color. A disabled button can mean the wrong contributor, a stale review, an archived report, local approval, or missing saved context.

In group traits, green ITA means higher than that sailor's prior ITA; yellow means lower. Equal values or no prior ITA receive no comparison highlight. Gray prior-value cells are references, not editable current marks.

---

# 4. Create, edit, and save reports

## 4.1 Create a draft

1. Open **Reports** and choose the new-report action.
2. Choose EVAL, CHIEFEVAL, or FITREP. Select a command member when using the member picker, or enter the required identity information.
3. Confirm the member, report type, reporting period, and summary-group context.
4. Work through **Front Page** and **Back Page**, following the official block captions.
5. Choose **Save draft** and confirm the saved-revision message.

All seven performance traits are grouped on Front Page for editing. This does not change their official printed-page associations. The pages use fixed field geometry. Resize or scroll the viewing area to reach fields; the report fields do not reflow to a different block order.

## 4.2 Understand the save states

| State | Meaning and next action |
| --- | --- |
| Pending changes | Current editor values differ from saved data. Save when ready. |
| Saved revision | The report's committed values and change history were recorded together. |
| Recovery checkpoint | A separate copy of pending work; it is not a committed revision. |
| Save failed | Pending values remain available. Correct the indicated entry and retry. |
| No pending changes | No unsaved report changes are detected. A no-op save does not create history. |

Incomplete drafts can be saved when their entered values meet storage requirements. A saved draft is not a certification that every policy check is satisfied.

Leaving the editor attempts a save. If raw entries cannot be saved, Swift keeps
them in a recovery checkpoint and allows navigation once that checkpoint succeeds.
Use **Settings & Help > Recover pending edits** to resume them. If the saved
report changed meanwhile, recover as a new draft. A failed checkpoint keeps the
editor open so pending work is not lost.

## 4.3 Move between pages or leave

Page changes preserve pending report values without creating revisions. Leaving a changed report normally saves through the application workflow. If a save fails, correct the values or deliberately undo pending changes; do not assume the failed save committed. Use **Undo pending changes** only when you intend to abandon those unsaved report edits.

Locally approved or archived reports are read-only. Reopen local approval with a reason, or recover an archived report, before editing. Notes can continue after local approval; archived reports remain read-only for notes too.

---

# 5. Add notes and replies

## 5.1 Purpose and screen areas

Report notes are shared communication among people working in the same authoritative workspace. They are independent of approval and report revisions. There is no resolve/reopen requirement. Posting a note does not accept a review step, unlock a report, or change its saved revision.

| Area | What it shows |
| --- | --- |
| Conversations - all contributors | Existing conversation roots, contributor, and posting time. |
| Selected conversation and replies | Full text of the selected conversation and its replies. |
| Your unfinished notes - not posted | Your private saved drafts for this report. |
| Writing area | A new note, a reply, or a resumed private draft. |

## 5.2 Post a new note

1. Open the saved report and choose **Report tools > Notes and replies**.
2. Read existing conversations to avoid duplicating the discussion.
3. Choose **Write a new note** and enter your message.
4. Check the active contributor name, then choose **Post note / reply**.
5. Confirm the note appears in the shared conversation list and reading area.

## 5.3 Reply to a conversation

1. Select a conversation in the left list and read its full text on the right.
2. Choose **Reply to conversation**.
3. Enter the reply and choose **Post note / reply**.

The reply stays with its selected conversation even if you save the unfinished reply and resume it later. Conversations remain usable after local approval or later report revisions.

## 5.4 Keep or discard unfinished text

Choose **Save draft** in the notes area to keep a private unfinished note. This is different from the report's Save draft button. Select an item under **Your unfinished notes - not posted**, then choose **Continue selected draft** to resume it. **Discard draft** deliberately removes that private checkpoint. Refreshing or changing views does not post the text.

Draft checkpoint failures block navigation that would lose the note. Unposted drafts remain private to their contributor profile; that privacy is an application convention, not authentication. Old notes, comments, replies, and attribution are retained when upgrading.

---

# 6. Check validation and calculations

## 6.1 Review report findings

1. Open the report and choose **Report tools > Validation checks**.
2. Select a finding. Read the message, affected block, and available policy reference or application rule.
3. Choose **Go to affected report block** to return to the relevant entry.
4. Correct the report. Refresh the checks and save the report when ready.

Checks include pending editor values. They do not create saved report revisions merely by running. Coverage is limited to implemented cases; a clean list is not complete policy or signature certification.

## 6.2 Keep findings beside the report

Choose **Validation issues** in the report action area, or **Pop out validation issues** from the embedded validation view. The popout is modeless: it can remain open while you return to report fields. Move it beside the report using normal window controls. Refresh or use its field navigation as needed. Close the validation window when finished; closing the report also closes its report popout.

## 6.3 Interpret categories

| Finding | How to handle it |
| --- | --- |
| Invalid entry | Correct the format, range, or storage problem. |
| Incomplete draft | Supply missing information when available; drafting may continue. |
| Policy issue | Review the cited rule and correct or substantiate the report as appropriate. |
| Needs review | A person must verify applicability, an exception, or unavailable evidence. |

Summary-group checks also include grouping, membership, and provisional-capacity reminders. The Demo's clean group has zero member-report findings, but keeps group reminders. A group reminder does not mean every member report has an entry error.

## 6.4 Inspect trait arithmetic

Open **Report tools > Trait calculations**. The view separates current editor values, latest saved values, and summary-group calculation context. Pending and saved figures can differ until you save. NOB and missing traits affect the calculation basis; read the explanation rather than treating every blank as zero. Group projections use saved member traits and the selected reporting-senior history.

---

# 7. Use revision history and recover work

## 7.1 Inspect or restore saved history

1. Open a report and choose **History**.
2. Select the revision of interest. Compare it with the current saved report to see recorded field changes and attribution.
3. If an earlier version is needed, use the restore action and confirm its scope.
4. Verify the restored values and the new saved revision.

Restore appends a revision; it does not erase later history. A meaningful saved change creates history, while an unchanged save does not. Older revisions may lack fields introduced later; inspect the restore preview rather than assuming missing historical fields contain today's values.

## 7.2 Recover pending report edits

1. From **Overview** or **Settings & Help**, open **Recover pending edits**.
2. Select the relevant report checkpoint and review its member, time, and baseline context.
3. Use the available recovery action. If the underlying saved report has changed, recover as a new draft rather than overwriting unrelated newer work.
4. Inspect the recovered report and choose **Save draft** when ready.

Checkpoints preserve work separately from committed revisions. They do not replace a backup of the whole database. Invalid raw entries can remain in recovery for correction instead of being committed as valid report data.

## 7.3 Archive, duplicate, and undo

Use the report-selection actions in **Reports** to archive or recover reports. Archive removes an item from active views while retaining its history. Adjust the archive filter to find it again. Duplicate creates a separate report identity; it is not a revision of the original.

**Settings & Help > Operation history** offers guarded undo for supported operations. Later changes can make an operation ineligible for undo. Read the affected-record scope before applying it.

**Keep the three protections distinct:** a revision records a saved report change; a checkpoint retains pending editor or note work; a backup protects the workspace file and its complete collection of records.

---

# 8. Manage command members and prior evaluations

## 8.1 Build the member directory

Use the command-member tools associated with report creation to maintain reusable identities. Confirm spelling, identifiers, rank/rate, and paygrade before creating reports. A member-directory record and a report are related but distinct records; prior evaluations remain historical information rather than silently replacing today's report.

## 8.2 Create reports from members

1. Open the member-based report creation workflow.
2. Filter by paygrade, such as E6, or by a specific rank/rate, such as YN1.
3. Select the intended member and report type, then review the proposed report context.
4. Create the report, inspect its Front/Back pages, and save it.

Changing a member filter clears the earlier selection. Re-select the member after filtering so a hidden previous selection is not used accidentally.

## 8.3 Create a batch

1. Open the batch report creation workflow and apply the appropriate member filters.
2. Select the intended members and provide the common report context.
3. Preview the batch. Check member names, types, periods, and group placement.
4. Apply only after the preview matches the intended scope, then inspect the created reports.

Changing the filters or source selections invalidates an earlier preview. Generate a new preview before applying. Batch creation reduces repeated entry; it does not verify individual reporting eligibility or remove the need to review each report.

## 8.4 Retain prior evaluations

Record prior evaluations against the correct command member using the prior-history tools. Confirm the prior report's date, type, traits, ITA, and promotion recommendation before relying on comparisons. An incorrect member link or missing historical entry can explain why a prior-value column is blank.

In the summary-group trait table, **Prior ITA** and **Prior Promotion** are read-only references. Correct the underlying member's prior evaluation when those references are wrong. Do not change current report marks merely to compensate for an incorrect historical value.

**Expected result:** Reusable member information supports current report creation, while historical evaluation data remains distinguishable from current saved report values.

---

# 9. Create and manage summary groups

## 9.1 Create and open a group

1. Open **Summary groups** and choose **Create Summary Group**.
2. Enter the group context, including its title, report type, and paygrade as requested.
3. Open the group and add the intended reports or create reports in its context.
4. Check the roster and group findings against the applicable summary-group criteria.

The group workspace provides Members, Defaults, Review, and History. Use its return-to-list action to choose another group. Selecting another main menu closes the working group view; opening a different group should show that group's roster and title.

## 9.2 Apply defaults deliberately

1. Open the group's **Defaults** or the selected-group defaults action.
2. Change the intended shared report block.
3. Read the confirmation's affected-report scope, including customized values and local approvals.
4. Confirm only if that scope is intended. Check the saved member values afterward.

These shared-field changes save through their field confirmation; there is no additional general Preview/Apply stage for defaults. Group defaults can replace that field in active members and affect future reports. Global defaults can also replace it across active groups. Archived groups/reports are excluded by this workflow. A failure rolls the change back together.

Do not assume a default only fills blanks. The confirmation exists because a change can replace customized values and return affected local approvals to editable work.

## 9.3 Interpret list information

The group list shows report counts, archive state, local group status, and closeout date. Closeout date comes from member period-ending values: **Mixed** means different dates; **Incomplete** means one or more are missing.

Archiving a summary group hides the group from active views but does not delete its member reports. Group archive state is different from report archive state. Use the archive filter and recovery action when needed.

Group status is local tracking. It does not certify that every member is complete or automatically accept the members' review steps. Chapter 14 explains completion and reporting-senior count effects.

---

# 10. Edit traits and compare members

## 10.1 Use the group trait table

1. Open the group's **Members** view and its trait table.
2. Check the member names, report states, and prior-value references before editing.
3. Edit the intended current trait or recommendation cell.
4. Leave the cell and confirm that the change saved. Valid direct edits save as report revisions.

Failed cells remain pending for correction, retry, or deliberate discard. Resolve pending cells before relying on saved group calculations, changing context, or reordering rows. Table headings sort rows; pending cells save before a new order is applied.

## 10.2 Read the comparison cues

| Cue | Meaning |
| --- | --- |
| Green current ITA | Higher than this sailor's prior ITA. |
| Yellow current ITA | Lower than this sailor's prior ITA. |
| No comparison highlight | Equal ITA or no prior ITA available. |
| Gray prior cells | Historical references that are not editable in the current trait table. |

Comparison color is a drafting aid, not a required direction of change. A higher or lower mark still needs an accurate report and appropriate judgment.

## 10.3 Distinguish ITA, group average, and PSG RSCA

ITA describes an individual report's applicable numeric traits. Group trait average describes the saved group basis. PSG RSCA projects the selected reporting senior's historical sum/count together with eligible group traits. These figures have different denominators and can legitimately differ.

Read counts and NOB exclusions beside the result. Whole-report NOB, missing traits, partial-NOB entries, and incomplete saved context can prevent a complete comparison or projection. Do not treat a displayed average as proof that group eligibility or promotion limits have been independently verified.

## 10.4 Use staged comparison proposals

The group comparison tools can preview staged mark proposals. Review the affected reports and before/after values before applying. Changes to membership, saved revisions, or locks require a fresh preview. Applying a proposal is distinct from editing a live trait cell and from changing a group default.

---

# 11. Manage reporting seniors and PSG RSCA

## 11.1 Record reusable senior information

Open **Resources > Reporting seniors and RSCA history**. Create or select the reporting senior and enter the reusable details. Applying those details to reports is a separate deliberate action; editing a directory record does not silently rewrite saved reports.

## 11.2 Enter historical averages

1. Open the selected senior's RSCA history table.
2. For each applicable paygrade, enter the historical RSCA, previous graded trait count, and date verified or pulled from BOL.
3. Leave unused paygrades blank. The count is graded traits, not the number of reports.
4. Choose **Save RSCA table**. Use **View prior baselines** to inspect older observations.

An entered rounded average and count produce an estimated historical sum. They do not establish an exact sum. Preserve the source and verification date so later users understand the basis of a projection.

## 11.3 Configure a group's projection

1. Open the summary group and choose **Reporting senior / PSG RSCA**.
2. Select the intended senior and matching historical paygrade baseline.
3. Confirm whether the baseline excludes this group's contribution, as requested by the calculator.
4. Save or confirm the selections through the provided controls and inspect the displayed PSG RSCA and calculation explanation.

Senior, baseline, and exclusion selections persist through their controls. Changing the senior or baseline clears the exclusion confirmation. A newer history entry does not silently replace a group's deliberate selection.

If PSG RSCA is unavailable, check the senior, matching paygrade history, baseline/exclusion confirmation, and complete eligible saved traits. Having a senior name alone does not establish every calculation input.

## 11.4 Account for completed groups

Marking a group **Completed** appends its eligible local contribution to the senior's paygrade totals. NOB traits are excluded. Re-completion applies differences rather than counting the same group twice. Reopening or archiving retains recorded totals until corrected completion. Later local totals are not a fresh BOL verification; independent BOL observations require deliberate reconciliation.

The group projection excludes its own already-counted contribution where its selected history contains that completion. Do not manually add the group a second time.

---

# 12. Run an ordered review route

## 12.1 Start-to-finish workflow

**Draft report > Save > Check findings > Start route > Accept/return each step > Final local approval > Group closeout > Authorized final processing.**

Notes and replies can occur throughout. They do not advance this sequence. Individual assignments, described in chapter 13, are another local review mechanism and do not automatically replace the ordered route.

## 12.2 Start review

1. Save the report and address the findings appropriate to the command's review stage.
2. Open **Review setup** or **Report tools > Review steps**.
3. Use **Route steps** mode. Select an available ordered route template. Maintain templates through the route-template tools if a new reviewer sequence is needed.
4. Choose **Start review**. Confirm the saved report context, selected route, and first reviewer.

Review actions use saved values. Save or undo pending report changes before recording a review action. Starting a route does not send email or transmit a report to a remote service.

## 12.3 Accept or return a step

1. The assigned contributor opens the current task from **My Work** or the report Review steps view.
2. Read the saved report, route instructions, findings, and shared notes.
3. Enter the decision note and choose **Accept step** or **Return report**. A return requires an explanation.
4. Confirm the route's new status and next eligible reviewer.

Future steps do not become the current reviewer's actionable work early. Unsubmitted decision text is separate from report conversation text; record or clear it before leaving when the editor blocks closure.

## 12.4 Handle changes and approval

After corrections are saved, an older pending or accepted route can become stale. Its requester restarts it using current saved values; earlier decisions remain in history. A group-scoped route can also become stale after membership changes. Requester cancellation and restart have their own eligibility checks.

Once all steps are accepted on current values, the final reviewer records **Local approval**. The report becomes read-only. **Reopen approval**, with a reason, permits further edits. Local approval is not an official signature or submission.

---

# 13. Use My Work and individual assignments

## 13.1 Open eligible work

1. Open **My Work** and confirm the active contributor.
2. Select a row and read the report identity, period, status, contributor, and next action.
3. Choose **Open selected work** to open that report's appropriate **Review steps** view and selected route or assignment.
4. Use **Back to My Work** when finished; it returns focus to the existing My Work tab. **Open report** opens the report for normal report work.

The review queue includes current assigned reviews, current ordered-route steps, ready local approvals, and requester actions for stale work. Future route steps are not actionable yet. Assignment due dates appear when supplied; routes do not acquire invented due dates.

## 13.2 Assign a saved revision

1. Open the assignment controls in My Work and select the intended saved report.
2. Select the reviewer. Enter optional due date and instructions.
3. Confirm the saved-revision context and choose the assignment action.
4. The assigned contributor opens the work in **Assigned reviews** mode, reviews it, and accepts or returns it with the appropriate decision note.

One pending individual assignment per report is allowed. Only the assigned contributor can accept/return eligible current work. A return needs a reason. The requester can cancel pending work. A later saved revision makes the older assignment stale; cancel it and assign the latest saved revision instead of applying an outdated acceptance.

Acceptance records that assignment's outcome. It is not ordered-route final approval, official signature, or Navy submission. Archived or locally approved reports restrict review decisions according to their state.

## 13.3 Read conversations separately

Choose My Work's **Report notes** tab to browse workspace conversations, then **Open selected report conversation**. It opens the report's Notes and replies view. Notes are not unresolved approval tasks and need not be closed to complete a review.

Use the completed/all-history controls to inspect older review work. A changed report can make historical work stale without deleting its original decision. If an action is disabled, check active contributor, current saved revision, archive state, local approval, and route step before retrying.

---

# 14. Complete a group and prepare final output

## 14.1 Understand local group status

| Group status | Intended local use |
| --- | --- |
| Initialized | Group created; roster and report preparation are beginning. |
| In-Work | Reports and shared group information are actively being prepared. |
| Ready for Signature | Local tracking that the group is ready for the command's signature process. |
| Completed | Local closeout; eligible group traits are incorporated into reporting-senior history. |

These statuses are separate from each report's route, local approval, and archive state. Changing group status is not an automatic certification that all report or Navy requirements are satisfied.

## 14.2 Close out deliberately

1. Confirm the roster, report types, paygrades, periods, and reporting-senior context.
2. Review individual report findings and all summary-group findings. Independently verify eligibility, roster completeness, and provisional promotion capacities.
3. Confirm the intended local report reviews and approvals are complete on current saved values.
4. Check the reporting-senior baseline, eligible trait counts, and PSG RSCA explanation.
5. In the group list, choose the intended **Local group status** and **Set group status**. Use Completed when the command's local closeout criteria are met.
6. Verify the recorded senior history and retain any required supporting records.

Completed groups update local RSCA totals. Re-completion posts differences; reopening or archiving does not simply erase an earlier contribution. If a group needs correction, inspect its completion/history context rather than manually creating a second historical entry for the same traits.

## 14.3 Prepare official processing separately

Save report edits before exchange. Export the report or group using the supported NAVFIT98A workflow, retain its baseline, and open the exported file in NAVFIT98A for native validation and printing. Follow the command's authorized signature and submission process outside Swift.

Swift has no verified equivalent of NAVFIT98A's native final print engine. Its report-editing layout and a clean finding list do not prove final printed fit, signature completion, or acceptance by an official system.

---

# 15. Exchange data with NAVFIT98A

## 15.1 Import reports or groups

1. Open **Reports > Import reports** and select the NAVFIT database workflow.
2. Choose the source database. Preview the available records and grouping context.
3. Select the intended items. Use **Select all** when the whole displayed selection is intended.
4. Review the import preview and apply it. Inspect the results and open the imported reports to verify identities, periods, traits, and narrative.

Imported native provenance is retained. Unsupported or unknown encodings still need review. The supported mapping does not make every possible native database or policy case equivalent to a fully verified report.

## 15.2 Export a summary group

1. Select the summary group in **Summary groups**, or open its Members page.
2. Choose **Export selected group to NAVFIT98A** or **Export group to NAVFIT98A**.
3. Confirm the displayed group and choose an export directory. Swift supplies its internal native template automatically.
4. Choose **Export summary group to NAVFIT98A**. All saved group members are included. Validation findings, incomplete content and report status do not prevent export. Pending edits remain separate; save them first if they belong in the export.
5. Open the native output in NAVFIT98A, validate it, and use its appropriate report/folder/summary printing actions.

Imported native-only values are retained. Exact PSG RSCA is written when complete traits and an exact historical sum support it; otherwise existing native values are preserved. A rounded-average estimate is not an exact historical basis. Export does not approve, sign or submit reports. Keep accompanying Swift baseline/provenance files locally; send the native report database.

## 15.3 Retained exchange history

The export screen is dedicated to creating a summary-group NAVFIT database. It does not present returned-file reconciliation or saved-return actions. Existing provenance, baselines, saved previews and historical records are preserved. Returned-group reconciliation and arbitrary-source matching remain unsupported.

---

# 16. Import PDFs and use narrative resources

## 16.1 Import a supported structured PDF

Extraction runs locally inside Swift NAVFIT. Acrobat, Reader, Python, external utilities and an internet connection are not required. Only the supported May 2025 fillable NAVPERS EVAL, CHIEFEVAL and FITREP templates are accepted. Flattened or scanned PDFs require manual entry. Use the extraction check to test a selected file. Password-required files, extraction-prohibited files, and unsupported encryption are rejected; the supplied CHIEFEVAL form opens without a password and permits extraction.

1. Open **Reports > Import reports** and choose the structured PDF workflow.
2. Choose the PDF and run **Check PDF extraction on this computer** if needed.
3. Leave the merge target blank to create a new draft, or choose the intended existing draft.
4. For an existing target, choose **Fill empty fields** or **Replace mapped values** deliberately.
5. Preview all mapped before/after values, including replacements and retained fields.
6. Apply, inspect Results, and open the imported or merged report for validation and saving as appropriate.

Close the target editor and resolve pending recovery before merging. Changed or locked targets require a fresh preview. Blank or missing source PDF fields do not clear existing values. Source or mapping-choice changes invalidate the preview.

The supported mappings cover defined structured fields for the three report types. They do not promise extraction from arbitrary scans or every PDF layout. Normal PDF import uses SSN; legacy DoDID mapping requires explicit compatibility selection. Inspect identifiers and leading zeros after import.

## 16.2 Use snippets and references

Open **Resources** for narrative snippets and reference tools. Filter by rate, inspect attribution and applicability, then use the available insertion workflow for the report narrative. Legacy untagged snippets appear as Unassigned until tagged.

NavyWriter collection links open the complete external source when selected. Only supported reviewed excerpts can be imported through the catalog tools. Check that suggested wording fits the member's actual work, grade, period, and the report's narrative capacity.

## 16.3 Keep communication out of official narrative

Use report Notes and replies for discussion. Use the official Front/Back report blocks for text intended to leave with the evaluation. A shared note, private note draft, route decision note, and official narrative have different purposes; posting one does not populate the others.

---

# 17. Upgrade, back up, and protect continuity

## 17.1 Transfer an older Swift workspace

1. Finish report work and close the old source workspace. Keep an intact backup.
2. Open a separate **new Blank** release, not the Demo or another populated workspace.
3. In setup or **Settings & Help**, choose **Import data from a previous Swift NAVFIT version**.
4. Select the old `.accdb`. Review the source version and record-count preview.
5. Choose **Import all workspace data**. Watch the running progress log for copying and verification phases. Do not start a second import because a large transfer takes time.
6. Wait for the committed-success message. Close and reopen the new destination.
7. Verify representative reports, groups, notes/replies, reviews, recovery, members, senior history, and settings before adopting the new file as authoritative.

Swift copies by table, restores relationships, and verifies original source values before committing. The source is read-only; source forms/code are not installed into the new application. Initial setup details in an otherwise blank destination are replaced by the source workspace's details.

This is a whole-workspace transfer into an empty destination, not a merge between populated files. Future schemas and unrecognized source tables are rejected. On a reported failure, read the log and retain both files; do not delete source tables to force acceptance. The r126 release was tested against retained schema 1-26 fixtures.

## 17.2 Backup and restore

Use the backup/restore tools in **Settings & Help** and follow their destination and confirmation prompts. For a file-copy backup, close Access cleanly first and retain a dated copy in the approved backup location. Protect exported native files and return baselines at the same data sensitivity as the workspace.

Verify a restored copy before replacing the authoritative working file. Keep older copies until the restore is checked. A restore to a previous whole-file backup can predate later report changes; report History is the better tool when only one earlier report revision is needed.

---

# 18. Troubleshooting and status reference

## 18.1 Find the next action

| Symptom or status | Check and next step |
| --- | --- |
| Button does nothing at startup | Confirm desktop Access and organizational VBA trust settings; read setup/error text. |
| Save fails | Correct the reported entry; pending values remain available. Save again and confirm success. |
| Field is read-only | Check archive state or local approval. Prior-value cells are intentionally read-only. |
| Accept/Return is disabled | Confirm active contributor, current saved revision, current route step, and report locks. |
| Review is stale / needs restart | Save corrections; requester restarts the route or cancels/reassigns an individual assignment. |
| PSG RSCA unavailable | Check matching senior/paygrade baseline, exclusion confirmation, and complete eligible saved traits. |
| Prior ITA is missing | Inspect the member link and prior evaluation records. Do not enter it in the read-only group reference cell. |
| Import looks busy | Read the timestamped progress log. Wait for success/failure instead of launching another import. |
| Workspace import rejected | Use an empty new Blank; check supported source version and log. Preserve the old file. |
| Cannot leave a report | Correct a failed save/checkpoint, or record/clear unfinished review decision text. Do not force-close to bypass it. |
| Report/group seems missing | Clear search and filters; check archive status and the selected summary group. |
| Notes are not in review queue | Open My Work's Report notes tab; communication is independent of approval tasks. |

## 18.2 Terms that can be confused

**Draft / Approved:** report editing workflow. Approved means locally locked after local review; it is not official Navy approval.

**Pending / Accepted / Returned / Cancelled:** a review assignment or route's local activity state. Accepted work can become stale if saved values or relevant group context change.

**Active / Archived:** visibility and editing lifecycle; archived items are retained and can be recovered.

**Initialized / In-Work / Ready for Signature / Completed:** summary-group tracking; separate from member-review decisions and external submission.

For a repeatable error, retain the file and note the version, action, report/group context, and full message. Use **Settings & Help > Support and limitations** for diagnostics. Avoid sending real member data when a synthetic reproduction is sufficient.

---

# 19. Practice exercise and glossary

## 19.1 Guided practice in a separate Demo copy

1. Open a fresh Demo copy and confirm the synthetic-data banner. Do not import real workspace data into it.
2. Open **DEMO - Clean reports (0 report issues)**. Inspect its members, ITA values, prior references, and reporting-senior projection.
3. Open a member report. Switch Front/Back pages, inspect Validation checks, and try the validation popout beside the report. Observe that group reminders are separate from member-report findings.
4. Post a new synthetic note, reply to it, and save another reply as a private unfinished draft. Switch views, then resume the draft.
5. Open **DEMO - Practice corrections (3 report issues)** and inspect the findings. Navigate to an affected block, correct a value, save, and inspect History.
6. Inspect existing synthetic review examples through My Work. Confirm that shared notes do not appear as unresolved review tasks.
7. Close the practice copy. Keep practice changes separate from the authoritative operational workspace.

## 19.2 Glossary

| Term | Meaning in Swift NAVFIT |
| --- | --- |
| Authoritative workspace | The one local database the command currently uses for its work. |
| Contributor | Local attribution profile used when saving, posting, or reviewing. |
| Revision | Committed report version with recorded changes. |
| Checkpoint | Pending work retained separately for recovery or private note continuation. |
| Conversation | Report-level shared note and replies, independent of approval. |
| Assignment | A named review task for a saved report revision, optionally with a due date. |
| Route | Ordered local review steps using a selected template. |
| Stale | Review context no longer matches current saved report or relevant group values. |
| ITA | Individual trait average using the report's applicable numeric traits. |
| PSG RSCA | Post-summary-group reporting-senior cumulative-average projection. |
| Baseline | Retained reference values for a calculation, export return, or guarded operation. |
| NOB | Not Observed; read the calculation's eligibility/exclusion explanation. |
| Provenance | Retained information about a value's source and import/exchange history. |

---

# 20. Scope, references, and document control

## 20.1 What this guide establishes

This manual describes the r126 local workflow. It is not a Navy instruction, authorization to sign, or certification of report policy compliance. Use applicable governing instructions and command procedures for eligibility, required content, signatures, and submission.

Swift does not provide BOL login, Navy role acknowledgment, online trusted-agent delegation, email delivery, automatic notifications, official signatures, or official submission. Contributor profiles and reporting-senior records are local data, not Navy account roles.

Native automated checks cover defined workflows and retained-schema transfers. They do not establish physical display/DPI acceptance, screen-reader acceptance, every policy case, every NAVFIT round trip, or final printed fit. Free Access Runtime is not a supported target. Consult current release evidence for other platform limitations.

## 20.2 Reference basis

| Reference | Use in this guide |
| --- | --- |
| eNavFit User Guide, Version 1.0, 10 December 2021; supplied 86-page PDF | Training-manual organization: overview, responsibilities, access, navigation, workflow, procedures, and reference material. Its online roles and submission process are not Swift features. |
| Swift NAVFIT r126 source and generated control definitions | Current captions, destinations, save guards, notes/review behavior, calculations, and migration procedures. |
| REPORT_TOOLS_R123_2026-10-20.md and associated verification | Consolidated views, preserved legacy history, schema26 migration, release checks, and boundaries. |
| Current workspace help and earlier feature documentation | Report editing, groups, members, defaults, RSCA, import/export, recovery, and backup context, reconciled to r126 navigation. |

The reference eNavFit manual is used as an organizational model, not reproduced as Swift instructions. Screen maps in this guide are schematic orientation aids; no eNavFit screenshot is presented as Swift's interface.

## 20.3 Maintenance

Guide version 1.0 was prepared on 20 October 2026 for r126, application 0.26.1-dev, schema26. The editable Markdown and reproducible PDF renderer accompany the PDF in the repository. Update procedures when controls or behavior change, then regenerate and visually inspect the PDF.

Start with this guide for current user procedures. Older feature guides retain historical service and verification detail and may describe superseded windows or review/comment behavior. The repository's current-state notes identify the latest delivered artifacts and remaining work.
