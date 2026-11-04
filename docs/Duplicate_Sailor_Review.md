# Review duplicate Sailors

Available in r127 / schema 27. This supplements the r126 user guide.

1. Close report, member-profile and prior-evaluation editors after saving or deliberately retaining their pending work.
2. Open **Resources > Command member directory > Review possible duplicate Sailors**.
3. Choose the **original member** whose links should move. The surname list suggests possible matches; double-click a candidate to select it. For a changed name, use the full **surviving member** list.
4. Choose the **surviving member** whose profile should remain active. If only one profile has an SSN, keep that profile as the survivor.
5. Select **Preview identities and all affected links**. Compare names, rank/rate, masked identifiers, report IDs and prior-evaluation IDs. Confirm independently that these are the same Sailor. A shared surname is not proof of identity.
6. Enter why you confirmed the match, without including SSNs or DoD IDs. Select **Apply reviewed merge**.

The original profile becomes archived and remains in the database. All its report and prior-evaluation links move to the survivor's stable member ID, along with its roster history, departures and dismissed report-due items. Where the survivor already has the same entry (listed on the same roster, an open departure, or the same dismissed due date), the original keeps its own. The survivor's profile fields stay as entered; values from the two profiles are not silently combined. Report content, revisions, notes, approvals, prior traits and their IDs remain intact. Existing provisional or Needs review link states are preserved. Where prior evaluations overlap, the history rules can still flag ambiguity rather than choose a record arbitrarily.

Conflicting SSNs, including those on linked reports, or conflicting linked-report DoD IDs stop the merge. Correct identity problems through the separate report/profile workflows after investigation, then preview again. Apply also rejects an outdated preview. A member involved in an active merge cannot be merged again until that merge is undone.

To reverse a merge, select it in **Merge history**, then **Undo selected merge**. Undo restores the original links, link metadata and profile availability. It keeps the merge audit and advances version counters. Later changes to either profile, linked report revisions, links, prior evaluations or traits block automatic undo so those changes are preserved. If undo is blocked, inspect the later changes; do not edit the tables to bypass the guard.

Surnames only suggest candidates. Existing confirmed member IDs remain the authoritative relationship. No report is matched or merged automatically by surname.

The workspace continues to use its folder's access controls. This feature masks identifiers in the review display and keeps identifiers out of generated merge journal messages; it does not encrypt the database or remove identifiers from retained source/history data.
