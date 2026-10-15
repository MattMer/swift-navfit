# r109 — group actions and one Members view

Reporting senior / PSG RSCA and Export group to NAVFIT98A now follow the visible fields and status on Members, Defaults, Review and History. The footer reflows when the active group page or viewport changes. The former header gap is removed.

The redundant Comparison tab is removed from the group forms. The editable comparison table remains the Members view. Older page-3 links resolve to Members. Internal comparison services and proposal/history functions remain available to existing workflows.

Separate demo and blank r109 candidates preserve r108 and existing workspaces. App 0.25.0-dev/schema 25 remain unchanged. Automated coverage includes footer placement, removed controls, legacy routing, group actions, live member edits and defaults context switching. Physical display/keyboard acceptance remains manual. No upload or GitHub push performed for r109.

Verification: 119 resource/footer assertions, 61 live-member assertions and 198 recursive-default assertions passed (378 total). Both candidates compile and open startup forms; all 48 user-data tables in the blank are empty. Exact paths and hashes: GROUP_NAVIGATION_R109_VERIFICATION.json. An initial pre-fix defaults run exposed a remaining five-tab loop; it was corrected and the complete defaults suite rerun successfully.
