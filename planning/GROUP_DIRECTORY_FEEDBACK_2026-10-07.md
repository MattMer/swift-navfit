# Group navigation and RUAD profiles — r90

Implemented the four requested changes in the native Access build, retaining the author-edited r89 baseline. The group Add report dialog supports an existing directory member or a new member, displays the group's report type/paygrade, offers date selectors and uses saved global/group defaults. Creation commits report revision, directory link and group membership atomically. The report editor returns to the group's member list. Cancel does not create a draft.

All group pages share Members, Defaults for new reports, Review, Comparison and History tabs and Add report. Members and Comparison use the editable member grid; switching those tabs does not create another snapshot or refresh away pending state. Global defaults do not display group-specific actions. Pending defaults must be applied or explicitly discarded before changing context.

Overview places contributor selection and its status above the work cards, with Unposted notes below recovery actions and above recent reports. Group footer status was moved to avoid covering action buttons.

The member profile includes every column from the supplied 35-column RUAD. Five pages cover Unit, Member, Assignment, Mobilization and Contact. NAME uses the existing member name; report rank/rate and official report name remain separate from raw roster fields. Page changes preserve pending values. Schema 25 adds one nullable memo field, CommandMember.RUADData, with known-key XML storage, null handling, length validation and atomic version-checked profile saves. No roster identity is inferred from names or BIN. Existing basic profile saves preserve roster values.

Two explicitly synthetic IT2/LT profiles demonstrate complete roster data. The user's edited source, r88 and r89 are preserved. Existing rows are compared field by field; only the documented schema metadata change, nullable new field and two demo profile/journal additions are allowed.

Verification includes native compile, 129 group/RUAD/Overview checks, approved-report retry checks, author-baseline checks, directory/history checks, startup checks and schema-20-to-25 plus schema-25-to-25 exact transfer/reopen tests. Final admin and responsive suite results and artifact hashes are recorded separately in GROUP_DIRECTORY_VERIFICATION_2026-10-07.json.

Earlier native visual inspection confirmed the Overview, common group tabs, Add report dialog and complete directory profile pages. Further computer use was stopped at the user's request. Automation tests do not establish screen-reader behavior or physical display/scaling acceptance.

Still pending: full RUAD Choose → Preview → Apply → Results import and refresh with reviewed identity matching, provenance and undo. Smaller observed UX follow-ups include cramped default-field reset/checkbox labels, directory list width, trait-grid header wrapping and clearer archive-filter wording. These were not expanded into this change.
