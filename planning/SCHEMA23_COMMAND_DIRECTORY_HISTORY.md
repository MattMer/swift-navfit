# Schema 23 - Command directory and prior evaluation history

Four additive tables retain command identities, report links, external prior headers and their seven trait slots. Historical evaluation details and stable identities cannot be represented by summary-group membership or current report fields. No existing report column or revision is replaced.

SSN is a nine-digit nullable unique identity key with leading zeros preserved. Name-only records stay provisional. Conflicting identity data needs explicit confirmation, rather than silently merging people. Profiles keep display and official report names separate. Profile edits do not rewrite saved reports.

Prior-history selection uses confirmed member links and the latest ending strictly before the current report start. Tied dates are ambiguous; no source is guessed. NOB traits are excluded from numeric ITA, missing traits show Incomplete, and whole-report NOB remains NOB. These selection choices are working assumptions pending user clarification.

Services use version guards and DAO transactions. Report-directory links participate in report saves. Current report revisions are unaffected by external prior-history edits. Schema 22-to-23 transfer creates directory links only after verifying original rows and fields; schema 23-to-23 transfer retains identity timestamps exactly.

Verified so far: 17 service and 13 native UI assertions. Migration, directory batch creation, RUAD mapping, physical-display and screen-reader acceptance remain outstanding.

Implementation update (r66): directory batch creation and explicit identity-confirmation screens are implemented. Twenty batch assertions and sixteen native directory/history UI assertions pass. Schema 22-to-23 and 23-to-23 transfers verify every original row and field, with unchanged source hashes. RUAD mapping and physical-display/screen-reader acceptance remain outstanding.
