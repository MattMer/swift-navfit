# Navigation and contextual help consolidation

Implemented 2026-10-19 on a separate copy of the batch 6 candidate.

## Behavior

- A small `?` at the top right of each standalone work window opens the existing modeless pop-out workspace guide. The guide names the originating window, selects its related topic, and provides a window-specific introduction for report editing, shared fields, members, review, directory/history, notes, narrative resources, imports and settings.
- The guide retains its topic list so users can move to related workflows. Closing it returns to the open work window. Opening help does not run editor-close/navigation logic, flush pending edits or create revisions.
- Legacy Saving/Folder help links and bare `frmHelp` calls resolve to the same current guide. Redundant legacy inline help buttons are hidden; Settings retains the complete workspace-guide entry. Legacy form definitions remain for compatibility but are no longer the visible help destination.
- Top-right help placement follows window resize and reserves space in top-line labels. The button is excluded from general body scaling. Embedded field/table forms and the decision dialog do not receive redundant help buttons.
- Group tabs use one explicit group-context resolver with page validation before mutation. The resolver reads the originating window, with the workbench's own selected group as a compatibility fallback. It never borrows another window's hidden context. Existing pending-default/cell guards and hidden adapter hosts remain in place.

## Evidence and release scope

See NAVIGATION_HELP_VERIFICATION.json for final native results, artifact hash, button inventory and source alignment. Tests use disposable copies. The previous batch 6 and populated r116 workspace retain their hashes.

The standard builder installs this consolidation after theme/priority-action application. The incremental builder refuses to overwrite an output and checks the template hash. This is a development candidate, not a published release. A full standard form-generation build, physical display/scaling, keyboard/screen-reader, Access Runtime and x86 acceptance remain pending. The broader removal of hidden group adapter hosts is not part of this change; their existing pending-edit safeguards remain necessary.

Final native results: 57 focused help, 39 group correctness, 200 defaults, 61 live-member and 264 fixed-page assertions passed. Structured application smoke Status=PASS matches the candidate hash. Its log includes passing injected-failure/rollback test names; those keyword matches are retained in the verification record. All 61 button bindings and six installed source modules were verified. Candidate: build/navigation-help-final.accdb; manifest: build/navigation-help-final.manifest.json.
