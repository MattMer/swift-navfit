# Defaults context switching — r107 (14 October 2026)

Fixed a stale shared-editor context when returning to Summary groups without closing Defaults.

- Removed the global-default entry point's refusal to open while `frmUXBulk` is loaded.
- Every defaults/shared-fields opening now finishes the old active control, closes the existing unbound host and opens a fresh context. Values, report type, group tabs, heading, OpenArgs and confirmation baseline reload together. Confirmed changes remain saved.
- The shared form remains shared. No extra copies of the UI or data schema were added.
- Regression coverage follows the actual list actions: open selected group defaults, focus Summary groups without Close, open global defaults, switch to another group, then switch back. Assert values/types/tabs and commit routing. The new checks include changing global and group defaults and verifying the other scope was not changed accidentally.
- Fixed a null comparison in the new test and removed unnecessary failure cleanup that obscured the original test error. Scratch fixtures are isolated and disposed by their runner.

Separate local candidates preserve r106 and existing user workspaces:

- `output/ux/Swift_NAVFIT_UI_Demo_2026-10-14-r107.accdb`
- `output/ux/Swift_NAVFIT_UI_Blank_2026-10-14-r107.accdb`

App 0.25.0-dev; schema 25. GitHub publication not requested. Native Access assertions and final artifact hashes/startup/blank-data checks are recorded in `RECURSIVE_DEFAULTS_R107_VERIFICATION.json`. Automated focus and injected confirmation choices do not replace manual keyboard/display/accessibility acceptance.
