# Guided import workflow

Implemented in r22, version 0.22.0-ux-preview, schema 21. This changes import
interaction and adds an optional reviewed-source hash guard to narrative import;
it does not add mappings or expand the supported template scope.

## Current journey

Reports > Import reports opens a chooser for NAVFIT databases and structured
PDFs. Both paths show Choose, Preview, Apply and Results, full before/after
values and supported mapping scope. These workflows create new partial drafts;
the existing-draft reconciliation comparison remains separate.

NAVFIT source selectors show member names with hidden structural IDs. Preview
maps form type and official narrative, with the independently entered local
display name. Other fields remain unmapped and retained in provenance. Source
field inspection remains available without replacing the mapping preview.

PDF preview uses official form-specific captions, date conversion and explicitly
chosen identifier interpretation. Normal options are SSN and Retain only.
DoDID appears only after explicit legacy-compatibility selection. Six-digit
dates require an explicit century. Unsupported/flattened PDFs remain rejected;
Acrobat automation remains required for PDF extraction.

Changing member selection, display name, PDF source or mapping choices
invalidates Apply. Applying requires the current reviewed source/hash and
choices. Success disables repeated Apply, retains the reviewed full values and
shows Results with an Open imported report action. Failures preserve the source
choice and entered options; a fresh preview is required. Reopening PDF import
clears hidden preview state. Existing transactional rollback, duplicate guards,
source preservation, receipts and provenance services remain in use.

## Verification

Native guided-flow tests exercise stale choices/source rejection, full mapping
values, legacy opt-in, apply/result states, repeated apply rejection, created
report navigation and reopen state. Existing PDF tests cover all three report
types; NAVFIT commit tests cover quarantine, duplicate rejection, late-failure
rollback and all 2,472 provenance values. Read-only source inspection tests are
retained. Human display, keyboard and assistive-technology acceptance remains
pending; no printed fit or complete interchange compliance is claimed.

Preview builds run the NAVFIT flow gate without Acrobat. Add
`-IncludePDFIntegrationChecks` to build-preview.ps1 for the complete PDF flow.
Candidate packaging includes the complete PDF integration gate on the verified
build host, with logs retained in the package.
