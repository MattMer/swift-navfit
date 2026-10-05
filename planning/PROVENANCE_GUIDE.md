# Retained import provenance

> Current native navigation: use the [workspace guide](UX_WORKSPACE_GUIDE.md) and
> [current state](CURRENT_STATE.md). This guide retains service details and
> earlier verification history; legacy screen instructions may differ.

Swift NAVFIT 0.6.1 adds a read-only viewer for retained import receipts and source values. From Home, open **Operations and import undo**, select an import receipt, then choose **Inspect selected import provenance**.

The receipt shows the recorded source path, SHA256, import time, contributor, adapter and scalar encoding, before/applied revision IDs and undo status. It does not open or depend on the source file. Undone imports keep their evidence.

Choose Reports, Folders, Summary or All. Optionally enter a source row and a field-name search, then Apply filters. Source rows are zero-based capture ordinals, not NAVFIT report IDs. Every receipt retains the whole source database's three data tables, not just the report selected for narrative import. Select a field to read its complete stored value. NULL and empty text have distinct labels; memo content is shown in a locked scrolling textbox.

Dates are retained as OLE numeric values and scalar values use the receipt's recorded VBA CStr encoding and decimal separator. The viewer displays the captured representation without converting it. Field ordinals and DAO type codes are retained evidence. These values are not current Swift report values, a full import mapping or an export adapter.

The release uses schema 6. To carry existing data into the new app, run scripts/upgrade-workspace.ps1 with the existing workspace as SourcePath and a new DestinationPath. The script copies and verifies data without changing the source file. Do not replace a populated workspace with the blank distribution.

Verification: 24 viewer assertions including scope, apostrophe filters, zero-based rows, invalid rows, long text, null/empty labels, no saved data changes and reopening; 54 integrated workflow assertions; 46 smoke checks. Schema-6-to-6 upgrade compared every original row and field, preserved 2475 retained values in the synthetic test fixture and reopened successfully. Full visual inspection and Access Runtime/x86 remain untested.
