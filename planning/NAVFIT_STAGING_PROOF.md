# NAVFIT staging proof — 3 October 2026

`scripts/stage-navfit-proof.ps1` creates an isolated, quarantined development batch. It does not modify the Swift application or import reports into its draft library.

The supplied TEST_DB passed comparison of all 824 cells across Folders (4 rows), Reports (4 rows), and Summary (0 rows). Field types, nulls and ordinal string values were compared. The staged database reopened successfully and the original SHA256 remained unchanged. Three report rows are pending review; the null-type row is quarantined. Parent values remain verbatim, including the unresolved `a ` prefix.

Evidence is in `build/navfit-staging-verified4/E732EE41A384495AC37BC009C59516A73461C7DB99E251B5384CE917652386E7/verification.json`. Data stays in the local batch; the JSON contains counts and status only.

The harness recognizes the observed data-table names, ordered field names/types/sizes and index definitions, and rejects additional data tables, relationships or saved queries. This is recognition of this specimen's structural shape, not certification of every Patch 2 database. Schema rejection, duplicate-batch rejection and injected-failure tests remain pending.

Native database copying preserves decimal precision and field properties. Tables are renamed RawFolders, RawReports and RawSummary, with a separate ReviewQueue. The copy can retain other source database objects; DAO does not execute source macros or VBA. Do not open the quarantined database as a trusted application. It is not a distributable blank template. Failed attempts are retained for diagnosis and never reported as passing batches.

Next: prove the logical parent encoding and form-specific field meanings, then implement review and deliberate commit in Swift. Export and NAVFIT reimport/printing remain unimplemented and unverified. The tested Swift 0.2.1 artifact remains the current application delivery.
