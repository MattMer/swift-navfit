# Bundled NAVFIT resources

The user supplied a local NAVFIT98A folder on 3 October 2026. Executables, databases, logos and FAQ remain local and are excluded from Git. No executable was launched by this inspection. Databases were inspected through read-only DAO scratch copies; original SHA256 values were unchanged.

## Identity and template

NAVFIT98A.exe has file/product version 2.2.0.332 and SHA256 `5366F1C1BB147F1706094C5E3E1D62981F34FC6CB301E4006A892ACFBFAB58CF`, matching the previously inspected installed executable. Its configuration declares .NET Framework 4.8.1, ACE OLEDB 16.0, the version, a displayed release date of August 9, 2026 and bundled data/template paths. These are package settings, not verification of workstation prerequisites or authentication behavior.

Data/NF98A_empty.accdb has SHA256 `43348B36198AF6D8E3D34A5F40090BB06AE1D0ADC01295066BAA8799A9078252`. Ordered names, DAO types and sizes match the inspected TEST_DB tables: Folders has 78 fields/one row, Reports 128 fields/one row, Summary six fields/no rows. The default Reports row is retained; it is not assumed disposable. A full property fingerprint and actual Swift export proof remain pending.

## Lookup catalogue

Data/NF98A_Dte.accdb has SHA256 `ABB8CA070891CB7BF76D3CB93599F9BBEC40FAAB7BFAF0124B594E2A3E1577A8`. Its application tables are:

| Table | Rows | Useful structure |
| --- | ---: | --- |
| RankRating | 1130 | Rank/rate, officer/enlisted indicator, grade |
| OtherRanks | 36 | Other-service rank, display description, grade |
| WarfareDevices | 21 | Warfare designations |
| PerformanceTraits | 21 | Form kind, trait name, official block number/title/descriptions |
| TraitPerformance | 6 | Performance grade/value |
| TraitPerfRel | 126 | Performance-to-trait descriptions |
| PerDates | 69 | Periodic class, grade, date code |
| PromotionRecom | 6 | Recommendation display/value |
| PromotionStatus | 4 | Promotion-status descriptions |
| PhysicalReadiness | 10 | Readiness descriptions |
| BilletSubCat | 37 | Billet subcategory descriptions |
| StateTbl | 55 | State code/descriptions |

These are reusable implementation references for code lists, form-specific trait captions and grade relationships. They should prevent reconstructing dropdowns and trait descriptions from screenshots. They are not yet embedded in Swift or treated as a Navy policy rule engine. Future adapters must preserve nulls, codes and form applicability, verify joins/relationships, and cross-check policy applicability against the governing instruction before enforcing rules. Package data alone does not establish recommendation limits or historical rule applicability.

The FAQ and logos were inventoried; FAQ contents were not reviewed in this inspection. Raw local schema evidence is under planning/evidence/navfit and is intentionally not committed.
