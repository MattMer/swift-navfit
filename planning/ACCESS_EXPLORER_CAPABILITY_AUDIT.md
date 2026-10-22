# Access Explorer capability audit

21 October 2026. Audit of the installed backend and currently callable MCP tool schemas; not certification of every mutation tool.

## Outcome

**66 tools are callable in this chat.** The previously recorded tool-discovery limitation no longer applies. Installed backend commit is `a53753d62c04a703d5ee0c586a62dbebcbe1028a`; setup identifies backend 0.7.64 and Access Explorer 1.2.0. No extension upgrade or configuration change was made.

Best fit: targeted native inspection, installed-code/control/relationship discovery, and disposable-copy diagnostics. Keep tracked source, manifest generation, production migration and focused workflow tests authoritative.

## Evidence and verification boundaries

Inspected the installed tool descriptions and backend modules: security.py, server.py, dispatcher.py, sql.py, compile.py, core.py, plus SECURITY.md. These live in the user-managed Access Explorer runtime, not maintained Swift source.

Representative live probes used a new copy of the r128 blank under:
`build/tests/access-capability-0e84bb6f67ab4549a0a283f6fd1b3045/scratch.accdb`.

| Probe | Result |
| --- | --- |
| access_list_objects (modules) | Success; native module inventory returned. |
| access_execute_sql (AppMetadata SELECT) | 0.27.1-dev / schema 27. |
| access_table_info (Report) | Fields/types and table metadata returned. |
| access_list_relationships | 111 relationships returned. |
| access_vbe_get_proc (modSwift.NewIdentity) | Procedure text, boundaries and declaration line returned. |
| access_list_controls (frmEditor) | 242 controls returned with selected fields. |
| access_lint_form (frmEditor, heuristic) | Tool succeeded; verdict FAIL, 2 errors, 2,377 warnings, 292 info findings. Of these, 2,365 were overlap warnings. |
| access_close | Closed owned sessions normally; no Access processes remained at the first post-probe process check. |

Initial procedure lookup mistakenly requested CurrentRevision in modSwift; the backend returned a descriptive error. Corrected lookup for NewIdentity succeeded. **The error response carried isError=false**: validate text/JSON error fields too.

Original r128 blank SHA-256 remained `55CC0C7C6B686A5F312FA52FC0DEA036D5FCA1F630B39F1ACDD39EAE266C8277` after the seven-probe run. Original was never passed to an Explorer tool. No mutation, repair, compilation, UI action, transfer or code-execution tools were exercised. The final corrected procedure read also targeted only the scratch.

## Material findings

1. **Not a read-only connection.** No enforced readonly/safe/full family mode is implemented. Tools accept database paths and operate with the launching user's privileges. Inspection labels below describe intended operation, not immutable-file guarantees.
2. **Arbitrary execution gate is narrow.** access_run_vba, access_eval_vba and access_run_macro are absent from the 66-tool inventory. They are gated by MCP_ACCESS_ALLOW_CODE_EXEC. Source reads the process environment on each call; changing launch configuration requires a process restart to affect that environment. Do not confuse the gate with a general write or event-execution prohibition.
3. **Forms/reports and UI can execute events.** Screenshot/open, report export and UI actions can run existing event code even with the arbitrary-execution gate closed. Opening a database uses startup mitigations, including AutomationSecurity and a global SHIFT bypass; these are not a read-only sandbox.
4. **Compile has a significant side effect.** access_compile_vba automatically decompiles a database not yet decompiled in that session, then compiles/saves. It can discard unsaved VBA. access_vbe_check_syntax is less intrusive but not a compiler.
5. **SQL batches are not atomic transactions.** The inspected loop executes statements individually without BeginTrans/CommitTrans/Rollback. stop_on_error stops later statements, not earlier committed writes. DELETE/DROP/ALTER confirmation is an accident guard; INSERT/UPDATE and many other writes remain available.
6. **Static UI lint needs interpretation.** Fixed Front/Back controls and state-dependent visibility can share exported coordinates. The 2,671 findings are not 2,671 verified visible defects; the two contrast errors also need actual-state review. No UI remediation is authorized by this audit.
7. **Search is broad but not semantic reachability.** find_usages plus search_controls(properties=['all']), query/macro reads and VBA searches cover more native surfaces than repository text alone. They still cannot prove that dynamically constructed names, external consumers or every callback are unused.
8. **Maintenance/session tools are not passive.** Compact/decompile replace files or launch processes. Owned-session close has a force-kill fallback after timeout; avoid sessions holding user edits. No force-kill was needed during the probes.
9. **Import/export is not migration acceptance.** Excel/CSV transfer and text object imports lack Swift's row/field/ID/history/provenance/rollback guarantees. Report PDF export requires an existing Access report and does not satisfy the official-template output backlog.
10. **No breakpoint/step debugger or test-runner capability is advertised.** Procedure editing/search and compile support do not replace native workflow tests, save/recovery acceptance, physical accessibility or print checks.

## Complete tool inventory

Inspect tools can still open/attach to Access or interact with native object state. Export/UI tools can write files and/or open event-bearing objects. Read/write tools depend on the selected action. All write operations belong on disposable development copies with changes brought back to tracked source.

| Tool | Boundary | Capability / practical constraint |
| --- | --- | --- |
| access_alter_table | Write/maintenance | Modifies the structure of an Access table via DAO. |
| access_build_form | Write/maintenance | Generic data-entry layout builder, unsuitable as replacement for fixed NAVPERS block geometry. |
| access_clone_object | Write/maintenance | Clones a form, report, module, class_module, query or macro to a new name. |
| access_close | Session | Owned sessions quit; backend has force-kill fallback after timeout; attached sessions are released rather than quit. |
| access_compact_repair | Write/maintenance | Replaces the database via a temporary compacted file; scratch/backup only. |
| access_compile_vba | Write/maintenance | Compiles/saves, but first auto-decompiles; use only scratch copies, not unsaved user sessions. |
| access_create_control | Write/maintenance | Creates a control on a form/report via COM. |
| access_create_database | Write/maintenance | Creates an empty Access database (.accdb). |
| access_create_form | Write/maintenance | Creates a new form in the database. |
| access_create_relationship | Write/maintenance | Creates a relationship between two tables. |
| access_create_table | Write/maintenance | Creates an Access table via DAO with full support: types, defaults, descriptions and primary key — all in a single call. |
| access_decompile_compact | Write/maintenance | File/process-changing maintenance; percentage reduction is an upstream claim, not measured here. |
| access_delete_control | Write/maintenance | Deletes a control from a form/report via COM.. |
| access_delete_object | Write/maintenance | Deletes an Access object (module, form, report, query, macro). |
| access_delete_relationship | Write/maintenance | Deletes a relationship between tables by name. |
| access_execute_batch | Read/write | Sequential statements, NOT a batch transaction; stop-on-error does not roll back earlier writes. |
| access_execute_sql | Read/write | SELECT inspection or DML/DDL; INSERT/UPDATE do not require destructive confirmation. |
| access_export_structure | Export/UI | Generates Markdown with database structure: modules with signatures, forms, reports, queries, macros. |
| access_export_text | Export/UI | Exports form/report/module as text (SaveAsText). |
| access_find_definition | Inspect | Go-to-definition for a VBA symbol — the mirror of access_find_usages. |
| access_find_usages | Inspect | Text/regex across VBA, SQL and selected control properties; not proof of complete dynamic reachability. |
| access_get_code | Inspect | Reads code/definition of an Access object. |
| access_get_control | Inspect | Full definition (Begin...End) of a control by name. |
| access_get_db_property | Inspect | Reads a database property (CurrentDb.Properties) or Access option (GetOption).. |
| access_get_field_properties | Inspect | Reads all properties of a field: DefaultValue, ValidationRule, Description, Format, etc.. |
| access_get_form_property | Inspect | Reads properties of a form or report (RecordSource, Caption, DefaultView, HasModule, etc.). |
| access_import_text | Write/maintenance | Replaces existing objects; text import is a write. |
| access_lint_form | Inspect | Static exported geometry/contrast checks; no physical accessibility or printed-fit certification. |
| access_list_controls | Inspect | Lists controls of a form/report with name, type, caption, control_source, position. |
| access_list_indexes | Inspect | Lists indexes of a table: name, fields, primary, unique, foreign.. |
| access_list_linked_tables | Inspect | Use password masking; connection strings can contain credentials. |
| access_list_objects | Inspect | Lists database objects by type (table, module, form, report, query, macro, all). |
| access_list_references | Inspect | Lists VBA references: name, GUID, path, is_broken, built_in.. |
| access_list_relationships | Inspect | Lists relationships between tables: name, tables, fields, cascade flags.. |
| access_list_startup_options | Inspect | Lists the 14 common startup options with their current values.. |
| access_manage_index | Write/maintenance | Creates or deletes an index. |
| access_manage_query | Read/write | Manages QueryDefs: create, modify, delete (requires confirm=true), rename, get_sql. |
| access_manage_reference | Write/maintenance | Adds (add) or removes (remove) a VBA reference. |
| access_manage_tab_order | Read/write | Reads or sets the TabIndex of controls on a form/report. |
| access_output_report | Export/UI | Opens existing native Access reports; may execute events; does not implement official NAVPERS output. |
| access_relink_table | Write/maintenance | Changes the connect string of a linked table and refreshes. |
| access_screenshot | Export/UI | Can open a form/report and pump timers; not a passive file read. |
| access_search_controls | Inspect | Use properties=['all'] to include event expressions; complements find_usages. |
| access_search_data | Inspect | Searches a text string in any Text/Memo field of any LOCAL table. |
| access_search_queries | Inspect | Searches for text or regex in the SQL of ALL queries. |
| access_set_code | Write/maintenance | Overwrites/creates native objects; built-in failure restore is not a workspace backup; normalize input to LF. |
| access_set_control_props | Write/maintenance | Modifies properties of a control via COM. |
| access_set_db_property | Write/maintenance | Sets a database property or Access option. |
| access_set_field_property | Write/maintenance | Sets a field property. |
| access_set_form_property | Write/maintenance | Sets form/report-level properties (RecordSource, Caption, DefaultView, HasModule, etc.) via COM in Design view. |
| access_set_multiple_controls | Write/maintenance | Modifies properties of multiple controls on a form/report in a single operation. |
| access_table_info | Inspect | Table structure via DAO: fields, types, size, required, record_count, is_linked.. |
| access_tips | Inspect | Tips and gotchas for working with Access via MCP. |
| access_transfer_data | Write/maintenance | Excel/CSV transfer, not Swift's verified workspace migration engine. |
| access_ui_click | Write/maintenance | Can trigger event code and change data; coordinate/focus dependent. |
| access_ui_type | Write/maintenance | Can trigger event code and change data; keyboard/focus dependent. |
| access_vbe_append | Write/maintenance | Appends code to the end of a VBA module.. |
| access_vbe_check_syntax | Inspect | Structural check only, not identifier/type resolution or native compilation. |
| access_vbe_find | Inspect | Searches for text or regex in ONE specific VBA module. |
| access_vbe_get_lines | Inspect | Reads a range of lines from a VBA module via VBE COM. |
| access_vbe_get_proc | Inspect | Code of a VBA procedure by name. |
| access_vbe_module_info | Inspect | Index of procedures in a VBA module: total_lines, procs [{name, start_line, body_line, count}]. |
| access_vbe_patch_proc | Write/maintenance | Surgical find/replace WITHIN a VBA procedure. |
| access_vbe_replace_lines | Write/maintenance | Replaces lines in a VBA module via VBE. |
| access_vbe_replace_proc | Write/maintenance | Replaces an entire VBA procedure by name (boundaries found via VBE). |
| access_vbe_search_all | Inspect | Searches for text or regex in ALL VBA modules (modules, forms, reports) in the database. |

## Recommended use in Swift NAVFIT

- Use object/schema/index/relationship/reference inspection to check a built artifact against tracked inputs.
- Use procedure reads and definition/search tools for native diagnostics; inspect control event expressions and macros before retiring callbacks.
- Use exported control properties and targeted lint rules to investigate geometry. Keep fixed Front/Back layouts and the left rail; do not introduce the generic declarative form builder as a second UI generator.
- Prefer direct tracked-source edits. Surgical native edits are experimental and limited to disposable copies; verify each relevant mutation individually before routine adoption, then port useful results to authoritative source and run applicable established checks. Schemas alone do not establish reliability. Avoid accidental CRLF doubling on text imports.
- Keep production workspace migration and release checks in the existing project scripts.
- Treat current capabilities as a practical inspection/editing toolbox, not justification to retain historical test implementations. The two serve different purposes.

## Changes and exclusions

This audit adds this report only. No server settings, dependencies, production VBA, build generators or user workspaces changed. All 66 schemas were reviewed; seven representative operation types were successfully probed, not all 66 tools executed.
