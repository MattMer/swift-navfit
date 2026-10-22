# Access Explorer MCP connection

Access Explorer's MCP backend is installed in its user-level managed runtime
folder, outside the source repository. It uses isolated Python 3.12.3 with MCP
1.30.0, pywin32 312 and Pillow 13.0.0. The upstream checkout is pinned at commit
`a53753d62c04a703d5ee0c586a62dbebcbe1028a` (mcp-msaccess-database 0.7.64).

Access Explorer user settings point to that interpreter and server. VS Code
user MCP registration uses `access-explorer-local`; Codex user registration
uses `access_explorer`. Existing user settings and Codex configuration were
backed up before adding the entries. No existing registrations were removed.

The smoke test in [test-access-mcp.py](../scripts/test-access-mcp.py) used a
uniquely named copy of the verified fresh r127 development blank. It completed
MCP initialization, discovered 66 tools, listed modules, queried metadata
(0.27.0-dev/schema 27) and ReportCount=0, read modSwift, and closed Access through
`access_close`. The source template hash is unchanged. See
[machine evidence](ACCESS_MCP_SETUP_2026-10-21.json) for paths and exact results.

Arbitrary code-execution tools remain disabled by the backend's default gate.
No Office security settings were changed. This is backend protocol verification;
the existing chat's tool inventory did not hot-refresh during setup. Start the
registered VS Code MCP server and open a new chat when needed for agent access.
Interactive Access Explorer connections use the same backend settings.

Use disposable copies for investigation and bring changes back into tracked
source. This setup neither replaces the standard builder/importer nor publishes
a release. Existing workspaces, earlier releases and unrelated FORMS files are
preserved; owned Access processes are closed after testing.
