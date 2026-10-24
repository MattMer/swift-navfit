@AGENTS.md

## Claude Code notes

- The Access Explorer MCP and the `swift_files` reader described in
  docs/Development_Workflow.md are registered for Codex and VS Code, not for
  Claude Code. Read source directly. For native Access work, use the existing
  PowerShell runners, which drive Access through COM.
- Run Access scripts with `powershell.exe -NoProfile -File ...` (Windows
  PowerShell 5.1), not `pwsh` or WSL.
- `scripts/develop.ps1 -Action BuildAndTest` takes about 11 minutes. Run it in
  the background and wait for the completion notification.
- [.claude/settings.json](.claude/settings.json) enforces two AGENTS.md rules
  with hooks, plus a permissions allowlist:
  - `.claude/hooks/guard.ps1` denies edits under `output/`,
    `FORMS_REFERENCES_TEMPLATES/` and `NAVFIT98A*`, and denies name-based
    Access kills. It asks before shell commands that delete, move or write into
    those folders.
  - `.claude/hooks/vba-check.ps1` runs `scripts/check-vba-source.ps1` after
    each edit to a `src/` or `tests/` `.bas` file and reports its errors back.
  - Put personal overrides in `.claude/settings.local.json`, which is git-ignored.
