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
