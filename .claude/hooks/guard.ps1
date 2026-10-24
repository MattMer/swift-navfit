# PreToolUse guard for the AGENTS.md protected-state and Access-session rules.
# Denies edits to protected folders and name-based Access kills. Asks before shell
# commands that delete, move or write into protected folders.
$ErrorActionPreference='Stop'
$event=[Console]::In.ReadToEnd()|ConvertFrom-Json
$root=if($env:CLAUDE_PROJECT_DIR){$env:CLAUDE_PROJECT_DIR}else{Split-Path (Split-Path $PSScriptRoot -Parent) -Parent}
$protected='output','FORMS_REFERENCES_TEMPLATES','NAVFIT98A','NAVFIT98A_v33p2_Upgrade_Tool'
function Decide([string]$decision,[string]$reason){
 @{hookSpecificOutput=@{hookEventName='PreToolUse';permissionDecision=$decision;permissionDecisionReason=$reason}}|ConvertTo-Json -Compress
 exit 0
}
$tool=[string]$event.tool_name
if($tool -in 'Edit','Write','NotebookEdit'){
 $target=[string]$event.tool_input.file_path
 if(-not $target){$target=[string]$event.tool_input.notebook_path}
 if($target){
  $full=[IO.Path]::GetFullPath($(if([IO.Path]::IsPathRooted($target)){$target}else{Join-Path $root $target}))
  foreach($folder in $protected){
   $base=[IO.Path]::GetFullPath((Join-Path $root $folder)).TrimEnd('\')+'\'
   if($full.StartsWith($base,[StringComparison]::OrdinalIgnoreCase)){Decide 'deny' "$folder/ is protected (published releases or supplied reference material; see AGENTS.md). Work on a copy outside it."}
  }
 }
}
elseif($tool -in 'Bash','PowerShell'){
 $command=[string]$event.tool_input.command
 # Access sessions are closed by owned PID only; a name-based kill can destroy the user's unsaved work.
 if($command -match '(?i)(taskkill\b[^|;&\r\n]*/im\s+"?msaccess|stop-process\b[^|;&\r\n]*-(name|processname)\s+"?msaccess|msaccess[^;&\r\n]*\|\s*(stop-process|kill)\b|\b(pkill|killall)\b[^|;&\r\n]*msaccess|\bkill\s+-name\s+"?msaccess)'){
  Decide 'deny' 'Name-based termination of Microsoft Access is blocked (AGENTS.md: close only sessions you own, by verified PID, after a normal save/close attempt). Use Stop-Process -Id <owned PID> or taskkill /PID <owned PID>.'
 }
 $paths=($protected|ForEach-Object {[regex]::Escape($_)}) -join '|'
 $mutating='(?i)(\b(rm|rmdir|del|erase|rd|mv|move|ren|rename|Remove-Item|Move-Item|Rename-Item|Set-Content|Add-Content|Out-File|Clear-Content|New-Item)\b[^|;&\r\n]*(^|[\s"''/\\])('+$paths+')([/\\"''\s]|$))|(>{1,2}\s*"?[^\s|;&]*(^|[/\\"\s])('+$paths+')[/\\])'
 if($command -match $mutating){
  Decide 'ask' 'This command may delete, move or write inside a protected folder (output/, FORMS_REFERENCES_TEMPLATES/, NAVFIT98A*). AGENTS.md allows that only for an explicitly requested release.'
 }
}
exit 0
