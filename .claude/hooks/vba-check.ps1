# PostToolUse: run the fast VBA source check on an edited src/ or tests/ module and
# send any errors back to Claude so they are fixed before a 9-minute build.
$ErrorActionPreference='Stop'
$event=[Console]::In.ReadToEnd()|ConvertFrom-Json
$root=if($env:CLAUDE_PROJECT_DIR){$env:CLAUDE_PROJECT_DIR}else{Split-Path (Split-Path $PSScriptRoot -Parent) -Parent}
$target=[string]$event.tool_input.file_path
if(-not $target -or $target -notmatch '(?i)\.bas$'){exit 0}
$full=[IO.Path]::GetFullPath($(if([IO.Path]::IsPathRooted($target)){$target}else{Join-Path $root $target}))
$inScope=$false
foreach($folder in 'src','tests'){
 if($full.StartsWith([IO.Path]::GetFullPath((Join-Path $root $folder)).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){$inScope=$true}
}
if(-not $inScope){exit 0}
$output=& powershell.exe -NoProfile -File (Join-Path $root 'scripts\check-vba-source.ps1') -Path $full 2>&1
if($LASTEXITCODE -ne 0){
 @{decision='block';reason=('VBA source check failed. Fix before building:'+"`n"+(($output|Where-Object {$_ -like 'ERROR *'}) -join "`n"))}|ConvertTo-Json -Compress
}
exit 0
