param([string[]]$Path)
# Fast static checks for VBA module sources, without Access.
# Errors are faults that break LoadFromText import or VBA compilation. Warnings are cosmetic.
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
if(-not $Path){$Path=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Filter '*.bas'|ForEach-Object FullName)}
$errors=0;$warnings=0
function Report([string]$level,[string]$file,[int]$line,[string]$message){
 $rel=$file;if($file.StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){$rel=$file.Substring($root.Length).TrimStart('\','/')}
 Write-Output ('{0} {1}:{2} {3}' -f $level,$rel,$line,$message)
}
$procedure=[regex]'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Function|Sub|Property\s+(?:Get|Let|Set))\s+(\w+)'
foreach($file in $Path){
 $file=[IO.Path]::GetFullPath($file)
 if(-not (Test-Path -LiteralPath $file)){Report 'ERROR' $file 0 'File not found.';$errors++;continue}
 $bytes=[IO.File]::ReadAllBytes($file)
 $text=[Text.Encoding]::GetEncoding(1252).GetString($bytes)
 $lines=[regex]::Split($text,'\r?\n')
 $name=[IO.Path]::GetFileNameWithoutExtension($file)
 # Access imports module text as ANSI; UTF-8 punctuation becomes mojibake in the installed module.
 for($i=0;$i -lt $bytes.Length;$i++){if($bytes[$i] -gt 127){$at=([Text.Encoding]::GetEncoding(1252).GetString($bytes,0,$i) -split '\n').Count;Report 'ERROR' $file $at ('Non-ASCII byte 0x{0:X2}; use ASCII (Chr/ChrW for special characters).' -f $bytes[$i]);$errors++;break}}
 if($lines[0] -ne ('Attribute VB_Name = "'+$name+'"')){Report 'ERROR' $file 1 ('First line must be Attribute VB_Name = "'+$name+'".');$errors++}
 if(-not ($lines|Where-Object {$_ -match '^\s*Option\s+Explicit\s*$'})){Report 'ERROR' $file 1 'Missing Option Explicit.';$errors++}
 $crlf=([regex]::Matches($text,"`r`n")).Count;$lf=([regex]::Matches($text,"(?<!`r)`n")).Count
 if($crlf -gt 0 -and $lf -gt 0){Report 'WARN' $file 0 ("Mixed line endings ($crlf CRLF, $lf LF); git normalizes on commit.");$warnings++}
 $run=0;$seen=@{};$start=0
 # A line starting "Name: ..." makes Name a line label, so a procedure written there
 # is never called. Real labels are GoTo/Resume targets; calls need "Call Name:".
 $labelTargets=@([regex]::Matches(($lines -join "`n"),'\b(?:GoTo|Resume|GoSub)\s+([A-Za-z_]\w*)')|ForEach-Object {$_.Groups[1].Value})
 for($n=0;$n -lt $lines.Count;$n++){
  $line=$lines[$n]
  if($procedure.IsMatch($line)){$start=$n}
  # Threat model E3: every XML parse refuses DTDs and external resources explicitly.
  if($line -match '\.LoadXML\b'){
   $body=$lines[$start..$n] -join "`n"
   if($body -notmatch '\.resolveExternals\s*=\s*False' -or $body -notmatch '"ProhibitDTD",\s*True'){Report 'ERROR' $file ($n+1) 'LoadXML without resolveExternals = False and ProhibitDTD True earlier in the procedure.';$errors++}
  }
  # Keywords cannot be variable or parameter names; VBA reports a syntax error only
  # when the module first runs, which a build does not do.
  $declared=[regex]::Matches($line,'(?i)(?:\b(?:Dim|Private|Public|Static|ByVal|ByRef|Optional)\s+|,\s*)([A-Za-z_]\w*)\s+As\b')
  foreach($d in $declared){if($d.Groups[1].Value -in @('Open','Imp','Eqv','Xor','Mod','Is','Like','Not','And','Or','Put','Get','Lock','Unlock','Input','Each','In','To','Then','Loop','Wend','Until','Shared')){Report 'ERROR' $file ($n+1) ("'"+$d.Groups[1].Value+"' is a VBA keyword and cannot name a variable or parameter.");$errors++}}
  $label=[regex]::Match($line,'^\s*([A-Za-z_]\w*):(?!=)\s*\S')
  if($label.Success -and $label.Groups[1].Value -notin @('Loop','Next','Else','Do','Wend','Exit','End','Resume','Case','Then','Rem') -and $label.Groups[1].Value -notin $labelTargets){Report 'ERROR' $file ($n+1) ("'"+$label.Groups[1].Value+":' starts the line, so VBA reads it as a label and never calls it. Write 'Call "+$label.Groups[1].Value+":' instead.");$errors++}
  if($line.Length -gt 1023){Report 'ERROR' $file ($n+1) ("Line is $($line.Length) characters; VBA allows 1023.");$errors++}
  if($line.TrimEnd() -match '\s_$'){$run++;if($run -eq 25){Report 'ERROR' $file ($n+1) 'More than 24 line continuations in one statement.';$errors++}}else{$run=0}
  $m=$procedure.Match($line)
  if($m.Success){
   $key=($m.Groups[1].Value -replace '\s+',' ').ToLowerInvariant()+' '+$m.Groups[2].Value.ToLowerInvariant()
   if($key -like 'function *' -or $key -like 'sub *'){$key='proc '+$m.Groups[2].Value.ToLowerInvariant()}
   if($seen.ContainsKey($key)){Report 'ERROR' $file ($n+1) ("Duplicate procedure '$($m.Groups[2].Value)' (first at line $($seen[$key])); VBA reports Ambiguous name.");$errors++}else{$seen[$key]=$n+1}
  }
 }
}
Write-Output ("{0} file(s): {1} error(s), {2} warning(s)." -f @($Path).Count,$errors,$warnings)
if($errors -gt 0){exit 1}
