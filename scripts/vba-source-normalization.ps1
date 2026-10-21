# Verify a reopened database, never just in-memory design objects.
function NormalizeSwiftVBA([string]$Text){
 $Text=[regex]::Replace($Text,'(?m)^Attribute .*\r?\n','')
 $Text=[regex]::Replace($Text,'"(?:[^"]|"")*"|\x27[^\r\n]*|[^"\x27]+',{param($m) if($m.Value.StartsWith('"') -or $m.Value.StartsWith("'")){$m.Value}else{[regex]::Replace($m.Value,'(?i)(\bCall\s+\w+)\s*\(\s*\)','$1')}})
 $parts=[regex]::Matches($Text,'"(?:[^"]|"")*"|\x27[^\r\n]*|\s+|[^\s"\x27]+')
 ($parts|ForEach-Object {$v=$_.Value;if($v.StartsWith('"')){$v}elseif(-not $v.StartsWith("'") -and -not [string]::IsNullOrWhiteSpace($v)){$v.ToLowerInvariant()}}) -join ''
}
