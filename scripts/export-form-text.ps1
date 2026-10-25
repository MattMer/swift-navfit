param([Parameter(Mandatory=$true)][string]$BuildPath,[string[]]$Form,[string]$OutDir)
# Forms are source. After changing a form in Access on a disposable dev build,
# export it here to write normalized text into src/forms (or -OutDir). Without
# -Form every form is exported. The build itself is never opened; a copy is.
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
if(-not $OutDir){$OutDir=Join-Path $root 'src/forms'}
[void](New-Item -ItemType Directory -Force -Path $OutDir)
#Per-build noise that LoadFromText does not need: the checksum, generated GUID,
#name-map and record-source-date blocks, and the BaseInfo combo cache (it embeds
#generated index names). Everything else is kept exactly.
function NormalizeFormText([string]$text){
 $out=[Collections.Generic.List[string]]::new();$skip=-1;$inBase=$false
 foreach($line in $text -split "`r`n|`n"){
  $indent=$line.Length-$line.TrimStart().Length
  if($skip -ge 0){if($line.Trim() -eq 'End' -and $indent -eq $skip){$skip=-1};continue}
  if($line -match '^\s*(GUID|NameMap|RecSrcDt)\s*=\s*Begin\s*$'){$skip=$indent;continue}
  if($line.StartsWith('Checksum =')){continue}
  if($inBase -and $line.TrimStart().StartsWith('"')){continue}
  $inBase=$line.TrimStart().StartsWith('BaseInfo =')
  if($inBase){continue}
  $out.Add($line)
 }
 while($out.Count -and $out[$out.Count-1] -eq ''){$out.RemoveAt($out.Count-1)}
 return ($out -join "`r`n")+"`r`n"
}
$copy=Join-Path $root ('build/export-forms-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $copy
$staging=Join-Path $root ('build/export-forms-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $staging)
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($copy,$true)
 foreach($open in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$open,2)}
 $all=@($app.CurrentProject.AllForms|ForEach-Object {$_.Name})
 $names=if($Form){$Form}else{$all}
 $unknown=@($names|Where-Object {$all -notcontains $_});if($unknown.Count){throw ('No such form in the build: '+($unknown -join ', '))}
 foreach($name in $names){
  $raw=Join-Path $staging ($name+'.txt');$app.SaveAsText(2,$name,$raw)
  [IO.File]::WriteAllText((Join-Path $OutDir ($name+'.form.txt')),(NormalizeFormText ([IO.File]::ReadAllText($raw))),[Text.UTF8Encoding]::new($false))
 }
 Write-Output ('Exported '+@($names).Count+' form(s) to '+$OutDir+'. Add any new form to "forms" in build-manifest.json.')
}finally{
 $app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
 Remove-Item -LiteralPath $staging -Recurse -ErrorAction SilentlyContinue
 for($i=0;$i -lt 20 -and (Test-Path -LiteralPath $copy);$i++){try{Remove-Item -LiteralPath $copy -ErrorAction Stop}catch{Start-Sleep -Milliseconds 250}}
}
