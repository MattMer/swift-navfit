param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$evidence=Join-Path $root ('build/tests/review-build-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $evidence)
$audit=Join-Path $evidence 'audit.accdb';Copy-Item -LiteralPath $BuildPath -Destination $audit
function NormalizeVBA([string]$text){
$text=($text.Replace("`r`n","`n") -replace '(?m)^Attribute VB_Name = .*\n','').Trim()
$parts=[regex]::Matches($text,'"(?:[^"]|"")*"|[^"]+')
$result=($parts|ForEach-Object{if($_.Value.StartsWith('"')){$_.Value}else{
$code=$_.Value.ToLowerInvariant()
$code=$code -replace '(?m)(call [a-z0-9_]+)\(\)(?=\s*(?:$|:))','$1'
$code -replace '(?m)^[ \t]+','' -replace ':[ \t]+',':'
}}) -join ''
return $result.Trim()
}
$hash=(Get-FileHash -LiteralPath $BuildPath).Hash;$app=New-Object -ComObject Access.Application;$modules=0
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($audit,$true)
foreach($file in Get-ChildItem (Join-Path $root 'src') -Filter '*.bas'){
$export=Join-Path $evidence ($file.BaseName+'.bas');$app.SaveAsText(5,$file.BaseName,$export)
if((NormalizeVBA ([IO.File]::ReadAllText($export))) -cne (NormalizeVBA ([IO.File]::ReadAllText($file.FullName)))){throw ('Installed code differs: '+$file.BaseName)};$modules++
}
foreach($form in @('frmEditor','frmUXReview','frmLibrary')){
$app.DoCmd.OpenForm($form,1);$f=$app.Forms.Item($form)
switch($form){
'frmEditor'{if($f.OnClose -ne '=UXQueueReportReturn()' -or $f.OnUnload -ne '[Event Procedure]'){throw 'Editor native return/unload guards missing'}}
'frmUXReview'{if($f.OnLoad -ne '=UXLoadReview()' -or $f.OnResize -ne '=UXReviewLayout()'){throw 'Review load/resize hooks missing'}}
'frmLibrary'{if($f.OnTimer -ne '=UXFinishReportReturn()'){throw 'Native delayed Reports return hook missing'}}
}
$app.DoCmd.Close(2,$form,2)
}
$app.DoCmd.RunCommand(126)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $hash){throw 'Audit changed deliverable'}
[pscustomobject]@{BuildPath=[IO.Path]::GetFullPath($BuildPath);SHA256=$hash;StandardModules=$modules;NativeEvents='PASS';SourceMatch='PASS';SourceUnchanged=$true}|ConvertTo-Json|Set-Content (Join-Path $evidence 'verification.json')
Write-Output ('PASS: '+$modules+' installed modules match current source; native review/return events and compile verified; deliverable unchanged. Evidence: '+$evidence)
