# Forms are source: src/forms holds each form as normalized Access text (UTF-8
# in git; Access loads UTF-16). Dot-source in a scope that defines $app and $root.
function InstallSwiftFormText([string[]]$Names){
 $folder=Join-Path $root 'src/forms'
 $files=@(Get-ChildItem -LiteralPath $folder -Filter '*.form.txt'|Where-Object {-not $Names -or $Names -contains $_.Name.Substring(0,$_.Name.Length-9)})
 if($Names -and $files.Count -ne @($Names).Count){throw ('Missing form text: '+((@($Names)|Where-Object {-not (Test-Path -LiteralPath (Join-Path $folder ($_+'.form.txt')))})-join ', '))}
 $staging=Join-Path $root ('build/forms-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $staging)
 try{
  $existing=@($app.CurrentProject.AllForms|ForEach-Object {$_.Name})
  foreach($file in $files){
   $name=$file.Name.Substring(0,$file.Name.Length-9)
   $staged=Join-Path $staging $file.Name
   [IO.File]::WriteAllText($staged,[IO.File]::ReadAllText($file.FullName,[Text.Encoding]::UTF8),[Text.Encoding]::Unicode)
   if($existing -contains $name){$app.DoCmd.Close(2,$name,2);$app.DoCmd.DeleteObject(2,$name)}
   try{$app.LoadFromText(2,$name,$staged)}catch{throw ('Form text failed to load: '+$file.Name+': '+$_.Exception.Message)}
  }
 }finally{Remove-Item -LiteralPath $staging -Recurse -ErrorAction SilentlyContinue}
 return $files.Count
}
