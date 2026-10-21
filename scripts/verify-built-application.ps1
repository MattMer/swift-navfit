# Verify a reopened database, never just in-memory design objects.
. (Join-Path $PSScriptRoot 'vba-source-normalization.ps1')
$swiftBuildVerification=[ordered]@{SourceModules=0;Forms=0;NativeUnloadGuards=$false;RetiredFormsAbsent=$false;Identity=$false}
$checkDB=$app.CurrentDb();$meta=$checkDB.OpenRecordset('SELECT AppVersion,SchemaVersion FROM AppMetadata',4)
if($meta.Fields.Item('AppVersion').Value -ne $swiftBuildPlan.AppVersion -or $meta.Fields.Item('SchemaVersion').Value -ne $swiftBuildPlan.SchemaVersion){throw 'Built application identity mismatch'}
$meta.Close();$checkDB.Close();$swiftBuildVerification.Identity=$true
$exportFolder=Join-Path $root ('build/source-check-'+[guid]::NewGuid());[void](New-Item -ItemType Directory $exportFolder)
$functions=@{}
foreach($source in $swiftBuildPlan.Manifest.sourceModules){
 $name=[IO.Path]::GetFileNameWithoutExtension($source);$export=Join-Path $exportFolder ($name+'.bas')
 $app.SaveAsText(5,$name,$export)
 $expected=Get-Content (Join-Path $root $source) -Raw;$actual=Get-Content $export -Raw
 if((NormalizeSwiftVBA $actual) -cne (NormalizeSwiftVBA $expected)){throw ('Installed module differs from current source: '+$name)}
 $swiftBuildVerification.SourceModules++
 foreach($m in [regex]::Matches($actual,'(?im)^Public (?:Function|Sub) (\w+)')){$functions[$m.Groups[1].Value]=$true}
}
foreach($name in $swiftBuildPlan.Manifest.generatedModules){
 if(@($app.CurrentProject.AllModules|ForEach-Object {$_.Name}) -notcontains $name){throw ('Missing generated module: '+$name)}
 $export=Join-Path $exportFolder ($name+'.bas');$app.SaveAsText(5,$name,$export)
 foreach($m in [regex]::Matches((Get-Content $export -Raw),'(?im)^Public (?:Function|Sub) (\w+)')){$functions[$m.Groups[1].Value]=$true}
}
foreach($pair in @(@('frmEditor','If Not UXEditorUnload() Then Cancel = True'),@('frmUXComparison','If Not UXGridUnload(Me) Then Cancel = True'))){
 $app.DoCmd.OpenForm($pair[0],1);$f=$app.Forms.Item($pair[0]);$code=$f.Module.Lines(1,$f.Module.CountOfLines)
 if($f.OnUnload -ne '[Event Procedure]' -or -not $code.Contains($pair[1]) -or ([regex]::Matches($code,'Private Sub Form_Unload\(')).Count -ne 1){throw ('Native unload guard missing/duplicated: '+$pair[0])}
 $app.DoCmd.Close(2,$pair[0],2)
}
$swiftBuildVerification.NativeUnloadGuards=$true
$retired=@($swiftBuildPlan.Manifest.retiredForms.PSObject.Properties.Name)
$formNames=@($app.CurrentProject.AllForms|ForEach-Object {$_.Name})
foreach($name in $retired){if($formNames -contains $name){throw ('Retired form still installed: '+$name)}}
$swiftBuildVerification.RetiredFormsAbsent=$true
foreach($name in $formNames){
 $app.DoCmd.OpenForm($name,1);$f=$app.Forms.Item($name)
 foreach($obj in @(,$f)+@($f.Controls)){
  foreach($event in @('OnClick','OnDblClick','AfterUpdate','BeforeUpdate','OnChange','OnEnter','OnExit','OnOpen','OnLoad','OnClose','OnUnload','OnActivate','OnResize','OnTimer','OnMouseDown')){
   try{$expression=[string]$obj.Properties.Item($event).Value}catch{continue}
   $match=[regex]::Match($expression,'^=(\w+)\(')
   if($match.Success -and -not $functions.ContainsKey($match.Groups[1].Value) -and $match.Groups[1].Value -notin @('IIf','Nz','Date','Now','Eval')){throw ('Unbound form expression: '+$name+'.'+$obj.Name+'.'+$event+' '+$expression)}
  }
 }
 $app.DoCmd.Close(2,$name,2);$swiftBuildVerification.Forms++
}
Write-Output ('Verified current source modules, event bindings and retired routes across '+$swiftBuildVerification.Forms+' forms.')
