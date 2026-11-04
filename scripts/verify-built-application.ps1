# Verify a reopened database, never just in-memory design objects.
. (Join-Path $PSScriptRoot 'vba-source-normalization.ps1')
$swiftBuildVerification=[ordered]@{SourceModules=0;Forms=0;NativeUnloadGuards=$false;RetiredFormsAbsent=$false;Identity=$false}
$checkDB=$app.CurrentDb();$meta=$checkDB.OpenRecordset('SELECT AppVersion,SchemaVersion FROM AppMetadata',4)
if($meta.Fields.Item('AppVersion').Value -ne $swiftBuildPlan.AppVersion -or $meta.Fields.Item('SchemaVersion').Value -ne $swiftBuildPlan.SchemaVersion){throw 'Built application identity mismatch'}
$meta.Close();$swiftBuildVerification.Identity=$true
$majorSchema=Get-Content (Join-Path $root 'src/major-schema.json') -Raw|ConvertFrom-Json
$relationshipChecks=0
foreach($sql in $majorSchema.SQL){
 if($sql -notmatch '^ALTER TABLE (\w+) ADD CONSTRAINT (FK_\w+) FOREIGN KEY \((\w+)\) REFERENCES (\w+) \((\w+)\)$'){continue}
 $child=$Matches[1];$name=$Matches[2];$childField=$Matches[3];$parent=$Matches[4];$parentField=$Matches[5]
 $relation=$checkDB.Relations.Item($name)
 if($relation.Table -ne $parent -or $relation.ForeignTable -ne $child -or $relation.Fields.Count -ne 1 -or $relation.Fields.Item(0).Name -ne $parentField -or $relation.Fields.Item(0).ForeignName -ne $childField -or ($relation.Attributes -band 2)){
  throw ('Relationship contract differs: '+$name)
 }
 $relationshipChecks++
}
if(@($checkDB.Relations|Where-Object {$_.Name -match '^FK_Major\d+$'}).Count){throw 'Opaque major relationship names remain installed'}
foreach($definition in $majorSchema.Queries){
 $query=$checkDB.QueryDefs.Item($definition.Name)
 $actual=[regex]::Replace($query.SQL,'[\[\];\s]','')
 $expected=[regex]::Replace($definition.SQL,'[\[\];\s]','')
 if($actual -ine $expected){throw ('Revision-ID alias definition differs: '+$definition.Name)}
 $query.Close()
}
$swiftBuildVerification.Relationships=$relationshipChecks
$swiftBuildVerification.RevisionAliases=@($majorSchema.Queries).Count
$checkDB.Close()
$exportFolder=Join-Path $root ('build/source-check-'+[guid]::NewGuid());[void](New-Item -ItemType Directory $exportFolder)
$functions=@{};$viewReferences=@{};$installedCode=''
foreach($source in $swiftBuildPlan.Manifest.sourceModules){
 $name=[IO.Path]::GetFileNameWithoutExtension($source);$export=Join-Path $exportFolder ($name+'.bas')
 $app.SaveAsText(5,$name,$export)
 $expected=Get-Content (Join-Path $root $source) -Raw;$actual=Get-Content $export -Raw
 $installedCode+=$actual+"`n"
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
# Read literal contract data from the already verified VBA, never execute it.
# Fresh-file trust settings may intentionally prevent Eval after a native reopen.
$viewContracts=@{};$constantNames=@{}
foreach($match in [regex]::Matches($installedCode,'(?im)^Public Const (\w+) As String = "([^"]*)"')){$constantNames[$match.Groups[1].Value]=$match.Groups[2].Value}
$contractCode=[regex]::Match($installedCode,'(?ms)^Public Function UXViewContracts\b.*?^End Function').Value
foreach($match in [regex]::Matches($contractCode,'(?im)^\s*Case ("[^"]+"|\w+): UXViewContracts = Array\(([^\r\n]*)\)')){
 $key=$match.Groups[1].Value
 if($key -eq 'Else'){continue}
 if(([regex]::Replace($match.Groups[2].Value,'"(?:[^"]|"")*"','')) -notmatch '^[\s,]*$'){throw 'View contracts must contain literal string data'}
 if($key.StartsWith('"')){$key=$key.Trim('"')}elseif($constantNames.ContainsKey($key)){$key=$constantNames[$key]}else{throw ('Unknown view contract identifier: '+$key)}
 $viewContracts[$key]=@([regex]::Matches($match.Groups[2].Value,'"(?:[^"]|"")*"')|ForEach-Object {$_.Value.Substring(1,$_.Value.Length-2).Replace('""','"')})
}
if($viewContracts.Count -eq 0 -or $viewContracts.Count -ne [regex]::Matches($contractCode,'(?im)^\s*Case (?!Else\b)').Count){throw 'Required view contracts could not be read completely'}
foreach($name in $viewContracts.Keys){if($formNames -notcontains $name){throw ('Required view missing: '+$name)}}
$canonicalControls=@((Get-Content (Join-Path $root 'src/field-definitions.json') -Raw|ConvertFrom-Json).fields|ForEach-Object {$_.control})
. (Join-Path $PSScriptRoot 'view-control-contracts.ps1')
$dynamicControls=GetSwiftDynamicViewControls $root
foreach($view in $dynamicControls.Keys){if($formNames -notcontains $view){throw ('Required dynamic view missing: '+$view)}}
foreach($name in $formNames){
 $app.DoCmd.OpenForm($name,1);$f=$app.Forms.Item($name)
 $controlNames=@($f.Controls|ForEach-Object {$_.Name})
 $required=@();if($viewContracts.ContainsKey($name)){$required=@($viewContracts[$name])}
 if($dynamicControls.ContainsKey($name)){$required+=@($dynamicControls[$name])}
 if($name -eq 'frmEditor'){$required+=@($canonicalControls)}
 if($name -eq 'frmUXTraitRows'){foreach($block in 33..39){$required+=('Trait'+$block+'|AfterUpdate|=UXGridQueue("Trait'+$block+'")')}}
 foreach($entry in $required){
  $parts=$entry -split '\|'
  if($controlNames -notcontains $parts[0]){throw ('Required view control missing: '+$name+'.'+$parts[0])}
  if($parts.Count -eq 3){
   $actual=[string]$f.Controls.Item($parts[0]).Properties.Item($parts[1]).Value;$expected=$parts[2]
   if($parts[1] -eq 'SourceObject'){$actual=$actual -replace '^Form\.','';$expected=$expected -replace '^Form\.',''}
   if($actual -ine $expected){throw ('Required view binding differs: '+$name+'.'+$parts[0]+'.'+$parts[1])}
  }
 }
 foreach($obj in @(,$f)+@($f.Controls)){
  foreach($event in @('OnClick','OnDblClick','AfterUpdate','BeforeUpdate','OnChange','OnEnter','OnExit','OnGotFocus','OnLostFocus','OnOpen','OnLoad','OnClose','OnUnload','OnActivate','OnResize','OnTimer','OnMouseDown')){
   try{$expression=[string]$obj.Properties.Item($event).Value}catch{continue}
   $match=[regex]::Match($expression,'^=(\w+)\(')
   if($match.Success -and -not $functions.ContainsKey($match.Groups[1].Value) -and $match.Groups[1].Value -notin @('IIf','Nz','Date','Now','Eval')){throw ('Unbound form expression: '+$name+'.'+$obj.Name+'.'+$event+' '+$expression)}
  }
 }
 $app.DoCmd.Close(2,$name,2);$swiftBuildVerification.Forms++
}
# Check literal cross-object references even when their callbacks were renamed or
# removed together. Dynamic names are covered by the explicit runtime contracts.
foreach($match in [regex]::Matches($installedCode,'Forms!(\w+)!(\w+)')){
 $viewReferences[$match.Groups[1].Value+'.'+$match.Groups[2].Value]=@($match.Groups[1].Value,$match.Groups[2].Value)
}
foreach($view in @($viewReferences.Values|ForEach-Object {$_[0]}|Sort-Object -Unique)){
 if($formNames -notcontains $view){throw ('Referenced form missing: '+$view)}
 $app.DoCmd.OpenForm($view,1);$f=$app.Forms.Item($view)
 foreach($pair in @($viewReferences.Values|Where-Object {$_[0] -eq $view})){
  if(@($f.Controls|ForEach-Object {$_.Name}) -notcontains $pair[1]){throw ('Required literal form/control reference missing: '+$view+'.'+$pair[1])}
 }
 $app.DoCmd.Close(2,$view,2)
}
foreach($macro in @($app.CurrentProject.AllMacros)){
 $export=Join-Path $exportFolder ($macro.Name+'.macro.txt');$app.SaveAsText(4,$macro.Name,$export)
 foreach($match in [regex]::Matches((Get-Content $export -Raw),'(?im)FunctionName\s*=\s*"=?([A-Za-z_]\w*)\(')){
  if(-not $functions.ContainsKey($match.Groups[1].Value)){throw ('Unbound macro callback: '+$macro.Name+'.'+$match.Groups[1].Value)}
 }
}
$reportNames=@($app.CurrentProject.AllReports|ForEach-Object {$_.Name})
if($reportNames -notcontains 'rptPDFOutput'){throw 'Required print report missing: rptPDFOutput'}
foreach($report in @($app.CurrentProject.AllReports)){
 $app.DoCmd.OpenReport($report.Name,1);$rpt=$app.Reports.Item($report.Name)
 foreach($obj in @(,$rpt)+@($rpt.Controls)){
  foreach($event in @('OnOpen','OnClose','OnClick','OnFormat','OnPrint','OnRetreat')){
   try{$expression=[string]$obj.Properties.Item($event).Value}catch{continue}
   $match=[regex]::Match($expression,'^=(\w+)\(')
   if($match.Success -and -not $functions.ContainsKey($match.Groups[1].Value)){throw ('Unbound report expression: '+$report.Name+'.'+$obj.Name+'.'+$event)}
  }
 }
 if($report.Name -eq 'rptPDFOutput'){
  $slots=[int]([regex]::Match((Get-Content (Join-Path $root 'src/modPDFOutput.bas') -Raw),'PDF_OUTPUT_SLOTS As Long = (\d+)').Groups[1].Value)
  $reportControlNames=@($rpt.Controls|ForEach-Object {$_.Name})
  foreach($controlName in @('txtPageNo')+@(1..$slots|ForEach-Object {'v'+$_.ToString('000')})){
   if($reportControlNames -notcontains $controlName){throw ('Required print control missing: '+$controlName)}
  }
 }
 $app.DoCmd.Close(3,$report.Name,2)
}
Write-Output ('Verified current source modules, event bindings and retired routes across '+$swiftBuildVerification.Forms+' forms.')
