# Expand finite dynamic control-name families without executing installed VBA.
# Source checks and reopened native build checks use the same inventory.
function GetSwiftDynamicViewControls([string]$Root){
 $result=@{}
 function AddControls([string]$view,[object[]]$names){
  if(-not $result.ContainsKey($view)){$result[$view]=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)}
  foreach($name in $names){[void]$result[$view].Add([string]$name)}
 }
 function LiteralKeys([string]$module,[string]$function){
  $text=[IO.File]::ReadAllText((Join-Path $Root ('src/'+$module+'.bas')))
  $body=[regex]::Match($text,'(?ms)^Public Function '+[regex]::Escape($function)+'\b.*?^End Function').Value
  $match=[regex]::Match($body,[regex]::Escape($function)+' = Array\(([^\r\n]*)\)')
  if(-not $match.Success -or ([regex]::Replace($match.Groups[1].Value,'"[^"]*"','')) -notmatch '^[\s,]*$'){throw ('Dynamic view contract needs literal keys: '+$function)}
  @([regex]::Matches($match.Groups[1].Value,'"([^"]*)"')|ForEach-Object {$_.Groups[1].Value})
 }
 foreach($pair in @(@('frmMemberDetails','modMemberDetails','MemberDetailsKeys'),@('frmDecisionDetails','modDecisionDetails','DecisionDetailsKeys'),@('frmPartyNames','modPartyNames','PartyNamesKeys'),@('frmReportingSenior','modReportingSenior','ReportingSeniorKeys'),@('frmReportPeople','modReportPeople','ReportPeopleKeys'),@('frmServiceDetails','modServiceDetails','ServiceDetailsKeys'))){
  AddControls $pair[0] @(LiteralKeys $pair[1] $pair[2]|ForEach-Object {'txt'+$_})
 }
 $context=[IO.File]::ReadAllText((Join-Path $Root 'src/modReportContext.bas'))
 $body=[regex]::Match($context,'(?ms)^Public Function ReportContextPanelKeys\b.*?^End Function').Value
 $panels=@([regex]::Matches($body,'Case "(\w+)": ReportContextPanelKeys = Array\(([^\r\n]*)\)'))
 if($panels.Count -ne 3){throw 'Dynamic context panel inventory changed; update the contract expansion.'}
 foreach($panel in $panels){
  $keys=@([regex]::Matches($panel.Groups[2].Value,'"([^"]+)"')|ForEach-Object {'txt'+$_.Groups[1].Value})
  AddControls ('frmContext'+$panel.Groups[1].Value) $keys
  AddControls 'frmEditor' $keys
 }
 AddControls 'frmUXSeniorEditor' @(@(LiteralKeys 'modRSDirectory' 'RSSeniorKeys')+@(LiteralKeys 'modRSDirectory' 'RSRecordOnlyKeys')|ForEach-Object {'rs'+$_})
 $ruad=[IO.File]::ReadAllText((Join-Path $Root 'src/modRUADProfiles.bas'))
 $keys=[regex]::Match($ruad,'RUADKeys = Split\("([^"]+)"').Groups[1].Value -split '\|'
 if($keys.Count -lt 30){throw 'Dynamic RUAD key inventory unavailable'}
 for($i=0;$i -lt $keys.Count;$i++){if($keys[$i] -ne 'NAME'){AddControls 'frmUXMemberEditor' @(('ruad'+$i),('ruadLabel'+$i))}}
 AddControls 'frmUXMemberEditor' @(0..4|ForEach-Object {'ruadPage'+$_})
 $fields=(Get-Content -LiteralPath (Join-Path $Root 'src/field-definitions.json') -Raw|ConvertFrom-Json).fields
 foreach($field in $fields){
  AddControls 'frmEditor' @($field.control)
  if($field.key -ne 'FormKind'){AddControls 'frmUXEditorFields' @($field.control,('label'+$field.key),('err'+$field.key))}
 }
 AddControls 'frmUXEditorFields' @(0..2|ForEach-Object {'uxGroupBox'+$_})
 AddControls 'frmUXEditorFields' @(33..39|ForEach-Object {'uxChiefGuidance'+$_})
 AddControls 'frmEditor' @(0..9|ForEach-Object {'uxSection'+$_})
 foreach($view in @('frmEditor','frmUXPriorEditor')){AddControls $view @(33..39|ForEach-Object {'lblTrait'+$_;'cboTrait'+$_})}
 AddControls 'frmUXTraitRows' @(33..39|ForEach-Object {'heading'+$_;'Trait'+$_})
 foreach($grade in @(LiteralKeys 'modUXSeniorTable' 'RSTableGrades')){
  AddControls 'frmUXRSCATable' @('grade','avg','count','date','src'|ForEach-Object {$_+$grade})
 }
 $catalogue=[IO.File]::ReadAllText((Join-Path $Root 'scripts/apply-ux-field-catalogue.ps1'))
 $shared=[regex]::Match($catalogue,'\$uxShared=@\(([^\r\n]*)\)').Groups[1].Value
 if(-not $shared){throw 'Dynamic shared-field inventory unavailable'}
 foreach($key in @([regex]::Matches($shared,"'([^']+)'")|ForEach-Object {$_.Groups[1].Value})){
  AddControls 'frmUXBulk' @('uxValue'+$key)
  AddControls 'frmUXSharedFields' @(('value'+$key),('label'+$key))
 }
 $routes=LiteralKeys 'modUX' 'UXWorkspaceRoutes'
 foreach($form in Get-ChildItem -LiteralPath (Join-Path $Root 'src/forms') -Filter '*.form.txt'){
  $design=[IO.File]::ReadAllText($form.FullName);$view=$form.Name -replace '\.form\.txt$',''
  if($design -match 'Name ="uxCompactNav"'){
   AddControls $view @(0..($routes.Count-1)|ForEach-Object {'uxNav'+$_})
  }
  if($design -match 'Name ="uxGroupTab0"'){AddControls $view @(0,1,2,4|ForEach-Object {'uxGroupTab'+$_})}
  foreach($list in [regex]::Matches($design,'(?ms)^                Begin ListBox\r?\n(.*?)(?=^                (?:Begin|End))')){
   $name=[regex]::Match($list.Groups[1].Value,'Name ="(\w+)"').Groups[1].Value
   if($design.Contains('Name ="sort_'+$name+'_0"')){
    $count=[regex]::Match($list.Groups[1].Value,'ColumnCount =(\d+)').Groups[1].Value
    $count=if($count){[int]$count}else{1}
    AddControls $view @(0..($count-1)|ForEach-Object {'sort_'+$name+'_'+$_})
   }
  }
 }
 return $result
}
