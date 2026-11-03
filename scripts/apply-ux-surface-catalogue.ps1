$surfaceModule=[Collections.Generic.List[string]]::new()
$surfaceModule.Add('Attribute VB_Name = "modUXSurfaceFields"');$surfaceModule.Add('Option Explicit')
function SurfaceProperty($functionName,$type,$property,$default){
    $surfaceModule.Add('Public Function '+$functionName+'(ByVal key As String) As '+$type)
    $surfaceModule.Add('Select Case key')
    for($i=0;$i -lt $uxFields.Count;$i++){
        $field=$uxFields[$i]
        $value=if($property -eq 'control'){'"'+$uxNames[$i]+'"'}elseif($type -eq 'String'){'"'+([string]$field.$property).Replace('"','""')+'"'}elseif($type -eq 'Boolean'){if($field.$property){'True'}else{'False'}}else{if($null -eq $field.$property){'0'}else{[string]$field.$property}}
        if($property -eq 'options'){$value='"'+(($field.options -join ';').Replace('"','""'))+'"'}
        $surfaceModule.Add('Case "'+$field.key+'": '+$functionName+' = '+$value)
    }
    $surfaceModule.Add('Case Else: '+$functionName+' = '+$default);$surfaceModule.Add('End Select');$surfaceModule.Add('End Function')
}
SurfaceProperty 'UXFieldControl' 'String' 'control' '""'
SurfaceProperty 'UXFieldSection' 'Long' 'section' '7'
SurfaceProperty 'UXFieldLimit' 'Long' 'limit' '0'
SurfaceProperty 'UXFieldDigits' 'Long' 'digits' '0'
SurfaceProperty 'UXFieldDate' 'Boolean' 'date' 'False'
SurfaceProperty 'UXFieldFlag' 'Boolean' 'checkbox' 'False'
SurfaceProperty 'UXFieldMultiline' 'Boolean' 'multiline' 'False'
SurfaceProperty 'UXFieldLegacy' 'Boolean' 'legacyHidden' 'False'
SurfaceProperty 'UXFieldOptions' 'String' 'options' '""'
$surfaceModule.Add('Public Function UXFieldApplies(ByVal key As String, ByVal kind As String) As Boolean')
$surfaceModule.Add('UXFieldApplies = True: Select Case key')
foreach($field in $uxFields|Where-Object forms){$surfaceModule.Add('Case "'+$field.key+'": UXFieldApplies = ('+(($field.forms|ForEach-Object{'kind = "'+$_+'"'})-join ' Or ')+')')}
$surfaceModule.Add('End Select: End Function')
$surfaceModule.Add('Public Function UXFieldCaption(ByVal key As String, ByVal kind As String) As String')
$surfaceModule.Add('Dim label As String, block As String: label = UXFieldLabel(key): Select Case key')
foreach($field in $uxFields){
    $surfaceModule.Add('Case "'+$field.key+'": block = "'+$field.block+'"')
    if($field.blocks){foreach($kind in @('EVAL','CHIEFEVAL','FITREP')){$surfaceModule.Add('If kind = "'+$kind+'" Then block = "'+$field.blocks.$kind+'"')}}
}
$surfaceModule.Add('End Select')
$surfaceModule.Add('If key = "Designator" Then block = "3"')
$surfaceModule.Add('If key = "RankRate" And kind = "CHIEFEVAL" Then label = "Rating"')
#NAVPERS 1616/26 block 12 is Promotion/Frocking; the Chief and officer forms use it for detachment of the reporting senior.
$surfaceModule.Add('If key = "OccasionFrocking" And kind = "EVAL" Then label = "Promotion / Frocking"')
$surfaceModule.Add('If IsTraitKey(key) Then label = Replace(Replace(CatalogueTraitTitle(kind, CLng(Mid$(key, 6))), vbCr, ""), vbLf, " ")')
$surfaceModule.Add('If Len(block) > 0 Then label = block & ". " & label')
$surfaceModule.Add('UXFieldCaption = label: End Function')
$surfaceModule.Add('Public Function UXSectionKeys(ByVal section As Long, ByVal kind As String) As Variant')
$surfaceModule.Add('Select Case kind & ":" & section')
foreach($kind in @('EVAL','CHIEFEVAL','FITREP')){for($section=0;$section -le 7;$section++){
    $ordered=@($uxFields|Where-Object {$_.section -eq $section}|Sort-Object @{Expression={ $block=if($_.blocks){$_.blocks.$kind}else{$_.block};if($_.key -eq 'Designator'){$block='3'};if($block -match '^\d+'){[int]$Matches[0]}else{999}}})
    $surfaceModule.Add('Case "'+$kind+':'+$section+'": UXSectionKeys = Array('+(($ordered|ForEach-Object{'"'+$_.key+'"'})-join ',')+')')
}}
$surfaceModule.Add('Case Else: UXSectionKeys = Array(): End Select: End Function')
$surfaceModulePath=Join-Path $root 'build/modUXSurfaceFields.bas'
[IO.File]::WriteAllText($surfaceModulePath,($surfaceModule -join "`r`n"),[Text.UTF8Encoding]::new($false))
$app.LoadFromText(5,'modUXSurfaceFields',$surfaceModulePath)
if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modUXSurface',(Join-Path $root 'src/modUXSurface.bas'))}
if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modUXDashboard',(Join-Path $root 'src/modUXDashboard.bas'))}
