# Executed in build.ps1's Access-generation scope. Reuses the unbound field
# controls and native services; this file does not touch an existing workspace.
$fieldDefinitions=Get-Content (Join-Path $root 'src/field-definitions.json') -Raw -Encoding UTF8|ConvertFrom-Json
if($fieldDefinitions.catalogueVersion -lt 1){throw 'Unsupported field definition catalogue'}
$uxFields=@($fieldDefinitions.fields|Sort-Object ordinal)
foreach($field in $uxFields){$field.label=$field.label -replace '[\u2013\u2014]','-'}
$uxNames=@($uxFields|ForEach-Object {$_.control})
$uxCatalogue=Get-Content (Join-Path $root 'src/trait-catalogue.json') -Raw|ConvertFrom-Json
$uxShared=@('CommandEmployment','UnitUIC','ShipStation','PromotionStatus','PeriodStart','PeriodEnd','DateReported','OccasionPeriodic','OccasionDetachment','OccasionFrocking','OccasionSpecial','ReportRegular','ReportConcurrent','ReportOperationalCommander','ObservationStatus','ComponentActive','ComponentTAR','ComponentInactive','ComponentATADSW','BilletSubcategory','ReadinessSummary','ReadinessDetail','ReadinessDate')
$uxSections=@('Admin data','Admin data','Command & duties','Counseling','Performance traits','Recommendations & comments','Final report details','Draft notes','Review','History')
$uxModule=[Collections.Generic.List[string]]::new()
$uxModule.Add('Attribute VB_Name = "modUXFields"');$uxModule.Add('Option Explicit')
$uxModule.Add('Public Const FIELD_DEFINITION_VERSION As Long = '+$fieldDefinitions.catalogueVersion)
$uxModule.Add('Public Function CatalogueFieldKeys() As Variant')
$uxModule.Add('Dim keys As String')
for($i=0;$i -lt $uxFields.Count;$i+=12){$uxModule.Add('keys = keys & "'+(($uxFields[$i..([Math]::Min($i+11,$uxFields.Count-1))].key)-join '|')+$(if($i+12 -lt $uxFields.Count){'|'})+'"')}
$uxModule.Add('CatalogueFieldKeys = Split(keys, "|")')
$uxModule.Add('End Function')
$uxModule.Add('Public Function CatalogueFieldLimits(ByVal keys As Variant) As Variant')
$uxModule.Add('Dim limits() As Variant, i As Long: ReDim limits(LBound(keys) To UBound(keys))')
$uxModule.Add('For i = LBound(keys) To UBound(keys): limits(i) = UXFieldLimit(CStr(keys(i))): Next')
$uxModule.Add('CatalogueFieldLimits = limits: End Function')
foreach($property in @(@('CatalogueControls','control'),@('CatalogueRecoveryKeys','recoveryKey'))){
    $uxModule.Add('Public Function '+$property[0]+'() As Variant')
    $uxModule.Add('Dim keys As String')
    for($i=0;$i -lt $uxFields.Count;$i+=12){$uxModule.Add('keys = keys & "'+(($uxFields[$i..([Math]::Min($i+11,$uxFields.Count-1))].($property[1]))-join '|')+$(if($i+12 -lt $uxFields.Count){'|'})+'"')}
    $uxModule.Add($property[0]+' = Split(keys, "|"): End Function')
}
$uxModule.Add('Public Function CatalogueStorageType(ByVal key As String) As String')
$uxModule.Add('Select Case key')
foreach($field in $uxFields){$uxModule.Add('Case "'+$field.key+'": CatalogueStorageType = "'+$field.storageType+'"')}
$uxModule.Add('End Select: End Function')
$uxModule.Add('Public Function CatalogueFieldVersion(ByVal key As String) As Long')
$uxModule.Add('Select Case key')
foreach($field in $uxFields){$uxModule.Add('Case "'+$field.key+'": CatalogueFieldVersion = '+$field.definitionVersion)}
$uxModule.Add('End Select: End Function')
foreach($adapter in @('navfit','pdf')){
    $fn='Catalogue'+$adapter+'Rows'
    $uxModule.Add('Public Function '+$fn+'(ByVal kind As String) As Collection')
    $uxModule.Add('Dim result As New Collection: Select Case kind')
    foreach($kind in @('EVAL','CHIEFEVAL','FITREP')){
        $uxModule.Add('Case "'+$kind+'"')
        $bindingRows=@(foreach($field in $uxFields){foreach($binding in $field.bindings.$kind.$adapter){[pscustomobject]@{key=$field.key;ordinal=$binding.ordinal;source=$binding.source;conversion=$binding.conversion}}})
        foreach($binding in $bindingRows|Sort-Object ordinal){$uxModule.Add('result.Add "'+$binding.key+'|'+$binding.source+'|'+$binding.conversion+'"')}
    }
    $uxModule.Add('End Select: Set '+$fn+' = result: End Function')
}
$uxModule.Add('Public Function CataloguePage(ByVal kind As String, ByVal key As String) As Long')
$uxModule.Add('Select Case kind & ":" & key')
foreach($kind in @('EVAL','CHIEFEVAL','FITREP')){foreach($field in $uxFields){if($null -ne $field.bindings.$kind.page){$uxModule.Add('Case "'+$kind+':'+$field.key+'": CataloguePage = '+$field.bindings.$kind.page.page)}}}
$uxModule.Add('Case Else: CataloguePage = 0: End Select: End Function')
$uxModule.Add('Public Function UXFieldOnForm(ByVal kind As String, ByVal key As String) As Boolean');$uxModule.Add('Select Case kind & ":" & key')
foreach($kind in @('EVAL','CHIEFEVAL','FITREP')){$onForm=@($uxFields|Where-Object {$null -ne $_.bindings.$kind.page}|ForEach-Object {'"'+$kind+':'+$_.key+'"'});for($i=0;$i -lt $onForm.Count;$i+=20){$uxModule.Add('Case '+(($onForm[$i..([Math]::Min($i+19,$onForm.Count-1))])-join ', ')+': UXFieldOnForm = True')}}
$uxModule.Add('End Select: End Function')
$uxModule.Add('Public Function CataloguePageRows(ByVal kind As String, ByVal page As Long) As Collection')
$uxModule.Add('Dim result As New Collection: Select Case kind & ":" & page')
foreach($layout in $fieldDefinitions.fixedLayout.PSObject.Properties){
    $uxModule.Add('Case "'+$layout.Name+'"')
    foreach($field in $layout.Value.fields){$uxModule.Add('result.Add "'+$field.key+'|'+$field.x+'|'+$field.y+'|'+$field.w+'|'+$field.h+'"')}
}
$uxModule.Add('End Select: Set CataloguePageRows = result: End Function')
$uxModule.Add('Public Function UXSectionName(ByVal section As Long) As String')
$uxModule.Add('Dim names As Variant: names = Array('+ (($uxSections|ForEach-Object {'"'+$_.Replace('"','""')+'"'})-join ',')+')')
$uxModule.Add('UXSectionName = names(section)');$uxModule.Add('End Function')
$uxModule.Add('Public Function UXSharedKeys() As Variant');$uxModule.Add('UXSharedKeys = Array('+ (($uxShared|ForEach-Object {'"'+$_+'"'})-join ',')+')');$uxModule.Add('End Function')
$uxModule.Add('Public Function UXFieldLabel(ByVal key As String) As String');$uxModule.Add('Select Case key')
foreach($field in $uxFields){$uxModule.Add('Case "'+$field.key+'": UXFieldLabel = "'+$field.label.Replace('"','""')+'"')}
$uxModule.Add('Case Else: UXFieldLabel = key');$uxModule.Add('End Select');$uxModule.Add('End Function')
$uxModule.Add('Public Function UXFieldIsCheckbox(ByVal key As String) As Boolean');$uxModule.Add('Select Case key');$uxModule.Add('Case '+(($uxFields|Where-Object {$_.checkbox}|ForEach-Object {'"'+$_.key+'"'})-join ', ')+': UXFieldIsCheckbox = True');$uxModule.Add('End Select');$uxModule.Add('End Function')
$uxModulePath=Join-Path $root 'build/modUXFields.bas';[IO.File]::WriteAllText($uxModulePath,($uxModule -join "`r`n")+"`r`n",[Text.UTF8Encoding]::new($false))
$app.LoadFromText(5,'modUXFields',$uxModulePath);$app.LoadFromText(5,'modUX',(Join-Path $root 'src/modUX.bas'))
