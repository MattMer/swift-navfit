# Executed in build.ps1's Access-generation scope. Reuses the unbound field
# controls and native services; this file does not touch an existing workspace.
$uxData=[IO.File]::ReadAllText((Join-Path $root 'prototype/fields.mjs'))
$uxFields=ConvertFrom-Json -InputObject (($uxData -replace '^export const fields = ','').Trim().TrimEnd(';'))
foreach($field in $uxFields){$field.label=$field.label -replace '[\u2013\u2014]','-'}
$uxRecovery=[IO.File]::ReadAllText((Join-Path $root 'src/modRecovery.bas'))
$uxNames=@([regex]::Matches(([regex]::Match($uxRecovery,'(?s)ControlsList = Array\((.*?)End Function').Groups[1].Value),'"([A-Za-z0-9]+)"')|ForEach-Object{$_.Groups[1].Value})
if($uxNames.Count -ne 79 -or $uxFields.Count -ne 79){throw 'UX field/control coverage mismatch.'}
$uxCatalogue=Get-Content (Join-Path $root 'src/trait-catalogue.json') -Raw|ConvertFrom-Json
$uxShared=@('CommandEmployment','UnitUIC','ShipStation','PromotionStatus','PeriodStart','PeriodEnd','DateReported','OccasionPeriodic','OccasionDetachment','OccasionFrocking','OccasionSpecial','ReportRegular','ReportConcurrent','ReportOperationalCommander','ObservationStatus','ComponentActive','ComponentTAR','ComponentInactive','ComponentATADSW','BilletSubcategory','ReadinessSummary','ReadinessDetail','ReadinessDate','ReportingSeniorName','ReportingSeniorGrade','ReportingSeniorDesignator','ReportingSeniorTitle','ReportingSeniorUIC','ReportingSeniorSSN','ReportingSeniorAddress1','ReportingSeniorAddress2','ReportingSeniorCity','ReportingSeniorState','ReportingSeniorZip','ReportingSeniorPhone','ReportingSeniorDSN','PrimaryDuty','DutyDescription','ReportingSeniorLastName','ReportingSeniorFirstInitial','ReportingSeniorMiddleInitial')
$uxSections=@('Admin data','Admin data','Command & duties','Counseling','Performance traits','Recommendations & comments','Final report details','Draft notes','Review','History')
$uxModule=[Collections.Generic.List[string]]::new()
$uxModule.Add('Attribute VB_Name = "modUXFields"');$uxModule.Add('Option Explicit')
$uxModule.Add('Public Function UXSectionName(ByVal section As Long) As String')
$uxModule.Add('Dim names As Variant: names = Array('+ (($uxSections|ForEach-Object {'"'+$_.Replace('"','""')+'"'})-join ',')+')')
$uxModule.Add('UXSectionName = names(section)');$uxModule.Add('End Function')
$uxModule.Add('Public Function UXSharedKeys() As Variant');$uxModule.Add('UXSharedKeys = Array('+ (($uxShared|ForEach-Object {'"'+$_+'"'})-join ',')+')');$uxModule.Add('End Function')
$uxModule.Add('Public Function UXFieldLabel(ByVal key As String) As String');$uxModule.Add('Select Case key')
foreach($field in $uxFields){$uxModule.Add('Case "'+$field.key+'": UXFieldLabel = "'+$field.label.Replace('"','""')+'"')}
$uxModule.Add('Case Else: UXFieldLabel = key');$uxModule.Add('End Select');$uxModule.Add('End Function')
$uxModulePath=Join-Path $root 'build/modUXFields.bas';[IO.File]::WriteAllText($uxModulePath,($uxModule -join "`r`n")+"`r`n",[Text.UTF8Encoding]::new($false))
$app.LoadFromText(5,'modUXFields',$uxModulePath);$app.LoadFromText(5,'modUX',(Join-Path $root 'src/modUX.bas'))
