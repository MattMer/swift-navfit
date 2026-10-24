if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modUXNAVFIT',(Join-Path $root 'src/modUXNAVFIT.bas'))}
# Embed the already-supported, tracked native blank; export never asks a user to
# supply implementation resources. Small procedures stay below VBA size limits.
$nativeTemplate=Join-Path $root 'FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb'
if((Get-FileHash -LiteralPath $nativeTemplate).Hash -cne 'DD3806E2DAAAE54B3F9B4EC07D59FB0654B46AADECEDCFF54F57DCC75167B2FA'){throw 'Review the changed native blank before embedding it'}
$encoded=[Convert]::ToBase64String([IO.File]::ReadAllBytes($nativeTemplate))
$templateCode=[Collections.Generic.List[string]]::new();$templateCode.Add('Attribute VB_Name = "modNAVFITTemplate"');$templateCode.Add('Option Explicit')
$templateCode.Add('Public Function NAVFITTemplateHash() As String');$templateCode.Add(' NAVFITTemplateHash = "'+(Get-FileHash -LiteralPath $nativeTemplate).Hash+'"');$templateCode.Add('End Function')
$chunks=[int][math]::Ceiling($encoded.Length/16384)
$templateCode.Add('Public Function NAVFITTemplateData() As String');$templateCode.Add(' Dim result As String')
for($i=0;$i -lt $chunks;$i++){$templateCode.Add(' result = result & NativeTemplateChunk'+$i+'()')}
$templateCode.Add(' NAVFITTemplateData = result');$templateCode.Add('End Function')
for($i=0;$i -lt $chunks;$i++){
 $part=$encoded.Substring($i*16384,[math]::Min(16384,$encoded.Length-$i*16384))
 $templateCode.Add('Private Function NativeTemplateChunk'+$i+'() As String');$templateCode.Add(' Dim result As String')
 for($offset=0;$offset -lt $part.Length;$offset+=512){$templateCode.Add(' result = result & "'+$part.Substring($offset,[math]::Min(512,$part.Length-$offset))+'"')}
 $templateCode.Add(' NativeTemplateChunk'+$i+' = result');$templateCode.Add('End Function')
}
$templateModule=Join-Path $root 'build/modNAVFITTemplate.bas';[IO.File]::WriteAllLines($templateModule,$templateCode,[Text.UTF8Encoding]::new($false));$app.LoadFromText(5,'modNAVFITTemplate',$templateModule)
$app.DoCmd.OpenForm('frmImportReview',1);$f=$app.Forms.Item('frmImportReview')
foreach($existing in @('lblNAVFITCentury','cboNAVFITCentury')){
 if(@($f.Controls|ForEach-Object{$_.Name}) -contains $existing){$app.DeleteControl($f.Name,$existing)}
}
$f.Controls('lblScope').Caption='Choose > Preview > Apply > Results. Map all supported Swift report values and retain every native source column. Imported reports become separate Swift drafts.'
$f.Controls('cmdPreview').Caption='Preview full report mapping';$f.Controls('cmdPreview').OnClick='=UXPreviewFullNAVFIT()'
$f.Controls('cmdImport').OnClick='=UXApplyFullNAVFIT()'
$f.Controls('cmdOpenResult').OnClick='=UXOpenFullNAVFITResult()'
[void](Control $f 100 'lblNAVFITCentury' 400 7000 4700 500 'For two-digit counseling years, choose the century explicitly')
[void](Control $f 111 'cboNAVFITCentury' 400 7550 4700 400);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='1900;2000';$c.LimitToList=$true;$c.AfterUpdate='=UXNarrativeInputChanged()'
SaveScreen $f 'frmImportReview'
NewScreen 'frmNAVFITExchange' 'NAVFIT export and returned reports' 12200;$f=$script:swiftForm;$f.ScrollBars=2
[void](Control $f 100 'lblScope' 400 950 12600 600 'Export saved reports or all saved group members, including incomplete drafts. Findings do not block export; export is not approval, signature or submission. Keep Swift baselines locally. Returns support single-report exports.')
[void](Control $f 100 'lblReport' 400 1650 2500 350 'Saved report')
[void](Control $f 111 'cboReport' 3000 1650 10000 400);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;9000';$c.LimitToList=$true
foreach($item in @(@('txtTemplate','Verified blank NAVFIT template',2150),@('txtExportPath','New export file',2700),@('txtReturned','Returned NAVFIT file',3900),@('txtBaseline','Original Swift baseline',4450))){
 [void](Control $f 100 ('lbl'+$item[0]) 400 $item[2] 3200 350 $item[1])
 [void](Control $f 109 $item[0] 3800 $item[2] 7300 400);$script:swiftControl.AfterUpdate='=UXNAVFITInputChanged()'
 [void](Control $f 104 ('cmd'+$item[0]) 11400 $item[2] 1600 400 'Choose...' ('=UXChooseNAVFITPath("'+$item[0]+'",'+$(if($item[0] -eq 'txtExportPath'){'True'}else{'False'})+')'))
}
[void](Control $f 104 'cmdExport' 3800 3250 9200 450 'Export saved report to NAVFIT' '=UXExportNAVFIT()')
[void](Control $f 100 'lblReturnCentury' 400 5000 2500 450 'Counseling century (if new)')
[void](Control $f 111 'cboReturnCentury' 2900 5000 1600 450);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='1900;2000';$c.LimitToList=$true;$c.AfterUpdate='=UXNAVFITInputChanged()'
[void](Control $f 104 'cmdPreview' 4800 5000 8200 450 'Preview returned file against export baseline' '=UXPreviewNAVFITReturn()')
[void](Control $f 110 'lstChoices' 400 5650 5500 3500);$c=$script:swiftControl;$c.ColumnCount=3;$c.ColumnWidths='0;3500;1800';$c.AfterUpdate='=UXNAVFITReturnDetail()'
[void](Control $f 109 'txtDetail' 6200 5650 6800 3500);$c=$script:swiftControl;$c.Locked=$true;$c.ScrollBars=2
[void](Control $f 104 'cmdKeep' 6200 9350 3100 450 'Keep current Swift value' '=UXResolveNAVFITReturn("Keep current")')
[void](Control $f 104 'cmdAccept' 9600 9350 3400 450 'Accept returned value' '=UXResolveNAVFITReturn("Accept returned")')
[void](Control $f 109 'txtStatus' 400 10000 12600 900);$c=$script:swiftControl;$c.Locked=$true;$c.ScrollBars=2
[void](Control $f 104 'cmdApply' 6500 11100 6500 450 'Apply resolved return' '=UXApplyNAVFITReturn()');$script:swiftControl.Enabled=$false
[void](Control $f 104 'cmdBack' 400 11100 4500 450 'Back to Reports' '=UXNavigate("frmLibrary")')
$f.OnUnload='=NAVFITClearReturned(False)'
foreach($c in $f.Controls){if($c.Top -ge 3900){$c.Top=[int]($c.Top+1800)}}
$f.Section(0).Height=14000
[void](Control $f 100 'lblGroup' 400 3800 2500 350 'Summary group')
[void](Control $f 111 'cboGroup' 3000 3800 10000 400);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;9000';$c.LimitToList=$true;$c.RowSource="SELECT GroupID,Title & ' | ' & FormKind & ' | ' & Paygrade & IIf(IsArchived,' | Archived','') FROM SummaryGroup ORDER BY Title";$c.AfterUpdate='=UXNAVFITInputChanged()'
[void](Control $f 106 'chkPSG' 400 4350 3700 400 'Write exact PSG RSCA');$script:swiftControl.AfterUpdate='=UXNAVFITInputChanged()';$script:swiftControl.DefaultValue='False'
[void](Control $f 104 'cmdPreviewGroup' 4300 4350 8700 450 'Preview saved group members and counts' '=UXPreviewNAVFITGroup()')
[void](Control $f 104 'cmdExportGroup' 4300 4900 8700 450 'Export previewed summary group to NAVFIT' '=UXExportNAVFITGroup()');$script:swiftControl.Enabled=$false
foreach($existingControl in $f.Controls){if($existingControl.Top -ge 5700){$existingControl.Top=[int]($existingControl.Top+550)}}
$f.Section(0).Height=14550
[void](Control $f 104 'cmdResumeReturn' 400 5500 6000 400 'Resume saved return preview' '=UXResumeNAVFITReturn()')
[void](Control $f 104 'cmdCancelReturn' 6700 5500 6300 400 'Cancel saved return preview' '=UXCancelNAVFITReturn()')
SaveScreen $f 'frmNAVFITExchange'
$app.DoCmd.OpenForm('frmMajorWorkbench',1);$f=$app.Forms.Item('frmMajorWorkbench')
if(@($f.Controls|ForEach-Object{$_.Name}) -contains 'uxNAVFITExport'){$app.DeleteControl($f.Name,'uxNAVFITExport')}
[void](Control $f 104 'uxNAVFITExport' 2850 8050 5000 450 'Export group to NAVFIT98A' '=UXOpenNAVFITGroupExport()')
SaveScreen $f 'frmMajorWorkbench'
$app.DoCmd.OpenForm('frmUXImportChoose',1);$f=$app.Forms.Item('frmUXImportChoose')
if(@($f.Controls|ForEach-Object{$_.Name}) -contains 'cmdNAVFITExchange'){$app.DeleteControl($f.Name,'cmdNAVFITExchange')}
$f.Section(0).Height=6200
$f.Controls('cmdBack').Top=5400
SaveScreen $f 'frmUXImportChoose'
