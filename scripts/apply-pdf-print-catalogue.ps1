# Build-time conversion of reviewed print metadata; no runtime JSON/PDF dependency.
$layoutPath=Join-Path $root 'src/pdf-print-layout.json'
$layout=Get-Content -LiteralPath $layoutPath -Raw -Encoding UTF8|ConvertFrom-Json
if($layout.version -ne 1 -or $layout.templates.Count -ne 4){throw 'Unsupported print catalogue'}
function PDFVB([string]$s){'"'+$s.Replace('"','""')+'"'}
$lines=[Collections.Generic.List[string]]::new()
foreach($s in @('Attribute VB_Name = "modPDFPrintCatalogue"','Option Explicit','Public Function PDFPrintCatalogueHash() As String',(' PDFPrintCatalogueHash = "'+(Get-FileHash $layoutPath).Hash+'"'),'End Function','Public Function PDFPrintRows(ByVal kind As String, ByVal page As Long) As Collection',' Dim rows As New Collection',' Select Case kind & "_" & page')){$lines.Add($s)}
$backgrounds=@()
foreach($template in $layout.templates){
 foreach($page in $template.pages){
  $key=$template.kind+'_'+$page.page;$backgrounds+=$key
  $image=Join-Path $root ('src/pdf-templates/'+$page.image)
  if((Get-FileHash -LiteralPath $image).Hash -ne $page.imageSHA256){throw ('Print background differs: '+$image)}
  $lines.Add(' Case '+(PDFVB $key))
  foreach($field in $page.fields){
   if($field.box.Count -ne 4 -or $field.box[0] -lt 0 -or $field.box[1] -lt 0 -or $field.box[2] -le 0 -or $field.box[3] -le 0 -or ($field.box[0]+$field.box[2]) -gt 612 -or ($field.box[1]+$field.box[3]) -gt 792){throw ('Print box outside Letter page: '+$key+'/'+$field.source)}
   $data=@($field.target,$field.rule,$field.selector)+@($field.box|ForEach-Object {$_.ToString([Globalization.CultureInfo]::InvariantCulture)})+@($field.font,[string]$field.size,$(if($field.check){'1'}else{'0'}),$field.source)
   if(@($data|Where-Object {$_ -match '[|\r\n]'}).Count){throw 'Invalid print field delimiter'}
   $lines.Add('  rows.Add '+(PDFVB ($data -join '|')))
  }
 }
}
foreach($s in @(' End Select',' Set PDFPrintRows = rows','End Function','Public Function PDFPrintBackgrounds() As Variant',(' PDFPrintBackgrounds = Array('+(($backgrounds|ForEach-Object {PDFVB $_}) -join ',')+')'),'End Function','Public Function PDFPrintAdditionalField(ByVal key As String) As Boolean',' Select Case key',' Case "FormKind", "Title", "MemberName", "DraftNotes", "MemberDoDID", "MemberFirstName", "MemberMiddleInitial", "MemberLastName", "MemberSuffix", "ReportingSeniorLastName", "ReportingSeniorFirstInitial", "ReportingSeniorMiddleInitial", "CounselorLastName", "CounselorFirstInitial", "CounselorMiddleInitial"','  PDFPrintAdditionalField = False',' Case Else','  PDFPrintAdditionalField = True',' End Select','End Function','Public Function PDFSummaryRows(ByVal template As Long) As Variant',' If template = 1 Then','  PDFSummaryRows = Array(192.12,202.92,213.72,224.52,235.32,246.12,256.92,267.72,278.52,289.32,300.12,311.64,323.16,334.68,346.2,357,367.8,378.6,389.4,400.2,411,421.8,432.6,444.12)',' Else','  PDFSummaryRows = Array(192.12,202.92,213.72,224.52,235.32,246.12,256.92,267.72,278.52,289.32,300.12,311.64,323.16,334.68,346.2,357,367.8,378.6,389.4,400.2,411,421.8,432.6,443.4,454.2,465,475.8,487.32,498.12,509.64,520.44,531.24,542.76,553.56,564.36,575.16,586.68,597.48,608.28,619.08,629.88,640.68,652.2,663,674.52,685.32,696.12,706.92,718.44,729.96)',' End If','End Function')){$lines.Add($s)}
$target=Join-Path $root 'build/modPDFPrintCatalogue.bas'
[IO.File]::WriteAllLines($target,$lines,[Text.UTF8Encoding]::new($false))
$app.LoadFromText(5,'modPDFPrintCatalogue',$target)
