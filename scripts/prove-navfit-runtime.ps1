param([Parameter(Mandatory=$true)][string]$Database,[Parameter(Mandatory=$true)][string]$OutputDirectory,[switch]$SaveMarker)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Windows.Forms,System.Drawing,System.Data
$a=[Reflection.Assembly]::LoadFrom('D:\Development\Swift NAVFIT\NAVFIT98A\NAVFIT98A.exe')
$flags=[Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
function Prop($obj,$name){$obj.GetType().GetProperty($name,$flags).GetValue($obj,$null)}
function Invoke-Native($type,$method,$obj,$arguments=@()){try{$unwrapped=New-Object object[] $arguments.Count;for($i=0;$i -lt $arguments.Count;$i++){if($null -eq $arguments[$i]){$unwrapped[$i]=$null}else{$unwrapped[$i]=$arguments[$i].PSObject.BaseObject}};$a.GetType('NAVFIT98A.'+$type).GetMethod($method,$flags).Invoke($obj,$unwrapped)}catch{throw ($type+'.'+$method+': '+$_.Exception.InnerException.ToString())}}
$settings=$a.GetType('NAVFIT98A.My.MySettings').GetProperty('Default',$flags).GetValue($null,$null)
$reference='D:\Development\Swift NAVFIT\NAVFIT98A\Data\NF98A_Dte.accdb'
if(-not (Test-Path -LiteralPath $reference)){throw 'Missing NAVFIT reference database'}
$settings.DBPath=$Database;$settings.DBProvider='Microsoft.ACE.OLEDB.16.0';$settings.DBACCDBPathDte=[IO.Path]::GetFileName($reference);$settings.DBACCDBPathRef=[IO.Path]::GetDirectoryName($reference)+'\'
$referenceConnection=New-Object Data.OleDb.OleDbConnection('Provider=Microsoft.ACE.OLEDB.16.0;Data Source='+$reference);$referenceConnection.Open();$referenceConnection.Close()
'Native reference path configured in memory'
$forms=$a.GetType('NAVFIT98A.My.MyProject').GetProperty('Forms',$flags).GetValue($null,$null)
$main=Prop $forms 'NF_Main';(Prop $main 'tbFolderSelection').Text='1';(Prop $main 'tbFolderTraitCnt').Text='7';(Prop $main 'tbFolderTraitSum').Text='28'
foreach($name in @('tbPromSP','tbPromPG','tbPromPM','tbPromMP','tbPromEP')){(Prop $main $name).Text='0'}
$conn=New-Object Data.OleDb.OleDbConnection('Provider=Microsoft.ACE.OLEDB.16.0;Data Source='+$Database)
$conn.Open();$cmd=$conn.CreateCommand();$cmd.CommandText='SELECT ReportID,ReportType,Comments,Parent FROM Reports WHERE ReportType Is Not Null ORDER BY ReportID';$reader=$cmd.ExecuteReader();$records=@()
while($reader.Read()){$records+=@{ID=[int]$reader['ReportID'];Kind=[string]$reader['ReportType'];Comments=[string]$reader['Comments'];Parent=[string]$reader['Parent']}}
$reader.Close();$conn.Close();[void](New-Item -ItemType Directory -Force -Path $OutputDirectory)
$results=@()
try {
foreach($record in $records){
 $spec=switch($record.Kind){'Eval'{@('NF_Eval','EvalControllerModule','FillComboBoxesEval','LoadEvaluationForm','UpdateEvaluation','tbEComments','NF_EvalPage2Data')}'Chief'{@('NF_Chief','ChiefControllerModule','FillComboBoxesChief','LoadChiefForm','UpdateChief','tbCComments','NF_ChiefPage2Data')}'FitRep'{@('NF_FitRep','FitrepControllerModule','FillComboBoxesFitrep','LoadFitRepForm','UpdateFitRep','tbFComments','NF_FitRepPage2Data')}}
 $form=Prop $forms $spec[0];[void](Invoke-Native $spec[0] $spec[2] $form)
 foreach($property in $form.GetType().GetProperties($flags)|Where-Object {$_.Name -match '^tb.*Summary(SP|Prog|Prom|MP|EP)$'}){$property.GetValue($form,$null).Text='0'}
 (Prop $main 'tbRptId').Text=[string]$record.ID;(Prop $main 'tbFolderSelection').Text=($record.Parent -replace '^a ','')
 try{[void](Invoke-Native $spec[1] $spec[3] $null @([string]$record.ID))}catch{ 'ID='+$record.ID;foreach($property in $form.GetType().GetProperties($flags)|Where-Object {$_.Name -match '^tb.*Summary(SP|Prog|Prom|MP|EP)$'}){$property.Name+'='+$property.GetValue($form,$null).Text};throw}
 $comments=Prop $form $spec[5]
 if($comments.Text -cne $record.Comments){throw ('Native load changed comments for '+$record.Kind)}
 $png=Join-Path $OutputDirectory ($record.Kind+'-'+$record.ID+'-page2.png')
 $bitmap=New-Object Drawing.Bitmap(850,1100);$bitmap.SetResolution(100,100);$g=[Drawing.Graphics]::FromImage($bitmap);$g.Clear([Drawing.Color]::White)
 $page=New-Object Drawing.Printing.PageSettings
 $event=New-Object Drawing.Printing.PrintPageEventArgs($g,(New-Object Drawing.Rectangle(0,0,850,1100)),(New-Object Drawing.Rectangle(0,0,850,1100)),$page)
 [void](Invoke-Native 'PrintModule' $spec[6] $null @($null,$event))
 $bitmap.Save($png,[Drawing.Imaging.ImageFormat]::Png);$g.Dispose();$bitmap.Dispose()
 if($SaveMarker){$comments.Text=$comments.Text+"`r`nNATIVE_RUNTIME_RETURN_PROOF";[void](Invoke-Native $spec[1] $spec[4] $null)}
 $results+=@{Kind=$record.Kind;ID=$record.ID;Loaded=$true;NativeSave=[bool]$SaveMarker;Font=$comments.Font.Name;FontPoints=$comments.Font.Size;Width=$comments.Width;Height=$comments.Height;MaxLength=$comments.MaxLength;NativePage=$png}
}
$results|ConvertTo-Json -Depth 5|Set-Content (Join-Path $OutputDirectory 'native-runtime-proof.json') -Encoding UTF8
'PASS: native controllers loaded '+$results.Count+' reports; page-data draw methods executed; save='+[bool]$SaveMarker
}finally{foreach($name in @('NF_Eval','NF_Chief','NF_FitRep','NF_Main')){ $f=Prop $forms $name;$f.Dispose()}}
