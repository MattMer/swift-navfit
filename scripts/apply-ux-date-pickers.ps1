# Native calendar buttons, preserving the unbound raw-value editor/recovery buffer.
if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modRSDirectoryUI',(Join-Path $root 'src/modRSDirectoryUI.bas'))}
$dateNames=@('txtDue','txtAsOf','txtFrom','txtTo')
for($i=0;$i -lt $uxFields.Count;$i++){
    if($uxFields[$i].date){$dateNames+=@($uxNames[$i],('value'+$uxFields[$i].key))}
}
foreach($formName in @($app.CurrentProject.AllForms|ForEach-Object {$_.Name})){
    $app.DoCmd.OpenForm($formName,1);$f=$app.Forms.Item($formName)
    foreach($c in $f.Controls){
        if($c.ControlType -eq 109 -and $dateNames.Contains($c.Name)){
            $c.Format='yyyy-mm-dd';$c.ShowDatePicker=1;$c.InputMask=''
            $c.ControlTipText='Choose a date from the calendar or enter YYYY-MM-DD.'
            if($c.Name -match 'CounselingDate'){$c.ControlTipText+=' Authorized counseling status text is also supported.'}
        }
    }
    SaveScreen $f $formName
}
