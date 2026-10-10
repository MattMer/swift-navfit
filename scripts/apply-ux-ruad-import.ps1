foreach($module in @('modRUADProfiles','modCommandDirectory','modRUADImport','modUXRUADImport')){$app.LoadFromText(5,$module,(Join-Path $root ('src/'+$module+'.bas')))}
NewScreen 'frmUXRUADImport' 'Import or refresh command RUAD' 9500
$f=$script:swiftForm;$f.PopUp=$true;$f.Width=13800;$f.ScrollBars=2
[void](Control $f 100 'lblSteps' 400 900 12600 350 'Choose CSV → Review matches and full values → Preview → Apply → Results')
[void](Control $f 109 'txtPath' 400 1350 9400 420);$script:swiftControl.Locked=$true
[void](Control $f 104 'cmdChoose' 10100 1350 2900 450 'Choose RUAD CSV' '=UXRUADImportChoose()')
[void](Control $f 104 'cmdNewProfiles' 400 1950 5100 450 'Select new profiles for unmatched names' '=UXRUADImportNew()')
[void](Control $f 100 'lblColumns' 400 2500 12600 300 'Member / vacant billet                         DESG           RUIC           Action')
[void](Control $f 110 'lstRows' 400 2900 5400 3500);$c=$script:swiftControl
$c.RowSourceType='Value List';$c.ColumnCount=5;$c.BoundColumn=1;$c.ColumnWidths='0;2600;750;750;1200';$c.AfterUpdate='=UXRUADImportSelect()'
[void](Control $f 100 'lblAction' 6100 2500 1800 300 'Selected row action')
[void](Control $f 111 'cboAction' 6100 2850 2700 420);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='Skip;Create new;Refresh existing';$c.LimitToList=$true;$c.AfterUpdate='=UXRUADImportSelection()'
[void](Control $f 100 'lblMatch' 9200 2500 3700 300 'Explicit existing member match')
[void](Control $f 111 'cboMatch' 9200 2850 3800 420);$c=$script:swiftControl;$c.ColumnCount=3;$c.BoundColumn=1;$c.ColumnWidths='0;3000;750';$c.LimitToList=$true;$c.AfterUpdate='=UXRUADImportSelection()'
[void](Control $f 106 'chkReplaceManual' 6100 3450 250 250);$script:swiftControl.AfterUpdate='=UXRUADImportSelection()'
[void](Control $f 100 'lblReplace' 6480 3390 6500 500 'Use incoming RUAD values for manually changed roster fields (core identity fields stay unchanged)')
[void](Control $f 109 'txtDetails' 6100 4100 6900 2300);$c=$script:swiftControl;$c.Locked=$true;$c.ScrollBars=2;$c.EnterKeyBehavior=$true
[void](Control $f 100 'lblStatus' 400 6600 12600 1000 'Choose a CSV. Nothing is imported until you review and apply a preview. Vacant and absent members are never automatically deleted or archived.')
[void](Control $f 104 'cmdPreview' 400 7900 3800 500 'Preview selected changes' '=UXRUADImportPreview()')
[void](Control $f 104 'cmdApply' 4500 7900 3800 500 'Apply reviewed changes' '=UXRUADImportApply()');$script:swiftControl.Enabled=$false
[void](Control $f 104 'cmdClose' 8600 7900 4400 500 'Close' '=UXDirectoryClose("frmUXRUADImport")')
$tab=0;foreach($name in @('cmdChoose','cmdNewProfiles','lstRows','cboAction','cboMatch','chkReplaceManual','txtDetails','cmdPreview','cmdApply','cmdClose')){$f.Controls($name).TabIndex=[int]$tab;$tab++}
$f.Section(0).BackColor=15397112
foreach($c in $f.Controls){switch([int]$c.ControlType){100{$c.ForeColor=2501670}104{$c.UseTheme=$false;$c.BackColor=7254527;$c.ForeColor=2501670;$c.BorderColor=5948108;if($c.Name -eq 'cmdApply'){$c.BackColor=5948108}}109{$c.BackColor=16777215;$c.ForeColor=2501670;$c.BorderColor=9682379}110{$c.BackColor=16777215;$c.ForeColor=2501670;$c.BorderColor=9682379}111{$c.BackColor=16777215;$c.ForeColor=2501670;$c.BorderColor=9682379}}}
$f.OnLoad='=RUADReset()'
SaveScreen $f 'frmUXRUADImport'
$app.DoCmd.OpenForm('frmUXDirectory',1);$f=$app.Forms.Item('frmUXDirectory')
if(-not @($f.Controls|ForEach-Object {$_.Name}).Contains('cmdRUADImport')){[void](Control $f 104 'cmdRUADImport' 400 8000 12600 500 'Import / refresh RUAD CSV' '=UXRUADImportOpen()')}
$f.Controls('cmdRUADImport').UseTheme=$false;$f.Controls('cmdRUADImport').BackColor=7254527;$f.Controls('cmdRUADImport').ForeColor=2501670;$f.Controls('cmdRUADImport').BorderColor=5948108
$f.Controls('cmdRUADImport').Top=8000;$f.Controls('cmdRUADImport').Width=12600
if($f.Section(0).Height -lt 8800){$f.Section(0).Height=8800}
SaveScreen $f 'frmUXDirectory'
