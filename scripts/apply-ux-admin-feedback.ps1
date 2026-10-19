# Run after the existing UI helpers; retain all backend control identities.
. (Join-Path $PSScriptRoot 'apply-ux-block28-schema.ps1')
foreach($m in @('modUXReadRows','modSummaryGroups','modOperations','modUXGroupLifecycle','modUXBlockTools','modUXSurface','modUX','modUXFeedback','modUXMyWork','modUXTraitGrid','modSnippets','modMeetings','modRouteTemplates')){
    $app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))
}
$layout=Get-Content (Join-Path $root 'scripts/apply-ux-layout.ps1') -Raw
Invoke-Expression ($layout.Substring(0,$layout.IndexOf('function UXChrome')))
$workspace=Get-Content (Join-Path $root 'scripts/apply-ux-workspace.ps1') -Raw
$begin=$workspace.IndexOf('$surfaceModule=');$end=$workspace.IndexOf("NewScreen 'frmUXEditorFields'")
Invoke-Expression ($workspace.Substring($begin,$end-$begin))
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor')
$f.Controls('uxSection0').Caption='Admin data';$f.Controls('uxSection1').Visible=$false
$f.Controls('uxSeniorRecord').Tag='LOCAL:0';$f.Controls('uxUseSenior').Tag='LOCAL:0'
$f.Controls('uxCompactSection').RowSource=((0..9|Where-Object {$_ -ne 1}|ForEach-Object {[string]$_+';'+$uxSections[$_]}) -join ';')
SaveScreen $f 'frmEditor'
$app.DoCmd.OpenForm('frmMajorWorkbench',1);$f=$app.Forms.Item('frmMajorWorkbench')
$f.Controls('uxMode1').Caption='Defaults for new reports';$f.Controls('uxMode1').OnClick='=UXOpenGroupDefaults()'
$f.Controls('uxMode3').Caption='Members & comparison';$f.Controls('uxMode3').OnClick='=UXOpenComparison()'
$f.Controls('lstReports').OnDblClick='=UXGroupRowOpen()'
$f.Controls('uxDefaultsEditor').Caption='Defaults for new reports'
$f.Controls('uxBulk').Caption='Change blocks on existing reports'
SaveScreen $f 'frmMajorWorkbench'
$app.DoCmd.OpenForm('frmUXTraitRows',1);$f=$app.Forms.Item('frmUXTraitRows')
$f.Controls('Member').OnDblClick='=UXGridOpenReport()'
SaveScreen $f 'frmUXTraitRows'
$app.DoCmd.OpenForm('frmMyWork',1);$f=$app.Forms.Item('frmMyWork')
if(-not @($f.Controls|ForEach-Object {$_.Name}).Contains('uxContributor')){
    [void](Control $f 100 'uxContributorLabel' 9000 1800 4000 300 'Active contributor')
    [void](Control $f 111 'uxContributor' 9000 2150 4000 420)
    $c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;3500';$c.RowSourceType='Table/Query';$c.LimitToList=$true;$c.AfterUpdate='=UXWorkContributor()';$c.ControlTipText='Active contributor';$c.TabIndex=[int]$f.Controls('chkAll').TabIndex
}
SaveScreen $f 'frmMyWork'
$app.DoCmd.OpenForm('frmSnippets',1);$f=$app.Forms.Item('frmSnippets')
if(-not @($f.Controls|ForEach-Object {$_.Name}).Contains('uxSources')){
    [void](Control $f 104 'uxSources' 10300 1350 2500 360 'Browse sources' '=UXOpenSnippetSources()')
}
SaveScreen $f 'frmSnippets'

$app.DoCmd.OpenForm('frmUXEditorFields',1);$f=$app.Forms.Item('frmUXEditorFields')
if(-not @($f.Controls|ForEach-Object {$_.Name}).Contains('txtCommandEmployment')){
    [void](Control $f 109 'txtCommandEmployment' 200 200 3000 2600);$c=$script:swiftControl;$c.Tag='SURF:CommandEmployment';$c.AfterUpdate='=UXSurfaceEdited("CommandEmployment")';$c.EnterKeyBehavior=$true;$c.ScrollBars=2
    [void](Control $f 100 'labelCommandEmployment' 200 100 3000 300 '28. Command employment and command achievements')
    [void](Control $f 100 'errCommandEmployment' 200 2900 3000 250 '')
}
SaveScreen $f 'frmUXEditorFields'
$app.DoCmd.OpenForm('frmUXBulk',1);$f=$app.Forms.Item('frmUXBulk')
if(-not @($f.Controls|ForEach-Object {$_.Name}).Contains('uxValueCommandEmployment')){
    [void](Control $f 109 'uxValueCommandEmployment' 10 10 100 100);$script:swiftControl.Visible=$false
}
$f.Controls('uxBlockSection1').Visible=$false
$f.Controls('uxBlockSection0').Caption='Admin data'
SaveScreen $f 'frmUXBulk'
$app.DoCmd.OpenForm('frmUXSharedFields',1);$f=$app.Forms.Item('frmUXSharedFields')
if(-not @($f.Controls|ForEach-Object {$_.Name}).Contains('valueCommandEmployment')){
    [void](Control $f 109 'valueCommandEmployment' 150 500 3000 1800);$c=$script:swiftControl;$c.AfterUpdate='=UXBlockEdit("CommandEmployment")';$c.EnterKeyBehavior=$true;$c.ScrollBars=2
    [void](Control $f 100 'labelCommandEmployment' 150 150 3000 300 '28. Command employment and command achievements')
    [void](Control $f 104 'clearCommandEmployment' 3300 500 450 330 'X' '=UXBlockClear("CommandEmployment")')
    [void](Control $f 104 'resetCommandEmployment' 2400 500 800 330 'Inherit' '=UXResetDefault("CommandEmployment")')
    [void](Control $f 100 'stateCommandEmployment' 150 2450 3000 260 'Unchanged')
}
SaveScreen $f 'frmUXSharedFields'

$app.DoCmd.OpenForm('frmUXNavyWriter',1);$f=$app.Forms.Item('frmUXNavyWriter')
$f.Controls('cmdBack').OnClick='=UXReturnNarrativeSources()'
SaveScreen $f 'frmUXNavyWriter'

$app.DoCmd.OpenForm('frmUXSeniorEditor',1);$f=$app.Forms.Item('frmUXSeniorEditor')
foreach($key in @('ReportingSeniorLastName','ReportingSeniorFirstInitial','ReportingSeniorMiddleInitial')){
    if(-not @($f.Controls|ForEach-Object {$_.Name}).Contains('rs'+$key)){
        [void](Control $f 109 ('rs'+$key) 10 10 100 100)
    }
    $f.Controls('rs'+$key).Visible=$false;$f.Controls('rs'+$key).Tag=''
    if(@($f.Controls|ForEach-Object {$_.Name}).Contains('lblRS'+$key)){$f.Controls('lblRS'+$key).Visible=$false;$f.Controls('lblRS'+$key).Tag=''}
}
SaveScreen $f 'frmUXSeniorEditor'

. (Join-Path $PSScriptRoot 'apply-ux-date-pickers.ps1')

. (Join-Path $PSScriptRoot 'apply-ui-author-baseline.ps1')

. (Join-Path $PSScriptRoot 'apply-ux-group-directory.ps1')

. (Join-Path $PSScriptRoot 'apply-ux-comparison-autosave.ps1')

. (Join-Path $PSScriptRoot 'apply-ux-ruad-import.ps1')
