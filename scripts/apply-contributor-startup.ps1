# Session-only contributor choice. No tables/history are changed by this stage.
. (Join-Path $PSScriptRoot 'access-properties.ps1')
$app.LoadFromText(2,'frmContributorStart',(Join-Path $root 'src/BlankForm.txt'))
$app.DoCmd.OpenForm('frmContributorStart',1);$f=$app.Forms.Item('frmContributorStart')
$f.Caption='Choose your contributor';$f.Width=8600;$f.Section(0).Height=6300
$f.PopUp=$true;$f.Modal=$true;$f.BorderStyle=3;$f.AutoCenter=$true
$f.CloseButton=$false;$f.RecordSelectors=$false;$f.NavigationButtons=$false;$f.ScrollBars=0
$f.OnLoad='=LoadContributorStartup()';$f.OnResize=''
function StartupControl([int]$type,[string]$name,[int]$x,[int]$y,[int]$w,[int]$h,[string]$caption='',[string]$event='',[string]$parent=''){
 $control=$app.CreateControl($f.Name,$type,0,$parent,'',$x,$y,$w,$h)
 $control.Name=$name;$control.FontName='Segoe UI';$control.FontSize=11
 if($type -in @(100,104)){$control.Caption=$caption}
 if($event){$control.OnClick=$event}
 return $control
}
$c=StartupControl 100 'lblHeading' 400 300 7800 600 'Who is contributing today?';$c.FontSize=20
[void](StartupControl 100 'lblInfo' 400 1050 7800 850 'Choose a contributor for this launch. This name records local authorship; it is not sign-in or proof of identity.')
$c=StartupControl 111 'cboContributor' 400 2450 7800 450;$c.ColumnCount=2;$c.ColumnWidths='0;7400';$c.BoundColumn=1;$c.LimitToList=$true;$c.AutoExpand=$true
[void](StartupControl 100 'lblExisting' 400 2050 7800 350 'Existing contributor' '' 'cboContributor')
[void](StartupControl 104 'cmdContinue' 400 3100 7800 500 'Continue as selected contributor' '=AcceptContributorStartup(False)')
[void](StartupControl 109 'txtNewContributor' 400 4300 5100 450)
[void](StartupControl 100 'lblNew' 400 3900 7800 350 'Or add your name as a new contributor' '' 'txtNewContributor')
[void](StartupControl 104 'cmdAdd' 5700 4300 2500 450 'Add and continue' '=AcceptContributorStartup(True)')
[void](StartupControl 100 'lblStatus' 400 5050 7800 650 '')
[void](StartupControl 104 'cmdCancel' 400 5800 7800 400 'Exit Swift NAVFIT' '=CancelContributorStartup()')
$app.DoCmd.Close(2,'frmContributorStart',1)
# Setup is already an explicit new-contributor choice; cancellation adds no rows.
$app.DoCmd.OpenForm('frmSetup',1);$f=$app.Forms.Item('frmSetup')
if(@($f.Controls|ForEach-Object {$_.Name}).Contains('cmdCancelSetup')){$app.DeleteControl('frmSetup','cmdCancelSetup')}
$bottom=0
foreach($control in $f.Controls){if($control.Visible){$bottom=[Math]::Max([int]$bottom,[int]($control.Top+$control.Height))}}
$c=StartupControl 104 'cmdCancelSetup' 400 ($bottom+250) 9950 500 'Exit without setting up' '=CancelContributorStartup()'
SetAccessProperty ($f.Section(0)) 'Height' ([Math]::Max([int]$f.Section(0).Height,[int]($bottom+1050)));$f.ScrollBars=2
$app.DoCmd.Close(2,'frmSetup',1)
