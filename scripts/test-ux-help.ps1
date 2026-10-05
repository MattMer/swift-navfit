param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/help-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$topics=Get-Content (Join-Path $root 'src/ux-help-topics.json') -Raw | ConvertFrom-Json
$guide=Get-Content (Join-Path $root 'planning/UX_WORKSPACE_GUIDE.md') -Raw
$app=New-Object -ComObject Access.Application;$checks=0
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true);$app.DoCmd.RunCommand(126)
$app.DoCmd.OpenForm('frmUXHelp');if($app.Eval('Forms!frmUXHelp!lstTopics.ListCount') -ne $topics.Count){throw 'Guide topic list incomplete'};$checks++
for($i=0;$i -lt $topics.Count;$i++){
if($app.Eval('UXHelpText('+$i+')') -cne $topics[$i].Body){throw ('Native guide differs from canonical topic '+$i)};$checks++
if(-not $guide.Contains($topics[$i].Body)){throw ('Workspace guide differs from topic '+$i)};$checks++
$app.Forms.Item('frmUXHelp').Controls.Item('lstTopics').Value=$i
[void]$app.Eval('UXHelpSelected()')
if($app.Eval('Forms!frmUXHelp!txtGuide') -cne $topics[$i].Body){throw ('Selection did not render topic '+$i)};$checks++
}
if(-not $app.Eval('Forms!frmUXHelp!txtGuide.Locked And Forms!frmUXHelp!txtGuide.ScrollBars=2')){throw 'Long guide text is not read-only/scrollable'};$checks++
$app.DoCmd.Close(2,'frmUXHelp',2)
$app.DoCmd.OpenForm('frmUXSettings',1)
if($app.Forms.Item('frmUXSettings').Controls.Item('cmdWorkspaceGuide').OnClick -ne '=OpenScreen("frmUXHelp")'){throw 'Settings guide entry missing'};$checks++
Write-Output ('PASS: '+$checks+' native help/documentation consistency assertions')
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
