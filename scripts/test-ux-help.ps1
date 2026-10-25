param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/help-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$topics=Get-Content (Join-Path $root 'src/ux-help-topics.json') -Raw | ConvertFrom-Json
$app=New-Object -ComObject Access.Application;$checks=0
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true);$app.DoCmd.RunCommand(126)
$app.DoCmd.OpenForm('frmUXHelp');if($app.Eval('Forms!frmUXHelp!lstTopics.ListCount') -ne $topics.Count){throw 'Guide topic list incomplete'};$checks++
for($i=0;$i -lt $topics.Count;$i++){
$body=$topics[$i].Body.Replace('{APP_VERSION}',[string]$app.Eval('DLookup("AppVersion","AppMetadata")')).Replace('{SCHEMA_VERSION}',[string]$app.Eval('DLookup("SchemaVersion","AppMetadata")'))
if($app.Eval('UXHelpText('+$i+')') -cne $body){throw ('Native guide differs from canonical topic '+$i)};$checks++
$app.Forms.Item('frmUXHelp').Controls.Item('lstTopics').Value=$i
[void]$app.Eval('UXHelpSelected()')
if($app.Eval('Forms!frmUXHelp!txtGuide') -cne $body){throw ('Selection did not render topic '+$i)};$checks++
}
if(-not $app.Eval('Forms!frmUXHelp!txtGuide.Locked And Forms!frmUXHelp!txtGuide.ScrollBars=2')){throw 'Long guide text is not read-only/scrollable'};$checks++
$app.DoCmd.Close(2,'frmUXHelp',2)
#The Settings guide entry opens the shared guide through contextual window help.
$app.DoCmd.OpenForm('frmUXSettings')
$guideButton=$app.Forms.Item('frmUXSettings').Controls.Item('cmdWorkspaceGuide')
if($guideButton.OnClick -ne '=UXWindowHelp("frmUXSettings")' -or -not $guideButton.Visible){throw 'Settings guide entry missing'};$checks++
[void]$app.Eval('UXWindowHelp("frmUXSettings")')
if(-not ([string]$app.Eval('Forms!frmUXHelp!txtGuide')).EndsWith([string]$app.Eval('UXHelpText(UXWindowHelpTopic("frmUXSettings"))'))){throw 'Settings guide entry did not open its guide topic'};$checks++
Write-Output ('PASS: '+$checks+' native help/documentation consistency assertions')
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
