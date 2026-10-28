param([Parameter(Mandatory=$true)][string]$BuildPath)
# Section 508 / WCAG 2.0 AA checks that can be automated:
#  - text contrast of every color pair UXApplyVisual uses (1.4.3, 4.5:1);
#  - every user-facing input control has a linked label (1.3.1, 4.1.2), which is
#    what a screen reader announces. Link the visible label (LabelName); where
#    there is none, attach a 15x15-twip label named "an<Control>" whose caption
#    is the spoken name;
#  - every command button has a name (caption or ScreenTip) (4.1.2).
# Known gaps live in tests/accessibility-gaps.json with a reason each; the list
# may only shrink. Screen-reader behavior itself still needs a person to verify.
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent;$checks=0
function Check([bool]$ok,[string]$message){if(-not $ok){throw $message};$script:checks++}

#--- Contrast, from the palette in source --------------------------------------
$palette=@{}
foreach($m in [regex]::Matches((Get-Content (Join-Path $root 'src/modUXVisual.bas') -Raw),'Case "(\w+)": UXColor = RGB\((\d+), (\d+), (\d+)\)')){$palette[$m.Groups[1].Value]=@([int]$m.Groups[2].Value,[int]$m.Groups[3].Value,[int]$m.Groups[4].Value)}
$palette['White']=@(255,255,255)
function Luminance($rgb){$l=@($rgb|ForEach-Object {$c=$_/255;if($c -le 0.03928){$c/12.92}else{[Math]::Pow(($c+0.055)/1.055,2.4)}});return 0.2126*$l[0]+0.7152*$l[1]+0.0722*$l[2]}
function Ratio($fg,$bg){$a=Luminance $palette[$fg];$b=Luminance $palette[$bg];return ([Math]::Max($a,$b)+0.05)/([Math]::Min($a,$b)+0.05)}
#Foreground on background, as UXApplyVisual assigns them.
#Muted on Neutral is also read-only reference text in the trait grid, not only disabled.
foreach($pair in @('Ink/Canvas','Ink/White','Ink/Neutral','Muted/Neutral','Muted/White','White/Primary','White/Rail','White/Active','HelpInk/Help','ErrorInk/Error','WarningInk/Warning','SuccessInk/Success')){
 $fg,$bg=$pair -split '/';$r=Ratio $fg $bg
 Check ($r -ge 4.5) ('Text contrast below 4.5:1: '+$pair+' = '+[Math]::Round($r,2))
}

#--- Labels and button names, from the built forms -----------------------------
$gaps=Get-Content (Join-Path $root 'tests/accessibility-gaps.json') -Raw|ConvertFrom-Json
$known=@{};foreach($p in $gaps.labels.PSObject.Properties){$known[$p.Name]=$p.Value};foreach($p in $gaps.buttons.PSObject.Properties){$known[$p.Name]=$p.Value}
$retired=@((Get-Content (Join-Path $root 'scripts/build-manifest.json') -Raw|ConvertFrom-Json).retiredForms.PSObject.Properties.Name)
$scratch=Join-Path $root ('build/tests/accessibility-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application;$missing=[Collections.Generic.List[string]]::new();$labelled=0;$buttons=0;$names=0
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 for($wait=0;$wait -lt 30 -and $app.Forms.Count -eq 0;$wait++){Start-Sleep -Milliseconds 100}
 foreach($open in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$open,2)}
 foreach($name in @($app.CurrentProject.AllForms|ForEach-Object {$_.Name}|Where-Object {$retired -notcontains $_})){
  $app.DoCmd.OpenForm($name,1);$f=$app.Forms.Item($name)
  foreach($c in $f.Controls){
   $tag=[string]$c.Tag
   if($c.ControlType -eq 100 -and $c.Name -clike 'an*' -and $c.Width -le 30){
    #Name label: an attached caption only a screen reader uses; it must stay out of sight.
    $names++;Check ($c.Height -le 30 -and ([string]$c.Caption).Trim() -ne '' -and $c.Parent.Name -ne $name -and $c.Visible -eq $c.Parent.Visible) ('Name label is oversized, empty, unattached or shown apart from its control: '+$name+'.'+$c.Name)
    continue
   }
   #Technical controls nobody reads: 100-twip state holders, focus/sort helpers,
   #the editor's hidden buffer and the bulk screen's hidden mirrors.
   if(($c.Width -le 120 -and $c.Height -le 120) -or $c.Name -like 'uxSortActivate*' -or $c.Name -eq 'uxFocusAnchor' -or $tag -eq 'BUFFER' -or $tag.StartsWith('BLX:')){continue}
   if($c.ControlType -in 109,110,111,106,105){
    $linked=$false;try{$linked=$c.Controls.Count -gt 0}catch{}
    if($linked){$labelled++}else{$missing.Add($name+'.'+$c.Name)}
   }elseif($c.ControlType -eq 104){
    $buttons++
    if(([string]$c.Caption).Trim() -eq '' -and ([string]$c.ControlTipText).Trim() -eq ''){$missing.Add($name+'.'+$c.Name)}
   }
  }
  $app.DoCmd.Close(2,$name,2)
 }
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
$unexpected=@($missing|Where-Object {-not $known.ContainsKey($_)});$fixed=@($known.Keys|Where-Object {$missing -notcontains $_})
Check ($unexpected.Count -eq 0) ('Controls without an accessible name: '+($unexpected -join ', '))
Check ($fixed.Count -eq 0) ('Known accessibility gap is fixed; remove it from tests/accessibility-gaps.json: '+($fixed -join ', '))
Write-Output ('PASS: '+$checks+' checks; 12 color pairs at 4.5:1 or better; '+$labelled+' labelled input controls ('+$names+' by name labels); '+$buttons+' named buttons; '+$known.Count+' known gaps listed')
