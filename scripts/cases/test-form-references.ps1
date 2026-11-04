param()
# M05 form/control coupling contract, read from source without Access. Every form
# and control the VBA names literally must exist in src/forms: Forms!f!c,
# Forms("f"), With Forms!f blocks, form variables bound with Set x = Forms!f (or
# any form when unbound), Me! in form modules, Controls("c"), UXHasControl probes,
# OpenForm/Close/AllForms targets, form-property expressions and SORT tags.
# Names built at run time are covered by the runtime view contracts (UXAssertView).
# Mutated inventories then prove that a renamed or removed control or form fails.
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$checks=0
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'view-control-contracts.ps1')
$dynamicViewControls=GetSwiftDynamicViewControls $root
function Check([bool]$ok,[string]$message){if(-not $ok){throw $message};$script:checks++}

# Fields of a record source set at run time that code reads by name with no
# control of their own. Each entry is part of that form's contract.
$recordFields=@{'frmUXTraitRows'=@('PriorPeriod','PriorSource')}

function ReadSource([string]$Root){
    $forms=@{};$modules=[ordered]@{}
    foreach($file in Get-ChildItem (Join-Path $Root 'src/forms') -Filter '*.form.txt'){
        $name=$file.Name -replace '\.form\.txt$',''
        $text=[IO.File]::ReadAllText($file.FullName)
        $split=$text.IndexOf("`nCodeBehindForm")
        $design=if($split -ge 0){$text.Substring(0,$split)}else{$text}
        $controls=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach($m in [regex]::Matches($design,'(?m)^\s*Name ="([^"]+)"')){[void]$controls.Add($m.Groups[1].Value)}
        $subforms=@{}
        foreach($m in [regex]::Matches($design,'(?ms)^\s*Begin Subform\s*$(.*?)^\s*End\s*$')){
            $n=[regex]::Match($m.Groups[1].Value,'(?m)^\s*Name ="([^"]+)"').Groups[1].Value
            $s=[regex]::Match($m.Groups[1].Value,'(?m)^\s*SourceObject ="(?:Form\.)?([^"]+)"').Groups[1].Value
            if($n -and $s){$subforms[$n]=$s}
        }
        if($recordFields.ContainsKey($name)){foreach($field in $recordFields[$name]){[void]$controls.Add($field)}}
        $forms[$name]=[pscustomobject]@{Controls=$controls;Design=$design;Subforms=$subforms}
        if($split -ge 0){$modules['Form_'+$name]=[pscustomobject]@{Form=$name;Code=$text.Substring($split+15)}}
    }
    foreach($file in Get-ChildItem (Join-Path $Root 'src') -Filter '*.bas'){
        $modules[$file.BaseName]=[pscustomobject]@{Form=$null;Code=[IO.File]::ReadAllText($file.FullName)}
    }
    [pscustomobject]@{Forms=$forms;Modules=$modules}
}

# Drop the comment from one logical VBA line, keeping string literals intact.
function StripComment([string]$Line){
    $inString=$false
    for($i=0;$i -lt $Line.Length;$i++){
        if($Line[$i] -eq '"'){$inString=-not $inString}
        elseif($Line[$i] -eq "'" -and -not $inString){return $Line.Substring(0,$i)}
    }
    if($Line -match '^\s*Rem\s'){return ''}
    $Line
}

function Get-FormReferenceFailures($Source){
    $forms=$Source.Forms;$failures=New-Object Collections.Generic.List[string]
    $anyControl=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach($f in $forms.Values){$anyControl.UnionWith($f.Controls)}
    $constants=@{}
    foreach($module in $Source.Modules.Values){
        foreach($m in [regex]::Matches($module.Code,'(?im)^\s*(?:Public |Private )?Const (\w+) As String = "(\w+)"')){$constants[$m.Groups[1].Value]=$m.Groups[2].Value}
    }
    $script:refCount=0
    function FormOf([string]$Literal,[string]$Constant){
        if($Literal){return $Literal}
        if($constants.ContainsKey($Constant)){return $constants[$Constant]}
        $null
    }
    function NeedForm([string]$Where,[string]$Form){
        $script:refCount++
        if(-not $forms.ContainsKey($Form)){$failures.Add($Where+': form '+$Form+' does not exist');return $false}
        $true
    }
    function NeedControl([string]$Where,[string]$Form,[string]$Control){
        $script:refCount++
        if($Form){
            if(-not $forms.ContainsKey($Form)){$failures.Add($Where+': form '+$Form+' does not exist')}
            elseif(-not $forms[$Form].Controls.Contains($Control)){$failures.Add($Where+': control '+$Form+'.'+$Control+' does not exist')}
        }elseif(-not $anyControl.Contains($Control)){$failures.Add($Where+': control '+$Control+' exists on no form')}
    }
    $formExpr='Forms(?:!\[?(\w+)\]?|\("(\w+)"\)|\((\w+)\))'
    foreach($moduleName in $Source.Modules.Keys){
        $module=$Source.Modules[$moduleName]
        $raw=$module.Code -split "\r?\n";$lines=New-Object Collections.Generic.List[object];$buffer='';$start=0
        for($i=0;$i -lt $raw.Count;$i++){
            if(-not $buffer){$start=$i+1}
            $piece=StripComment $raw[$i]
            if($piece -match '\s_\s*$'){$buffer+=($piece -replace '\s_\s*$',' ');continue}
            $lines.Add(@($start,($buffer+$piece)));$buffer=''
        }
        $formVars=@{};$moduleVars=@{};$withStack=New-Object Collections.Generic.List[object];$inProc=$false
        # First pass per procedure: which form each variable is bound to.
        $procBindings=@{};$proc=0
        foreach($entry in $lines){
            $line=$entry[1]
            if($line -match '^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+\w+'){$proc++;$procBindings[$proc]=@{}}
            foreach($m in [regex]::Matches($line,'\bSet (\w+) = '+$formExpr+'(?:!(\w+)\.Form)?(?![!.(\w])')){
                $form=FormOf ($m.Groups[2].Value+$m.Groups[3].Value) $m.Groups[4].Value
                #A subform's form is the form its SourceObject names.
                if($m.Groups[5].Value){
                    if($form -and $forms.ContainsKey($form) -and $forms[$form].Subforms.ContainsKey($m.Groups[5].Value)){$form=$forms[$form].Subforms[$m.Groups[5].Value]}else{$form=''}
                }
                $var=$m.Groups[1].Value
                if(-not $procBindings.ContainsKey($proc)){$procBindings[$proc]=@{}}
                $b=$procBindings[$proc]
                if($b.ContainsKey($var) -and $b[$var] -ne $form){$b[$var]=''}else{$b[$var]=$form}
            }
        }
        $proc=0
        foreach($entry in $lines){
            $at=$moduleName+':'+$entry[0];$line=$entry[1]
            if($line -match '^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Function|Sub|Property\s+(?:Get|Let|Set))\s+\w+'){
                $proc++;$inProc=$true;$formVars=@{}+$moduleVars;$withStack.Clear()
            }
            if($line -match '^\s*End (Function|Sub|Property)\b'){$inProc=$false;continue}
            # Variables typed as forms, in parameters and Dim lines.
            foreach($m in [regex]::Matches($line,'(?:ByVal |ByRef |Optional |Dim |Private |Public |,\s*|\(\s*)(\w+) As (?:Access\.)?Form\b')){
                $var=$m.Groups[1].Value;$bound=$null
                if($procBindings.ContainsKey($proc) -and $procBindings[$proc].ContainsKey($var)){$bound=$procBindings[$proc][$var]}
                if($inProc){$formVars[$var]=$bound}else{$moduleVars[$var]=$null;$formVars[$var]=$null}
            }
            # With blocks: a form, a form variable or anything else.
            if($line -match '^\s*With\s+(.+?)\s*$'){
                $target=$Matches[1];$withForm=$false;$withName=$null
                if($target -match ('^'+$formExpr+'$')){$withName=FormOf ($Matches[1]+$Matches[2]) $Matches[3];$withForm=$true;if($withName){[void](NeedForm $at $withName)}}
                elseif($target -eq 'Me' -and $module.Form){$withName=$module.Form;$withForm=$true}
                elseif($formVars.ContainsKey($target)){$withName=$formVars[$target];$withForm=$true}
                $withStack.Add(@($withForm,$withName));continue
            }
            if($line -match '^\s*End With\b'){if($withStack.Count){$withStack.RemoveAt($withStack.Count-1)};continue}
            # Literal and constant form names.
            foreach($m in [regex]::Matches($line,$formExpr+'(?:!\[?(\w+)\]?|\.Controls\("(\w+)"\))?')){
                $form=FormOf ($m.Groups[1].Value+$m.Groups[2].Value) $m.Groups[3].Value
                if(-not $form){continue}
                $control=$m.Groups[4].Value+$m.Groups[5].Value
                if($control){NeedControl $at $form $control}else{[void](NeedForm $at $form)}
            }
            foreach($m in [regex]::Matches($line,'(?:DoCmd\.OpenForm|DoCmd\.Close acForm,|AllForms\(|UXAssertView\(|UXCheckView\(|DoCmd\.SelectObject acForm,)\s*"(\w+)"(?!\s*&)')){[void](NeedForm $at $m.Groups[1].Value)}
            # Me in form modules.
            if($module.Form){
                foreach($m in [regex]::Matches($line,'\bMe(?:!\[?(\w+)\]?|\.Controls\("(\w+)"\)|\("(\w+)"\))')){NeedControl $at $module.Form ($m.Groups[1].Value+$m.Groups[2].Value+$m.Groups[3].Value)}
            }
            # Form variables.
            foreach($var in @($formVars.Keys)){
                foreach($m in [regex]::Matches($line,'(?<![\w.!])'+$var+'(?:!\[?(\w+)\]?|\.Controls\("(\w+)"\)|\("(\w+)"\))')){NeedControl $at $formVars[$var] ($m.Groups[1].Value+$m.Groups[2].Value+$m.Groups[3].Value)}
            }
            # Optional-control probes still name controls that must exist somewhere.
            foreach($m in [regex]::Matches($line,'UXHasControl\((\w+(?:!\w+)?|Forms\("\w+"\)|Forms\(\w+\)),\s*"(\w+)"\)')){
                $arg=$m.Groups[1].Value;$form=$null
                if($arg -match '^Forms!(\w+)$'){$form=$Matches[1]}
                elseif($arg -match '^Forms\("(\w+)"\)$'){$form=$Matches[1]}
                elseif($arg -match '^Forms\((\w+)\)$'){$form=FormOf '' $Matches[1]}
                elseif($arg -eq 'Me'){$form=$module.Form}
                elseif($formVars.ContainsKey($arg)){$form=$null}
                NeedControl $at $(if($form){$form}else{$null}) $m.Groups[2].Value
            }
            # Bang references inside With blocks over a form.
            if($withStack.Count -and $withStack[$withStack.Count-1][0]){
                $withName=$withStack[$withStack.Count-1][1]
                foreach($m in [regex]::Matches($line,'(?<![\w\)\]!.])!\[?(\w+)\]?')){NeedControl $at $withName $m.Groups[1].Value}
            }
        }
    }
    # Form properties: expressions naming other forms' controls, and sort tags.
    foreach($name in $forms.Keys){
        $design=$forms[$name].Design
        foreach($m in [regex]::Matches($design,'Forms!\[?(\w+)\]?!\[?(\w+)\]?')){NeedControl ('form '+$name) $m.Groups[1].Value $m.Groups[2].Value}
        foreach($m in [regex]::Matches($design,'Tag ="SORT:(\w+):\d+"')){NeedControl ('form '+$name+' sort tag') $name $m.Groups[1].Value}
    }
    foreach($view in $dynamicViewControls.Keys){
        foreach($control in $dynamicViewControls[$view]){NeedControl 'dynamic view contract' $view $control}
    }
    [pscustomobject]@{Failures=$failures;References=$script:refCount}
}

$source=ReadSource $root
$result=Get-FormReferenceFailures $source
if($result.Failures.Count){throw ('Unresolved form/control references ('+$result.Failures.Count+'): '+(($result.Failures|Select-Object -First 40) -join '; '))}
Check ($result.References -gt 1000) ('Too few references resolved: '+$result.References)

# Detection: each mutation must fail and name what it broke.
function Mutated([string]$Form,[scriptblock]$Change){
    $copy=ReadSource $root
    & $Change $copy.Forms[$Form]
    $copy
}
$renamed=Mutated 'frmUXSeniorEditor' {param($f) [void]$f.Controls.Remove('txtMemberID');[void]$f.Controls.Add('txtMemberIDRenamed')}
Check (@((Get-FormReferenceFailures $renamed).Failures|Where-Object {$_ -like '*frmUXSeniorEditor.txtMemberID *'}).Count -gt 0) 'Renamed literal control was not detected'
$probe=Mutated 'frmEditor' {param($f) [void]$f.Controls.Remove('uxContext')}
foreach($formName in @($source.Forms.Keys)){if($formName -ne 'frmEditor'){[void]$probe.Forms[$formName].Controls.Remove('uxContext')}}
Check (@((Get-FormReferenceFailures $probe).Failures|Where-Object {$_ -like '*uxContext*'}).Count -gt 0) 'Control removed from every form was not detected'
$gone=ReadSource $root;$gone.Forms.Remove('frmUXSeniorPicker')
Check (@((Get-FormReferenceFailures $gone).Failures|Where-Object {$_ -like '*form frmUXSeniorPicker does not exist*'}).Count -gt 0) 'Removed form was not detected'
$sorted=Mutated 'frmUXSeniors' {param($f) [void]$f.Controls.Remove('lstSeniors')}
Check (@((Get-FormReferenceFailures $sorted).Failures|Where-Object {$_ -like '*frmUXSeniors.lstSeniors*'}).Count -gt 0) 'Sort tag list removal was not detected'
Write-Output ('PASS: '+$result.References+' literal form/control references resolve; '+$checks+' contract and mutation checks')
foreach($pair in @(@('frmLibrary','uxNav5'),@('frmContextDuties','txtDutyDescription'),@('frmUXEditorFields','errOfficialName'),@('frmUXSeniorEditor','rsReportingSeniorEmail'),@('frmUXMemberEditor','ruadPage4'),@('frmUXSharedFields','valueUnitUIC'),@('frmUXRSCATable','countO10'),@('frmUXMemberHistory','sort_lstCounseling_4'),@('frmServiceDetails','txtReadinessDetail'),@('frmUXPriorEditor','lblTrait39'),@('frmEditor','uxSection9'))){
    $copy=ReadSource $root;[void]$copy.Forms[$pair[0]].Controls.Remove($pair[1])
    $needle=$pair[0]+'.'+$pair[1]
    Check (@((Get-FormReferenceFailures $copy).Failures|Where-Object {$_ -like ('*dynamic view contract*'+$needle+'*')}).Count -gt 0) ('Dynamic control removal was missed: '+$needle)
}
Write-Output ('PASS: dynamic form/control families expanded from catalogues; '+$checks+' total mutation checks')
