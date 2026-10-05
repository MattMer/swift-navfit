param([string]$Executable='D:\Development\Swift NAVFIT\NAVFIT98A\NAVFIT98A.exe')
# Metadata/IL inspection only: no target application methods are invoked.
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$assembly=[Reflection.Assembly]::LoadFrom($Executable)
$opcodes=@{}
foreach($field in [Reflection.Emit.OpCodes].GetFields([Reflection.BindingFlags]'Public,Static')){
    $op=$field.GetValue($null);$opcodes[[int]$op.Value -band 65535]=$op
}
$result=[ordered]@{Executable=$Executable;SHA256=(Get-FileHash -LiteralPath $Executable).Hash;Version=$assembly.GetName().Version.ToString();AssemblyAttributes=@($assembly.GetCustomAttributesData()|Where-Object{$_.AttributeType.Name -match 'ComVisible|Guid'}|ForEach-Object{$_.ToString()});Methods=@()}
foreach($typeName in @('PrintModule','SummaryControllerModule','CalculationFunctions','ProceduresModule','NF_Main')){
    $type=$assembly.GetType('NAVFIT98A.'+$typeName)
    foreach($method in $type.GetMethods([Reflection.BindingFlags]'Public,NonPublic,Static,Instance,DeclaredOnly')){
        if($method.Name -notmatch 'Print|Prt|Summary|Export|Address'){continue}
        $body=$method.GetMethodBody();if(-not $body){continue};$bytes=$body.GetILAsByteArray();$instructions=@();$position=0
        while($position -lt $bytes.Length){
            $offset=$position;$number=[int]$bytes[$position];$position++
            if($number -eq 254){$number=65024+[int]$bytes[$position];$position++}
            $op=$opcodes[[int]$number]; if (-not $op) { throw "Unknown opcode $number at $offset in $($method.Name)" };$operand=$null;$length=0
            switch($op.OperandType.ToString()){
                'InlineNone' {}
                'ShortInlineI' {$length=1;$operand=[int]$bytes[$position]}
                'ShortInlineVar' {$length=1;$operand=[int]$bytes[$position]}
                'ShortInlineBrTarget' {$length=1}
                'InlineVar' {$length=2}
                'InlineI8' {$length=8}
                'InlineR' {$length=8}
                'InlineSwitch' {$length=4+4*[BitConverter]::ToInt32($bytes,$position)}
                default {
                    $length=4;$token=[BitConverter]::ToInt32($bytes,$position)
                    switch($op.OperandType.ToString()){
                        'InlineString' {$operand=$method.Module.ResolveString($token)}
                        'InlineMethod' {$operand=$method.Module.ResolveMethod($token).ToString()}
                        'InlineField' {$operand=$method.Module.ResolveField($token).ToString()}
                        'InlineI' {$operand=$token}
                    }
                }
            }
            $position+=$length
            $instructions+=@{Offset=$offset;Op=$op.Name;Operand=$operand}
        }
        $result.Methods+=@{Type=$type.FullName;Method=$method.Name;Instructions=$instructions}
    }
}
$path=Join-Path $root 'planning/evidence/navfit/native-print-export.json'
$result|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $path -Encoding UTF8
Write-Output ('Inspected '+$result.Methods.Count+' methods without executing target code: '+$path)
