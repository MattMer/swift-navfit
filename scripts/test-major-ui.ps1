param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/major-ui-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null;$bounds=0
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    foreach($module in @('MajorHarness','MajorUIHarness')){$app.LoadFromText(5,$module,(Join-Path $root ('tests/'+$module+'.bas')))}
    $result=$app.Eval('TestMajor()');if(-not $result.StartsWith('PASS:')){throw $result}
    $result=$app.Eval('TestMajorUI()');if(-not $result.StartsWith('PASS:')){throw $result}
    foreach($name in @('frmHome','frmMajorWorkbench','frmPDFImport','frmBackup')){
        [void]$app.Eval('OpenScreen("'+$name+'")');$form=$app.Forms.Item($name)
        foreach($control in $form.Controls){
            if($control.Visible){
                if($control.Left -lt 0 -or $control.Top -lt 0 -or $control.Left+$control.Width -gt $form.Width -or $control.Top+$control.Height -gt $form.Section(0).Height){throw ('Control outside form: '+$name+'.'+$control.Name)}
                $bounds++
            }
        }
    }
    Write-Output $result
    Write-Output ('PASS: '+$bounds+' visible controls within new screen bounds')
}finally{if($app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
