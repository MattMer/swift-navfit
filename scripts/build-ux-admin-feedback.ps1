param([Parameter(Mandatory=$true)][string]$TemplatePath,[Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$OutputPath=[IO.Path]::GetFullPath($OutputPath)
if(Test-Path -LiteralPath $OutputPath){throw 'Refusing overwrite'}
Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath
$app=New-Object -ComObject Access.Application
try {
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($OutputPath,$true)
    foreach($openForm in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$openForm,2)}
    $source=Get-Content (Join-Path $root 'scripts/build.ps1') -Raw
    Invoke-Expression ([regex]::Match($source,'(?s)    function Control\(.*?(?=    NewScreen ''frmSetup'')').Value)
    . (Join-Path $PSScriptRoot 'apply-ux-admin-feedback.ps1')
    foreach($formName in @('frmEditor','frmUXComparison')){
        $app.DoCmd.OpenForm($formName,1);$form=$app.Forms.Item($formName)
        $form.HasModule=$true;$form.OnUnload='[Event Procedure]';$module=$form.Module
        $statement=if($formName -eq 'frmEditor'){'    If Not UXEditorUnload() Then Cancel = True'}else{'    UXGridUnload Me'}
        if($module.Lines(1,$module.CountOfLines) -notmatch [regex]::Escape($statement.Trim())){
            $line=$module.CreateEventProc('Unload','Form');$module.InsertLines($line+1,$statement)
        }
        $app.DoCmd.Close(2,$formName,1)
    }
    $app.LoadFromText(4,'AutoExec',(Join-Path $root 'src/AutoExec.txt'))
    if(-not @($app.CurrentProject.AllMacros|ForEach-Object {$_.Name}).Contains('AutoExec')){throw 'Missing startup macro'}
    $app.DoCmd.RunCommand(126);$app.CloseCurrentDatabase()
    # Locally generated, source-reviewed output; keep the supplied template's origin marker intact.
    Unblock-File -LiteralPath $OutputPath
    Write-Output ('Compiled admin feedback build: '+$OutputPath)
} finally { if($app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)} }


