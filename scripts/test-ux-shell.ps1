param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/ux-shell-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null;$checks=0
try{
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.DoCmd.RunCommand(126)
    foreach($name in @('frmHome','frmLibrary','frmEditor','frmUXGroups','frmMyWork','frmUXResources','frmUXSettings','frmMajorWorkbench','frmUXBulk','frmUXComparison','frmUXSeniors','frmUXSeniorEditor','frmUXGroupRSCA')){
        $app.DoCmd.OpenForm($name,1);$f=$app.Forms.Item($name)
        foreach($c in $f.Controls){
            if($c.Name -match '^uxNav\d+$'){
                if($c.Left -ge 2450 -or -not $c.Visible){throw "$name navigation is unavailable"}
                $checks++
            }elseif($c.Visible -and $c.ControlType -ne 101 -and $c.Name -notin @('uxBrand','uxNavCaption')){
                if($c.Left -lt 2450 -or $c.Left+$c.Width -gt $f.Width){throw "$name/$($c.Name) exceeds its content area"}
                $checks++
            }
        }
        if($name -eq 'frmEditor' -and -not $f.Controls('cboKind').Visible){throw 'Report type selector is hidden'}
        if($name -eq 'frmLibrary' -and $f.Controls('lstFolders').Visible){throw 'Legacy folder tree still visible'}
        $app.DoCmd.Close(2,$name,2)
    }
    Write-Output "PASS: $checks native navigation/content bounds checks; compiled successfully. Visual and assistive-technology testing still required."
}finally{if($null -ne $app){$app.Quit(2)}}
