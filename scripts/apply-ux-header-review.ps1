if(-not $script:swiftManagedBuild){foreach($m in @('modUXHeaderReview','modUXReview','modUXSurface')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor')
[void](Control $f 104 'uxReviewSetup' 5900 2650 1800 420 'Review setup' '=UXEditorReview("Routes")')
[void](Control $f 104 'uxReviewAccept' 7800 2650 1800 420 'Accept step' '=UXHeaderReview("Accept")')
[void](Control $f 104 'uxReviewReturn' 9700 2650 1800 420 'Return report' '=UXHeaderReview("Return")')
[void](Control $f 104 'uxReviewApproval' 11600 2650 2200 420 'Approve / reopen' '=UXHeaderReview("Approval")')
foreach($name in @('uxReviewAccept','uxReviewReturn','uxReviewApproval')){$f.Controls($name).ControlTipText='Applies to saved values. Requires the assigned contributor and a current route; local approval requires the final reviewer. Unavailable actions explain the reason above the fields.'}
SaveScreen $f 'frmEditor'
