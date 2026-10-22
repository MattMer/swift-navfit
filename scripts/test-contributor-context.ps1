# Test-only helper: explicitly choose the retained synthetic contributor in a
# disposable test database. Production code never calls this helper.
function SelectSyntheticContributor($application){
 if($application.Eval('DCount("*","Workspace")') -eq 0){return}
 [void]$application.Eval('SwitchContributor(CLng(DLookup("ActiveProfileID","Workspace","WorkspaceID=1")))')
 if($application.Eval('CurrentProject.AllForms("frmContributorStart").IsLoaded')){
  $application.DoCmd.Close(2,'frmContributorStart',2)
 }
}
