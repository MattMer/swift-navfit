Attribute VB_Name = "UXNavyWriterHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6922, , message
    n = n + 1
End Sub
Public Function TestUXNavyWriter() As String
    Dim before As Long, body As String, row As Long, id As Long, unassigned As Long, priorIT As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC CATALOG", "Tester"
    unassigned = DCount("*", "NarrativeSnippet", "SnippetRate([Attribution])='UNSPECIFIED'")
    priorIT = DCount("*", "NarrativeSnippet", "SnippetRate([Attribution])='IT'")
    DoCmd.OpenForm "frmUXNavyWriter"
    Check Forms!frmUXNavyWriter!lstCatalog.ListCount = NavyWriterCount() And NavyWriterCount() > 80, "Expanded catalog entries missing"
    Check Not UXImportNavyWriter(), "Import did not require selection"
    Forms!frmUXNavyWriter!lstCatalog = 0
    Check UXSelectNavyWriter(), "Example preview failed"
    Check InStr(Forms!frmUXNavyWriter!txtSource, "navywriter.com") > 0, "Source URL absent"
    body = Forms!frmUXNavyWriter!txtBody: before = DCount("*", "NarrativeSnippet")
    Check UXImportNavyWriter(), "Reviewed example import failed"
    Check DCount("*", "NarrativeSnippet") = before + 1, "Example not saved"
    Check UXImportNavyWriter() And DCount("*", "NarrativeSnippet") = before + 1, "Duplicate example imported twice"
    Forms!frmUXNavyWriter!txtBody = "Original synthetic adapted accomplishment"
    Check UXImportNavyWriter() And DCount("*", "NarrativeSnippet") = before + 2, "Adapted example could not be saved"
    Forms!frmUXNavyWriter!txtSearch = "FCA": UXLoadNavyWriter
    Check Forms!frmUXNavyWriter!lstCatalog.ListCount = 2, "Catalog search failed"
    Forms!frmUXNavyWriter!txtSearch = "": Forms!frmUXNavyWriter!cboRate = "IT": UXLoadNavyWriter
    Check Forms!frmUXNavyWriter!lstCatalog.ListCount = 2, "IT filter included unrelated rates"
    For row = 0 To Forms!frmUXNavyWriter!lstCatalog.ListCount - 1
        id = Forms!frmUXNavyWriter!lstCatalog.ItemData(row)
        Check NavyWriterMatches(id, "IT", "") And Not NavyWriterMatches(id, "BM", ""), "Rate boundary incorrect"
        If NavyWriterValue(id, 4) = "Source collection" Then
            Forms!frmUXNavyWriter!lstCatalog = id: UXSelectNavyWriter
            Check Not Forms!frmUXNavyWriter!cmdImport.Enabled And Forms!frmUXNavyWriter!txtBody.Locked And Not UXImportNavyWriter(), "Collection description could be imported as narrative"
        End If
    Next
    Forms!frmUXNavyWriter!txtSearch = "naval messages": UXLoadNavyWriter
    Check Forms!frmUXNavyWriter!lstCatalog.ListCount = 1, "Rate/search intersection incorrect"
    Forms!frmUXNavyWriter!lstCatalog = Forms!frmUXNavyWriter!lstCatalog.ItemData(0): UXSelectNavyWriter
    Check Forms!frmUXNavyWriter!cmdImport.Enabled And Not Forms!frmUXNavyWriter!txtBody.Locked, "Excerpt is not adaptable"
    Check UXImportNavyWriter(), "Rate-specific example import failed"
    Check InStr(DLookup("Attribution", "NarrativeSnippet", "SnippetTitle='NavyWriter - IT naval messages'"), "Rate: IT") > 0, "Rate/source attribution lost"
    Forms!frmUXNavyWriter!cboRate = "CWT": UXLoadNavyWriter
    Check IsNull(Forms!frmUXNavyWriter!lstCatalog) And Len(Forms!frmUXNavyWriter!txtBody) = 0 And Not Forms!frmUXNavyWriter!cmdImport.Enabled, "Filtered selection retained stale import"
    Forms!frmUXNavyWriter!txtSearch = "": Forms!frmUXNavyWriter!cboRate = "GENERAL": UXLoadNavyWriter
    Check Forms!frmUXNavyWriter!lstCatalog.ListCount = 5, "General examples mixed with rated examples"
    DoCmd.OpenForm "frmSnippets"
    Forms!frmSnippets!cboRateFilter = "IT": LoadSnippets
    Check Forms!frmSnippets!lstSnippets.ListCount = priorIT + 2, "Saved snippet rate filter failed"
    Forms!frmSnippets!txtSearch = "does-not-exist": LoadSnippets
    Check Forms!frmSnippets!lstSnippets.ListCount = 0, "Saved snippet search/filter intersection failed"
    Forms!frmSnippets!txtSearch = "": Forms!frmSnippets!cboRateFilter = "UNSPECIFIED": LoadSnippets
    Check Forms!frmSnippets!lstSnippets.ListCount = unassigned, "Imported tagged excerpts incorrectly treated as legacy"
    Check SnippetRate("Old untagged attribution") = "UNSPECIFIED", "Legacy snippet rate invented"
    Check SnippetRate(SnippetRateAttribution("Source [Rate: IT]", "CWT")) = "CWT", "Rate assignment did not replace existing tag"
    NewSnippetUI: Forms!frmSnippets!txtTitle = "Manual rate example": Forms!frmSnippets!txtBody = "Synthetic original text": Forms!frmSnippets!txtAttribution = "Command author": Forms!frmSnippets!cboSnippetRate = "IT"
    Check SaveSnippetUI(), "Manual snippet rate save failed"
    Check SnippetRate(DLookup("Attribution", "NarrativeSnippet", "SnippetTitle='Manual rate example'")) = "IT", "Manual rate not retained"
    Forms!frmSnippets!cboRateFilter = "IT": LoadSnippets
    Check Forms!frmSnippets!lstSnippets.ListCount = priorIT + 3, "Manual rate snippet absent"
    DoCmd.Close acForm, "frmSnippets", acSaveNo
    Check UXCloseResource("frmUXNavyWriter"), "Catalog return failed"
    Check CurrentProject.AllForms("frmUXResources").IsLoaded, "Catalog did not return to Resources"
    TestUXNavyWriter = "PASS: " & n & " catalog preview/import assertions": Exit Function
Failed:
    TestUXNavyWriter = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
