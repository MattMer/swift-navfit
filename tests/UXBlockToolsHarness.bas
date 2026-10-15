Attribute VB_Name = "UXBlockToolsHarness"
Option Explicit
' Compatibility entry for the former staged-block suite. Direct recursive edits
' replaced that workflow; live comparison coverage resides in its own suite.
Public Function TestUXBlockTools() As String
    TestUXBlockTools = TestRecursiveDefaults()
End Function
