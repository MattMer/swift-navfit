Attribute VB_Name = "modEvalmanLinks"
Option Compare Database
Option Explicit
' Where EVALMAN explains a report field or a finding. Sections are found by the start
' of their title, so regenerating modEvalmanText keeps the links; chapter 1 comes first,
' so a block always opens its paragraph 1-2 instructions.

' The first section whose title starts with the given text, or -1.
Public Function EvalmanSectionByTitle(ByVal titleStart As String) As Long
    Dim i As Long
    EvalmanSectionByTitle = -1
    If Len(titleStart) = 0 Then Exit Function
    For i = 0 To EvalmanCount() - 1
        If Left$(CStr(EvalmanHeader(i, 1)), Len(titleStart)) = titleStart Then EvalmanSectionByTitle = i: Exit Function
    Next
End Function

' The start of the EVALMAN section title for a report field, or "" when no block applies.
' Block sections cover all three forms; FITREP and CHIEFEVAL numbers are given in each.
Public Function EvalmanTitleForField(ByVal key As String) As String
    Dim t As String
    Select Case key
        Case "OfficialName": t = "Block 1:"
        Case "RankRate": t = "Block 2:"
        Case "Designator": t = "Block 3:"
        Case "MemberSSN": t = "Block 4:"
        Case "ComponentActive", "ComponentTAR", "ComponentInactive", "ComponentATADSW": t = "Block 5:"
        Case "UnitUIC", "ShipStation": t = "Block 6:"
        Case "PromotionStatus": t = "Block 8:"
        Case "DateReported": t = "Block 9:"
        Case "OccasionPeriodic", "OccasionDetachment", "OccasionFrocking", "OccasionSpecial": t = "Blocks 10-13:"
        Case "PeriodStart", "PeriodEnd": t = "Blocks 14-15:"
        Case "ObservationStatus": t = "Block 16:"
        Case "ReportRegular", "ReportConcurrent", "ReportOperationalCommander": t = "Blocks 17-19:"
        Case "ReadinessSummary", "ReadinessDetail", "ReadinessDate": t = "Block 20:"
        Case "BilletSubcategory": t = "Block 21:"
        Case "ReportingSeniorName": t = "Block 22:"
        Case "ReportingSeniorGrade": t = "Block 23:"
        Case "ReportingSeniorDesignator": t = "Block 24:"
        Case "ReportingSeniorTitle": t = "Block 25:"
        Case "ReportingSeniorUIC": t = "Block 26:"
        Case "ReportingSeniorSSN": t = "Block 27:"
        Case "CommandEmployment": t = "Block 28:"
        Case "PrimaryDuty", "DutyDescription": t = "Block 29:"
        Case "CounselingDate": t = "Block 30:"
        Case "CounselorName": t = "Block 31:"
        Case "AssignmentRecommendationA", "AssignmentRecommendationB": t = "EVAL Block 41:"
        Case "RaterName", "RaterDate": t = "EVAL Block 42:"
        Case "OfficialComments": t = "EVAL Block 43:"
        Case "Qualifications": t = "EVAL Block 44:"
        Case "PromotionRecommendation": t = "EVAL Block 45:"
        Case "RetentionYes", "RetentionNo": t = "EVAL Block 47:"
        Case "ReportingSeniorAddress1", "ReportingSeniorAddress2", "ReportingSeniorCity", "ReportingSeniorState", "ReportingSeniorZip", "ReportingSeniorPhone", "ReportingSeniorDSN": t = "EVAL Block 48:"
        Case "SeniorRaterName", "SeniorRaterDate": t = "EVAL Block 49:"
        Case "StatementYes", "StatementNo": t = "EVAL Block 51:"
        Case Else
            If key Like "Trait3#" Or key Like "Trait3#Comment#" Then t = "Blocks 33-39:"
    End Select
    EvalmanTitleForField = t
End Function

Public Function EvalmanSectionForField(ByVal key As String) As Long
    EvalmanSectionForField = EvalmanSectionByTitle(EvalmanTitleForField(key))
End Function

' A finding opens the paragraph its rule cites, or else its field's block.
Public Function EvalmanSectionForFinding(ByVal code As String, ByVal key As String) As Long
    Dim t As String
    Select Case code
        Case "VR-CONTINUITY": t = "3-2."
        Case "VR-PERIODIC": t = "3-4."
        Case "VR-OPS-COUNSEL": t = "5-6."
        Case Else
            If Left$(code, 3) = "RJ-" Then t = RejectRuleSection(code)
    End Select
    If Len(t) > 0 Then EvalmanSectionForFinding = EvalmanSectionByTitle(t) Else EvalmanSectionForFinding = EvalmanSectionForField(key)
End Function

' Paragraph 1-2, for a field with no block of its own.
Public Function EvalmanBlocksSection() As Long
    EvalmanBlocksSection = EvalmanSectionByTitle("1-2.")
End Function
