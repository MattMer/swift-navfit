Attribute VB_Name = "modValidationReferences"
Option Explicit
' Excerpts below are generated only after matching the fingerprinted PDF extraction.
Public Function ExpandedPolicyReference(ByVal code As String) As String
 Dim page As String, locator As String, excerpt As String, pdfPage As Long
 Select Case code
  Case "VR-DUTIES": pdfPage = 40: page = "1-11": locator = "paragraph 1-2, blocks 28-29; also page 1-12": excerpt = "Enter a brief statement of command employment and significant command"
  Case "VR-RETENTION": pdfPage = 51: page = "1-22": locator = "paragraph 1-2, EVAL block 47": excerpt = "Select or place an ""X"" in Not Recommended or Recommended."
  Case "VR-RATER": pdfPage = 44: page = "1-15": locator = "paragraph 1-2, EVAL block 42; senior rater page 1-23": excerpt = "Type the last name, initials, and grade or rate of the rater and obtain"
  Case "VR-OPS-COUNSEL": pdfPage = 89: page = "5-2": locator = "paragraph 5-6, FITREP/CHIEFEVAL counseling blocks 30-32": excerpt = "BLOCKS 30-32 Leave blank."
  Case "VR-PFA": pdfPage = 34: page = "1-5": locator = "paragraph 1-2, block 20; continues pages 1-6 to 1-7; Concurrent exception 4-12, page 4-5": excerpt = "Ensure data entered here matches data entered into the official Navy Physical"
  Case "VR-RS-NAME": pdfPage = 38: page = "1-9": locator = "paragraph 1-2, block 22": excerpt = "Place a comma after the last name"
  Case "VR-RS-DESIG": pdfPage = 39: page = "1-10": locator = "paragraph 1-2, blocks 23-24": excerpt = "Enter reporting senior's four-digit officer designator."
  Case "VR-CONTACT": pdfPage = 52: page = "1-23": locator = "paragraph 1-2, EVAL48/FITREP44/CHIEFEVAL51; OMPF/member-copy distinction": excerpt = "reporting senior's address and phone number for"
  Case "VR-PERIODIC": pdfPage = 17: page = "12": locator = "enclosure (1), table 1 through 2026; table 1A page 12a/PDF18 from 2027; exceptions enclosure (2) 3-4/3-5": excerpt = "A periodic report may be omitted if a Service member has received a graded Regular report within the"
  Case "VR-CONTINUITY": pdfPage = 70: page = "3-1": locator = "paragraph 3-2, Regular report continuity": excerpt = "Begin each Regular report on the day following the ending date of the previous"
  Case Else
   ExpandedPolicyReference = NativeValidationReference(code): Exit Function
 End Select
 ExpandedPolicyReference = "BUPERSINST 1610.10H CH-2 (26 May 2026 compilation)" & vbCrLf & "Printed page " & page & "; " & locator & " (PDF page " & pdfPage & ")." & vbCrLf & "Excerpt: " & Chr$(34) & excerpt & Chr$(34) & vbCrLf & "Source SHA-256: " & NAVY_RULE_SOURCE_HASH
End Function

Public Function NativeValidationReference(ByVal code As String) As String
 Dim method As String
 Select Case code
  Case "NV-SSN": method = "ReportCheckSocSecNum / ReportCheckRSSsn"
  Case "NV-RATER": method = "ReportCheckRater"
  Case "NV-COUNSEL": method = "ReportCheckCounseler"
  Case "NV-PFA": method = "ReportCheckPhysicalReadinessDT"
  Case "NV-CONTACT": method = "ReportCheckRSAddress"
  Case "NV-RATE": method = "ReportCheckRate"
  Case "NV-DESIG": method = "ReportCheckDesig"
  Case "NV-SUBCAT": method = "ReportCheckBilletSubcat"
  Case "NV-RRS": method = "ReportCheckRrsInfo"
  Case "NV-GROUP": method = "FolderCheckPeriodic / FolderCheckSpecial"
  Case "NV-RS-GROUP": method = "FolderCheckRSGradeConsistency / FolderCheckRSDesigConsistency / FolderCheckRSTitleConsistency"
  Case Else: NativeValidationReference = "A source excerpt has not been verified for this rule.": Exit Function
 End Select
 NativeValidationReference = "NAVFIT98A compatibility review, not an instruction-based policy failure." & vbCrLf & "Version 2.2.0.332, ValidationModule." & method & vbCrLf & "REA evidence ev_8a845642608c3e4eda53106f794403a54b5904eaecf5004a18bdbc639e913c39; executable SHA-256 5366f1c1bb147f1706094c5e3e1d62981f34fc6cb301e4006a892acfbfab58cf."
End Function
