"""Generate verified source excerpts for the additional readiness catalogue."""
from pathlib import Path
import re
ROOT = Path(__file__).resolve().parents[1]
text = (ROOT/'planning/evidence/BUPERSINST 1610.10.txt').read_text(encoding='utf-8-sig')
def norm(s):
    return re.sub(r'\s+', ' ', s.replace('“','"').replace('”','"')).strip()
rows = [
 ('VR-DUTIES',40,'1-11','paragraph 1-2, blocks 28-29; also page 1-12','Enter a brief statement of command employment and significant command'),
 ('VR-RETENTION',51,'1-22','paragraph 1-2, EVAL block 47','Select or place an "X" in Not Recommended or Recommended.'),
 ('VR-RATER',44,'1-15','paragraph 1-2, EVAL block 42; senior rater page 1-23','Type the last name, initials, and grade or rate of the rater and obtain'),
 ('VR-OPS-COUNSEL',89,'5-2','paragraph 5-6, FITREP/CHIEFEVAL counseling blocks 30-32','BLOCKS 30-32 Leave blank.'),
 ('VR-PFA',34,'1-5','paragraph 1-2, block 20; continues pages 1-6 to 1-7; Concurrent exception 4-12, page 4-5','Ensure data entered here matches data entered into the official Navy Physical'),
 ('VR-RS-NAME',38,'1-9','paragraph 1-2, block 22','Place a comma after the last name'),
 ('VR-RS-DESIG',39,'1-10','paragraph 1-2, blocks 23-24','Enter reporting senior\'s four-digit officer designator.'),
 ('VR-CONTACT',52,'1-23','paragraph 1-2, EVAL48/FITREP44/CHIEFEVAL51; OMPF/member-copy distinction','reporting senior\'s address and phone number for'),
 ('VR-PERIODIC',17,'12','enclosure (1), table 1 through 2026; table 1A page 12a/PDF18 from 2027; exceptions enclosure (2) 3-4/3-5','A periodic report may be omitted if a Service member has received a graded Regular report within the'),
 ('VR-CONTINUITY',70,'3-1','paragraph 3-2, Regular report continuity','Begin each Regular report on the day following the ending date of the previous'),
]
lines=['Attribute VB_Name = "modValidationReferences"','Option Explicit',"' Excerpts below are generated only after matching the fingerprinted PDF extraction.",'Public Function ExpandedPolicyReference(ByVal code As String) As String',' Dim page As String, locator As String, excerpt As String, pdfPage As Long',' Select Case code']
q=lambda s:'"'+s.replace('"','""')+'"'
for code,pdf,page,loc,excerpt in rows:
    body=re.search(r'(?s)--- PDF PAGE '+str(pdf)+r' ---\s*(.*?)(?=--- PDF PAGE|$)',text).group(1)
    if norm(excerpt) not in norm(body): raise ValueError((code,pdf,excerpt))
    lines.append(f'  Case {q(code)}: pdfPage = {pdf}: page = {q(page)}: locator = {q(loc)}: excerpt = {q(excerpt)}')
lines += ['  Case Else','   ExpandedPolicyReference = NativeValidationReference(code): Exit Function',' End Select',' ExpandedPolicyReference = "BUPERSINST 1610.10H CH-2 (26 May 2026 compilation)" & vbCrLf & "Printed page " & page & "; " & locator & " (PDF page " & pdfPage & ")." & vbCrLf & "Excerpt: " & Chr$(34) & excerpt & Chr$(34) & vbCrLf & "Source SHA-256: " & NAVY_RULE_SOURCE_HASH','End Function',
'','Public Function NativeValidationReference(ByVal code As String) As String',' Dim method As String',' Select Case code',
'  Case "NV-SSN": method = "ReportCheckSocSecNum / ReportCheckRSSsn"',
'  Case "NV-RATER": method = "ReportCheckRater"',
'  Case "NV-COUNSEL": method = "ReportCheckCounseler"',
'  Case "NV-PFA": method = "ReportCheckPhysicalReadinessDT"',
'  Case "NV-CONTACT": method = "ReportCheckRSAddress"',
'  Case "NV-RATE": method = "ReportCheckRate"',
'  Case "NV-DESIG": method = "ReportCheckDesig"',
'  Case "NV-SUBCAT": method = "ReportCheckBilletSubcat"',
'  Case "NV-RRS": method = "ReportCheckRrsInfo"',
'  Case "NV-GROUP": method = "FolderCheckPeriodic / FolderCheckSpecial"',
'  Case "NV-RS-GROUP": method = "FolderCheckRSGradeConsistency / FolderCheckRSDesigConsistency / FolderCheckRSTitleConsistency"',
'  Case Else: NativeValidationReference = "A source excerpt has not been verified for this rule.": Exit Function',' End Select',
' NativeValidationReference = "NAVFIT98A compatibility review, not an instruction-based policy failure." & vbCrLf & "Version 2.2.0.332, ValidationModule." & method & vbCrLf & "REA evidence ev_8a845642608c3e4eda53106f794403a54b5904eaecf5004a18bdbc639e913c39; executable SHA-256 5366f1c1bb147f1706094c5e3e1d62981f34fc6cb301e4006a892acfbfab58cf."','End Function']
(ROOT/'src/modValidationReferences.bas').write_text('\n'.join(lines)+'\n',encoding='utf-8')
print(f'PASS: {len(rows)} additional source excerpts match their PDF pages')
