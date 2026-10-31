Attribute VB_Name = "modUXResourceTools"
Option Explicit
' Editing tools flush queued cells; saved-value export preserves pending buffers.
Public Function UXCurrentGroupTool(ByVal screen As String, ByVal action As String) As Boolean
 Dim id As Long, message As String
 On Error GoTo Failed
 id = Nz(Forms(screen)!uxGroupContext, 0): If id = 0 Then Exit Function
 If action <> "Export" Then If Not UXGridDrain() Then Exit Function
 Select Case action
  Case "Export": UXCurrentGroupTool = UXOpenNAVFITGroupExport(id)
  Case "Senior": UXCurrentGroupTool = RSOpenGroupCalculator(id)
 End Select
 Exit Function
Failed:
 'UXHasControl sets its own handler, which clears Err.
 message = Err.Description
 If UXHasControl(Forms(screen), "lblStatus") Then
  Forms(screen)!lblStatus.Caption = message
 ElseIf UXHasControl(Forms(screen), "uxStatus") Then
  Forms(screen)!uxStatus.Caption = message
 End If
End Function
Public Function UXGroupSeniorCaption(ByVal id As Long, Optional ByVal preparedPSG As Object = Nothing) As String
 Dim title As String, db As DAO.Database, r As DAO.Recordset, psg As Object
 title = Nz(MajorScalar("SELECT s.Title FROM GroupRSCABaseline AS g INNER JOIN ReportingSeniorRecord AS s ON g.SeniorID=s.SeniorID WHERE g.GroupID=" & id), "")
 If title = "" Then
  Set db = CurrentDb: Set r = db.OpenRecordset("SELECT DISTINCT r.ReportingSeniorName FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & id, dbOpenSnapshot)
  If Not r.EOF Then
   title = Nz(r!ReportingSeniorName, "Not set"): If title = "" Then title = "Not set"
   r.MoveNext: If Not r.EOF Then title = "Mixed report values"
  End If
  r.Close
 End If
 If title = "" Then title = "Not selected"
 If preparedPSG Is Nothing Then Set psg = RSGroupPSG(id) Else Set psg = preparedPSG
 UXGroupSeniorCaption = "Reporting senior: " & title & " | PSG RSCA: "
 If IsNull(psg("Value")) Then
  UXGroupSeniorCaption = UXGroupSeniorCaption & "Unavailable: " & psg("Reason")
 Else
  UXGroupSeniorCaption = UXGroupSeniorCaption & Format$(psg("Value"), "0.00") & IIf(psg("Estimated"), " (estimate)", "") & IIf(psg("Partial"), " (partial traits)", "")
 End If
End Function
Public Sub UXRefreshGroupSenior()
 If CurrentProject.AllForms("frmUXComparison").IsLoaded Then
  If UXHasControl(Forms!frmUXComparison, "uxSeniorCaption") Then Forms!frmUXComparison!uxSeniorCaption.Caption = UXGroupSeniorCaption(Nz(Forms!frmUXComparison!uxGroupContext, 0))
 End If
End Sub
' Resources opens a dedicated template view, without a summary-group workbench.
Public Function UXOpenRouteTemplates() As Boolean
 If Not UXLeaveEditor() Then Exit Function
 DoCmd.OpenForm "frmUXRouteTemplates"
 UXOpenRouteTemplates = True
End Function
Public Function UXLoadRouteTemplates() As Boolean
 With Forms!frmUXRouteTemplates
  !lstTemplates.RowSource = "SELECT TemplateID,Title FROM RouteTemplate WHERE IsArchived=False ORDER BY Title"
  !cboReviewer.RowSource = "SELECT UserProfileID,DisplayName FROM UserProfile ORDER BY DisplayName,UserProfileID"
 End With
 UXLoadRouteTemplates = True
End Function
Public Function UXShowTemplate() As Boolean
 With Forms!frmUXRouteTemplates
  !lstSteps.RowSource = "SELECT s.StepOrdinal AS [Step],u.DisplayName AS Contributor,s.Instructions FROM RouteTemplateStep AS s INNER JOIN UserProfile AS u ON s.AssigneeID=u.UserProfileID WHERE s.TemplateID=" & Nz(!lstTemplates, 0) & " ORDER BY s.StepOrdinal"
 End With
 UXShowTemplate = True
End Function
Public Function UXTemplateReviewer(ByVal clear As Boolean) As Boolean
 Dim f As Access.Form
 Set f = Forms!frmUXRouteTemplates
 If clear Then
  f!txtIDs = Null: f!txtReviewers = Null
 ElseIf Nz(f!cboReviewer, 0) > 0 Then
  f!txtIDs = Nz(f!txtIDs, "") & IIf(Len(Nz(f!txtIDs, "")) > 0, ",", "") & f!cboReviewer
  f!txtReviewers = Nz(f!txtReviewers, "") & IIf(Len(Nz(f!txtReviewers, "")) > 0, " > ", "") & f!cboReviewer.Column(1)
 End If
 UXTemplateReviewer = True
End Function
Public Function UXSaveTemplate() As Boolean
 Dim id As Long
 On Error GoTo Failed
 With Forms!frmUXRouteTemplates
  id = CreateRouteTemplate(Nz(!txtTitle, ""), Nz(!txtIDs, ""), Nz(!txtInstructions, ""))
  !lstTemplates.Requery: !lstTemplates = id: UXShowTemplate
  !txtTitle = Null: !txtIDs = Null: !txtReviewers = Null: !txtInstructions = Null
  !lblStatus.Caption = "Template saved. Use it from a report's Review section."
 End With
 UXSaveTemplate = True: Exit Function
Failed:
 Forms!frmUXRouteTemplates!lblStatus.Caption = Err.Description & " Entries retained."
End Function
Public Function UXCloseTemplates() As Boolean
 DoCmd.Close acForm, "frmUXRouteTemplates", acSaveNo
 UXCloseTemplates = UXNavigate("frmUXResources")
End Function

' Inspect the selected reusable senior record directly from its group calculator.
Public Function RSViewGroupSenior() As Boolean
 Dim id As Long
 id = Nz(Forms!frmUXGroupRSCA!cboSenior, 0)
 If id = 0 Then Forms!frmUXGroupRSCA!lblStatus.Caption = "Choose a reporting senior record first.": Exit Function
 If Not RSLeaveSeniorEditor() Then Exit Function
 DoCmd.OpenForm "frmUXSeniorEditor", , , , , , CStr(id)
 RSViewGroupSenior = True
End Function
' Navy terms that Office spelling flags (about 800: ratings and paygrades, grades,
' units and ship types, admin and evaluation terms), drafted with a fast model and reviewed.
Public Function NavyWordList() As Variant
    Dim w As String
    w = "3M AB AB1 AB2 AB3 ABC ABCM ABCS ABE ABF ABH AC AC1 AC2 AC3 ACC ACCM ACCS ACDUTRA ACU AD1 AD2 AD3 ADC ADCM ADCS ADM ADOS ADSW AE AE1 AE2 AE3 AEC AECM AECS Aerographer Aerographers AFQT AG AG1 AG2 AG3"
    w = w & " AGC AGCM AGCS AGR AH1Z AIMD airdale airwing airwings AM1 AM2 AM3 AMC AMCM AMCS AME AO AO1 AO2 AO3 AOC AOCM AOCS AOE AOR ARG ARS AS1 AS2 AS3 ASC ASCM ASCS ASR ASVAB AT1 AT2 AT3 ATC ATCM ATCS ATF ATS"
    w = w & " AWF AWO AWR AWS AWV AZ AZ1 AZ2 AZ3 AZC AZCM AZCS BCA blackshoe BM BM1 BM2 BM3 BMC BMCM BMCS BMOW BU BU1 BU2 BU3 BUC BUCM BUCS BUMED BUPERS BUPERSINST C2A CAC CAPT CASREP CDO CDR CE CE1 CE2 CE3 CEC"
    w = w & " CECM CECS CFL CG CGN CH53E CHIEFEVAL CM CM1 CM2 CM3 CMC CMCM CMCS CMDCM CMEO CNAF CNIC CNO CNPC CNRC CNRFC COMNAVAIRFOR COMNAVSURFOR COMPHIBGRU COMPHIBRON COMSUBFOR COMSUBRON CONREP Controlman"
    w = w & " Controlmen CONUS CPO CPOA CPPD Cryptologic Cryptologist Cryptologists CS CS1 CS2 CS3 CSADD CSC CSCM CSCS CSG CSOSA CTF CTG CTI CTI1 CTI2 CTI3 CTIC CTICM CTICS CTM CTM1 CTM2 CTM3 CTMC CTMCM CTMCS CTN"
    w = w & " CTN1 CTN2 CTN3 CTNC CTNCM CTNCS CTR CTR1 CTR2 CTR3 CTRC CTRCM CTRCS CTT CTT1 CTT2 CTT3 CTTC CTTCM CTTCS CTU CV CVN CVW CWO CWO2 CWO3 CWO4 CWO5 CY CYBERCOM cybersecurity cyberwarfare DAPA DC DC1 DC2"
    w = w & " DC3 DCC DCCM DCCS DCPO DD DDG deckplate deckplates DEOMI DEP deployability DESRON DFA DH DIVO DM DODI DODINST DRRS DSRV DTS E2D E6B EA1 EA18G EA2 EA3 EAC EACM EACS EAOS EAWS EEO EIWS EM1 EM2 EM3 EMC"
    w = w & " EMCM EMCS EMR EN1 EN2 EN3 ENC ENCM ENCS Engineman Enginemen ENS EO EO1 EO2 EO3 EOC EOCM EOCS EOD EODMU EOOW EPF ESB ESG ESR ESWP ESWS ET ET1 ET2 ET3 ETC ETCM ETCS ETR ETV EVAL EVALS EXW F18 F35 F35C"
    w = w & " FC FC1 FC2 FC3 FCC FCCM FCCS FEP FFG FFSC FITREP FITREPS flightdeck flightdecks FLTCM FOD FORCM FRC frocking FT1 FT2 FT3 FTC FTCM FTCS FTS FY GM GM1 GM2 GM3 GMC GMCM GMCS GMT GSE GSE1 GSE2 GSE3 GSEC"
    w = w & " GSECM GSECS GSM GSM1 GSM2 GSM3 GSMC GSMCM GSMCS HM1 HM2 HM3 HMC HMCM HMCS homeport homeported homeporting homeports Hospitalman Hospitalmen hotwash HSC HSL HSM HSV HT HT1 HT2 HT3 HTC HTCM HTCS IC IC1"
    w = w & " IC2 IC3 ICC ICCM ICCS IDT INDOPACOM INSURV IS1 IS2 IS3 ISC ISCM ISCS ISIC IT1 IT2 IT3 ITA ITC ITCM ITCS ITSA ITSN ITSR JO JOOD JQR ladderwell ladderwells LANTFLT LCAC LCC LCDR LCPO LCS LCU LDO LHA LHD"
    w = w & " LIMDU LN LN1 LN2 LN3 LNC LNCM LNCS LOGREQ LPD LPO LS LS1 LS2 LS3 LSC LSCM LSCS LSD LSS LT LTJG MA1 MA2 MA3 MAC MACM MACS MALS MC MC1 MC2 MC3 MCC MCCM MCCS MCM MCPO MCPON messdecks MEU MH53E MH60R"
    w = w & " MH60S MHC MILPERSMAN Mineman Minemen Missileman Missilemen MM MM1 MM2 MM3 MMA MMC MMCM MMCS MMW MN MN1 MN2 MN3 MNC MNCM MNCS MR1 MR2 MR3 MRC MRCM MRCS MSC MT MT1 MT2 MT3 MTC MTCM MTCS MU MU1 MU2 MU3"
    w = w & " MUC MUCM MUCS multimission multiplatform mustering MV22 MWR NAC NAVADMIN NAVAIR NAVCENT NAVEDTRA NAVFAC NAVFIT NAVFIT98A NAVMED NAVMEDINST NAVPERS NAVPERSCOM NAVRES NAVRESFOR NAVSAFECEN NAVSEA NAVSTA"
    w = w & " NAVSUP NAVWAR NCIS NCOIC NCTAMS ND NDAA NEC NECC NETC netcentric NFO NIOC NJP NKO NMCB NMCI NOB nonjudicial nonpunitive nonrated NPC NRD NROTC NSIPS NSW NSWG OCONUS OCS OMPF OOD OPNAV OPNAVINST OPREP"
    w = w & " OPSEC OPTEMPO Opticalman Opticalmen OS1 OS2 OS3 OSC OSCM OSCS P3C P8A PACFLT PACT passdown passdowns PC PCS PCU Personnelman Personnelmen PERSTEMPO PERSUPP PFA pipedown pipedowns plankholder"
    w = w & " plankholders plankowner plankowners PMS PN PO1 PO2 PO3 POOW PQS PR PR1 PR2 PR3 PRC PRCM PRCS PRD promotability PRT PS PS1 PS2 PS3 PSC PSCM PSCS PSD PSG QA QM QM1 QM2 QM3 QMC QMCM QMCS qualed quals"
    w = w & " quarterdeck quarterdecks RADM RAS RDML recommendability reenlist reenlisted reenlisting reenlistment reenlistments requal requals retainability RP RP1 RP2 RP3 RPC RPCM RPCS RRS RS RSCA SA SAPR SARC SB"
    w = w & " SCPO SCW SDO SECDEF SECNAV SECNAVINST selectee selectees SELRES SGA SGLI SH1 SH2 SH3 SHC SHCM SHCS shiphandling shipmate shipmates shipwide SK SK1 SK2 SK3 SN SNCO SO1 SO2 SO3 SOC SOCM SOCS Sonarman"
    w = w & " Sonarmen sonobuoy sonobuoys SR SRB SSBN SSGN SSK SSN Steelworker Steelworkers sternwatch sternwatches STG STG1 STG2 STG3 STGC STGCM STGCS STS STS1 STS2 STS3 STSC STSCM STSCS subspecialty SW SW1 SW2"
    w = w & " SW3 SWC SWCC SWCM SWCS SWO TAGOS TDY tiltrotor Torpedoman Torpedomen TYCOM UA UCMJ UH1Y UIC underways UNREP UNREPS UPC USFF USN USNA USNR UT UT1 UT2 UT3 UTC UTCM UTCS VADM VAQ VAW VERTREP VFA VP VQ VR"
    w = w & " VRC VUP wardroom warfighter warfighters warfighting wargaming watchbill watchbills watchstander watchstanders watchstanding watchstation watchstations weatherdeck weatherdecks WO1 WO2 WO3 WO4 WO5"
    w = w & " workcenter workcenters XO YN YN1 YN2 YN3 YNC YNCM YNCS YNS"
    NavyWordList = Split(w, " ")
End Function
' Save the list as an Office custom dictionary (UTF-16 with byte-order mark, one word per
' line). The user adds it under File > Options > Proofing; Swift never changes Office settings.
Public Function SaveNavyWordList(ByVal path As String) As Long
    Dim stream As Object, words As Variant
    words = NavyWordList()
    Set stream = CreateObject("ADODB.Stream"): stream.Type = 2: stream.Charset = "unicode": stream.Open
    stream.WriteText Join(words, vbCrLf) & vbCrLf
    stream.SaveToFile path, 2: stream.Close
    SaveNavyWordList = UBound(words) + 1
End Function
Public Function UXSaveNavyWords() As Boolean
    Dim picker As Object, path As String, count As Long
    On Error GoTo Failed
    Set picker = Application.FileDialog(4): picker.Title = "Choose a folder for the Navy word list"
    If picker.Show <> -1 Then Exit Function
    path = picker.SelectedItems(1) & "\Swift NAVFIT Navy words.dic"
    If Len(Dir$(path)) > 0 Then If MsgBox("Replace " & path & "?", vbYesNo + vbQuestion, "Navy word list") <> vbYes Then Exit Function
    count = SaveNavyWordList(path)
    MsgBox count & " Navy terms saved to " & path & "." & vbCrLf & vbCrLf & "To use them in spell check, choose File > Options > Proofing > Custom Dictionaries > Add in Access and pick this file. Swift does not change Office settings.", vbInformation, "Navy word list"
    UXSaveNavyWords = True
    Exit Function
Failed:
    MsgBox "The Navy word list could not be saved: " & Err.Description, vbExclamation, "Navy word list"
End Function
