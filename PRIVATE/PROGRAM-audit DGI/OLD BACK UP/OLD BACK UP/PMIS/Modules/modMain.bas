Attribute VB_Name = "modPMIOSMain"

Option Explicit
Private Const CB_LIMITTEXT = &H141
Private Const CB_FINDSTRING = &H14C
Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_ERR = (-1)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_FIRST = &H1000
'*******************************************************************************************
' FORM VARIABLE
Public frmMasterFile_Parts                             As New frmPMISMaster_Parts
Public frmMasterFile_Accessories                       As New frmPMISMaster_Parts
Public frmMasterFile_Material                          As New frmPMISMaster_Parts

Public frmMasterFile_Counter_Parts                     As New frmPMISMaster_Counter
Public frmMasterFile_Counter_Accessories               As New frmPMISMaster_Counter
Public frmMasterFile_Counter_Materials                 As New frmPMISMaster_Counter


Public frmPMIS_CounterInquiry_Parts                    As New frmPMISInquiry_CounterInquiry
Public frmPMIS_CounterInquiry_Accessories              As New frmPMISInquiry_CounterInquiry
Public frmPMIS_CounterInquiry_Materials                As New frmPMISInquiry_CounterInquiry

Public frmPMISInquiry_CheckPrevBal_Parts               As New frmPMISInquiry_CheckPrevBal
Public frmPMISInquiry_CheckPrevBal_Accessories         As New frmPMISInquiry_CheckPrevBal
Public frmPMISInquiry_CheckPrevBal_Materials           As New frmPMISInquiry_CheckPrevBal

Public frmPMISReports_Location_Parts                   As New frmPMISReports_Location
Public frmPMISReports_Location_Accessories             As New frmPMISReports_Location
Public frmPMISReports_Location_Materials               As New frmPMISReports_Location

Public frmPMISTrans_InventoryAdjustment_Parts          As New frmPMISTrans_InventoryAdjustment
Public frmPMISTrans_InventoryAdjustment_Accessories    As New frmPMISTrans_InventoryAdjustment
Public frmPMISTrans_InventoryAdjustment_Materials      As New frmPMISTrans_InventoryAdjustment


' END VARIABLE
'*******************************************************************************************
'Use this flag to disAble button in commandbars
Public FLAG                                            As Integer

Dim Pcnt                                               As Integer
Dim MCNT                                               As Integer
Dim ACNT                                               As Integer
Dim KCNT                                               As Integer

Dim DISCTOTAL                                          As Double
Public C_TYPE, DESC_TYPE                               As String
Attribute DESC_TYPE.VB_VarUserMemId = 1073741848
Public TOTJOBAMT                                       As Double
Attribute TOTJOBAMT.VB_VarUserMemId = 1073741850
Public TOTJOBDISC                                      As Double
Attribute TOTJOBDISC.VB_VarUserMemId = 1073741851
Public TOTJOBDISCVAL                                   As Double
Attribute TOTJOBDISCVAL.VB_VarUserMemId = 1073741852
Public TOTJOBTAX                                       As Double
Attribute TOTJOBTAX.VB_VarUserMemId = 1073741853
Public TOTPARTSAMT                                     As Double
Attribute TOTPARTSAMT.VB_VarUserMemId = 1073741854
Public TOTPARTSDISC                                    As Double
Attribute TOTPARTSDISC.VB_VarUserMemId = 1073741855
Public TOTPARTSDISCVAL                                 As Double
Attribute TOTPARTSDISCVAL.VB_VarUserMemId = 1073741856
Public TOTPARTSTAX                                     As Double
Attribute TOTPARTSTAX.VB_VarUserMemId = 1073741857
Public TOTMATAMT                                       As Double
Attribute TOTMATAMT.VB_VarUserMemId = 1073741858
Public TOTMATDISC                                      As Double
Attribute TOTMATDISC.VB_VarUserMemId = 1073741859
Public TOTMATDISCVAL                                   As Double
Attribute TOTMATDISCVAL.VB_VarUserMemId = 1073741860
Public TOTMATTAX                                       As Double
Attribute TOTMATTAX.VB_VarUserMemId = 1073741861
Public TOTACCAMT                                       As Double
Attribute TOTACCAMT.VB_VarUserMemId = 1073741862
Public TOTACCDISC                                      As Double
Attribute TOTACCDISC.VB_VarUserMemId = 1073741863
Public TOTACCDISCVAL                                   As Double
Attribute TOTACCDISCVAL.VB_VarUserMemId = 1073741864
Public TOTACCTAX                                       As Double
Attribute TOTACCTAX.VB_VarUserMemId = 1073741865


Dim JOBTOTAL                                           As Double
Attribute JOBTOTAL.VB_VarUserMemId = 1073741866
Dim JOBCOMTOTAL                                        As Double
Attribute JOBCOMTOTAL.VB_VarUserMemId = 1073741867
Dim JOBSALESTOTAL                                      As Double
Attribute JOBSALESTOTAL.VB_VarUserMemId = 1073741868
Dim JOBWARTOTAL                                        As Double
Attribute JOBWARTOTAL.VB_VarUserMemId = 1073741869
Dim JOBDISCTOTAL                                       As Double
Attribute JOBDISCTOTAL.VB_VarUserMemId = 1073741870
Dim JOBVATTOTAL                                        As Double
Attribute JOBVATTOTAL.VB_VarUserMemId = 1073741871

Dim PARTSTOTAL                                         As Double
Attribute PARTSTOTAL.VB_VarUserMemId = 1073741872
Dim PARTSCOMTOTAL                                      As Double
Attribute PARTSCOMTOTAL.VB_VarUserMemId = 1073741873
Dim PARTSSALESTOTAL                                    As Double
Attribute PARTSSALESTOTAL.VB_VarUserMemId = 1073741874
Dim PARTSWARTOTAL                                      As Double
Attribute PARTSWARTOTAL.VB_VarUserMemId = 1073741875
Dim PARTSDISCTOTAL                                     As Double
Attribute PARTSDISCTOTAL.VB_VarUserMemId = 1073741876
Dim PARTSVATTOTAL                                      As Double
Attribute PARTSVATTOTAL.VB_VarUserMemId = 1073741877

Dim MATTOTAL                                           As Double
Attribute MATTOTAL.VB_VarUserMemId = 1073741878
Dim MATCOMTOTAL                                        As Double
Attribute MATCOMTOTAL.VB_VarUserMemId = 1073741879
Dim MATSALESTOTAL                                      As Double
Attribute MATSALESTOTAL.VB_VarUserMemId = 1073741880
Dim MATWARTOTAL                                        As Double
Attribute MATWARTOTAL.VB_VarUserMemId = 1073741881
Dim MATDISCTOTAL                                       As Double
Attribute MATDISCTOTAL.VB_VarUserMemId = 1073741882
Dim MATVATTOTAL                                        As Double
Attribute MATVATTOTAL.VB_VarUserMemId = 1073741883

Dim ACCTOTAL                                           As Double
Attribute ACCTOTAL.VB_VarUserMemId = 1073741884
Dim ACCCOMTOTAL                                        As Double
Attribute ACCCOMTOTAL.VB_VarUserMemId = 1073741885
Dim ACCSALESTOTAL                                      As Double
Attribute ACCSALESTOTAL.VB_VarUserMemId = 1073741886
Dim ACCWARTOTAL                                        As Double
Attribute ACCWARTOTAL.VB_VarUserMemId = 1073741887
Dim ACCDISCTOTAL                                       As Double
Attribute ACCDISCTOTAL.VB_VarUserMemId = 1073741888
Dim ACCVATTOTAL                                        As Double
Attribute ACCVATTOTAL.VB_VarUserMemId = 1073741889

Dim COMTOTAL                                           As Double
Attribute COMTOTAL.VB_VarUserMemId = 1073741890
Dim SALESTOTAL                                         As Double
Attribute SALESTOTAL.VB_VarUserMemId = 1073741891
Dim WARTOTAL                                           As Double
Attribute WARTOTAL.VB_VarUserMemId = 1073741892
Dim INSTOTAL                                           As Double
Attribute INSTOTAL.VB_VarUserMemId = 1073741893

Dim VATTOTAL                                           As Double
Attribute VATTOTAL.VB_VarUserMemId = 1073741894
Dim ROTOTAL                                            As Double
Attribute ROTOTAL.VB_VarUserMemId = 1073741895

Dim PREVRONUMBER                                       As String
Attribute PREVRONUMBER.VB_VarUserMemId = 1073741896
Dim RO_RIV_TRANNO(100)                                 As Long
Attribute RO_RIV_TRANNO.VB_VarUserMemId = 1073741897
Dim RO_RIV_TRANNO_COUNTER                              As Integer
Attribute RO_RIV_TRANNO_COUNTER.VB_VarUserMemId = 1073741898
Dim RO_MRIS_TRANNO(100)                                As Long
Attribute RO_MRIS_TRANNO.VB_VarUserMemId = 1073741899
Dim RO_MRIS_TRANNO_COUNTER                             As Integer
Attribute RO_MRIS_TRANNO_COUNTER.VB_VarUserMemId = 1073741900

Dim RSCSMSORD_HIST                                     As ADODB.Recordset
Attribute RSCSMSORD_HIST.VB_VarUserMemId = 1073741901
Dim RSCSMSDAYTRAN                                      As ADODB.Recordset
Attribute RSCSMSDAYTRAN.VB_VarUserMemId = 1073741902
Dim RSCSMSORD_HD, RSCSMSTDAYTRAN                       As ADODB.Recordset
Attribute RSCSMSORD_HD.VB_VarUserMemId = 1073741903
Attribute RSCSMSTDAYTRAN.VB_VarUserMemId = 1073741903

Dim VPIS_NO_CHARGE_TO                                  As String
Attribute VPIS_NO_CHARGE_TO.VB_VarUserMemId = 1073741905

Public Sub Main()
    SERVERNAME = GetSetting("DMIS 2.0", "SETTINGS", "SERVERNAME")
    SQLSERVERNAME = GetSetting("DMIS 2.0", "SETTINGS", "SQLSERVERNAME")
    DATABASE = GetSetting("DMIS 2.0", "SETTINGS", "DATABASE")
    If SQLSERVERNAME = "" Or DATABASE = "" Then
        MsgBox "Application Not Yet Configured. Please Configure Server Setting From DSA.", vbCritical
        End
        Exit Sub
    End If
 

    ConnStr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & DATABASE & " " & " ;Data Source=" & SQLSERVERNAME
    DMIS_Connection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & DATABASE & " ;Data Source=" & SQLSERVERNAME
    DMIS_REPORT_Connection = "DSN=" & DATABASE & ";DSQ=" & SQLSERVERNAME
    DMIS_Audit_Connection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=DMIS_AUDIT ;Data Source=" & SQLSERVERNAME


    frmMain.Show
    frmMain.ZOrder 1
    frmSplash.Show
    frmSecurity.Show vbModal
    frmSecurity.ZOrder 1
    frmMainMenu.Show
    ReminderModule ""
End Sub

Public Function OpenSQLDb() As Boolean
    Screen.MousePointer = 11
    frmSecurity.Hide
    frmSplash.Show
    frmSplash.ZOrder 0
    frmSplash.labCon.Caption = "Connecting to SQL SERVER ... Please wait..."
    DoEvents
    ApplySecurityValidation = True
    On Error GoTo ConnErr
    Set gconDMIS = New ADODB.Connection
    gconDMIS.ConnectionString = DMIS_Connection
    frmSplash.labCon.Caption = "Connecting to PMIS Database... Please wait..."
    DoEvents
    gconDMIS.Mode = adModeReadWrite
    gconDMIS.CursorLocation = adUseClient
    gconDMIS.Open
    OpenSQLDb = True
    SetCompanyProfile

    Screen.MousePointer = 0
    frmSplash.Command1.Value = True
    Exit Function

ConnErr:
    MsgBox err.Description
    MsgBox "I can't open a connection!!! You may have to " & vbCrLf & _
           "LOG-IN again to connect to the SERVER to run this program. " & vbCrLf & _
           "If you don't have an account contact your friendly " & vbCrLf & _
           "neighborhood SysAdministrator.", _
           vbOKOnly + vbCritical, "ERROR"
    End
End Function

Public Sub SetUserSettings()
    Call SetUserPathSettings
    With frmMain
        .StatusBar1.Panels(1).Text = "User: " & LOGNAME
        .StatusBar1.Panels(2).Text = "Level: " & LOGLEVEL
        .StatusBar1.Panels(3).Text = "Date: " & Format(LOGDATE, "long date")
        .StatusBar1.Panels(4).Text = "Login Time: " & LOGTIME
        .StatusBar1.Panels(9).Text = "Server : " & SQLSERVERNAME & "-" & DATABASE
    End With

End Sub



Sub FillParts(vREP_OR As String)
    TOTPARTSAMT = 0: TOTPARTSDISC = 0: TOTPARTSDISCVAL = 0: TOTPARTSTAX = 0
    Pcnt = 0: PARTSCOMTOTAL = 0: PARTSSALESTOTAL = 0: PARTSWARTOTAL = 0
    Dim rsRo_det                                       As ADODB.Recordset
    Set rsRo_det = New ADODB.Recordset
    Set rsRo_det = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where rep_or = '" & vREP_OR & "' and livil = '2' order by LINE_NO asc")
    If Not rsRo_det.EOF And Not rsRo_det.BOF Then
        rsRo_det.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsRo_det.EOF
            Pcnt = Pcnt + 1
            If Null2String(rsRo_det!wCode) = "C" Then
                PARTSCOMTOTAL = PARTSCOMTOTAL + N2Str2Zero(rsRo_det!Det_AMT)
            ElseIf Null2String(rsRo_det!wCode) = "S" Then PARTSSALESTOTAL = PARTSSALESTOTAL + N2Str2Zero(rsRo_det!Det_AMT)
            ElseIf Null2String(rsRo_det!wCode) = "W" Then PARTSWARTOTAL = PARTSWARTOTAL + N2Str2Zero(rsRo_det!Det_AMT)
            Else
                TOTPARTSAMT = TOTPARTSAMT + N2Str2Zero(rsRo_det!Det_AMT)
                TOTPARTSDISC = TOTPARTSDISC + N2Str2Zero(rsRo_det!discount_2)
                TOTPARTSDISCVAL = TOTPARTSDISCVAL + N2Str2Zero(rsRo_det!disval)
                TOTPARTSTAX = TOTPARTSTAX + N2Str2Zero(rsRo_det!taxval)
            End If
            rsRo_det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRo_det = Nothing
    TOTPARTSAMT = Round(TOTPARTSAMT, 2): TOTPARTSDISC = Round(TOTPARTSDISC, 2): TOTPARTSDISCVAL = Round(TOTPARTSDISCVAL, 2): TOTPARTSTAX = Round(TOTPARTSTAX, 2)
End Sub

Sub FillJobs(vREP_OR As String)
    TOTJOBAMT = 0: TOTJOBDISC = 0: TOTJOBDISCVAL = 0: TOTJOBTAX = 0
    KCNT = 0: JOBCOMTOTAL = 0: JOBSALESTOTAL = 0: JOBWARTOTAL = 0
    Dim rsRo_det                                       As ADODB.Recordset

    Set rsRo_det = New ADODB.Recordset
    Set rsRo_det = gconDMIS.Execute("select discount_2,det_amt,wcode,disval,taxval from CSMS_RO_Det where rep_or = '" & vREP_OR & "' and livil = '1' order by LINE_NO asc")
    If Not rsRo_det.EOF And Not rsRo_det.BOF Then
        Screen.MousePointer = 11
        rsRo_det.MoveFirst
        Do While Not rsRo_det.EOF
            KCNT = KCNT + 1
            If Null2String(rsRo_det!wCode) = "C" Then
                JOBCOMTOTAL = JOBCOMTOTAL + N2Str2Zero(rsRo_det!Det_AMT)
            ElseIf Null2String(rsRo_det!wCode) = "S" Then JOBSALESTOTAL = JOBSALESTOTAL + N2Str2Zero(rsRo_det!Det_AMT)
            ElseIf Null2String(rsRo_det!wCode) = "W" Then JOBWARTOTAL = JOBWARTOTAL + N2Str2Zero(rsRo_det!Det_AMT)
            Else
                TOTJOBAMT = TOTJOBAMT + N2Str2Zero(rsRo_det!Det_AMT)
                TOTJOBDISC = TOTJOBDISC + N2Str2Zero(rsRo_det!discount_2)
                TOTJOBDISCVAL = TOTJOBDISCVAL + N2Str2Zero(rsRo_det!disval)
                TOTJOBTAX = TOTJOBTAX + N2Str2Zero(rsRo_det!taxval)
            End If
            rsRo_det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRo_det = Nothing
    TOTJOBAMT = Round(TOTJOBAMT, 2): TOTJOBDISC = Round(TOTJOBDISC, 2): TOTJOBDISCVAL = Round(TOTJOBDISCVAL, 2): TOTJOBTAX = Round(TOTJOBTAX, 2)
End Sub

Sub FillMaterials(vREP_OR As String)
    TOTMATAMT = 0: TOTMATDISC = 0: TOTMATDISCVAL = 0: TOTMATTAX = 0
    MCNT = 0: MATCOMTOTAL = 0: MATSALESTOTAL = 0: MATWARTOTAL = 0
    Dim rsRo_det                                       As ADODB.Recordset

    Set rsRo_det = New ADODB.Recordset
    Set rsRo_det = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where rep_or = '" & vREP_OR & "' and livil = '3' order by LINE_NO asc")
    If Not rsRo_det.EOF And Not rsRo_det.BOF Then
        Screen.MousePointer = 11
        rsRo_det.MoveFirst
        Do While Not rsRo_det.EOF
            MCNT = MCNT + 1
            If Null2String(rsRo_det!wCode) = "C" Then
                MATCOMTOTAL = MATCOMTOTAL + N2Str2Zero(rsRo_det!Det_AMT)
            ElseIf Null2String(rsRo_det!wCode) = "S" Then MATSALESTOTAL = MATSALESTOTAL + N2Str2Zero(rsRo_det!Det_AMT)
            ElseIf Null2String(rsRo_det!wCode) = "W" Then MATWARTOTAL = MATWARTOTAL + N2Str2Zero(rsRo_det!Det_AMT)
            Else
                TOTMATAMT = TOTMATAMT + N2Str2Zero(rsRo_det!Det_AMT)
                TOTMATDISC = TOTMATDISC + N2Str2Zero(rsRo_det!discount_2)
                TOTMATDISCVAL = TOTMATDISCVAL + N2Str2Zero(rsRo_det!disval)
                TOTMATTAX = TOTMATTAX + N2Str2Zero(rsRo_det!taxval)
            End If
            rsRo_det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRo_det = Nothing
    TOTMATAMT = Round(TOTMATAMT, 2): TOTMATDISC = Round(TOTMATDISC, 2): TOTMATDISCVAL = Round(TOTMATDISCVAL, 2): TOTMATTAX = Round(TOTMATTAX, 2)
End Sub

Sub FillAccessories(vREP_OR As String)
    TOTACCAMT = 0: TOTACCDISC = 0: TOTACCDISCVAL = 0: TOTACCTAX = 0
    ACNT = 0: ACCCOMTOTAL = 0: ACCSALESTOTAL = 0: ACCWARTOTAL = 0
    Dim rsRo_det                                       As ADODB.Recordset

    Set rsRo_det = New ADODB.Recordset
    Set rsRo_det = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where rep_or = '" & vREP_OR & "' and livil = '4' order by LINE_NO asc")
    If Not rsRo_det.EOF And Not rsRo_det.BOF Then
        Screen.MousePointer = 11
        rsRo_det.MoveFirst
        Do While Not rsRo_det.EOF
            ACNT = ACNT + 1
            If Null2String(rsRo_det!wCode) = "C" Then
                ACCCOMTOTAL = ACCCOMTOTAL + N2Str2Zero(rsRo_det!Det_AMT)
            ElseIf Null2String(rsRo_det!wCode) = "S" Then ACCSALESTOTAL = ACCSALESTOTAL + N2Str2Zero(rsRo_det!Det_AMT)
            ElseIf Null2String(rsRo_det!wCode) = "W" Then ACCWARTOTAL = ACCWARTOTAL + N2Str2Zero(rsRo_det!Det_AMT)
            Else
                TOTACCAMT = TOTACCAMT + N2Str2Zero(rsRo_det!Det_AMT)
                TOTACCDISC = TOTACCDISC + N2Str2Zero(rsRo_det!discount_2)
                TOTACCDISCVAL = TOTACCDISCVAL + N2Str2Zero(rsRo_det!disval)
                TOTACCTAX = TOTACCTAX + N2Str2Zero(rsRo_det!taxval)
            End If
            rsRo_det.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRo_det = Nothing
    TOTACCAMT = Round(TOTACCAMT, 2): TOTACCDISC = Round(TOTACCDISC, 2): TOTACCDISCVAL = Round(TOTACCDISCVAL, 2): TOTACCTAX = Round(TOTACCTAX, 2)
End Sub

Sub UpdateParticipation(vREP_OR As String)
    '    Screen.MousePointer = 11
    Dim rsCSMS_REPOR                                   As ADODB.Recordset

    Dim vPartLabor                                     As Double
    Dim vPartParts                                     As Double
    Dim vPartMaterials                                 As Double
    Dim vPartAccessories                               As Double

    Dim vPartTotal                                     As Double

    Set rsCSMS_REPOR = New ADODB.Recordset
    Set rsCSMS_REPOR = gconDMIS.Execute("Select * from CSMS_Repor Where Rep_Or = " & N2Str2Null(vREP_OR))
    If Not rsCSMS_REPOR.EOF And Not rsCSMS_REPOR.BOF Then
        vPartLabor = N2Str2Zero(rsCSMS_REPOR!PartLabor)
        vPartParts = N2Str2Zero(rsCSMS_REPOR!PartParts)
        vPartMaterials = N2Str2Zero(rsCSMS_REPOR!PartMaterials)
        vPartAccessories = N2Str2Zero(rsCSMS_REPOR!PartAccessories)

        vPartTotal = vPartLabor + vPartParts + vPartMaterials + vPartAccessories

        FillJobs vREP_OR
        ROTOTAL = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT - (NumericVal(vPartTotal))
        SQL_STATEMENT = "update CSMS_RepOr set" & _
                      " labor = " & Round(TOTJOBAMT - TOTJOBTAX, 2) - (NumericVal(vPartLabor)) & "," & _
                      " l_amtvalue = " & Round(TOTJOBAMT, 2) - (NumericVal(vPartLabor)) & "," & _
                      " l_disc = " & Round(TOTJOBDISCVAL, 2) & "," & _
                      " l_disc2 = " & Round(TOTJOBDISC * (VAT_RATE / 100), 2) & "," & _
                      " l_taxval = " & Round(TOTJOBTAX, 2) & "," & _
                      " l_discount = " & Round(TOTJOBDISC, 2) & "," & _
                      " amount = " & Round(ROTOTAL - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & "," & _
                      " rovat = " & Round(TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX, 2) & "," & _
                      " wl_amt = " & 0 & "," & _
                      " ro_amount = " & Round(ROTOTAL - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC, 2) & _
                      " where REP_OR = " & N2Str2Null(vREP_OR)
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT ----------------------------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREP_OR), "REP_OR", "CSMS_REPOR"), "", "POSTED IN PARTS W/ INS. PARTICIPATION", "", "")
        'NEW LOG AUDIT ----------------------------------------------------------------------

        FillParts vREP_OR
        ROTOTAL = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT - (NumericVal(vPartTotal))
        SQL_STATEMENT = "update CSMS_RepOr set" & _
                      " parts = " & TOTPARTSAMT - TOTPARTSTAX - (NumericVal(vPartParts)) & "," & _
                      " p_amtvalue = " & TOTPARTSAMT - NumericVal(vPartParts) & "," & _
                      " p_disc = " & TOTPARTSDISCVAL & "," & _
                      " p_disc2 = " & TOTPARTSDISC * (VAT_RATE / 100) & "," & _
                      " p_taxval = " & TOTPARTSTAX & "," & _
                      " p_discount = " & TOTPARTSDISC & "," & _
                      " amount = " & ROTOTAL - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                      " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                      " wp_amt = " & 0 & "," & _
                      " ro_amount = " & ROTOTAL - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
                      " where REP_OR = " & N2Str2Null(vREP_OR)
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT ----------------------------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREP_OR), "REP_OR", "CSMS_REPOR"), "", "POSTED IN PARTS W/ INS. PARTICIPATION", "", "")
        'NEW LOG AUDIT ----------------------------------------------------------------------

        FillMaterials vREP_OR
        ROTOTAL = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
        SQL_STATEMENT = "update CSMS_RepOr set" & _
                      " material = " & TOTMATAMT - TOTMATTAX - NumericVal(vPartMaterials) & "," & _
                      " m_amtvalue = " & TOTMATAMT - NumericVal(vPartMaterials) & "," & _
                      " m_disc = " & TOTMATDISCVAL & "," & _
                      " m_disc2 = " & TOTMATDISC * (VAT_RATE / 100) & "," & _
                      " m_taxval = " & TOTMATTAX & "," & _
                      " m_discount = " & TOTMATDISC & "," & _
                      " amount = " & ROTOTAL - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                      " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                      " wm_amt = " & 0 & "," & _
                      " ro_amount = " & ROTOTAL - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
                      " where REP_OR = " & N2Str2Null(vREP_OR)
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT ----------------------------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREP_OR), "REP_OR", "CSMS_REPOR"), "", "POSTED IN PARTS W/ INS. PARTICIPATION", "", "")
        'NEW LOG AUDIT ----------------------------------------------------------------------

        FillAccessories vREP_OR
        ROTOTAL = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
        SQL_STATEMENT = "update CSMS_RepOr set" & _
                      " Accessories = " & TOTACCAMT - TOTACCTAX - NumericVal(vPartAccessories) & "," & _
                      " A_amtvalue = " & TOTACCAMT - NumericVal(vPartAccessories) & "," & _
                      " A_disc = " & TOTACCDISCVAL & "," & _
                      " A_disc2 = " & TOTACCDISC * (VAT_RATE / 100) & "," & _
                      " A_taxval = " & TOTACCTAX & "," & _
                      " A_discount = " & TOTACCDISC & "," & _
                      " amount = " & ROTOTAL - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                      " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                      " WA_amt = " & 0 & "," & _
                      " ro_amount = " & ROTOTAL - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
                      " where REP_OR = " & N2Str2Null(vREP_OR)
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT ----------------------------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREP_OR), "REP_OR", "CSMS_REPOR"), "", "POSTED IN PARTS W/ INS. PARTICIPATION", "", "")
        'NEW LOG AUDIT ----------------------------------------------------------------------
    End If
    Screen.MousePointer = 0
End Sub


Sub ImportParts(vREP_OR As String)
    'On Error GoTo ERRORCODE

    Dim rsCSMS_REPOR                                   As ADODB.Recordset
    Set rsCSMS_REPOR = New ADODB.Recordset
    Set rsCSMS_REPOR = gconDMIS.Execute("Select DTE_COMP from CSMS_REPOR WHERE REP_OR = '" & vREP_OR & "'")
    If Not rsCSMS_REPOR.EOF And Not rsCSMS_REPOR.BOF Then
        If Null2Date(rsCSMS_REPOR!dte_comp) <> "" Then
            MsgBox "Repair Order is Already Billed. Transaction will not be imported!", vbInformation, vREP_OR & " Already Billed"
            Exit Sub
        End If
    End If
    Set rsCSMS_REPOR = Nothing


    Dim RONOFORMAT                                     As String
    Dim YZA                                            As Integer
    Dim TISOY                                          As String
    Dim KEIKEI                                         As String
    RONOFORMAT = ""

    KEIKEI = ""
    TISOY = ""
    YZA = 0

    For YZA = 1 To Len(vREP_OR)
        TISOY = Mid(vREP_OR, YZA, 1)
        KEIKEI = KEIKEI + TISOY
    Next

    RONOFORMAT = KEIKEI

    Dim VARPARTSLINE_NO                                As String
    Dim VarPartNo                                      As String
    Dim VARDESCRIPTION                                 As String
    Dim VARPARTCODE                                    As String
    Dim VARQTY                                         As Double
    Dim VARUNITCOST                                    As Double
    Dim VARUNITPRICE                                   As Double
    Dim VARPARTAMOUNT                                  As String
    Dim VARCHARGETO                                    As String
    Dim VARPARTDISCOUNT                                As String
    Dim PARTSREP_OR                                    As String
    Dim PARTSLEVEL                                     As String
    Dim PARTSLINE_NO                                   As String
    Dim PARTSDETCDE                                    As String
    Dim PARTSDETDSC                                    As String
    Dim PARTSDETUNT                                    As String
    Dim PARTSDETVOL                                    As Double
    Dim PARTSDETPRC                                    As Double
    Dim PARTSDETAMT                                    As Double
    Dim PARTSCODE                                      As String
    Dim PARTSWCODE                                     As String
    Dim PARTSTAXRATE                                   As Double
    Dim PARTSDISCRATE                                  As Double
    Dim PARTSTAXVAL                                    As Double
    Dim PARTSDISVAL                                    As Double
    Dim PARTSPOCODE                                    As String
    Dim PARTSREP_OR2                                   As String
    Dim PARTSDETAIL                                    As String
    Dim PARTSDET_AMT                                   As Double
    Dim PARTSDETCOST                                   As Double
    Dim PARTSDIS_VAL                                   As Double
    Dim PARTSDISCOUNT_2                                As Double
    Dim PARTSREMARKS                                   As String
    Dim REF_RIV_ADB                                    As String
    Dim RSRR_HDCHECK                                   As ADODB.Recordset
    Dim RSRR_HDTDAYTRANCHECK                           As ADODB.Recordset
    Dim VGJORBP                                        As String
    Dim MTRANNO                                        As String
    
    VPIS_NO_CHARGE_TO = "NULL"
    VGJORBP = "NULL"
    gconDMIS.Execute "delete from CSMS_RO_Det where livil = '2' and rep_or = " & N2Str2Null(vREP_OR)
    Pcnt = 0
    RO_RIV_TRANNO_COUNTER = 0
    Set RSCSMSORD_HIST = New ADODB.Recordset
    
    'Set RSCSMSORD_HIST = gconDMIS.Execute("select rono,trandate,trantype,tranno,refpisno from PMIS_ord_hist where [TYPE] = 'P' AND status <> 'C' and status <> 'N' and rono = '" & RONOFORMAT & "'")
     Set RSCSMSORD_HIST = gconDMIS.Execute("select rono,trandate,trantype,tranno,REFPISNO from PMIS_ord_hIST where [TYPE] = 'P' AND status <> 'C' AND ISNULL(STATUS2,'')<>'R' and status <> 'N' and trantype IN('RIV','ADB') and rono = '" & RONOFORMAT & "'")
    If Not RSCSMSORD_HIST.EOF And Not RSCSMSORD_HIST.BOF Then
        RSCSMSORD_HIST.MoveFirst
        Do While Not RSCSMSORD_HIST.EOF
            If Mid(Null2String(RSCSMSORD_HIST!refpisno), 5, 1) = "C" Then VPIS_NO_CHARGE_TO = "NULL"
            If Mid(Null2String(RSCSMSORD_HIST!refpisno), 5, 1) = "I" Then VPIS_NO_CHARGE_TO = "'C'"
            If Mid(Null2String(RSCSMSORD_HIST!refpisno), 5, 1) = "W" Then VPIS_NO_CHARGE_TO = "'W'"
            If Mid(Null2String(RSCSMSORD_HIST!refpisno), 4, 1) = "B" Then VGJORBP = "'BP'"
            If Mid(Null2String(RSCSMSORD_HIST!refpisno), 4, 1) = "G" Then VGJORBP = "'GJ'"

            Set RSCSMSDAYTRAN = New ADODB.Recordset
            Set RSCSMSDAYTRAN = gconDMIS.Execute("select itemno,trantype,tranno,stock_ord,stock_sup,tranqty,tranucost,tranuprice from PMIS_DayTran where [TYPE] = 'P' AND trantype = 'RIV' and tranno = " & N2Str2Null(RSCSMSORD_HIST!TRANNO) & " order by itemno asc")
            If Not RSCSMSDAYTRAN.EOF And Not RSCSMSDAYTRAN.BOF Then
                Screen.MousePointer = 11
                RSCSMSDAYTRAN.MoveFirst
                RO_RIV_TRANNO_COUNTER = RO_RIV_TRANNO_COUNTER + 1
                RO_RIV_TRANNO(RO_RIV_TRANNO_COUNTER) = Null2String(RSCSMSORD_HIST!TRANNO)
                Do While Not RSCSMSDAYTRAN.EOF
                    MTRANNO = Null2String(RSCSMSDAYTRAN!TRANNO)
                    Pcnt = Pcnt + 1
                    VARPARTSLINE_NO = "": VarPartNo = "": VARDESCRIPTION = ""
                    VARPARTCODE = "": VARQTY = 0: VARUNITPRICE = 0
                    VARPARTAMOUNT = "": VARCHARGETO = " ": VARPARTDISCOUNT = ZERO

                    VARPARTSLINE_NO = Format(Pcnt, "00")
                    VarPartNo = Null2String(RSCSMSDAYTRAN!STOCK_ORD)
                    VARDESCRIPTION = SetSTOCKDESC(Null2String(RSCSMSDAYTRAN!STOCK_ORD))
                    VARPARTCODE = "01"
                    VARQTY = Format(N2Str2IntZero(RSCSMSDAYTRAN!tranqty), "####0")
                    VARUNITCOST = N2Str2Zero(RSCSMSDAYTRAN!tranucost)
                    VARUNITPRICE = N2Str2Zero(RSCSMSDAYTRAN!TRANUPRICE)
                    VARPARTAMOUNT = N2Str2Zero(RSCSMSDAYTRAN!tranqty) * N2Str2Zero(RSCSMSDAYTRAN!TRANUPRICE)
                    VARCHARGETO = " "
                    VARPARTDISCOUNT = ZERO
                    Set RSRR_HDCHECK = New ADODB.Recordset
                    Set RSRR_HDCHECK = gconDMIS.Execute("select * from PMIS_RR_HD where [TYPE] = 'P' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(RSCSMSORD_HIST!TRANNO), "000000") & "'")
                    If Not RSRR_HDCHECK.EOF And Not RSRR_HDCHECK.BOF Then
                        Set RSRR_HDTDAYTRANCHECK = New ADODB.Recordset
                        Set RSRR_HDTDAYTRANCHECK = gconDMIS.Execute("select * from PMIS_TdayTran where [TYPE] = 'P' AND trantype = 'RR' and tranno = " & N2Str2Null(RSRR_HDCHECK!RRNO) & " order by Itemno asc")
                        If Not RSRR_HDTDAYTRANCHECK.EOF And Not RSRR_HDTDAYTRANCHECK.BOF Then
                            RSRR_HDTDAYTRANCHECK.MoveNext
                            Do While Not RSRR_HDTDAYTRANCHECK.EOF
                                If Null2String(RSRR_HDTDAYTRANCHECK!STOCK_ORD) = VarPartNo Then GoTo 10000
                                RSRR_HDTDAYTRANCHECK.MoveNext
                            Loop
                        End If
                    End If
                    Set RSRR_HDCHECK = New ADODB.Recordset
                    Set RSRR_HDCHECK = gconDMIS.Execute("select * from PMIS_REC_HIST where [TYPE] = 'P' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(RSCSMSORD_HIST!TRANNO), "000000") & "'")
                    If Not RSRR_HDCHECK.EOF And Not RSRR_HDCHECK.BOF Then
                        Set RSRR_HDTDAYTRANCHECK = New ADODB.Recordset
                        Set RSRR_HDTDAYTRANCHECK = gconDMIS.Execute("select * from PMIS_DayTran where [TYPE] = 'P' AND trantype = 'RR' and tranno = " & N2Str2Null(RSRR_HDCHECK!RRNO) & " order by Itemno asc")
                        If Not RSRR_HDTDAYTRANCHECK.EOF And Not RSRR_HDTDAYTRANCHECK.BOF Then
                            RSRR_HDTDAYTRANCHECK.MoveNext
                            Do While Not RSRR_HDTDAYTRANCHECK.EOF
                                If Null2String(RSRR_HDTDAYTRANCHECK!PARTNO) = VarPartNo Then GoTo 10000
                                RSRR_HDTDAYTRANCHECK.MoveNext
                            Loop
                        End If
                    End If
                    REF_RIV_ADB = "'RIV" & Format(Null2String(RSCSMSDAYTRAN!TRANNO), "000000") & Format(Null2String(RSCSMSDAYTRAN!itemno), "000") & "'"
                    PARTSDISVAL = 0: PARTSTAXVAL = 0: PARTSDETAMT = 0
                    PARTSDIS_VAL = 0: PARTSDISCOUNT_2 = 0: PARTSDISCRATE = 0

                    PARTSREP_OR = N2Str2Null(vREP_OR)
                    PARTSLEVEL = "'2'"
                    PARTSLINE_NO = N2Str2Null(Format(VARPARTSLINE_NO, "00"))
                    PARTSDETCDE = N2Str2Null(Left(VarPartNo, 50))
                    PARTSDETDSC = N2Str2Null(Mid(VARDESCRIPTION, 1, 100))
                    PARTSDETUNT = "NULL"
                    PARTSDETVOL = N2Str2Zero(VARQTY)
                    PARTSDETCOST = NumericVal(VARUNITCOST)
                    PARTSDETPRC = NumericVal(VARUNITPRICE)
                    PARTSDETAMT = Round(NumericVal(VARPARTAMOUNT) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                    PARTSCODE = "NULL"
                    PARTSWCODE = VPIS_NO_CHARGE_TO
                    PARTSTAXRATE = (VAT_RATE / 100)
                    PARTSDISCRATE = NumericVal(VARPARTDISCOUNT) / 100
                    PARTSDISVAL = Round((PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE), 2)
                    PARTSPOCODE = N2Str2Null(VARPARTCODE)
                    PARTSREP_OR2 = "NULL"
                    PARTSDETAIL = "NULL"
                    PARTSDET_AMT = NumericVal(VARPARTAMOUNT)
                    PARTSDIS_VAL = Round(PARTSDISVAL * PARTSTAXRATE, 2)
                    PARTSDISCOUNT_2 = Round(PARTSDET_AMT * PARTSDISCRATE, 2)
                    PARTSTAXVAL = Round((PARTSDETAMT - PARTSDISCOUNT_2) * PARTSTAXRATE, 2)

                    SQL_STATEMENT = "insert into CSMS_RO_Det " & _
                                    "(rep_or,livil,LINE_NO,JOBTYPE,detcde,detdsc,detunt,detvol,detcost,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,REF_RIV_ADB)" & _
                                  " values (" & PARTSREP_OR & ", " & PARTSLEVEL & ", " & PARTSLINE_NO & "," & _
                                  " " & VGJORBP & ", " & PARTSDETCDE & "," & PARTSDETDSC & "," & _
                                  " " & PARTSDETUNT & ", " & PARTSDETVOL & "," & _
                                  " " & PARTSDETCOST & ", " & PARTSDETPRC & ", " & PARTSDETAMT & ", " & PARTSCODE & _
                                    ", " & PARTSWCODE & ", " & PARTSTAXRATE * 100 & ", " & PARTSDISCRATE * 100 & _
                                    ", " & PARTSTAXVAL & ", " & PARTSDISVAL & ", " & PARTSPOCODE & _
                                    ", " & PARTSREP_OR2 & ", " & PARTSDETAIL & ", " & PARTSDET_AMT & _
                                    ", " & PARTSDIS_VAL & ", " & PARTSDISCOUNT_2 & "," & REF_RIV_ADB & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    'NEW LOG AUDIT ---------------------------------------------------------
                    Call NEW_LogAudit("AA", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREP_OR), "REP_OR", "CSMS_REPOR"), "P", "PART NO: " & Null2String(RSCSMSDAYTRAN!TRANNO) & " ,PART NO: " & Null2String(PARTSDETCDE) & " - HISTORY", "RIV", "")
                    'NEW LOG AUDIT ---------------------------------------------------------
                    Screen.MousePointer = 0
10000               RSCSMSDAYTRAN.MoveNext
                Loop
            End If
            RSCSMSORD_HIST.MoveNext
        Loop
    End If

    Set RSCSMSORD_HD = gconDMIS.Execute("select rono,tranno,trantype,REFPISNO from PMIS_ord_hd where [TYPE] = 'P' AND status <> 'C' AND ISNULL(STATUS2,'')<>'R' and status <> 'N' and trantype IN('RIV','ADB') and rono = '" & RONOFORMAT & "'")
    If Not RSCSMSORD_HD.EOF And Not RSCSMSORD_HD.BOF Then
        RSCSMSORD_HD.MoveFirst
        Do While Not RSCSMSORD_HD.EOF
            If Mid(Null2String(RSCSMSORD_HD!refpisno), 5, 1) = "C" Then VPIS_NO_CHARGE_TO = "NULL"
            If Mid(Null2String(RSCSMSORD_HD!refpisno), 5, 1) = "I" Then VPIS_NO_CHARGE_TO = "'C'"
            If Mid(Null2String(RSCSMSORD_HD!refpisno), 5, 1) = "W" Then VPIS_NO_CHARGE_TO = "'W'"
            If Mid(Null2String(RSCSMSORD_HD!refpisno), 4, 1) = "B" Then VGJORBP = "'BP'"
            If Mid(Null2String(RSCSMSORD_HD!refpisno), 4, 1) = "G" Then VGJORBP = "'GJ'"
            
            Set RSCSMSTDAYTRAN = gconDMIS.Execute("select itemno,trantype,tranno,STOCK_ord,STOCK_sup,tranqty,tranucost,tranuprice from PMIS_TdayTran where [TYPE] = 'P' AND trantype = '" & Null2String(RSCSMSORD_HD!TranType) & "' and tranno = " & N2Str2Null(RSCSMSORD_HD!TRANNO) & " order by itemno asc")
            If Not RSCSMSTDAYTRAN.EOF And Not RSCSMSTDAYTRAN.BOF Then
                Screen.MousePointer = 11
                RSCSMSTDAYTRAN.MoveFirst
                RO_RIV_TRANNO_COUNTER = RO_RIV_TRANNO_COUNTER + 1
                RO_RIV_TRANNO(RO_RIV_TRANNO_COUNTER) = Null2String(RSCSMSORD_HD!TRANNO)
                
                Do While Not RSCSMSTDAYTRAN.EOF
                    MTRANNO = Null2String(RSCSMSTDAYTRAN!TRANNO)
                    Pcnt = Pcnt + 1
                    VARPARTSLINE_NO = ""
                    VarPartNo = ""
                    VARDESCRIPTION = ""
                    VARPARTCODE = ""
                    VARQTY = 0
                    VARUNITPRICE = 0
                    VARPARTAMOUNT = ""
                    VARCHARGETO = " "
                    VARPARTDISCOUNT = ZERO
                    VARPARTSLINE_NO = Format(Pcnt, "00")
                    VarPartNo = Null2String(RSCSMSTDAYTRAN!STOCK_ORD)
                    VARDESCRIPTION = SetSTOCKDESC(Null2String(RSCSMSTDAYTRAN!STOCK_ORD))
                    VARPARTCODE = "01"
                    VARQTY = Format(N2Str2IntZero(RSCSMSTDAYTRAN!tranqty), "####0")
                    VARUNITCOST = N2Str2Zero(RSCSMSTDAYTRAN!tranucost)
                    VARUNITPRICE = N2Str2Zero(RSCSMSTDAYTRAN!TRANUPRICE)
                    VARPARTAMOUNT = N2Str2Zero(RSCSMSTDAYTRAN!tranqty) * N2Str2Zero(RSCSMSTDAYTRAN!TRANUPRICE)
                    VARCHARGETO = " "
                    VARPARTDISCOUNT = ZERO
                    
                    Set RSRR_HDCHECK = gconDMIS.Execute("select * from PMIS_RR_HD where [TYPE] = 'P' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(RSCSMSORD_HD!TRANNO), "000000") & "'")
                    If Not RSRR_HDCHECK.EOF And Not RSRR_HDCHECK.BOF Then
                        Set RSRR_HDTDAYTRANCHECK = New ADODB.Recordset
                        Set RSRR_HDTDAYTRANCHECK = gconDMIS.Execute("select * from PMIS_Tdaytran where [TYPE] = 'P' AND trantype = 'RR' and tranno = " & N2Str2Null(RSRR_HDCHECK!RRNO) & " order by Itemno asc")
                        If Not RSRR_HDTDAYTRANCHECK.EOF And Not RSRR_HDTDAYTRANCHECK.BOF Then
                            RSRR_HDTDAYTRANCHECK.MoveNext
                            Do While Not RSRR_HDTDAYTRANCHECK.EOF
                                If Null2String(RSRR_HDTDAYTRANCHECK!PARTNO) = VarPartNo Then GoTo 20000
                                RSRR_HDTDAYTRANCHECK.MoveNext
                            Loop
                        End If
                    End If
                    
                    Set RSRR_HDCHECK = gconDMIS.Execute("select * from PMIS_REC_HIST where [TYPE] = 'P' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(RSCSMSORD_HD!TRANNO), "000000") & "'")
                    If Not RSRR_HDCHECK.EOF And Not RSRR_HDCHECK.BOF Then
                        Set RSRR_HDTDAYTRANCHECK = New ADODB.Recordset
                        Set RSRR_HDTDAYTRANCHECK = gconDMIS.Execute("select * from PMIS_DayTran where [TYPE] = 'P' AND trantype = 'RR' and tranno = " & N2Str2Null(RSRR_HDCHECK!RRNO) & " order by Itemno asc")
                        If Not RSRR_HDTDAYTRANCHECK.EOF And Not RSRR_HDTDAYTRANCHECK.BOF Then
                            RSRR_HDTDAYTRANCHECK.MoveNext
                            Do While Not RSRR_HDTDAYTRANCHECK.EOF
                                If Null2String(RSRR_HDTDAYTRANCHECK!PARTNO) = VarPartNo Then GoTo 20000
                                RSRR_HDTDAYTRANCHECK.MoveNext
                            Loop
                        End If
                    End If
                    
                    REF_RIV_ADB = "'RIV" & Format(Null2String(RSCSMSTDAYTRAN!TRANNO), "000000") & Format(Null2String(RSCSMSTDAYTRAN!itemno), "000") & "'"
                    PARTSDISVAL = 0
                    PARTSTAXVAL = 0
                    PARTSDETAMT = 0
                    PARTSDIS_VAL = 0
                    PARTSDISCOUNT_2 = 0
                    PARTSDISCRATE = 0
                    PARTSREP_OR = N2Str2Null(vREP_OR)
                    PARTSLEVEL = "'2'"
                    PARTSLINE_NO = N2Str2Null(Format(VARPARTSLINE_NO, "00"))
                    PARTSDETCDE = N2Str2Null(Left(VarPartNo, 50))
                    PARTSDETDSC = N2Str2Null(Mid(VARDESCRIPTION, 1, 100))
                    PARTSDETUNT = "NULL"
                    PARTSDETVOL = N2Str2Zero(VARQTY)
                    PARTSDETCOST = NumericVal(VARUNITCOST)
                    PARTSDETPRC = NumericVal(VARUNITPRICE)
                    PARTSDETAMT = Round(NumericVal(VARPARTAMOUNT) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                    PARTSCODE = "NULL"
                    PARTSWCODE = VPIS_NO_CHARGE_TO    'N2Str2Null(VarChargeTo)
                    PARTSTAXRATE = (VAT_RATE / 100)
                    PARTSDISCRATE = Round(NumericVal(VARPARTDISCOUNT) / 100, 2)
                    PARTSDISVAL = Round((PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE), 2)
                    PARTSPOCODE = N2Str2Null(VARPARTCODE)
                    PARTSREP_OR2 = "NULL"
                    PARTSDETAIL = "NULL"
                    PARTSDET_AMT = NumericVal(VARPARTAMOUNT)
                    PARTSDIS_VAL = Round(PARTSDISVAL * PARTSTAXRATE, 2)
                    PARTSDISCOUNT_2 = Round(PARTSDET_AMT * PARTSDISCRATE, 2)
                    PARTSTAXVAL = Round((PARTSDETAMT - PARTSDISCOUNT_2) * PARTSTAXRATE, 2)

                    SQL_STATEMENT = "INSERT INTO CSMS_RO_DET " & _
                                    "(REP_OR,LIVIL,LINE_NO,JOBTYPE,DETCDE,DETDSC,DETUNT,DETVOL,DETCOST,DETPRC,DETAMT,CODE,WCODE,TAXRATE,DISCRATE,TAXVAL,DISVAL,POCODE,REP_OR2,DETAIL,DET_AMT,DIS_VAL,DISCOUNT_2,REF_RIV_ADB)" & _
                                  " VALUES (" & PARTSREP_OR & ", " & PARTSLEVEL & ", " & PARTSLINE_NO & "," & _
                                  " " & VGJORBP & ", " & PARTSDETCDE & "," & PARTSDETDSC & "," & _
                                  " " & PARTSDETUNT & ", " & PARTSDETVOL & "," & _
                                  " " & PARTSDETCOST & ", " & PARTSDETPRC & ", " & PARTSDETAMT & ", " & PARTSCODE & _
                                    ", " & PARTSWCODE & ", " & PARTSTAXRATE * 100 & ", " & PARTSDISCRATE * 100 & _
                                    ", " & PARTSTAXVAL & ", " & PARTSDISVAL & ", " & PARTSPOCODE & _
                                    ", " & PARTSREP_OR2 & ", " & PARTSDETAIL & ", " & PARTSDET_AMT & _
                                    ", " & PARTSDIS_VAL & ", " & PARTSDISCOUNT_2 & "," & REF_RIV_ADB & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    'NEW LOG AUDIT ---------------------------------------------------------
                    Call NEW_LogAudit("AA", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREP_OR), "REP_OR", "CSMS_REPOR"), "P", "TRAN NO: " & Null2String(RSCSMSTDAYTRAN!TRANNO) & " ,PART NO: " & Null2String(PARTSDETCDE), "RIV", "")
                    'NEW LOG AUDIT ---------------------------------------------------------
                    Screen.MousePointer = 0
20000               RSCSMSTDAYTRAN.MoveNext
                Loop
            End If
            RSCSMSORD_HD.MoveNext
        Loop
    End If

    Set RSCSMSORD_HD = New ADODB.Recordset
    Set RSCSMSORD_HD = gconDMIS.Execute("Select PARTICIPAT FROM CSMS_REPOR WHERE REP_OR = " & N2Str2Null(vREP_OR))
    If Not RSCSMSORD_HD.EOF And Not RSCSMSORD_HD.BOF Then
        If Null2String(RSCSMSORD_HD!PARTICIPAT) = "" Then
            FillJobs vREP_OR
            FillParts vREP_OR
            FillMaterials vREP_OR
            FillAccessories vREP_OR
            ROTOTAL = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
            SQL_STATEMENT = "update CSMS_RepOr set" & _
                          " parts = " & TOTPARTSAMT - TOTPARTSTAX & "," & _
                          " p_amtvalue = " & TOTPARTSAMT & "," & _
                          " p_disc = " & TOTPARTSDISCVAL & "," & _
                          " p_disc2 = " & TOTPARTSDISC * (VAT_RATE / 100) & "," & _
                          " p_taxval = " & TOTPARTSTAX & "," & _
                          " p_discount = " & TOTPARTSDISC & "," & _
                          " amount = " & ROTOTAL - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                          " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                          " wp_amt = " & 0 & "," & _
                          " ro_amount = " & ROTOTAL - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
                          " where rep_or = " & N2Str2Null(vREP_OR)
            gconDMIS.Execute SQL_STATEMENT
            'NEW LOG AUDIT ---------------------------------------------------------
            Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREP_OR), "REP_OR", "CSMS_REPOR"), "P", "TRAN NO: " & MTRANNO, "RIV", "")
            'NEW LOG AUDIT ---------------------------------------------------------
        Else
            UpdateParticipation vREP_OR
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub

Errorcode:
    Screen.MousePointer = 0:
    ShowVBError
    MsgBox err.Description
End Sub

Sub ImportMaterials(vREP_OR As String)
    'On Error GoTo ERRORCODE
    Dim rsCSMS_REPOR                                   As ADODB.Recordset
    Set rsCSMS_REPOR = gconDMIS.Execute("Select DTE_COMP from CSMS_REPOR WHERE REP_OR = '" & vREP_OR & "'")
    If rsCSMS_REPOR.EOF And Not rsCSMS_REPOR.BOF Then
        If Null2Date(rsCSMS_REPOR!dte_comp) <> "" Then
            MsgBox "Repair Order is Already Billed. Transaction will not be imported!", vbInformation, vREP_OR & " Already Billed"
            Exit Sub
        End If
    End If
    Set rsCSMS_REPOR = Nothing

    Dim RONOFORMAT                                     As String
    Dim YZA                                            As Integer
    Dim TISOY                                          As String
    Dim KEIKEI                                         As String
    RONOFORMAT = ""

    KEIKEI = ""
    TISOY = ""
    YZA = 0
    For YZA = 1 To Len(vREP_OR)
        TISOY = Mid(vREP_OR, YZA, 1)
        KEIKEI = KEIKEI + TISOY
    Next
    RONOFORMAT = KEIKEI
    Dim VARPARTSLINE_NO                                As String
    Dim VarPartNo                                      As String
    Dim VARDESCRIPTION                                 As String
    Dim VARPARTCODE                                    As String
    Dim VARQTY                                         As Double
    Dim VARUNITCOST                                    As Double
    Dim VARUNITPRICE                                   As Double
    Dim VARPARTAMOUNT, VARCHARGETO, VARPARTDISCOUNT    As String
    Dim PARTSREP_OR                                    As String
    Dim PARTSLEVEL                                     As String
    Dim PARTSLINE_NO                                   As String
    Dim PARTSDETCDE                                    As String
    Dim PARTSDETDSC                                    As String
    Dim PARTSDETUNT                                    As String
    Dim PARTSDETVOL                                    As Double
    Dim PARTSDETPRC                                    As Double
    Dim PARTSDETAMT                                    As Double
    Dim PARTSCODE                                      As String
    Dim PARTSWCODE                                     As String
    Dim PARTSTAXRATE                                   As Double
    Dim PARTSDISCRATE                                  As Double
    Dim PARTSTAXVAL                                    As Double
    Dim PARTSDISVAL                                    As Double
    Dim PARTSPOCODE                                    As String
    Dim PARTSREP_OR2                                   As String
    Dim PARTSDETAIL                                    As String
    Dim PARTSDET_AMT                                   As Double
    Dim PARTSDETCOST                                   As Double
    Dim PARTSDIS_VAL                                   As Double
    Dim PARTSDISCOUNT_2                                As Double
    Dim PARTSREMARKS                                   As String
    Dim REF_RIV_ADB                                    As String
    Dim RSRR_HDCHECK                                   As ADODB.Recordset
    Dim RSRR_HDTDAYTRANCHECK                           As ADODB.Recordset
    Dim VGJORBP                                        As String
    Dim MTRANNO                                        As String

    VPIS_NO_CHARGE_TO = "NULL"
    gconDMIS.Execute "DELETE FROM CSMS_RO_DET WHERE LIVIL = '3' AND REP_OR = " & N2Str2Null(vREP_OR)
    Pcnt = 0
    RO_RIV_TRANNO_COUNTER = 0
    Set RSCSMSORD_HIST = New ADODB.Recordset
    'Set RSCSMSORD_HIST = gconDMIS.Execute("select rono,trandate,trantype,tranno,REFPISNO from PMIS_ord_hist where [TYPE] = 'M' AND status <> 'C' and status <> 'N' and rono = '" & RONOFORMAT & "'")
    Set RSCSMSORD_HIST = gconDMIS.Execute("SELECT RONO,TRANDATE,TRANTYPE,TRANNO,REFPISNO FROM PMIS_ORD_HIST WHERE [TYPE] = 'M' AND STATUS <> 'C' AND ISNULL(STATUS2,'')<>'R' AND STATUS <> 'N' AND TRANTYPE IN('RIV','ADB') AND RONO = '" & RONOFORMAT & "'")

    If Not RSCSMSORD_HIST.EOF And Not RSCSMSORD_HIST.BOF Then
        RSCSMSORD_HIST.MoveFirst
        Do While Not RSCSMSORD_HIST.EOF
            If Mid(Null2String(RSCSMSORD_HIST!refpisno), 5, 1) = "C" Then VPIS_NO_CHARGE_TO = "NULL"
            If Mid(Null2String(RSCSMSORD_HIST!refpisno), 5, 1) = "I" Then VPIS_NO_CHARGE_TO = "'C'"
            If Mid(Null2String(RSCSMSORD_HIST!refpisno), 5, 1) = "W" Then VPIS_NO_CHARGE_TO = "'W'"

            If Mid(Null2String(RSCSMSORD_HIST!refpisno), 4, 1) = "B" Then VGJORBP = "'BP'"
            If Mid(Null2String(RSCSMSORD_HIST!refpisno), 4, 1) = "G" Then VGJORBP = "'GJ'"

            MTRANNO = Null2String(RSCSMSORD_HIST!TRANNO)


            Set RSCSMSDAYTRAN = gconDMIS.Execute("SELECT ITEMNO,TRANTYPE,TRANNO,STOCK_ORD,STOCK_SUP,TRANQTY,TRANUCOST,TRANUPRICE FROM PMIS_DAYTRAN WHERE [TYPE] = 'M' AND TRANTYPE = '" & Null2String(RSCSMSORD_HIST!TranType) & "'  AND TRANNO = " & N2Str2Null(RSCSMSORD_HIST!TRANNO) & " ORDER BY ITEMNO ASC")
            If Not RSCSMSDAYTRAN.EOF And Not RSCSMSDAYTRAN.BOF Then
                Screen.MousePointer = 11
                RSCSMSDAYTRAN.MoveFirst
                RO_RIV_TRANNO_COUNTER = RO_RIV_TRANNO_COUNTER + 1
                RO_RIV_TRANNO(RO_RIV_TRANNO_COUNTER) = Null2String(RSCSMSORD_HIST!TRANNO)
                Do While Not RSCSMSDAYTRAN.EOF
                    Pcnt = Pcnt + 1
                    VARPARTSLINE_NO = ""
                    VarPartNo = ""
                    VARDESCRIPTION = ""
                    VARPARTCODE = ""
                    VARQTY = 0
                    VARUNITPRICE = 0
                    VARPARTAMOUNT = ""
                    VARCHARGETO = " "
                    VARPARTDISCOUNT = ZERO
                    VARPARTSLINE_NO = Format(Pcnt, "00")
                    VarPartNo = Null2String(RSCSMSDAYTRAN!STOCK_ORD)
                    VARDESCRIPTION = SetSTOCKDESC(Null2String(RSCSMSDAYTRAN!STOCK_ORD))
                    VARPARTCODE = "01"
                    VARQTY = Format(N2Str2IntZero(RSCSMSDAYTRAN!tranqty), "####0")
                    VARUNITCOST = N2Str2Zero(RSCSMSDAYTRAN!tranucost)
                    VARUNITPRICE = N2Str2Zero(RSCSMSDAYTRAN!TRANUPRICE)
                    VARPARTAMOUNT = N2Str2Zero(RSCSMSDAYTRAN!tranqty) * N2Str2Zero(RSCSMSDAYTRAN!TRANUPRICE)
                    VARCHARGETO = " "
                    VARPARTDISCOUNT = ZERO
                    Set RSRR_HDCHECK = New ADODB.Recordset
                    Set RSRR_HDCHECK = gconDMIS.Execute("select * from PMIS_RR_HD where [TYPE] = 'M' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(RSCSMSORD_HIST!TRANNO), "000000") & "'")
                    If Not RSRR_HDCHECK.EOF And Not RSRR_HDCHECK.BOF Then
                        Set RSRR_HDTDAYTRANCHECK = New ADODB.Recordset
                        Set RSRR_HDTDAYTRANCHECK = gconDMIS.Execute("select * from PMIS_TdayTran where [TYPE] = 'M' AND trantype = 'RR' and tranno = " & N2Str2Null(RSRR_HDCHECK!RRNO) & " order by Itemno asc")
                        If Not RSRR_HDTDAYTRANCHECK.EOF And Not RSRR_HDTDAYTRANCHECK.BOF Then
                            RSRR_HDTDAYTRANCHECK.MoveNext
                            Do While Not RSRR_HDTDAYTRANCHECK.EOF
                                If Null2String(RSRR_HDTDAYTRANCHECK!STOCK_ORD) = VarPartNo Then GoTo 10000
                                RSRR_HDTDAYTRANCHECK.MoveNext
                            Loop
                        End If
                    End If
                    Set RSRR_HDCHECK = New ADODB.Recordset
                    Set RSRR_HDCHECK = gconDMIS.Execute("select * from PMIS_REC_HIST where [TYPE] = 'M' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(RSCSMSORD_HIST!TRANNO), "000000") & "'")
                    If Not RSRR_HDCHECK.EOF And Not RSRR_HDCHECK.BOF Then
                        Set RSRR_HDTDAYTRANCHECK = New ADODB.Recordset
                        Set RSRR_HDTDAYTRANCHECK = gconDMIS.Execute("select * from PMIS_DAYTRAN where [TYPE] = 'M' AND trantype = 'RR' and tranno = " & N2Str2Null(RSRR_HDCHECK!RRNO) & " order by Itemno asc")
                        If Not RSRR_HDTDAYTRANCHECK.EOF And Not RSRR_HDTDAYTRANCHECK.BOF Then
                            RSRR_HDTDAYTRANCHECK.MoveNext
                            Do While Not RSRR_HDTDAYTRANCHECK.EOF
                                If Null2String(RSRR_HDTDAYTRANCHECK!PARTNO) = VarPartNo Then GoTo 10000
                                RSRR_HDTDAYTRANCHECK.MoveNext
                            Loop
                        End If
                    End If
                    REF_RIV_ADB = "'RIV" & Format(Null2String(RSCSMSDAYTRAN!TRANNO), "000000") & Format(Null2String(RSCSMSDAYTRAN!itemno), "000") & "'"
                    PARTSDISVAL = 0: PARTSTAXVAL = 0: PARTSDETAMT = 0
                    PARTSDIS_VAL = 0: PARTSDISCOUNT_2 = 0: PARTSDISCRATE = 0

                    PARTSREP_OR = N2Str2Null(vREP_OR)
                    PARTSLEVEL = "'3'"
                    PARTSLINE_NO = N2Str2Null(Format(VARPARTSLINE_NO, "00"))
                    PARTSDETCDE = N2Str2Null(VarPartNo)
                    PARTSDETDSC = N2Str2Null(Mid(VARDESCRIPTION, 1, 100))
                    PARTSDETUNT = "NULL"
                    PARTSDETVOL = N2Str2Zero(VARQTY)
                    PARTSDETCOST = NumericVal(VARUNITCOST)
                    PARTSDETPRC = NumericVal(VARUNITPRICE)
                    PARTSDETAMT = Round(NumericVal(VARPARTAMOUNT) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                    PARTSCODE = "NULL"
                    PARTSWCODE = VPIS_NO_CHARGE_TO
                    PARTSTAXRATE = (VAT_RATE / 100)
                    PARTSDISCRATE = NumericVal(VARPARTDISCOUNT) / 100
                    PARTSDISVAL = Round((PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE), 2)
                    PARTSPOCODE = N2Str2Null(VARPARTCODE)
                    PARTSREP_OR2 = "NULL"
                    PARTSDETAIL = "NULL"
                    PARTSDET_AMT = NumericVal(VARPARTAMOUNT)
                    PARTSDIS_VAL = Round(PARTSDISVAL * PARTSTAXRATE, 2)
                    PARTSDISCOUNT_2 = Round(PARTSDET_AMT * PARTSDISCRATE, 2)
                    PARTSTAXVAL = Round((PARTSDETAMT - PARTSDISCOUNT_2) * PARTSTAXRATE, 2)

                    SQL_STATEMENT = "insert into CSMS_RO_Det " & _
                                    "(rep_or,livil,LINE_NO,JOBTYPE,detcde,detdsc,detunt,detvol,detcost,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,REF_RIV_ADB)" & _
                                  " values (" & PARTSREP_OR & ", " & PARTSLEVEL & ", " & PARTSLINE_NO & "," & _
                                  " " & VGJORBP & "," & PARTSDETCDE & "," & PARTSDETDSC & "," & _
                                  " " & PARTSDETUNT & ", " & PARTSDETVOL & "," & _
                                  " " & PARTSDETCOST & ", " & PARTSDETPRC & ", " & PARTSDETAMT & ", " & PARTSCODE & _
                                    ", " & PARTSWCODE & ", " & PARTSTAXRATE * 100 & ", " & PARTSDISCRATE * 100 & _
                                    ", " & PARTSTAXVAL & ", " & PARTSDISVAL & ", " & PARTSPOCODE & _
                                    ", " & PARTSREP_OR2 & ", " & PARTSDETAIL & ", " & PARTSDET_AMT & _
                                    ", " & PARTSDIS_VAL & ", " & PARTSDISCOUNT_2 & "," & REF_RIV_ADB & ")"

                    gconDMIS.Execute SQL_STATEMENT
                    Call NEW_LogAudit("AA", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREP_OR), "REP_OR", "CSMS_REPOR"), "", "TRAN NO: " & Null2String(RSCSMSORD_HIST!TRANNO), "", "")
                    Screen.MousePointer = 0
10000               RSCSMSDAYTRAN.MoveNext
                Loop
            End If
            RSCSMSORD_HIST.MoveNext
        Loop
    End If

    'Set RSCSMSORD_HD = gconDMIS.Execute("select rono,tranno,trantype,REFPISNO from PMIS_ord_hd where [TYPE] = 'M' AND status <> 'C' and status <> 'N' and rono = '" & RONOFORMAT & "'")
    Set RSCSMSORD_HD = gconDMIS.Execute("SELECT RONO,TRANNO,TRANTYPE,REFPISNO FROM PMIS_ORD_HD WHERE [TYPE] = 'M' AND STATUS <> 'C' AND ISNULL(STATUS2,'')<>'R' AND STATUS <> 'N' AND TRANTYPE IN('RIV','ADB') AND RONO = '" & RONOFORMAT & "'")
    If Not RSCSMSORD_HD.EOF And Not RSCSMSORD_HD.BOF Then
        RSCSMSORD_HD.MoveFirst
        Do While Not RSCSMSORD_HD.EOF
            If Mid(Null2String(RSCSMSORD_HD!refpisno), 5, 1) = "C" Then VPIS_NO_CHARGE_TO = "NULL"
            If Mid(Null2String(RSCSMSORD_HD!refpisno), 5, 1) = "I" Then VPIS_NO_CHARGE_TO = "'C'"
            If Mid(Null2String(RSCSMSORD_HD!refpisno), 5, 1) = "W" Then VPIS_NO_CHARGE_TO = "'W'"

            If Mid(Null2String(RSCSMSORD_HD!refpisno), 4, 1) = "B" Then VGJORBP = "'BP'"
            If Mid(Null2String(RSCSMSORD_HD!refpisno), 4, 1) = "G" Then VGJORBP = "'GJ'"

            MTRANNO = Null2String(RSCSMSORD_HD!TRANNO)
            Set RSCSMSTDAYTRAN = gconDMIS.Execute("SELECT ITEMNO,TRANTYPE,TRANNO,STOCK_ORD,STOCK_SUP,TRANQTY,TRANUCOST,TRANUPRICE FROM PMIS_TDAYTRAN WHERE [TYPE] = 'M' AND TRANTYPE = '" & Null2String(RSCSMSORD_HD!TranType) & "'  AND TRANNO = " & N2Str2Null(RSCSMSORD_HD!TRANNO) & " ORDER BY ITEMNO ASC")
            If Not RSCSMSTDAYTRAN.EOF And Not RSCSMSTDAYTRAN.BOF Then
                Screen.MousePointer = 11
                RSCSMSTDAYTRAN.MoveFirst
                RO_RIV_TRANNO_COUNTER = RO_RIV_TRANNO_COUNTER + 1
                RO_RIV_TRANNO(RO_RIV_TRANNO_COUNTER) = Null2String(RSCSMSORD_HD!TRANNO)
                Do While Not RSCSMSTDAYTRAN.EOF
                    Pcnt = Pcnt + 1
                    VARPARTSLINE_NO = "": VarPartNo = "": VARDESCRIPTION = ""
                    VARPARTCODE = "": VARQTY = 0: VARUNITPRICE = 0
                    VARPARTAMOUNT = "": VARCHARGETO = " ": VARPARTDISCOUNT = ZERO

                    VARPARTSLINE_NO = Format(Pcnt, "00")
                    VarPartNo = Null2String(RSCSMSTDAYTRAN!STOCK_ORD)
                    VARDESCRIPTION = SetSTOCKDESC(Null2String(RSCSMSTDAYTRAN!STOCK_ORD))
                    VARPARTCODE = "01"
                    VARQTY = Format(N2Str2IntZero(RSCSMSTDAYTRAN!tranqty), "####0")
                    VARUNITCOST = N2Str2Zero(RSCSMSTDAYTRAN!tranucost)
                    VARUNITPRICE = N2Str2Zero(RSCSMSTDAYTRAN!TRANUPRICE)
                    VARPARTAMOUNT = N2Str2Zero(RSCSMSTDAYTRAN!tranqty) * N2Str2Zero(RSCSMSTDAYTRAN!TRANUPRICE)
                    VARCHARGETO = " "
                    VARPARTDISCOUNT = ZERO

                    Set RSRR_HDCHECK = New ADODB.Recordset
                    Set RSRR_HDCHECK = gconDMIS.Execute("select * from PMIS_RR_HD where [TYPE] = 'M' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(RSCSMSORD_HD!TRANNO), "000000") & "'")
                    If Not RSRR_HDCHECK.EOF And Not RSRR_HDCHECK.BOF Then
                        Set RSRR_HDTDAYTRANCHECK = New ADODB.Recordset
                        Set RSRR_HDTDAYTRANCHECK = gconDMIS.Execute("select * from PMIS_Tdaytran where [TYPE] = 'M' AND trantype = 'RR' and tranno = " & N2Str2Null(RSRR_HDCHECK!RRNO) & " order by Itemno asc")
                        If Not RSRR_HDTDAYTRANCHECK.EOF And Not RSRR_HDTDAYTRANCHECK.BOF Then
                            RSRR_HDTDAYTRANCHECK.MoveNext
                            Do While Not RSRR_HDTDAYTRANCHECK.EOF
                                If Null2String(RSRR_HDTDAYTRANCHECK!PARTNO) = VarPartNo Then GoTo 20000
                                RSRR_HDTDAYTRANCHECK.MoveNext
                            Loop
                        End If
                    End If

                    Set RSRR_HDCHECK = New ADODB.Recordset
                    Set RSRR_HDCHECK = gconDMIS.Execute("select * from PMIS_REC_HIST where [TYPE] = 'M' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(RSCSMSORD_HD!TRANNO), "000000") & "'")
                    If Not RSRR_HDCHECK.EOF And Not RSRR_HDCHECK.BOF Then
                        Set RSRR_HDTDAYTRANCHECK = New ADODB.Recordset
                        Set RSRR_HDTDAYTRANCHECK = gconDMIS.Execute("select * from PMIS_DayTran where [TYPE] = 'M' AND trantype = 'RR' and tranno = " & N2Str2Null(RSRR_HDCHECK!RRNO) & " order by Itemno asc")
                        If Not RSRR_HDTDAYTRANCHECK.EOF And Not RSRR_HDTDAYTRANCHECK.BOF Then
                            RSRR_HDTDAYTRANCHECK.MoveNext
                            Do While Not RSRR_HDTDAYTRANCHECK.EOF
                                If Null2String(RSRR_HDTDAYTRANCHECK!PARTNO) = VarPartNo Then GoTo 20000
                                RSRR_HDTDAYTRANCHECK.MoveNext
                            Loop
                        End If
                    End If

                    REF_RIV_ADB = "'RIV" & Format(Null2String(RSCSMSTDAYTRAN!TRANNO), "000000") & Format(Null2String(RSCSMSTDAYTRAN!itemno), "000") & "'"
                    PARTSDISVAL = 0: PARTSTAXVAL = 0: PARTSDETAMT = 0
                    PARTSDIS_VAL = 0: PARTSDISCOUNT_2 = 0: PARTSDISCRATE = 0

                    PARTSREP_OR = N2Str2Null(vREP_OR)
                    PARTSLEVEL = "'3'"
                    PARTSLINE_NO = N2Str2Null(Format(VARPARTSLINE_NO, "00"))
                    PARTSDETCDE = N2Str2Null(VarPartNo)
                    PARTSDETDSC = N2Str2Null(Mid(VARDESCRIPTION, 1, 100))
                    PARTSDETUNT = "NULL"
                    PARTSDETVOL = N2Str2Zero(VARQTY)
                    PARTSDETCOST = NumericVal(VARUNITCOST)
                    PARTSDETPRC = NumericVal(VARUNITPRICE)
                    PARTSDETAMT = Round(NumericVal(VARPARTAMOUNT) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                    PARTSCODE = "NULL"
                    PARTSWCODE = VPIS_NO_CHARGE_TO
                    PARTSTAXRATE = (VAT_RATE / 100)
                    PARTSDISCRATE = Round(NumericVal(VARPARTDISCOUNT) / 100, 2)
                    PARTSDISVAL = Round((PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE), 2)
                    PARTSPOCODE = N2Str2Null(VARPARTCODE)
                    PARTSREP_OR2 = "NULL"
                    PARTSDETAIL = "NULL"
                    PARTSDET_AMT = NumericVal(VARPARTAMOUNT)
                    PARTSDIS_VAL = Round(PARTSDISVAL * PARTSTAXRATE, 2)
                    PARTSDISCOUNT_2 = Round(PARTSDET_AMT * PARTSDISCRATE, 2)
                    PARTSTAXVAL = Round((PARTSDETAMT - PARTSDISCOUNT_2) * PARTSTAXRATE, 2)

                    SQL_STATEMENT = "insert into CSMS_RO_Det " & _
                                    "(rep_or,livil,LINE_NO,JOBTYPE,detcde,detdsc,detunt,detvol,detcost,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,REF_RIV_ADB)" & _
                                  " values (" & PARTSREP_OR & ", " & PARTSLEVEL & ", " & PARTSLINE_NO & "," & _
                                  " " & VGJORBP & "," & PARTSDETCDE & "," & PARTSDETDSC & "," & _
                                  " " & PARTSDETUNT & ", " & PARTSDETVOL & "," & _
                                  " " & PARTSDETCOST & ", " & PARTSDETPRC & ", " & PARTSDETAMT & ", " & PARTSCODE & _
                                    ", " & PARTSWCODE & ", " & PARTSTAXRATE * 100 & ", " & PARTSDISCRATE * 100 & _
                                    ", " & PARTSTAXVAL & ", " & PARTSDISVAL & ", " & PARTSPOCODE & _
                                    ", " & PARTSREP_OR2 & ", " & PARTSDETAIL & ", " & PARTSDET_AMT & _
                                    ", " & PARTSDIS_VAL & ", " & PARTSDISCOUNT_2 & "," & REF_RIV_ADB & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    Call NEW_LogAudit("AA", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREP_OR), "REP_OR", "CSMS_REPOR"), "", "TRAN NO: " & Null2String(RSCSMSORD_HD!TRANNO), "", "")
                    Screen.MousePointer = 0
20000               RSCSMSTDAYTRAN.MoveNext
                Loop
            End If
            RSCSMSORD_HD.MoveNext
        Loop
    End If
    Set RSCSMSORD_HD = New ADODB.Recordset
    Set RSCSMSORD_HD = gconDMIS.Execute("Select PARTICIPAT FROM CSMS_REPOR WHERE REP_OR = " & N2Str2Null(vREP_OR))
    If Not RSCSMSORD_HD.EOF And Not RSCSMSORD_HD.BOF Then
        If Null2String(RSCSMSORD_HD!PARTICIPAT) = "" Then
            FillJobs vREP_OR
            FillParts vREP_OR
            FillMaterials vREP_OR
            FillAccessories vREP_OR
            ROTOTAL = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
            SQL_STATEMENT = "update CSMS_RepOr set" & _
                          " Material = " & TOTMATAMT - TOTMATTAX & "," & _
                          " M_amtvalue = " & TOTPARTSAMT & "," & _
                          " M_disc = " & TOTMATDISCVAL & "," & _
                          " M_disc2 = " & TOTMATDISC * (VAT_RATE / 100) & "," & _
                          " M_taxval = " & TOTMATTAX & "," & _
                          " M_discount = " & TOTMATDISC & "," & _
                          " amount = " & ROTOTAL - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                          " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                          " wm_amt = " & 0 & "," & _
                          " ro_amount = " & ROTOTAL - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
                          " where rep_or = " & N2Str2Null(vREP_OR)
            gconDMIS.Execute SQL_STATEMENT

            Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREP_OR), "REP_OR", "CSMS_REPOR"), "", "TRAN NO: " & MTRANNO, "", "")
        Else
            UpdateParticipation vREP_OR
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub

Errorcode:
    Screen.MousePointer = 0: ShowVBError
End Sub


Sub ImportAccessories(vREP_OR As String)
    'On Error GoTo ERRORCODE
    Dim rsCSMS_REPOR                                   As ADODB.Recordset
    Set rsCSMS_REPOR = New ADODB.Recordset
    Set rsCSMS_REPOR = gconDMIS.Execute("Select DTE_COMP from CSMS_REPOR WHERE REP_OR = '" & vREP_OR & "'")
    If rsCSMS_REPOR.EOF And Not rsCSMS_REPOR.BOF Then
        If Null2Date(rsCSMS_REPOR!dte_comp) <> "" Then
            MsgBox "Repair Order is Already Billed. Transaction will not be imported!", vbInformation, vREP_OR & " Already Billed"
            Exit Sub
        End If
    End If
    Set rsCSMS_REPOR = Nothing

    Dim RONOFORMAT                                     As String
    Dim YZA                                            As Integer
    Dim TISOY, KEIKEI                                  As String
    RONOFORMAT = ""

    KEIKEI = "": TISOY = "": YZA = 0
    For YZA = 1 To Len(vREP_OR)
        TISOY = Mid(vREP_OR, YZA, 1)
        KEIKEI = KEIKEI + TISOY
    Next
    RONOFORMAT = KEIKEI
    Dim VARPARTSLINE_NO, VarPartNo, VARDESCRIPTION     As String
    Dim VARPARTCODE                                    As String
    Dim VARQTY, VARUNITCOST, VARUNITPRICE              As Double
    Dim VARPARTAMOUNT, VARCHARGETO, VARPARTDISCOUNT    As String

    Dim PARTSREP_OR, PARTSLEVEL, PARTSLINE_NO          As String
    Dim PARTSDETCDE, PARTSDETDSC, PARTSDETUNT          As String
    Dim PARTSDETVOL, PARTSDETPRC, PARTSDETAMT          As Double
    Dim PARTSCODE, PARTSWCODE                          As String
    Dim PARTSTAXRATE, PARTSDISCRATE, PARTSTAXVAL       As Double
    Dim PARTSDISVAL                                    As Double
    Dim PARTSPOCODE, PARTSREP_OR2, PARTSDETAIL         As String
    Dim PARTSDET_AMT, PARTSDETCOST, PARTSDIS_VAL, PARTSDISCOUNT_2 As Double
    Dim PARTSREMARKS, REF_RIV_ADB                      As String
    Dim RSRR_HDCHECK                                   As ADODB.Recordset
    Dim RSRR_HDTDAYTRANCHECK                           As ADODB.Recordset
    VPIS_NO_CHARGE_TO = "NULL"
    Dim VGJORBP                                        As String
    Dim MTRANNO                                        As String

    VGJORBP = "NULL"
    gconDMIS.Execute "delete from CSMS_RO_Det where livil = '4' and rep_or = " & N2Str2Null(vREP_OR)
    Pcnt = 0
    RO_RIV_TRANNO_COUNTER = 0
    Set RSCSMSORD_HIST = New ADODB.Recordset
    Set RSCSMSORD_HIST = gconDMIS.Execute("select rono,trandate,trantype,tranno,REFPISNO from PMIS_ord_hist where [TYPE] = 'A' AND status <> 'C' and status <> 'N' and rono = '" & RONOFORMAT & "'")
    If Not RSCSMSORD_HIST.EOF And Not RSCSMSORD_HIST.BOF Then
        RSCSMSORD_HIST.MoveFirst
        Do While Not RSCSMSORD_HIST.EOF
            If Mid(Null2String(RSCSMSORD_HIST!refpisno), 5, 1) = "C" Then VPIS_NO_CHARGE_TO = "NULL"
            If Mid(Null2String(RSCSMSORD_HIST!refpisno), 5, 1) = "I" Then VPIS_NO_CHARGE_TO = "'C'"
            If Mid(Null2String(RSCSMSORD_HIST!refpisno), 5, 1) = "W" Then VPIS_NO_CHARGE_TO = "'W'"

            If Mid(Null2String(RSCSMSORD_HIST!refpisno), 4, 1) = "B" Then VGJORBP = "'BP'"
            If Mid(Null2String(RSCSMSORD_HIST!refpisno), 4, 1) = "G" Then VGJORBP = "'GJ'"
            MTRANNO = Null2String(RSCSMSORD_HIST!TRANNO)

            Set RSCSMSDAYTRAN = New ADODB.Recordset
            Set RSCSMSDAYTRAN = gconDMIS.Execute("select itemno,trantype,tranno,stock_ord,stock_sup,tranqty,tranucost,tranuprice from PMIS_DayTran where [TYPE] = 'A' AND trantype = 'RIV' and tranno = " & N2Str2Null(RSCSMSORD_HIST!TRANNO) & " order by itemno asc")
            If Not RSCSMSDAYTRAN.EOF And Not RSCSMSDAYTRAN.BOF Then
                Screen.MousePointer = 11
                RSCSMSDAYTRAN.MoveFirst
                RO_RIV_TRANNO_COUNTER = RO_RIV_TRANNO_COUNTER + 1
                RO_RIV_TRANNO(RO_RIV_TRANNO_COUNTER) = Null2String(RSCSMSORD_HIST!TRANNO)
                Do While Not RSCSMSDAYTRAN.EOF
                    Pcnt = Pcnt + 1
                    VARPARTSLINE_NO = "": VarPartNo = "": VARDESCRIPTION = ""
                    VARPARTCODE = "": VARQTY = 0: VARUNITPRICE = 0
                    VARPARTAMOUNT = "": VARCHARGETO = " ": VARPARTDISCOUNT = ZERO

                    VARPARTSLINE_NO = Format(Pcnt, "00")
                    VarPartNo = Null2String(RSCSMSDAYTRAN!STOCK_ORD)
                    VARDESCRIPTION = SetSTOCKDESC(Null2String(RSCSMSDAYTRAN!STOCK_SUP))
                    VARPARTCODE = "01"
                    VARQTY = Format(N2Str2IntZero(RSCSMSDAYTRAN!tranqty), "####0")
                    VARUNITCOST = N2Str2Zero(RSCSMSDAYTRAN!tranucost)
                    VARUNITPRICE = N2Str2Zero(RSCSMSDAYTRAN!TRANUPRICE)
                    VARPARTAMOUNT = N2Str2Zero(RSCSMSDAYTRAN!tranqty) * N2Str2Zero(RSCSMSDAYTRAN!TRANUPRICE)
                    VARCHARGETO = " "
                    VARPARTDISCOUNT = ZERO
                    Set RSRR_HDCHECK = New ADODB.Recordset
                    Set RSRR_HDCHECK = gconDMIS.Execute("select * from PMIS_RR_HD where [TYPE] = 'A' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(RSCSMSORD_HIST!TRANNO), "000000") & "'")
                    If Not RSRR_HDCHECK.EOF And Not RSRR_HDCHECK.BOF Then
                        Set RSRR_HDTDAYTRANCHECK = New ADODB.Recordset
                        Set RSRR_HDTDAYTRANCHECK = gconDMIS.Execute("select * from PMIS_TdayTran where [TYPE] = 'A' AND trantype = 'RR' and tranno = " & N2Str2Null(RSRR_HDCHECK!RRNO) & " order by Itemno asc")
                        If Not RSRR_HDTDAYTRANCHECK.EOF And Not RSRR_HDTDAYTRANCHECK.BOF Then
                            RSRR_HDTDAYTRANCHECK.MoveNext
                            Do While Not RSRR_HDTDAYTRANCHECK.EOF
                                If Null2String(RSRR_HDTDAYTRANCHECK!STOCK_ORD) = VarPartNo Then GoTo 10000
                                RSRR_HDTDAYTRANCHECK.MoveNext
                            Loop
                        End If
                    End If
                    Set RSRR_HDCHECK = New ADODB.Recordset
                    Set RSRR_HDCHECK = gconDMIS.Execute("select * from PMIS_REC_HIST where [TYPE] = 'A' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(RSCSMSORD_HIST!TRANNO), "000000") & "'")
                    If Not RSRR_HDCHECK.EOF And Not RSRR_HDCHECK.BOF Then
                        Set RSRR_HDTDAYTRANCHECK = New ADODB.Recordset
                        Set RSRR_HDTDAYTRANCHECK = gconDMIS.Execute("select * from PMIS_DayTran where [TYPE] = 'A' AND trantype = 'RR' and tranno = " & N2Str2Null(RSRR_HDCHECK!RRNO) & " order by Itemno asc")
                        If Not RSRR_HDTDAYTRANCHECK.EOF And Not RSRR_HDTDAYTRANCHECK.BOF Then
                            RSRR_HDTDAYTRANCHECK.MoveNext
                            Do While Not RSRR_HDTDAYTRANCHECK.EOF
                                If Null2String(RSRR_HDTDAYTRANCHECK!PARTNO) = VarPartNo Then GoTo 10000
                                RSRR_HDTDAYTRANCHECK.MoveNext
                            Loop
                        End If
                    End If
                    REF_RIV_ADB = "'RIV" & Format(Null2String(RSCSMSDAYTRAN!TRANNO), "000000") & Format(Null2String(RSCSMSDAYTRAN!itemno), "000") & "'"
                    PARTSDISVAL = 0: PARTSTAXVAL = 0: PARTSDETAMT = 0
                    PARTSDIS_VAL = 0: PARTSDISCOUNT_2 = 0: PARTSDISCRATE = 0

                    PARTSREP_OR = N2Str2Null(vREP_OR)
                    PARTSLEVEL = "'4'"
                    PARTSLINE_NO = N2Str2Null(Format(VARPARTSLINE_NO, "00"))
                    PARTSDETCDE = N2Str2Null(VarPartNo)
                    PARTSDETDSC = N2Str2Null(Mid(VARDESCRIPTION, 1, 100))
                    PARTSDETUNT = "NULL"
                    PARTSDETVOL = N2Str2Zero(VARQTY)
                    PARTSDETCOST = NumericVal(VARUNITCOST)
                    PARTSDETPRC = NumericVal(VARUNITPRICE)
                    PARTSDETAMT = Round(NumericVal(VARPARTAMOUNT) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                    PARTSCODE = "NULL"
                    PARTSWCODE = VPIS_NO_CHARGE_TO
                    PARTSTAXRATE = (VAT_RATE / 100)
                    PARTSDISCRATE = NumericVal(VARPARTDISCOUNT) / 100
                    PARTSDISVAL = Round((PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE), 2)
                    PARTSPOCODE = N2Str2Null(VARPARTCODE)
                    PARTSREP_OR2 = "NULL"
                    PARTSDETAIL = "NULL"
                    PARTSDET_AMT = NumericVal(VARPARTAMOUNT)
                    PARTSDIS_VAL = Round(PARTSDISVAL * PARTSTAXRATE, 2)
                    PARTSDISCOUNT_2 = Round(PARTSDET_AMT * PARTSDISCRATE, 2)
                    PARTSTAXVAL = Round((PARTSDETAMT - PARTSDISCOUNT_2) * PARTSTAXRATE, 2)

                    SQL_STATEMENT = "insert into CSMS_RO_Det " & _
                                    "(rep_or,livil,LINE_NO,JOBTYPE,detcde,detdsc,detunt,detvol,detcost,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,REF_RIV_ADB)" & _
                                  " values (" & PARTSREP_OR & ", " & PARTSLEVEL & ", " & PARTSLINE_NO & "," & _
                                  " " & VGJORBP & ", " & PARTSDETCDE & "," & PARTSDETDSC & "," & _
                                  " " & PARTSDETUNT & ", " & PARTSDETVOL & "," & _
                                  " " & PARTSDETCOST & ", " & PARTSDETPRC & ", " & PARTSDETAMT & ", " & PARTSCODE & _
                                    ", " & PARTSWCODE & ", " & PARTSTAXRATE * 100 & ", " & PARTSDISCRATE * 100 & _
                                    ", " & PARTSTAXVAL & ", " & PARTSDISVAL & ", " & PARTSPOCODE & _
                                    ", " & PARTSREP_OR2 & ", " & PARTSDETAIL & ", " & PARTSDET_AMT & _
                                    ", " & PARTSDIS_VAL & ", " & PARTSDISCOUNT_2 & "," & REF_RIV_ADB & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    Call NEW_LogAudit("AA", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREP_OR), "REP_OR", "CSMS_REPOR"), "", "TRAN NO: " & Null2String(RSCSMSORD_HIST!TRANNO), "HISTORY", "")

                    Screen.MousePointer = 0
10000               RSCSMSDAYTRAN.MoveNext
                Loop
            End If
            RSCSMSORD_HIST.MoveNext
        Loop
    End If

    Set RSCSMSORD_HD = New ADODB.Recordset
    Set RSCSMSORD_HD = gconDMIS.Execute("select rono,tranno,trantype,REFPISNO from PMIS_ord_hd where [TYPE] = 'A' AND status <> 'C' and status <> 'N' and trantype = 'RIV' and rono = '" & RONOFORMAT & "'")
    If Not RSCSMSORD_HD.EOF And Not RSCSMSORD_HD.BOF Then
        RSCSMSORD_HD.MoveFirst
        Do While Not RSCSMSORD_HD.EOF
            If Mid(Null2String(RSCSMSORD_HD!refpisno), 5, 1) = "C" Then VPIS_NO_CHARGE_TO = "NULL"
            If Mid(Null2String(RSCSMSORD_HD!refpisno), 5, 1) = "I" Then VPIS_NO_CHARGE_TO = "'C'"
            If Mid(Null2String(RSCSMSORD_HD!refpisno), 5, 1) = "W" Then VPIS_NO_CHARGE_TO = "'W'"

            If Mid(Null2String(RSCSMSORD_HD!refpisno), 4, 1) = "B" Then VGJORBP = "'BP'"
            If Mid(Null2String(RSCSMSORD_HD!refpisno), 4, 1) = "G" Then VGJORBP = "'GJ'"
            MTRANNO = Null2String(RSCSMSORD_HD!TRANNO)

            Set RSCSMSTDAYTRAN = New ADODB.Recordset
            Set RSCSMSTDAYTRAN = gconDMIS.Execute("select itemno,trantype,tranno,STOCK_ord,STOCK_sup,tranqty,tranucost,tranuprice from PMIS_TdayTran where [TYPE] = 'A' AND trantype = 'RIV' and tranno = " & N2Str2Null(RSCSMSORD_HD!TRANNO) & " order by itemno asc")
            If Not RSCSMSTDAYTRAN.EOF And Not RSCSMSTDAYTRAN.BOF Then
                Screen.MousePointer = 11
                RSCSMSTDAYTRAN.MoveFirst
                RO_RIV_TRANNO_COUNTER = RO_RIV_TRANNO_COUNTER + 1
                RO_RIV_TRANNO(RO_RIV_TRANNO_COUNTER) = Null2String(RSCSMSORD_HD!TRANNO)
                Do While Not RSCSMSTDAYTRAN.EOF
                    Pcnt = Pcnt + 1
                    VARPARTSLINE_NO = "": VarPartNo = "": VARDESCRIPTION = ""
                    VARPARTCODE = "": VARQTY = 0: VARUNITPRICE = 0
                    VARPARTAMOUNT = "": VARCHARGETO = " ": VARPARTDISCOUNT = ZERO

                    VARPARTSLINE_NO = Format(Pcnt, "00")
                    VarPartNo = Null2String(RSCSMSTDAYTRAN!STOCK_ORD)
                    VARDESCRIPTION = SetSTOCKDESC(Null2String(RSCSMSTDAYTRAN!STOCK_SUP))
                    VARPARTCODE = "01"
                    VARQTY = Format(N2Str2IntZero(RSCSMSTDAYTRAN!tranqty), "####0")
                    VARUNITCOST = N2Str2Zero(RSCSMSTDAYTRAN!tranucost)
                    VARUNITPRICE = N2Str2Zero(RSCSMSTDAYTRAN!TRANUPRICE)
                    VARPARTAMOUNT = N2Str2Zero(RSCSMSTDAYTRAN!tranqty) * N2Str2Zero(RSCSMSTDAYTRAN!TRANUPRICE)
                    VARCHARGETO = " "
                    VARPARTDISCOUNT = ZERO
                    Set RSRR_HDCHECK = New ADODB.Recordset
                    Set RSRR_HDCHECK = gconDMIS.Execute("select * from PMIS_RR_HD where [TYPE] = 'A' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(RSCSMSORD_HD!TRANNO), "000000") & "'")
                    If Not RSRR_HDCHECK.EOF And Not RSRR_HDCHECK.BOF Then
                        Set RSRR_HDTDAYTRANCHECK = New ADODB.Recordset
                        Set RSRR_HDTDAYTRANCHECK = gconDMIS.Execute("select * from PMIS_Tdaytran where [TYPE] = 'A' AND trantype = 'RR' and tranno = " & N2Str2Null(RSRR_HDCHECK!RRNO) & " order by Itemno asc")
                        If Not RSRR_HDTDAYTRANCHECK.EOF And Not RSRR_HDTDAYTRANCHECK.BOF Then
                            RSRR_HDTDAYTRANCHECK.MoveNext
                            Do While Not RSRR_HDTDAYTRANCHECK.EOF
                                If Null2String(RSRR_HDTDAYTRANCHECK!PARTNO) = VarPartNo Then GoTo 20000
                                RSRR_HDTDAYTRANCHECK.MoveNext
                            Loop
                        End If
                    End If
                    Set RSRR_HDCHECK = New ADODB.Recordset
                    Set RSRR_HDCHECK = gconDMIS.Execute("select * from PMIS_REC_HIST where [TYPE] = 'A' AND ClassCode = 'RRV' and RIV_Tranno = '" & Format(Null2String(RSCSMSORD_HD!TRANNO), "000000") & "'")
                    If Not RSRR_HDCHECK.EOF And Not RSRR_HDCHECK.BOF Then
                        Set RSRR_HDTDAYTRANCHECK = New ADODB.Recordset
                        Set RSRR_HDTDAYTRANCHECK = gconDMIS.Execute("select * from PMIS_DayTran where [TYPE] = 'A' AND trantype = 'RR' and tranno = " & N2Str2Null(RSRR_HDCHECK!RRNO) & " order by Itemno asc")
                        If Not RSRR_HDTDAYTRANCHECK.EOF And Not RSRR_HDTDAYTRANCHECK.BOF Then
                            RSRR_HDTDAYTRANCHECK.MoveNext
                            Do While Not RSRR_HDTDAYTRANCHECK.EOF
                                If Null2String(RSRR_HDTDAYTRANCHECK!PARTNO) = VarPartNo Then GoTo 20000
                                RSRR_HDTDAYTRANCHECK.MoveNext
                            Loop
                        End If
                    End If
                    REF_RIV_ADB = "'RIV" & Format(Null2String(RSCSMSTDAYTRAN!TRANNO), "000000") & Format(Null2String(RSCSMSTDAYTRAN!itemno), "000") & "'"
                    PARTSDISVAL = 0: PARTSTAXVAL = 0: PARTSDETAMT = 0
                    PARTSDIS_VAL = 0: PARTSDISCOUNT_2 = 0: PARTSDISCRATE = 0

                    PARTSREP_OR = N2Str2Null(vREP_OR)
                    PARTSLEVEL = "'4'"
                    PARTSLINE_NO = N2Str2Null(Format(VARPARTSLINE_NO, "00"))
                    PARTSDETCDE = N2Str2Null(VarPartNo)
                    PARTSDETDSC = N2Str2Null(Mid(VARDESCRIPTION, 1, 100))
                    PARTSDETUNT = "NULL"
                    PARTSDETVOL = N2Str2Zero(VARQTY)
                    PARTSDETCOST = NumericVal(VARUNITCOST)
                    PARTSDETPRC = NumericVal(VARUNITPRICE)
                    PARTSDETAMT = Round(NumericVal(VARPARTAMOUNT) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                    PARTSCODE = "NULL"
                    PARTSWCODE = VPIS_NO_CHARGE_TO
                    PARTSTAXRATE = (VAT_RATE / 100)
                    PARTSDISCRATE = Round(NumericVal(VARPARTDISCOUNT) / 100, 2)
                    PARTSDISVAL = Round((PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE), 2)
                    PARTSPOCODE = N2Str2Null(VARPARTCODE)
                    PARTSREP_OR2 = "NULL"
                    PARTSDETAIL = "NULL"
                    PARTSDET_AMT = NumericVal(VARPARTAMOUNT)
                    PARTSDIS_VAL = Round(PARTSDISVAL * PARTSTAXRATE, 2)
                    PARTSDISCOUNT_2 = Round(PARTSDET_AMT * PARTSDISCRATE, 2)
                    PARTSTAXVAL = Round((PARTSDETAMT - PARTSDISCOUNT_2) * PARTSTAXRATE, 2)

                    SQL_STATEMENT = "insert into CSMS_RO_Det " & _
                                    "(rep_or,livil,LINE_NO,JOBTYPE,detcde,detdsc,detunt,detvol,detcost,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,REF_RIV_ADB)" & _
                                  " values (" & PARTSREP_OR & ", " & PARTSLEVEL & ", " & PARTSLINE_NO & "," & _
                                  " " & VGJORBP & ", " & PARTSDETCDE & "," & PARTSDETDSC & "," & _
                                  " " & PARTSDETUNT & ", " & PARTSDETVOL & "," & _
                                  " " & PARTSDETCOST & ", " & PARTSDETPRC & ", " & PARTSDETAMT & ", " & PARTSCODE & _
                                    ", " & PARTSWCODE & ", " & PARTSTAXRATE * 100 & ", " & PARTSDISCRATE * 100 & _
                                    ", " & PARTSTAXVAL & ", " & PARTSDISVAL & ", " & PARTSPOCODE & _
                                    ", " & PARTSREP_OR2 & ", " & PARTSDETAIL & ", " & PARTSDET_AMT & _
                                    ", " & PARTSDIS_VAL & ", " & PARTSDISCOUNT_2 & "," & REF_RIV_ADB & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    Call NEW_LogAudit("AA", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREP_OR), "REP_OR", "CSMS_REPOR"), "", "TRAN NO: " & Null2String(RSCSMSORD_HD!TRANNO), "", "")

                    Screen.MousePointer = 0
20000               RSCSMSTDAYTRAN.MoveNext
                Loop
            End If
            RSCSMSORD_HD.MoveNext
        Loop
    End If

    Set RSCSMSORD_HD = New ADODB.Recordset
    Set RSCSMSORD_HD = gconDMIS.Execute("Select PARTICIPAT FROM CSMS_REPOR WHERE REP_OR = " & N2Str2Null(vREP_OR))
    If Not RSCSMSORD_HD.EOF And Not RSCSMSORD_HD.BOF Then
        If Null2String(RSCSMSORD_HD!PARTICIPAT) = "" Then
            FillJobs vREP_OR
            FillParts vREP_OR
            FillMaterials vREP_OR
            FillAccessories vREP_OR
            ROTOTAL = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
            SQL_STATEMENT = "update CSMS_RepOr set" & _
                          " Accessories = " & TOTACCAMT - TOTACCTAX & "," & _
                          " A_amtvalue = " & TOTACCAMT & "," & _
                          " A_disc = " & TOTACCDISCVAL & "," & _
                          " A_disc2 = " & TOTACCDISC * (VAT_RATE / 100) & "," & _
                          " A_taxval = " & TOTACCTAX & "," & _
                          " A_discount = " & TOTACCDISC & "," & _
                          " amount = " & ROTOTAL - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & "," & _
                          " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX & "," & _
                          " wa_amt = " & 0 & "," & _
                          " ro_amount = " & ROTOTAL - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC - TOTACCDISC & _
                          " where rep_or = " & N2Str2Null(vREP_OR)
            gconDMIS.Execute SQL_STATEMENT
            Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(vREP_OR), "REP_OR", "CSMS_REPOR"), "", "TRAN NO: " & MTRANNO & " POSTED", "", "")
        Else
            UpdateParticipation vREP_OR
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub


Errorcode:
    Screen.MousePointer = 0: ShowVBError
End Sub

Function SetSTOCKDESC(ppp As String)
    Dim RSPARTMAS                                      As ADODB.Recordset
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select STOCKNO,STOCKDESC,srp,mac,dnp from PMIS_STOCKMAS where STOCKNO= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKDESC = Null2String(RSPARTMAS!STOCKDESC)
    End If
End Function


Function CheckORNum(YYY As String, XCOUNTERTYPE, InvType As String) As String
    Dim rsCMIS_OFF_DT                                  As ADODB.Recordset
    Set rsCMIS_OFF_DT = New ADODB.Recordset
    Set rsCMIS_OFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE INVOICETYPE = '" & InvType & "' AND TRANTYPE = '" & XCOUNTERTYPE & "' AND INVOICENO = '" & YYY & "' AND isnull(STATUS,'') <>'C'")
    If Not rsCMIS_OFF_DT.EOF And Not rsCMIS_OFF_DT.BOF Then
        CheckORNum = UCase(Null2String(rsCMIS_OFF_DT!OR_NUM))
    End If
    Set rsCMIS_OFF_DT = Nothing
End Function

Function CheckSJNum(YYY As String, XCOUNTERTYPE) As String
    Dim rsAMIS_JournalSJ                               As ADODB.Recordset
    Set rsAMIS_JournalSJ = New ADODB.Recordset
    If COUNTERTYPE = "CHG" Or COUNTERTYPE = "CSH" Then
        Set rsAMIS_JournalSJ = gconDMIS.Execute("Select * from AMIS_JOURNAL_HD WHERE INVOICETYPE = '" & XCOUNTERTYPE & "' AND INVOICENO = '" & YYY & "' AND isnull(STATUS,'')<>'C' AND PAYTYPE='" & COUNTERTYPE & "'")
    Else
        Set rsAMIS_JournalSJ = gconDMIS.Execute("Select * from AMIS_JOURNAL_HD WHERE INVOICETYPE = '" & XCOUNTERTYPE & "' AND INVOICENO = '" & YYY & "' AND isnull(STATUS,'')<>'C'")
    End If
    If Not rsAMIS_JournalSJ.EOF And Not rsAMIS_JournalSJ.BOF Then
        CheckSJNum = UCase(Null2String(rsAMIS_JournalSJ!VOUCHERNO))
    End If
    Set rsAMIS_JournalSJ = Nothing
End Function

Function CheckAPJNum(YYY As String, XCOUNTERTYPE) As String
    Dim rsAMIS_JournalSJ                               As ADODB.Recordset
    Set rsAMIS_JournalSJ = New ADODB.Recordset
    Set rsAMIS_JournalSJ = gconDMIS.Execute("Select VOUCHERNO from amis_pv_detail  WHERE JTYPE = '" & XCOUNTERTYPE & "' AND MRR_NO = '" & YYY & "' AND STATUS<>'C'")
    If Not rsAMIS_JournalSJ.EOF And Not rsAMIS_JournalSJ.BOF Then
        CheckAPJNum = UCase(Null2String(rsAMIS_JournalSJ!VOUCHERNO))
    End If
    Set rsAMIS_JournalSJ = Nothing
End Function

Public Function flex_FillReportView(RS As ADODB.Recordset, grd As XtremeReportControl.ReportControl, Optional ByVal WithSN As Boolean = False)
    Dim fld                                            As ADODB.Field
    Dim J                                              As Long
    Dim REC                                            As XtremeReportControl.ReportRecord

    grd.Records.DeleteAll

    While Not RS.EOF
        J = J + 1

        Set REC = grd.Records.Add
        If WithSN = True Then
            REC.AddItem J
        End If
        For Each fld In RS.Fields
            REC.AddItem (Trim(fld.Value))
        Next
        RS.MoveNext
    Wend
    grd.Populate
    Set fld = Nothing
    Set REC = Nothing
    Set RS = Nothing
End Function

Public Sub AddColumnHeader(StringHeaders As String, LST As Object)
    Dim ar()                                           As String
    Dim cWidth                                         As Long
    Dim I                                              As Integer
    ar = Split(StringHeaders, ",")
    If TypeOf LST Is ListView Then
        cWidth = LST.Width
        LST.ColumnHeaders.Clear
        For I = LBound(ar) To UBound(ar)
            LST.ColumnHeaders.Add , , ar(I)
        Next
    ElseIf TypeOf LST Is ReportControl Then
        LST.Columns.DeleteAll
        For I = LBound(ar) To UBound(ar)
            LST.Columns.Add I, ar(I), 100, True
        Next
    End If

    Erase ar
    StringHeaders = vbNullString
End Sub
Public Sub ResizeColumnHeader(grd As Object, SizeArray As String)
    grd.Visible = False

    Dim ar()                                           As String
    Dim cWidth                                         As Long
    Dim I                                              As Integer
    Dim scwidth                                        As Long
    ar = Split(SizeArray, ",")
    cWidth = grd.Width

    If TypeOf grd Is ListView Then
        For I = LBound(ar) To UBound(ar)
            If I <= grd.ColumnHeaders.Count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.ColumnHeaders(I + 1).Width = scwidth
            End If
        Next
    ElseIf TypeOf grd Is ReportControl Then
        For I = LBound(ar) To UBound(ar)
            If I < grd.Columns.Count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.Columns(I).Width = scwidth
            End If
        Next

    End If

    Erase ar
    grd.Visible = True
End Sub
Sub flex_FillReportPaintManager(LST As ReportControl)
    With LST
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        '.PaintManager.HighlightBackColor = RGB(34, 133, 13)
        .PaintManager.ShadeSortColor = RGB(250, 251, 189)
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
        .PaintManager.GroupRowTextBold = True
        .PaintManager.GroupForeColor = vbBlue
        .PaintManager.ColumnStyle = xtpColumnOffice2003

    End With
End Sub

Public Sub FlexGrid_To_Excel(mgrdPhyCnt As MSFlexGrid, TheRows As Integer, TheCols As Integer, Optional GridStyle As Integer = 1, Optional WorkSheetName As String)

    Dim objXL                                          As New Excel.Application
    Dim wbXL                                           As New Excel.Workbook
    Dim wsXL                                           As New Excel.Worksheet
    Dim intRow                                         As Integer    ' counter
    Dim intCol                                         As Integer    ' counter

    If Not IsObject(objXL) Then
        MsgBox "You need Microsoft Excel to use this function", _
               vbExclamation, "Print to Excel"
        Exit Sub
    End If


    On Error Resume Next

    Set wbXL = objXL.Workbooks.Add
    Set wsXL = objXL.ActiveSheet


    With wsXL
        If Not WorkSheetName = "" Then
            .Name = WorkSheetName
        End If
    End With
    'update by: NVB
    'Update date: 01/27/08
    'Description: Put an header for the excel printing
    '               elimate $ sign in every column

    TheRows = TheRows + 3                             'adding 3 rows, Use in inserting row for headings

    For intRow = 1 To TheRows
        For intCol = 1 To TheCols
            With mgrdPhyCnt
                If intRow = 1 And intCol = 1 Then     ' this condition is set to Add header
                    With wsXL

                        wsXL.Range("A" & 1 & ":" & "C" & 1).Merge

                        Dim I                          As Integer
                        For I = 2 To 3
                            wsXL.Range("A" & I & ":" & "F" & I).Merge
                        Next
                        wsXL.Cells(1, "A") = "Company Name:    " & COMPANY_NAME
                        wsXL.Cells(2, "A") = "Compnay Address: " & COMPANY_ADDRESS
                    End With
                    intRow = intRow + 2
                Else

                    wsXL.Cells(intRow, intCol).Value = _
                    "" & CStr(.TextMatrix(intRow - 4, intCol - 1)) & "  "
                End If
            End With
        Next
    Next


    For intCol = 1 To TheCols
        'change Auto Format of data in worksheet
        wsXL.Columns(intCol).AutoFit
        'wsXL.Range("A1", Right(wsXL.Columns(TheCols).AddressLocal, 1) & TheRows).AutoFormat GridStyle
        wsXL.Range("A1", Right(wsXL.Columns(TheCols).AddressLocal, 1) & TheRows).AutoFormat xlRangeAutoFormatClassic2

    Next
    '------------------------------------------
    objXL.Visible = True
End Sub
Sub FormExistsShow(frmx As Form)

    On Error GoTo Errorcode
    Dim m_Exists                                       As Boolean
    Dim FRM                                            As Form
    frmx.Show
    For Each FRM In Forms
        If (UCase(FRM.Name) = UCase(frmx.Name)) Then
            m_Exists = True
            Exit For
        End If
    Next
    
    
    
    Set FRM = Nothing

    If m_Exists = True Then
        frmx.WindowState = 0
        frmx.ZOrder 0
    
    
    
    End If

    Exit Sub
Errorcode:
    err.Clear
End Sub




'Function ComputeStockMasMac(XX As String) As Double
'    Dim RSMAC                                          As ADODB.Recordset
'    Dim rsSTOCKMAS                                     As ADODB.Recordset
'    Dim MACX                                           As Double
'    Dim Qty                                            As Long
'    Dim UNITCOST                                       As Double
'    Dim LINECOST                                       As Double
'    Dim INVENTORYAMOUNT                                As Double
'    Dim BALANCE                                        As Long
'    Dim COMPUTEDMAC                                    As Double
'
'
'    Set RSMAC = gconDMIS.Execute("SELECT IN_OUT, TRANNO,TRANDATE,TRANTYPE , TRANQTY, TRANUCOST FROM PMIS_ALLDAYTRAN WHERE TRANTYPE IN('BEG','RR','RIV','CSH','CHG','DR','ADJ') AND STOCK_ORD='" & XX & " ' AND STATUS in('P','B') ORDER BY TRANDATE,ID ASC")
'    While Not RSMAC.EOF
'        UNITCOST = 0: LINECOST = 0: Qty = 0
'        Qty = N2Str2IntZero(RSMAC!tranqty)
'        Select Case UCase(Null2String(RSMAC!TranType))
'            Case "BEG", "RR"
'                UNITCOST = NumericVal(RSMAC!tranucost)
'                LINECOST = Qty * UNITCOST
'                INVENTORYAMOUNT = INVENTORYAMOUNT + LINECOST
'                BALANCE = BALANCE + Qty
'                If BALANCE > 0 Then
'                    COMPUTEDMAC = INVENTORYAMOUNT / BALANCE
'                Else
'                    COMPUTEDMAC = UNITCOST
'                End If
'            Case "DR", "RIV", "CSH", "CHG"
'                LINECOST = Qty * COMPUTEDMAC
'                INVENTORYAMOUNT = INVENTORYAMOUNT - LINECOST
'                BALANCE = BALANCE - Qty
'            Case "ADJ"
'                If UCase(Null2String(RSMAC!TRANNO)) = "111111" And UCase(Null2String(RSMAC!IN_OUT)) = "I" Then
'                    UNITCOST = NumericVal(RSMAC!tranucost)
'                    LINECOST = Qty * UNITCOST
'                    INVENTORYAMOUNT = INVENTORYAMOUNT + LINECOST
'                    BALANCE = BALANCE + Qty
'                    If BALANCE > 0 Then
'                        COMPUTEDMAC = INVENTORYAMOUNT / BALANCE
'                    Else
'                        COMPUTEDMAC = UNITCOST
'                    End If
'
'                ElseIf UCase(Null2String(RSMAC!TRANNO)) = "000000" And UCase(Null2String(RSMAC!IN_OUT)) = "O" Then
'                    LINECOST = Qty * COMPUTEDMAC
'                    INVENTORYAMOUNT = INVENTORYAMOUNT - LINECOST
'                    BALANCE = BALANCE - Qty
'                End If
'        End Select
'        RSMAC.MoveNext
'    Wend
'    ComputeStockMasMac = Round(COMPUTEDMAC, 2)
'End Function


'
'Function ComputeMacasofDate(XX As String, str_TRANDATE As String) As Double
'     Dim RSMAC                                          As ADODB.Recordset
'    Dim rsSTOCKMAS                                     As ADODB.Recordset
'    Dim MACX                                           As Double
'    Dim Qty                                            As Long
'    Dim UNITCOST                                       As Double
'    Dim LINECOST                                       As Double
'    Dim INVENTORYAMOUNT                                As Double
'    Dim BALANCE                                        As Long
'    Dim COMPUTEDMAC                                    As Double
'
'
'    Set RSMAC = gconDMIS.Execute("SELECT IN_OUT, TRANNO,TRANDATE,TRANTYPE , TRANQTY, TRANUCOST FROM PMIS_ALLDAYTRAN WHERE TRANTYPE IN('BEG','RR','RIV','CSH','CHG','DR','ADJ') AND STOCK_ORD='" & XX & " ' AND STATUS in('P','B') AND TRANDATE <=" & N2Str2Null(str_TRANDATE) & " ORDER BY TRANDATE,ID ASC")
'    While Not RSMAC.EOF
'        UNITCOST = 0: LINECOST = 0: Qty = 0
'        Qty = N2Str2IntZero(RSMAC!tranqty)
'        Select Case UCase(Null2String(RSMAC!TranType))
'            Case "BEG", "RR"
'                UNITCOST = NumericVal(RSMAC!tranucost)
'                LINECOST = Qty * UNITCOST
'                INVENTORYAMOUNT = INVENTORYAMOUNT + LINECOST
'                BALANCE = BALANCE + Qty
'                If BALANCE > 0 Then
'                    COMPUTEDMAC = INVENTORYAMOUNT / BALANCE
'                Else
'                    COMPUTEDMAC = UNITCOST
'                End If
'            Case "DR", "RIV", "CSH", "CHG"
'                LINECOST = Qty * COMPUTEDMAC
'                INVENTORYAMOUNT = INVENTORYAMOUNT - LINECOST
'                BALANCE = BALANCE - Qty
'            Case "ADJ"
'                If UCase(Null2String(RSMAC!TRANNO)) = "111111" And UCase(Null2String(RSMAC!IN_OUT)) = "I" Then
'                    UNITCOST = NumericVal(RSMAC!tranucost)
'                    LINECOST = Qty * UNITCOST
'                    INVENTORYAMOUNT = INVENTORYAMOUNT + LINECOST
'                    BALANCE = BALANCE + Qty
'                    If BALANCE > 0 Then
'                        COMPUTEDMAC = INVENTORYAMOUNT / BALANCE
'                    Else
'                        COMPUTEDMAC = UNITCOST
'                    End If
'
'                ElseIf UCase(Null2String(RSMAC!TRANNO)) = "000000" And UCase(Null2String(RSMAC!IN_OUT)) = "O" Then
'                    LINECOST = Qty * COMPUTEDMAC
'                    INVENTORYAMOUNT = INVENTORYAMOUNT - LINECOST
'                    BALANCE = BALANCE - Qty
'                End If
'        End Select
'        RSMAC.MoveNext
'    Wend
'    MsgBox COMPUTEDMAC
'    ComputeMacasofDate = Round(COMPUTEDMAC, 2)
'
'
'End Function


Function ComputeTransactionMac(xx As String, added_Qty As Long, added_UnitCost As Double, str_TRANDATE As String) As Double
    'Function ComputeTransactionMac(XX As String, added_Qty As Long, added_UnitCost As Double, LNID As Long) As Double

    Dim RSMAC                                          As ADODB.Recordset
    Dim rsSTOCKMAS                                     As ADODB.Recordset
    Dim MACX                                           As Double
    Dim QTY                                            As Long
    Dim UNITCOST                                       As Double
    Dim LINECOST                                       As Double
    Dim INVENTORYAMOUNT                                As Double
    Dim BALANCE                                        As Long
    Dim COMPUTEDMAC                                    As Double


    Set RSMAC = gconDMIS.Execute("SELECT IN_OUT, TRANNO,TRANDATE,TRANTYPE , TRANQTY, TRANUCOST FROM PMIS_ALLDAYTRAN WHERE TRANTYPE IN('BEG','RR','RIV','CSH','CHG','DR','ADJ') AND STOCK_ORD='" & xx & " ' AND STATUS in('P','B') and TRANDATE <='" & str_TRANDATE & "'  ORDER BY TRANDATE,ID ASC")
    'Set RSMAC = gconDMIS.Execute("SELECT IN_OUT, TRANNO,TRANDATE,TRANTYPE , TRANQTY, TRANUCOST FROM PMIS_ALLDAYTRAN WHERE TRANTYPE IN('BEG','RR','RIV','CSH','CHG','DR','ADJ') AND STOCK_ORD='" & XX & " ' AND STATUS in('P','B') and ID <=" & LNID & "  ORDER BY TRANDATE,ID ASC")
    'INVENTORYAMOUNT = added_UnitCost * added_Qty
    'BALANCE = added_Qty
    While Not RSMAC.EOF
        UNITCOST = 0: LINECOST = 0: QTY = 0
        QTY = N2Str2IntZero(RSMAC!tranqty)
        Select Case UCase(Null2String(RSMAC!TranType))
            Case "BEG", "RR"
                UNITCOST = N2Str2IntZero(RSMAC!tranucost)
                LINECOST = QTY * UNITCOST
                INVENTORYAMOUNT = INVENTORYAMOUNT + LINECOST
                BALANCE = BALANCE + QTY
                If BALANCE > 0 Then
                    COMPUTEDMAC = INVENTORYAMOUNT / BALANCE
                Else
                    COMPUTEDMAC = UNITCOST
                End If
            Case "DR", "RIV", "CSH", "CHG"
                LINECOST = QTY * COMPUTEDMAC
                INVENTORYAMOUNT = INVENTORYAMOUNT - LINECOST
                BALANCE = BALANCE - QTY
            Case "ADJ"
                If UCase(Null2String(RSMAC!TRANNO)) = "111111" And UCase(Null2String(RSMAC!IN_OUT)) = "I" Then
                    UNITCOST = N2Str2IntZero(RSMAC!tranucost)
                    LINECOST = QTY * UNITCOST
                    INVENTORYAMOUNT = INVENTORYAMOUNT + LINECOST
                    BALANCE = BALANCE + QTY
                    If BALANCE > 0 Then
                        COMPUTEDMAC = INVENTORYAMOUNT / BALANCE
                    Else
                        COMPUTEDMAC = UNITCOST
                    End If

                ElseIf UCase(Null2String(RSMAC!TRANNO)) = "000000" And UCase(Null2String(RSMAC!IN_OUT)) = "O" Then
                    LINECOST = QTY * COMPUTEDMAC
                    INVENTORYAMOUNT = INVENTORYAMOUNT - LINECOST
                    BALANCE = BALANCE - QTY
                End If
        End Select

        RSMAC.MoveNext
    Wend
    
    If BALANCE > 0 Then
        COMPUTEDMAC = (INVENTORYAMOUNT + (added_Qty * added_UnitCost)) / (BALANCE + added_Qty)
         
    Else
         COMPUTEDMAC = added_UnitCost
     End If

MsgBox COMPUTEDMAC
    ComputeTransactionMac = Round(COMPUTEDMAC, 2)


End Function

Public Sub SetComboWidth(c As ComboBox, xWidth As Long)
    Call SendMessage(c.hwnd, CB_SETDROPPEDWIDTH, xWidth, 0)
End Sub

Function COMPUTE_ONHANDASOFDATE(str_TRANDATE As String, str_Stockno As String, str_type As String) As Long
    Dim SQL                                            As String
    Dim RSTOTAL                                        As ADODB.Recordset

    SQL = "DECLARE @STOCKNO NVARCHAR(30)  " & vbCrLf
    SQL = SQL & "DECLARE @TYPE NVARCHAR(1) " & vbCrLf
    SQL = SQL & "DECLARE @TRANDATE SMALLDATETIME " & vbCrLf
    SQL = SQL & "SET @STOCKNO='" & str_Stockno & "' " & vbCrLf
    SQL = SQL & "SET @TRANDATE='" & str_TRANDATE & "' " & vbCrLf
    SQL = SQL & "SET @TYPE='" & str_type & "' " & vbCrLf
    SQL = SQL & "SELECT SUM(TRANQTY) AS ONHANDASOF FROM( " & vbCrLf
    SQL = SQL & "SELECT 'BEG' AS TRANTYPE ,  1 * ISNULL(SUM(TRANQTY),0) AS TRANQTY  FROM PMIS_DAYTRAN WHERE STOCK_ORD=@STOCKNO AND TRANTYPE='BEG' AND TYPE=@TYPE AND STATUS IN('P','B') AND TRANDATE<=@TRANDATE " & vbCrLf
    SQL = SQL & "Union " & vbCrLf
    SQL = SQL & "SELECT 'ADJ-IN' ,  1 * ISNULL(ISNULL(SUM(TRANQTY),0),0) AS TRANQTY    FROM PMIS_ALLDAYTRAN WHERE STOCK_ORD=@STOCKNO AND TRANTYPE='ADJ' AND TRANNO='111111' AND STATUS IN('P','B')   AND TRANDATE<=@TRANDATE " & vbCrLf
    SQL = SQL & "Union " & vbCrLf
    SQL = SQL & "SELECT 'RR' ,  1 * ISNULL(SUM(TRANQTY),0) AS TRANQTY      FROM PMIS_ALLDAYTRAN WHERE STOCK_ORD=@STOCKNO AND TRANTYPE='RR' AND TYPE=@TYPE AND STATUS IN('P','B')  AND TRANDATE<=@TRANDATE " & vbCrLf
    SQL = SQL & "Union " & vbCrLf
    SQL = SQL & "SELECT 'ADJ-OUT' , -1 * ISNULL(SUM(TRANQTY),0) AS TRANQTY    FROM PMIS_ALLDAYTRAN WHERE STOCK_ORD=@STOCKNO AND TRANTYPE='ADJ' AND TRANNO='000000' AND STATUS IN('P','B')    AND TRANDATE<=@TRANDATE " & vbCrLf
    SQL = SQL & "Union " & vbCrLf
    SQL = SQL & "SELECT 'ISS' , -1 * ISNULL(SUM(TRANQTY),0) AS TRANQTY    FROM PMIS_ALLDAYTRAN WHERE STOCK_ORD=@STOCKNO AND TRANTYPE IN('CSH','CHG','RIV','DR') AND TYPE=@TYPE AND STATUS IN('P','B')  AND TRANDATE<=@TRANDATE " & vbCrLf
    SQL = SQL & ") T"

    Set RSTOTAL = gconDMIS.Execute(SQL)
    If Not RSTOTAL.EOF Or Not RSTOTAL.BOF Then
        COMPUTE_ONHANDASOFDATE = N2Str2IntZero(RSTOTAL(0).Value)
    End If
End Function


