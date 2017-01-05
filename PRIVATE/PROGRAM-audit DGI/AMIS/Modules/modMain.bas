Attribute VB_Name = "modAMISMain"
Option Explicit
Dim xEntity                                                 As String
Public LOAD_NEWJOURNAL                                      As Boolean
Dim rsCHECKUNPOSTED As ADODB.Recordset

Public Sub Main()
    On Error Resume Next
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
    DMIS_REPORT_Connection = "DSN=" & DATABASE & " ;DSQ=" & SQLSERVERNAME
    DMIS_Audit_Connection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=DMIS_AUDIT ;Data Source=" & SQLSERVERNAME


    frmMain.Show
    frmMain.ZOrder 1
    Load frmSecurity
    frmSecurity.ZOrder 0
    frmSecurity.Show vbModal
    frmMain.ZOrder 1
    frmMainMenu.Show
    ReminderModule ""
End Sub

Public Function OpenSQLDb() As Boolean
    On Error GoTo ConnErr
    Screen.MousePointer = 11
    frmSecurity.Hide
    frmSplash.Show
    frmSplash.ZOrder 0
    frmSplash.labCon.Caption = "Connecting to DMIS Databases... Please wait..."
    DoEvents
    ApplySecurityValidation = True
    BIR_RELIEF_Connection = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=true;Data Source=C:\BIR_RLF\DATA\BIR_RELIEF.MDB"

    CASH_SALES = "'41'"
    CHARGE_SALES = "'42'"
    CASH_DISCOUNT = "'51'"
    CHARGE_DISCOUNT = "'52'"
    CASH_COSTOFSALES = "'61'"
    CHARGE_COSTOFSALES = "'62'"
    OPERATIONAL_EXPENSE = "'71'"
    ADMIN_EXPENSE = "'40'"
    OTHER_INCOME = "'81'"
    OTHER_EXPENSE = "'91'"
    CURRENT_ASSET = "'11'"

    TAX_CREDITS = "1107"
    PROPERTY_EQUIPMENT = "1201"
    ACCUMULATED_DEPRECIATION = "1202"
    OTHER_ASSET = "1204"

    COA_AR_TRADE_UNITS = "'11-02100-00'"
    COA_AR_TRADE_SERVICE = "'11-02200-00'"
    COA_AR_TRADE_PARTS = "'11-02300-00'"

    COA_OUTPUT_TAX = "'21-05100-00'"

    Set gconDMIS = New ADODB.Connection
    gconDMIS.ConnectionString = DMIS_Connection
    gconDMIS.Mode = adModeReadWrite
    gconDMIS.CursorLocation = adUseClient
    frmSplash.labCon.Caption = "Connecting to DMIS Database... Please wait..."
    gconDMIS.Open
    OpenSQLDb = True
    SetCompanyProfile
    Screen.MousePointer = 0
    frmSplash.Command1.Value = True
    Exit Function

ConnErr:
    If Err.Number = 3704 Then
        Resume Next
    Else
        MsgBox Err.DESCRIPTION
        MsgBox "I can't open a connection!!! You may have to " & vbCrLf & _
               "LOG-IN again to connect to the (LOCAL) to run this program. " & vbCrLf & _
               "If you don't have an account contact your friendly " & vbCrLf & _
               "neighborhood SysAdministrator.", _
               vbOKOnly + vbCritical, "ERROR"
        End
    End If
End Function

Public Sub SetUserSettings()
    Call SetUserPathSettings
    With frmMain
        .StatusBar1.Panels(1).Text = "User: " & LOGNAME
        .StatusBar1.Panels(2).Text = "Level: " & LOGLEVEL
        .StatusBar1.Panels(3).Text = "Date: " & Format(LOGDATE, "long date")
        .StatusBar1.Panels(4).Text = "Login Time: " & LOGTIME
        .StatusBar1.Panels(9).Text = "Server Name: " & SQLSERVERNAME & "-" & DATABASE
    End With
End Sub

'Public Sub ShowReport(ReportName As Variant, ReportFolder As Variant, filter As Variant, ReportHeading As String, REPORT_DATE As String, WithDate As Boolean)
'    Screen.MousePointer = 11
'    Dim rsProfile                                                     As ADODB.Recordset
'    Dim CrystalRpt                                                    As Crystal.CrystalReport
'    frmMain.rptMain.Reset
'    Set CrystalRpt = frmMain.rptMain
'    Set rsProfile = New ADODB.Recordset
'    Set rsProfile = gconDMIS.Execute("Select * from ALL_Profile WHERE MODULENAME = '" & MODULENAME & "'")
'    If Not (rsProfile.EOF And rsProfile.BOF) Then
'        CrystalRpt.Reset
'        CrystalRpt.ReportFileName = AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt"
'        CrystalRpt.WindowShowPrintSetupBtn = True
'        CrystalRpt.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
'        CrystalRpt.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
'
'        'UPDATED BY: JUN/ARNOLD------------------------------------------------------------------------
'        'DATE UPDATED: 06-11-2009
'        If ReportName = "CustomersSubsidiaryLedger" Then
'                CrystalRpt.Formulas(55) = "CUST_OPENING = '" & ToDoubleNumber(xBALANCE) & "'"
'                CrystalRpt.Formulas(56) = "COB_DATE ='" & Format(BEG_BALANCE_DATE, "mmm dd yyyy") & "'"
'        End If
'        'UPDATED BY: JUN/ARNOLD------------------------------------------------------------------------
'
'
'        If COMPANY_CODE = "HGC" Then
'            ' Update By BTT : 07282008
'            If ReportName = "AccountsPayable" Or ReportName = "SalesJournal" Or ReportName = "GeneralJournal" Or ReportName = "CashDisbursement" Or ReportName = "CashReceipts" Then
'                If REPRINT_CAPTION = "YES" Then
'                    CrystalRpt.Formulas(3) = "Reprint= '" & "REPRINTED" & "'"
'                Else
'                    CrystalRpt.Formulas(3) = "Reprint= '" & "" & "'"
'                End If
'            End If
'        End If
'        If WithDate = True Then CrystalRpt.Formulas(2) = "ReportDate = '" & REPORT_DATE & "'"
'        CrystalRpt.ReportTitle = ReportHeading: CrystalRpt.WindowTitle = ReportHeading
'        If IsEmpty(filter) = True Then
'            PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", "", DMIS_REPORT_Connection, 1
'        Else
'            PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", filter, DMIS_REPORT_Connection, 1
'        End If
'        CrystalRpt.PageZoom 89
'    End If
'    Screen.MousePointer = 0
'
'End Sub

Public Sub ShowReport(ReportName As Variant, ReportFolder As Variant, filter As Variant, ReportHeading As String, REPORT_DATE As String, WithDate As Boolean)
    Screen.MousePointer = 11
    Dim rsProfile                                           As ADODB.Recordset
    Dim CrystalRpt                                          As Crystal.CrystalReport
    frmMain.rptMain.Reset
    Set CrystalRpt = frmMain.rptMain
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_Profile WHERE MODULENAME = '" & MODULENAME & "'")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        CrystalRpt.Reset
        CrystalRpt.ReportFileName = AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt"
        CrystalRpt.WindowShowPrintSetupBtn = True
        CrystalRpt.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        CrystalRpt.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        'CrystalRpt.Formulas(100) = "TIN = '" & "TIN: " & Null2String(rsProfile!COMPANYTINNO) & "'"
        '        If COMPANY_CODE = "HMH" Then
        '            CrystalRpt.Formulas(81) = "Preparedby = '" & Null2String(rsProfile!PreparedBy) & "'"
        '            CrystalRpt.Formulas(82) = "Checkedby = '" & Null2String(rsProfile!CheckedBy) & "'"
        '            CrystalRpt.Formulas(83) = "Approvedby = '" & Null2String(rsProfile!ApprovedBy) & "'"
        '        End If
        'UPDATED BY: JUN/ARNOLD------------------------------------------------------------------------
        'DATE UPDATED: 06-11-2009
        If ReportName = "CustomersSubsidiaryLedger" Then
            CrystalRpt.Formulas(55) = "CUST_OPENING = '" & ToDoubleNumber(xBALANCE) & "'"
            CrystalRpt.Formulas(56) = "COB_DATE ='" & Format(BEG_BALANCE_DATE, "mmm dd yyyy") & "'"
        End If

        'UPDATED BY: JUN/ARNOLD------------------------------------------------------------------------
        '        If ReportName = "GeneralJournal" Then
        '            If COMPANY_CODE = "DGI" Then
        '                GJ_REMARKS_XXX
        '                'CrystalRpt.Formulas(59) = "REMARK = '" & xEntity & "'"
        '            End If
        '        End If

        If COMPANY_CODE = "HGC" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "HCA" Or COMPANY_CODE = "HCC" Or COMPANY_CODE = "HQA" Or COMPANY_CODE = "HNE" Then
            ' Update By BTT : 07282008
            If ReportName = "AccountsPayable" Or ReportName = "SalesJournal" Or ReportName = "GeneralJournal" Or ReportName = "CashDisbursement" Or ReportName = "CashReceipts" Then
                If REPRINT_CAPTION = "YES" Then
                    CrystalRpt.Formulas(3) = "Reprint= '" & "REPRINTED" & "'"
                Else
                    CrystalRpt.Formulas(3) = "Reprint= '" & "" & "'"
                End If
                CrystalRpt.Formulas(99) = "PRINTEDBY='" & LOGNAME & "'"
            End If
        End If
        If WithDate = True Then CrystalRpt.Formulas(2) = "ReportDate = '" & REPORT_DATE & "'"
        CrystalRpt.ReportTitle = ReportHeading: CrystalRpt.WindowTitle = ReportHeading
        If IsEmpty(filter) = True Then
            PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", "", DMIS_REPORT_Connection, 1
        Else
            'PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", filter, DMIS_REPORT_Connection, 1
            'UPDATED BY: JUN
            'DATE UPDATED: 03-20-2009
            'DESCRIPTION: ASK REQUESTED BY MAM JOECY TO SEPARATE THE PRINTING OF THE CHECK TO CASH DISBURSEMENT VOUCHER
            If COMPANY_CODE = "HPI" Then
                If ReportName = "CashDisbursement" Then
                    If frmAMISJournalEntry_CDJ.optPrintVoucher = True Then
                        PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", filter, DMIS_REPORT_Connection, 1
                        PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\CashDisbursement_TEMPCHECK.rpt", filter, DMIS_REPORT_Connection, 1
                    Else
                        PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", filter, DMIS_REPORT_Connection, 1
                    End If
                Else
                    PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", filter, DMIS_REPORT_Connection, 1
                End If
            Else
                'THIS IS THE ORIGINAL CODE
                If xPAYEE_NAME <> "" Then
                    CrystalRpt.Formulas(80) = "NEW_PAYEE_NAME='" & xPAYEE_NAME & "'"
                End If
                PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", filter, DMIS_REPORT_Connection, 1
            End If
        End If
        CrystalRpt.PageZoom 89
    End If
    Screen.MousePointer = 0
End Sub

Function GJ_REMARKS_XXX() As String
    Dim GJ_REMARKS                                          As ADODB.Recordset
    Dim xENTITY2                                            As String
    xEntity = ""
    Set GJ_REMARKS = New ADODB.Recordset
    GJ_REMARKS.Open "SELECT DISTINCT ADJ_REMARKS,ENTITY FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & frmAMISJournalEntry_GJ.txtVoucherNo.Text & "' AND JTYPE = 'GJ'", gconDMIS, adOpenKeyset
    If Not GJ_REMARKS.EOF And Not GJ_REMARKS.BOF Then
        xEntity = Left(Null2String(GJ_REMARKS!ENTITY), 1)
        '                xENTITY2 = Right(Null2String(GJ_REMARKS!ENTITY), 6)
        Do While Not GJ_REMARKS.EOF
            If IsNull(GJ_REMARKS!ADJ_REMARKS) <> True Then
                GJ_REMARKS_XXX = GJ_REMARKS_XXX + Null2String(GJ_REMARKS!ADJ_REMARKS)
                GJ_REMARKS_XXX = GJ_REMARKS_XXX + Chr(13)
            End If
            GJ_REMARKS.MoveNext
        Loop
    End If
    'gconDMIS.Execute ("UPDATE AMIS_JOURNAL_HD SET REMARKS=NULL WHERE VOUCHERNO = '" & frmAMISJournalEntry_GJ.txtVoucherNo.Text & "' AND JTYPE = 'GJ'")
    gconDMIS.Execute ("UPDATE AMIS_JOURNAL_HD SET REMARKS='" & GJ_REMARKS_XXX & "' WHERE VOUCHERNO = '" & frmAMISJournalEntry_GJ.txtVoucherNo.Text & "' AND JTYPE = 'GJ'")

    '            If xENTITY = "C" Then
    '                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_HD SET CUSTOMERCODE='" & xENTITY2 & "' WHERE VOUCHERNO = '" & frmAMIS_GJ_JOURNAL_ENTRY.txtVoucherNo.Text & "' AND JTYPE = 'GJ'")
    '            Else
    '                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_HD SET VENDORCODE='" & xENTITY2 & "' WHERE VOUCHERNO = '" & frmAMIS_GJ_JOURNAL_ENTRY.txtVoucherNo.Text & "' AND JTYPE = 'GJ'")
    '            End If
    Set GJ_REMARKS = Nothing
End Function


Public Sub ShowRangeReport(V_From As String, V_To As String, ReportName As Variant, ReportFolder As Variant, filter As Variant, ReportHeading As String, WithDate As Boolean)
    On Error Resume Next
    Screen.MousePointer = 11
    Dim rsProfile                                           As ADODB.Recordset
    Dim CrystalRpt                                          As Crystal.CrystalReport
    frmMain.rptMain.Reset
    Set CrystalRpt = frmMain.rptMain
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE WHERE MODULENAME = '" & MODULENAME & "'")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        CrystalRpt.ReportFileName = AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt"
        CrystalRpt.WindowShowPrintSetupBtn = True
        CrystalRpt.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        CrystalRpt.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        'CrystalRpt.Formulas(6) = "TIN = '" & "TIN: " & Null2String(rsProfile!COMPANYTINNO) & "'"

        If WithDate = True Then CrystalRpt.Formulas(2) = "ReportDate = '" & LOGDATE & "'"
        CrystalRpt.Formulas(3) = "FromJDate = '" & V_From & "'"
        CrystalRpt.Formulas(4) = "ToJDate = '" & V_To & "'"
        CrystalRpt.Formulas(5) = "PRINTEDBY='" & LOGNAME & "' "
        CrystalRpt.ReportTitle = ReportHeading: CrystalRpt.WindowTitle = ReportHeading
        PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", filter, DMIS_REPORT_Connection, 1
        CrystalRpt.PageZoom 89
    End If
    Screen.MousePointer = 0
End Sub

Public Function CheckIfBookIsOpen(vJtype As String, vAcctngMonth As Integer, vAcctngYear As Integer) As Boolean
    Dim FieldName                                           As String
    If vJtype = "APJ" Or vJtype = "CDJ" Or vJtype = "SJ" Or vJtype = "CRJ" Or vJtype = "GJ" Then
        FieldName = Trim(vJtype & "Month" & vAcctngMonth)

        Dim rsCheckRecord                                   As ADODB.Recordset
        Set rsCheckRecord = New ADODB.Recordset
        Set rsCheckRecord = gconDMIS.Execute("Select " & FieldName & " AS FieldToCheck from AMIS_AcctngPeriod Where Yeer = " & vAcctngYear)
        If Not rsCheckRecord.EOF And Not rsCheckRecord.BOF Then
            If Null2Bit(rsCheckRecord!FieldToCheck) = 1 Then
                MsgBox "Accounting Book for this Period is already closed.", vbInformation, "Book already Closed"
                CheckIfBookIsOpen = False
            Else
                CheckIfBookIsOpen = True
            End If
        End If
    Else
        CheckIfBookIsOpen = True
    End If
End Function

Function ReturnAccountName(XXX As String) As String
    Dim SQL                                                 As String
    Dim RS                                                  As New ADODB.Recordset
    SQL = "SELECT Description FROM AMIS_ChartAccount where acctcode=" & XXX & ""
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)
    If Not RS.EOF And Not RS.BOF Then
        ReturnAccountName = Null2String(RS!DESCRIPTION)
    End If
    Set RS = Nothing
End Function

Sub FormExistsShow(frmx As Form)
    On Error GoTo ErrorCode
    Dim m_Exists                                            As Boolean
    Dim frm                                                 As Form
    frmx.Show
    For Each frm In Forms
        If (UCase(frm.Name) = UCase(frmx.Name)) Then
            m_Exists = True
            Exit For
        End If
    Next
    Set frm = Nothing

    If m_Exists = True Then
        frmx.WindowState = 0
        frmx.ZOrder 0
    End If

    Exit Sub
ErrorCode:
    Err.Clear
End Sub

Sub ReportControlAddColumnHeader(LST As ReportControl, StringHeaders As String)
    Dim ar()                                                As String
    Dim i                                                   As Integer


    ar = Split(StringHeaders, ",")
    LST.Columns.DeleteAll
    For i = LBound(ar) To UBound(ar)
        LST.Columns.Add i, ar(i), 100, True
    Next
    Erase ar
    StringHeaders = vbNullString

End Sub

Public Sub ResizeColumnHeader(grd As Object, SizeArray As String)
    grd.Visible = False
    Dim ar()                                                As String
    Dim cWidth                                              As Long
    Dim i                                                   As Integer
    Dim scwidth                                             As Long
    ar = Split(SizeArray, ",")
    cWidth = grd.Width

    If TypeOf grd Is ListView Then
        For i = LBound(ar) To UBound(ar)
            If i <= grd.ColumnHeaders.Count Then
                scwidth = cWidth * (CDec(ar(i)) / 100)
                grd.ColumnHeaders(i + 1).Width = scwidth
            End If
        Next
    ElseIf TypeOf grd Is ReportControl Then
        For i = LBound(ar) To UBound(ar)
            If i < grd.Columns.Count Then
                scwidth = cWidth * (CDec(ar(i)) / 100)
                grd.Columns(i).Width = scwidth
            End If
        Next

    End If

    Erase ar
    grd.Visible = True
End Sub

Public Sub AddColumnHeader(StringHeaders As String, lvGrid As ListView)
    Dim ar()                                                As String
    Dim cWidth                                              As Long
    Dim i                                                   As Integer

    ar = Split(StringHeaders, ",")
    cWidth = lvGrid.Width
    lvGrid.ColumnHeaders.Clear
    For i = LBound(ar) To UBound(ar)
        lvGrid.ColumnHeaders.Add , , ar(i)
    Next
    Erase ar
    StringHeaders = vbNullString
End Sub

Sub ReportControlPaintManager(LST As ReportControl)
    With LST
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.HighlightBackColor = RGB(34, 133, 13)
        .PaintManager.ShadeSortColor = RGB(250, 251, 189)
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
        .PaintManager.GroupRowTextBold = True
        .PaintManager.GroupForeColor = vbBlue
        .PaintManager.ColumnStyle = xtpColumnExplorer

    End With

End Sub

Function SelectCombo(C As ComboBox, STR As String, Optional ByVal ByItemData As Boolean = False) As Integer
    If C.ListCount = 0 Then: SelectCombo = -1: Exit Function
    Dim i                                                   As Long
    Dim ItemDataX                                           As Long
    If ByItemData = False Then
        For i = 0 To C.ListCount - 1
            If UCase(C.List(i)) = UCase(Trim(STR)) Then
                SelectCombo = i
                Exit Function
            End If
        Next
    Else
        If STR = vbNullString Then
            SelectCombo = -1
            Exit Function
        End If

        ItemDataX = CLng(STR)

        For i = 0 To C.ListCount - 1
            If C.ItemData(i) = STR Then
                SelectCombo = i
                Exit Function
            End If
        Next
    End If
    SelectCombo = -1
End Function

Sub FillcboNewYear(ByRef XXX As Object)
    Dim i                                                   As Integer
    XXX.Clear
    For i = 1990 To 2020
        XXX.AddItem i
    Next i
    XXX.Text = Year(LOGDATE)
End Sub

Function GetAcctcode(XXX As String) As String
    Dim rsAcctCode                                          As ADODB.Recordset
    Set rsAcctCode = New ADODB.Recordset
    rsAcctCode.Open "SELECT ACCT_CODE FROM AMIS_JOURNAL_DET WHERE ID = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsAcctCode.EOF And Not rsAcctCode.BOF Then
        GetAcctcode = Null2String(rsAcctCode!ACCT_CODE)
    End If
    Set rsAcctCode = Nothing
End Function

Function GetCUTOFF_DATE() As String
    Dim rsGET_CUT_OFF_DATE                                  As ADODB.Recordset
    Set rsGET_CUT_OFF_DATE = New ADODB.Recordset
    rsGET_CUT_OFF_DATE.Open "SELECT CUT_OFF_DATE FROM ALL_PROFILE WHERE MODULENAME = 'AMIS'", gconDMIS, adOpenKeyset
    If Not rsGET_CUT_OFF_DATE.EOF And Not rsGET_CUT_OFF_DATE.BOF Then
        GetCUTOFF_DATE = Null2String(rsGET_CUT_OFF_DATE!Cut_Off_Date)
    Else
        MessagePop InfoFriend, "SYSTEM MESSAGE", "Please Update the Cut-Off date in Amis Profile Module"
        Screen.MousePointer = 0
        Exit Function
    End If
    Set rsGET_CUT_OFF_DATE = Nothing
End Function

Function UNPOSTEDAPJ() As Boolean
    Dim rsUNPOSTEDAPJ                                       As ADODB.Recordset
    Set rsUNPOSTEDAPJ = New ADODB.Recordset
    If BATCHIMPORT = True Then
        rsUNPOSTEDAPJ.Open "SELECT * FROM (" & _
                           "SELECT CASE WHEN TYPE='P' THEN 'PARTS' WHEN TYPE='M' THEN 'MATERIALS' WHEN TYPE='A' THEN 'ACCESSORIES' END AS TYPE,RRNO,CAST(TTLRRAMT AS DECIMAL(18,2)) AS AMOUNT,RRDATE FROM PMIS_VW_RR_TRANS WHERE (CLASSCODE = 'PCG' OR CLASSCODE = 'PCS' OR CLASSCODE = 'IBT' ) AND STATUS='N' " & _
                           "UNION SELECT 'VEHICLES' AS TYPE,CODE AS RRNO,CAST(PURCHPRICE AS DECIMAL(18,2)) AS AMOUNT,DATERECEIVED AS RRDATE FROM SMIS_MRRINV LEFT OUTER JOIN CSMS_SELLINGDEALER ON SMIS_MRRINV.SOURCE = CSMS_SELLINGDEALER.DEALERCODE WHERE STATUS='N' " & _
                           "UNION SELECT 'SUBLET' AS TYPE,RC_NO AS RRNO,CAST(SUBLET_TOTAL_NET_AMT AS DECIMAL(18,2)) AS AMOUNT,RC_DATE AS RRDATE FROM CSMS_PO_RC_HD WHERE STATUS='N' " & _
                           ")T WHERE RRDATE BETWEEN '" & frmAPJImport.dtFrom.Value & "' AND '" & frmAPJImport.dtTo.Value & "'", gconDMIS, adOpenForwardOnly
    Else
        rsUNPOSTEDAPJ.Open "SELECT * FROM (" & _
                           "SELECT CASE WHEN TYPE='P' THEN 'PARTS' WHEN TYPE='M' THEN 'MATERIALS' WHEN TYPE='A' THEN 'ACCESSORIES' END AS TYPE,RRNO,CAST(TTLRRAMT AS DECIMAL(18,2)) AS AMOUNT,RRDATE FROM PMIS_VW_RR_TRANS WHERE (CLASSCODE = 'PCG' OR CLASSCODE = 'PCS' OR CLASSCODE = 'IBT' ) AND STATUS='N' " & _
                           "UNION SELECT 'VEHICLES' AS TYPE,CODE AS RRNO,CAST(PURCHPRICE AS DECIMAL(18,2)) AS AMOUNT,DATERECEIVED AS RRDATE FROM SMIS_MRRINV LEFT OUTER JOIN CSMS_SELLINGDEALER ON SMIS_MRRINV.SOURCE = CSMS_SELLINGDEALER.DEALERCODE WHERE STATUS='N' " & _
                           "UNION SELECT 'SUBLET' AS TYPE,RC_NO AS RRNO,CAST(SUBLET_TOTAL_NET_AMT AS DECIMAL(18,2)) AS AMOUNT,RC_DATE AS RRDATE FROM CSMS_PO_RC_HD WHERE STATUS='N' " & _
                           ")T WHERE RRDATE='" & frmAPJImport.dtpTranDate.Value & "'", gconDMIS, adOpenForwardOnly
    End If
    If Not rsUNPOSTEDAPJ.EOF And Not rsUNPOSTEDAPJ.BOF Then
        Do While Not rsUNPOSTEDAPJ.EOF
            frmListofUnposted.Grid1.AddItem 0 & Chr(9) & Null2String(rsUNPOSTEDAPJ!Type) & Chr(9) & Null2String(rsUNPOSTEDAPJ!RRNO) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsUNPOSTEDAPJ!amount))
            rsUNPOSTEDAPJ.MoveNext
        Loop
        UNPOSTEDAPJ = True
    Else
        UNPOSTEDAPJ = False
    End If
    Set rsUNPOSTEDAPJ = Nothing
End Function

Function UNPOSTEDSJ() As Boolean
    Dim rsUNPOSTEDSJ                                        As ADODB.Recordset
    Set rsUNPOSTEDSJ = New ADODB.Recordset
    If BATCHIMPORT = True Then
        rsUNPOSTEDSJ.Open "SELECT * FROM (" & _
                          "SELECT CASE WHEN TYPE='P' THEN 'PARTS' WHEN TYPE='M' THEN 'MATERIALS' WHEN TYPE='A' THEN 'ACCESSORIES' END AS REFERENCE,TRANTYPE+'-'+TRANNO AS INVNO,NETINVAMT AS AMOUNT,TRANDATE FROM PMIS_VW_ISS_HISTORY WHERE (TRANTYPE = 'CSH' OR  TRANTYPE = 'CHG' OR  TRANTYPE = 'DR') AND STATUS = 'N' " & _
                          "UNION SELECT CSMS_REPOR.REP_OR AS REFERENCE,CSMS_REPOR.INVOICE AS INVNO,CSMS_REPOR.RO_AMOUNT AS AMOUNT,DTE_COMP AS TRANDATE FROM CSMS_REPOR  WHERE INVOICE = 'INT RO' AND DTE_REL IS NULL " & _
                          "UNION SELECT CSMS_REPOR.REP_OR AS REFENCE,CSMS_REPOR.INVOICE AS INVNO,SUM(CSMS_RO_DET.DETPRC) AS AMOUNT,DTE_COMP AS TRANDATE FROM CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE INVOICE <> 'INT RO' AND RO_AMOUNT = 0 AND DETAMT > 0 AND WCODE IN ('S','C') AND INVOICE <> 'NO CHG' AND INVOICE <> 'PDI RO' AND DTE_REL IS NULL GROUP BY CSMS_REPOR.REP_OR,INVOICE,DTE_COMP " & _
                          "UNION SELECT REP_OR AS REFERENCE,INVOICE AS INVNO,AMOUNT,DTE_COMP FROM CSMS_REPOR WHERE RO_AMOUNT > 0 AND INVOICE <> 'INT RO' AND INVOICE <> 'NO CHG' AND INVOICE <> 'PDI RO' AND DTE_REL IS NULL " & _
                          "UNION SELECT CSMS_REPOR.REP_OR AS REFERENCE,CSMS_REPOR.INVOICE AS INVNO,SUM(CSMS_RO_DET.DETAMT) AS AMOUNT,DTE_COMP AS TRANDATE FROM CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE RO_AMOUNT = 0 AND DETAMT > 0 AND WCODE = 'W' AND INVOICE <> 'NO CHG' AND INVOICE <> 'PDI RO' AND INVOICE <> 'INT RO' AND DTE_REL IS NULL GROUP BY CSMS_REPOR.REP_OR,INVOICE,DTE_COMP " & _
                          "UNION SELECT IGNKEY_NO AS REFERENCE,VI_NO AS INVNO,NETSALESPRICE AS AMOUNT,CAST(CONVERT(VARCHAR, DATERELEASED, 101) AS SMALLDATETIME) AS TRANDATE FROM SMIS_PURCHAGREE WHERE STATUS = 'N' " & _
                          ")T WHERE TRANDATE BETWEEN '" & frmSALESImport.dtFrom.Value & "' AND '" & frmSALESImport.dtTo.Value & "'", gconDMIS, adOpenForwardOnly
    Else
        rsUNPOSTEDSJ.Open "SELECT * FROM (" & _
                          "SELECT CASE WHEN TYPE='P' THEN 'PARTS' WHEN TYPE='M' THEN 'MATERIALS' WHEN TYPE='A' THEN 'ACCESSORIES' END AS REFERENCE,TRANTYPE+'-'+TRANNO AS INVNO,NETINVAMT AS AMOUNT,TRANDATE FROM PMIS_VW_ISS_HISTORY WHERE (TRANTYPE = 'CSH' OR  TRANTYPE = 'CHG' OR  TRANTYPE = 'DR') AND STATUS = 'N' " & _
                          "UNION SELECT CSMS_REPOR.REP_OR AS REFERENCE,CSMS_REPOR.INVOICE AS INVNO,CSMS_REPOR.RO_AMOUNT AS AMOUNT,DTE_COMP AS TRANDATE FROM CSMS_REPOR  WHERE INVOICE = 'INT RO' AND DTE_REL IS NULL " & _
                          "UNION SELECT CSMS_REPOR.REP_OR AS REFENCE,CSMS_REPOR.INVOICE AS INVNO,SUM(CSMS_RO_DET.DETPRC) AS AMOUNT,DTE_COMP AS TRANDATE FROM CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE INVOICE <> 'INT RO' AND RO_AMOUNT = 0 AND DETAMT > 0 AND WCODE IN ('S','C') AND INVOICE <> 'NO CHG' AND INVOICE <> 'PDI RO' AND DTE_REL IS NULL GROUP BY CSMS_REPOR.REP_OR,INVOICE,DTE_COMP " & _
                          "UNION SELECT REP_OR AS REFERENCE,INVOICE AS INVNO,AMOUNT,DTE_COMP FROM CSMS_REPOR WHERE RO_AMOUNT > 0 AND INVOICE <> 'INT RO' AND INVOICE <> 'NO CHG' AND INVOICE <> 'PDI RO' AND DTE_REL IS NULL " & _
                          "UNION SELECT CSMS_REPOR.REP_OR AS REFERENCE,CSMS_REPOR.INVOICE AS INVNO,SUM(CSMS_RO_DET.DETAMT) AS AMOUNT,DTE_COMP AS TRANDATE FROM CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE RO_AMOUNT = 0 AND DETAMT > 0 AND WCODE = 'W' AND INVOICE <> 'NO CHG' AND INVOICE <> 'PDI RO' AND INVOICE <> 'INT RO' AND DTE_REL IS NULL GROUP BY CSMS_REPOR.REP_OR,INVOICE,DTE_COMP " & _
                          "UNION SELECT IGNKEY_NO AS REFERENCE,VI_NO AS INVNO,NETSALESPRICE AS AMOUNT,CAST(CONVERT(VARCHAR, DATERELEASED, 101) AS SMALLDATETIME) AS TRANDATE FROM SMIS_PURCHAGREE WHERE STATUS = 'N' " & _
                          ")T WHERE TRANDATE = '" & frmSALESImport.dtpTranDate.Value & "'", gconDMIS, adOpenForwardOnly
    End If
    If Not rsUNPOSTEDSJ.EOF And Not rsUNPOSTEDSJ.BOF Then
        Do While Not rsUNPOSTEDSJ.EOF
            frmListofUnposted.Grid1.AddItem 0 & Chr(9) & Null2String(rsUNPOSTEDSJ!Reference) & Chr(9) & Null2String(rsUNPOSTEDSJ!INVNO) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsUNPOSTEDSJ!amount))
            rsUNPOSTEDSJ.MoveNext
        Loop
        UNPOSTEDSJ = True
    Else
        UNPOSTEDSJ = False
    End If
    Set rsUNPOSTEDSJ = Nothing
End Function

Function CheckIfARAccount(xACCT_CODE As String) As Boolean
    Dim rsCheck                                             As ADODB.Recordset
    Set rsCheck = New ADODB.Recordset
    rsCheck.Open "SELECT * FROM AMIS_CHARTACCOUNT WHERE LEFT(ACCTCODE,5) IN ('11-02','11-03','11-04') AND IS_SCHEDULE_ACCNT=1 AND ACCTCODE = " & xACCT_CODE & "", gconDMIS, adOpenForwardOnly
    If Not rsCheck.EOF And Not rsCheck.BOF Then
        CheckIfARAccount = True
    Else
        CheckIfARAccount = False
    End If
    Set rsCheck = Nothing
End Function

Function CheckIfBalanceAR(xACCT_CODE As String, xAMOUNT As Double, xSJVOUCHERNO As String) As Boolean
    Dim rsAR                                                As ADODB.Recordset
    Set rsAR = New ADODB.Recordset
    rsAR.Open "SELECT * FROM (SELECT SUM(ISNULL(AMOUNT_TOPAY,0)) AS AR FROM AMIS_AR WHERE ACCOUNT_CODE='" & xACCT_CODE & "' AND SJVOUCHERNO='" & xSJVOUCHERNO & "')T WHERE AR='" & NumericVal(xAMOUNT) & "'", gconDMIS, adOpenForwardOnly
    If Not rsAR.EOF And Not rsAR.BOF Then
        CheckIfBalanceAR = True
    Else
        CheckIfBalanceAR = False
    End If
    Set rsAR = Nothing
End Function

Function CheckIfBalanceAP(xACCT_CODE As String, xAMOUNT As Double, xVOUCHERNO As String) As Boolean
    Dim rsAR                                                As ADODB.Recordset
    Set rsAR = New ADODB.Recordset
    rsAR.Open "SELECT * FROM (SELECT SUM(ISNULL(AMOUNT2PAY,0)) AS AP FROM AMIS_AP WHERE ACCT_CODE='" & xACCT_CODE & "' AND VOUCHERNO='" & xVOUCHERNO & "')T WHERE AP='" & NumericVal(xAMOUNT) & "'", gconDMIS, adOpenForwardOnly
    If Not rsAR.EOF And Not rsAR.BOF Then
        CheckIfBalanceAP = True
    Else
        CheckIfBalanceAP = False
    End If
    Set rsAR = Nothing
End Function

Function CheckIfBalanceARDetails(xACCT_CODE As String, xAMOUNT As Double, xVOUCHERNO As String) As Boolean
    Dim rsAR                                                As ADODB.Recordset
    Set rsAR = New ADODB.Recordset
    rsAR.Open "SELECT * FROM (SELECT SUM(ISNULL(INVOICEAMOUNT,0)) AS AR FROM AMIS_DETAIL WHERE ACCT_CODE='" & xACCT_CODE & "' AND JTYPE+'-'+VOUCHERNO='" & xVOUCHERNO & "')T WHERE AR='" & NumericVal(xAMOUNT) & "'", gconDMIS, adOpenForwardOnly
    If Not rsAR.EOF And Not rsAR.BOF Then
        CheckIfBalanceARDetails = True
    Else
        CheckIfBalanceARDetails = False
    End If
    Set rsAR = Nothing
End Function

Function CheckIfBalanceAPDetails(xACCT_CODE As String, xAMOUNT As Double, xVOUCHERNO) As Boolean
    Dim rsAR                                                As ADODB.Recordset
    Set rsAR = New ADODB.Recordset
    rsAR.Open "SELECT * FROM (SELECT SUM(ISNULL(AMOUNTPAID,0)) AS AP FROM AMIS_DETAILS WHERE ACCT_CODE='" & xACCT_CODE & "' AND JTYPE+'-'+VOUCHERNO='" & xVOUCHERNO & "')T WHERE AP='" & NumericVal(xAMOUNT) & "'", gconDMIS, adOpenForwardOnly
    If Not rsAR.EOF And Not rsAR.BOF Then
        CheckIfBalanceAPDetails = True
    Else
        CheckIfBalanceAPDetails = False
    End If
    Set rsAR = Nothing
End Function

Function CheckIfARDebitNotZero(ACCT_CODE As String, ar As Boolean, Debit As Double) As Boolean
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    rsChartAccount.Open "SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE IS_SCHEDULE_ACCNT=1 AND LEFT(ACCTCODE,5) IN ('11-02','11-03','11-04') AND ACCTCODE = '" & ACCT_CODE & "'", gconDMIS, adOpenForwardOnly
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        If ar = True And Debit > 0 Then
            CheckIfARDebitNotZero = True
        Else
            CheckIfARDebitNotZero = False
        End If
    End If
    Set rsChartAccount = Nothing
End Function

Function CheckIfAPDebitNotZero(ACCT_CODE As String, ar As Boolean, Debit As Double) As Boolean
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    rsChartAccount.Open "SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE IS_SCHEDULE_ACCNT=1 AND LEFT(ACCTCODE,5) IN ('21-01','21-02','21-06','21-07') AND ACCTCODE = '" & ACCT_CODE & "'", gconDMIS, adOpenForwardOnly
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        If ar = False And Debit > 0 Then
            CheckIfAPDebitNotZero = True
        Else
            CheckIfAPDebitNotZero = False
        End If
    End If
    Set rsChartAccount = Nothing
End Function

Function CheckIfARCreditNotZero(ACCT_CODE As String, ar As Boolean, Credit As Double) As Boolean
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    rsChartAccount.Open "SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE IS_SCHEDULE_ACCNT=1 AND LEFT(ACCTCODE,5) IN ('11-02','11-03','11-04') AND ACCTCODE = '" & ACCT_CODE & "'", gconDMIS, adOpenForwardOnly
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        If ar = True And Credit > 0 Then
            CheckIfARCreditNotZero = True
        Else
            CheckIfARCreditNotZero = False
        End If
    End If
    Set rsChartAccount = Nothing
End Function

Function CheckIfAPCreditNotZero(ACCT_CODE As String, ar As Boolean, Credit As Double) As Boolean
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    rsChartAccount.Open "SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE IS_SCHEDULE_ACCNT=1 AND LEFT(ACCTCODE,5) IN ('21-01','21-02','21-06','21-07') AND ACCTCODE = '" & ACCT_CODE & "'", gconDMIS, adOpenForwardOnly
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        If ar = False And Credit > 0 Then
            CheckIfAPCreditNotZero = True
        Else
            CheckIfAPCreditNotZero = False
        End If
    End If
    Set rsChartAccount = Nothing
End Function

Function CheckGLSLARDebit(xJType As String, xVOUCHERNO) As Boolean
    Dim rsJournalDT                                         As ADODB.Recordset
    Dim rsAR                                                As ADODB.Recordset
    Set rsJournalDT = New ADODB.Recordset
    rsJournalDT.Open "SELECT * FROM (SELECT ACCT_CODE,SUM(ISNULL(DEBIT,0)) AS DEBIT FROM AMIS_JOURNAL_DET WHERE LEFT(ACCT_CODE,5) IN ('11-02','11-03','11-04') AND JTYPE='" & xJType & "' AND VOUCHERNO = '" & xVOUCHERNO & "' AND DEBIT > 0 GROUP BY ACCT_CODE) DT INNER JOIN AMIS_CHARTACCOUNT AC ON DT.ACCT_CODE=AC.ACCTCODE WHERE IS_SCHEDULE_ACCNT=1", gconDMIS, adOpenForwardOnly
    If Not rsJournalDT.EOF And Not rsJournalDT.BOF Then
        Do While Not rsJournalDT.EOF
            Set rsAR = New ADODB.Recordset
            rsAR.Open "SELECT SUM(ISNULL(AMOUNT_TOPAY,0)) AS AR FROM AMIS_AR WHERE ACCOUNT_CODE = '" & rsJournalDT!ACCT_CODE & "' AND SJVOUCHERNO='" & xJType + "-" + xVOUCHERNO & "'", gconDMIS, adOpenForwardOnly
            If Not rsAR.EOF And Not rsAR.BOF Then
                If NumericVal(rsJournalDT!Debit) = NumericVal(rsAR!ar) Then
                    CheckGLSLARDebit = True
                Else
                    MessagePop InfoWarning, "System Message", "Please check. GL Amount not equal to SL" & Chr(13) & "GL " & rsJournalDT!ACCT_CODE & " => " & rsJournalDT!Debit & Chr(13) & "SL " & rsJournalDT!ACCT_CODE & " => " & rsAR!ar
                    CheckGLSLARDebit = False
                    Exit Function
                End If
            End If
            rsJournalDT.MoveNext
        Loop
    Else
        CheckGLSLARDebit = True
    End If
    Set rsJournalDT = Nothing
    Set rsAR = Nothing
End Function

Function CheckGLSLARCredit(xJType As String, xVOUCHERNO) As Boolean
    Dim rsJournalDT                                         As ADODB.Recordset
    Dim rsAR                                                As ADODB.Recordset
    Set rsJournalDT = New ADODB.Recordset
    rsJournalDT.Open "SELECT * FROM (SELECT ACCT_CODE,SUM(ISNULL(CREDIT,0)) AS CREDIT FROM AMIS_JOURNAL_DET WHERE LEFT(ACCT_CODE,5) IN ('11-02','11-03','11-04') AND JTYPE='" & xJType & "' AND VOUCHERNO = '" & xVOUCHERNO & "' AND CREDIT > 0 GROUP BY ACCT_CODE) DT INNER JOIN AMIS_CHARTACCOUNT AC ON DT.ACCT_CODE=AC.ACCTCODE WHERE IS_SCHEDULE_ACCNT=1", gconDMIS, adOpenForwardOnly
    If Not rsJournalDT.EOF And Not rsJournalDT.BOF Then
        Do While Not rsJournalDT.EOF
            Set rsAR = New ADODB.Recordset
            rsAR.Open "SELECT SUM(ISNULL(INVOICEAMOUNT,0)) AS AMOUNTPAID FROM AMIS_DETAIL WHERE ACCT_CODE = '" & rsJournalDT!ACCT_CODE & "' AND JTYPE='" & xJType & "' AND VOUCHERNO ='" & xVOUCHERNO & "'", gconDMIS, adOpenForwardOnly
            If Not rsAR.EOF And Not rsAR.BOF Then
                If NumericVal(rsJournalDT!Credit) = NumericVal(rsAR!AMOUNTPAID) Then
                    CheckGLSLARCredit = True
                Else
                    MessagePop InfoWarning, "System Message", "Please check. GL Amount not equal to SL" & Chr(13) & "GL " & rsJournalDT!ACCT_CODE & " => " & rsJournalDT!Credit & Chr(13) & "SL " & rsJournalDT!ACCT_CODE & " => " & rsAR!AMOUNTPAID
                    CheckGLSLARCredit = False
                    Exit Function
                End If
            End If
            rsJournalDT.MoveNext
        Loop
    Else
        CheckGLSLARCredit = True
    End If
    Set rsJournalDT = Nothing
    Set rsAR = Nothing
End Function

Function CheckGLSLAPDebit(xJType As String, xVOUCHERNO) As Boolean
    Dim rsJournalDT                                         As ADODB.Recordset
    Dim rsAP                                                As ADODB.Recordset
    Set rsJournalDT = New ADODB.Recordset
    rsJournalDT.Open "SELECT * FROM (SELECT ACCT_CODE,SUM(ISNULL(DEBIT,0)) AS DEBIT FROM AMIS_JOURNAL_DET WHERE LEFT(ACCT_CODE,5) IN ('21-01','21-02','21-06','21-07') AND JTYPE='" & xJType & "' AND VOUCHERNO = '" & xVOUCHERNO & "' AND DEBIT > 0 GROUP BY ACCT_CODE) DT INNER JOIN AMIS_CHARTACCOUNT AC ON DT.ACCT_CODE=AC.ACCTCODE WHERE IS_SCHEDULE_ACCNT=1", gconDMIS, adOpenForwardOnly
    If Not rsJournalDT.EOF And Not rsJournalDT.BOF Then
        Do While Not rsJournalDT.EOF
            Set rsAP = New ADODB.Recordset
            rsAP.Open "SELECT SUM(ISNULL(AMOUNTPAID,0)) AS PAYMENT FROM AMIS_DETAILS WHERE ACCT_CODE = '" & rsJournalDT!ACCT_CODE & "' AND JTYPE='" & xJType & "' AND VOUCHERNO='" & xVOUCHERNO & "'", gconDMIS, adOpenForwardOnly
            If Not rsAP.EOF And Not rsAP.BOF Then
                If NumericVal(rsJournalDT!Debit) = NumericVal(rsAP!PAYMENT) Then
                    CheckGLSLAPDebit = True
                Else
                    MessagePop InfoWarning, "System Message", "Please check. GL Amount not equal to SL" & Chr(13) & "GL " & rsJournalDT!ACCT_CODE & " => " & rsJournalDT!Debit & Chr(13) & "SL " & rsJournalDT!ACCT_CODE & " => " & rsAP!PAYMENT
                    CheckGLSLAPDebit = False
                    Exit Function
                End If
            End If
            rsJournalDT.MoveNext
        Loop
    Else
        CheckGLSLAPDebit = True
    End If
    Set rsJournalDT = Nothing
    Set rsAP = Nothing
End Function

Function CheckGLSLAPCredit(xJType As String, xVOUCHERNO) As Boolean
    Dim rsJournalDT                                         As ADODB.Recordset
    Dim rsAP                                                As ADODB.Recordset
    Set rsJournalDT = New ADODB.Recordset
    rsJournalDT.Open "SELECT * FROM (SELECT ACCT_CODE,SUM(ISNULL(CREDIT,0)) AS CREDIT FROM AMIS_JOURNAL_DET WHERE LEFT(ACCT_CODE,5) IN ('21-01','21-02','21-06','21-07') AND JTYPE='" & xJType & "' AND VOUCHERNO = '" & xVOUCHERNO & "' AND CREDIT > 0 GROUP BY ACCT_CODE) DT INNER JOIN AMIS_CHARTACCOUNT AC ON DT.ACCT_CODE=AC.ACCTCODE WHERE IS_SCHEDULE_ACCNT=1", gconDMIS, adOpenForwardOnly
    If Not rsJournalDT.EOF And Not rsJournalDT.BOF Then
        Do While Not rsJournalDT.EOF
            Set rsAP = New ADODB.Recordset
            rsAP.Open "SELECT SUM(ISNULL(AMOUNT2PAY,0)) AS AP FROM AMIS_AP WHERE ACCT_CODE = '" & rsJournalDT!ACCT_CODE & "' AND VOUCHERNO='" & xJType + "-" + xVOUCHERNO & "'", gconDMIS, adOpenForwardOnly
            If Not rsAP.EOF And Not rsAP.BOF Then
                If NumericVal(rsJournalDT!Credit) = NumericVal(rsAP!AP) Then
                    CheckGLSLAPCredit = True
                Else
                    MessagePop InfoWarning, "System Message", "Please check. GL Amount not equal to SL" & Chr(13) & "GL " & rsJournalDT!ACCT_CODE & " => " & rsJournalDT!Credit & Chr(13) & "SL " & rsJournalDT!ACCT_CODE & " => " & rsAP!AP
                    CheckGLSLAPCredit = False
                    Exit Function
                End If
            End If
            rsJournalDT.MoveNext
        Loop
    Else
        CheckGLSLAPCredit = True
    End If
    Set rsJournalDT = Nothing
    Set rsAP = Nothing
End Function

Function CheckIfSameAccount(xACCT_CODEDT As String, xACCT_CODECBO As String, xVOUCHERNO As String, xdebit As Double, xcredit As Double) As Boolean
    Dim rsAR                                                As ADODB.Recordset
    Dim rsAP                                                As ADODB.Recordset
    Set rsAR = New ADODB.Recordset
    If CheckIfARAccount(N2Str2Null(xACCT_CODECBO)) = True And NumericVal(xdebit) > 0 Then
        rsAR.Open "SELECT ACCOUNT_CODE AS ACCT_CODE FROM AMIS_AR WHERE ACCOUNT_CODE = '" & xACCT_CODEDT & "' AND ISNULL(AMOUNT_TOPAY,0) > 0 AND SJVOUCHERNO = '" & xVOUCHERNO & "' ", gconDMIS, adOpenForwardOnly
    ElseIf CheckIfARAccount(N2Str2Null(xACCT_CODECBO)) = True And NumericVal(xcredit) > 0 Then
        rsAR.Open "SELECT * FROM AMIS_DETAIL WHERE ACCT_CODE = '" & xACCT_CODEDT & "' AND ISNULL(INVOICEAMOUNT,0) > 0 AND JTYPE+'-'+VOUCHERNO = '" & xVOUCHERNO & "' ", gconDMIS, adOpenForwardOnly
    ElseIf CheckIfARAccount(N2Str2Null(xACCT_CODECBO)) = False And NumericVal(xdebit) > 0 Then
        rsAR.Open "SELECT * FROM AMIS_DETAILS WHERE ACCT_CODE = '" & xACCT_CODEDT & "' AND ISNULL(AMOUNTPAID,0) > 0 AND JTYPE+'-'+VOUCHERNO = '" & xVOUCHERNO & "' ", gconDMIS, adOpenForwardOnly
    ElseIf CheckIfARAccount(N2Str2Null(xACCT_CODECBO)) = False And NumericVal(xcredit) > 0 Then
        rsAR.Open "SELECT * FROM AMIS_AP WHERE ACCT_CODE = '" & xACCT_CODEDT & "' AND ISNULL(AMOUNT2PAY,0) > 0 AND VOUCHERNO = '" & xVOUCHERNO & "' ", gconDMIS, adOpenForwardOnly
    End If
    If Not rsAR.EOF And Not rsAR.BOF Then
        If rsAR!ACCT_CODE = xACCT_CODECBO Then
            CheckIfSameAccount = True
        Else
            CheckIfSameAccount = False
        End If
    Else
        CheckIfSameAccount = True
    End If
    Set rsAR = Nothing
End Function

Function CheckARDetails(xSJVOUCHERNO As String, xACCT_CODE As String, xJOURNAL_DET_ID As Long) As Boolean
    Dim rsAR                                                As ADODB.Recordset
    Set rsAR = New ADODB.Recordset
    rsAR.Open "SELECT SJVOUCHERNO FROM AMIS_AR WHERE SJVOUCHERNO='" & xSJVOUCHERNO & "' AND ACCOUNT_CODE='" & xACCT_CODE & "' AND JOURNAL_DET_ID = " & xJOURNAL_DET_ID & "", gconDMIS, adOpenForwardOnly
    If Not rsAR.EOF And Not rsAR.BOF Then
        CheckARDetails = True
    Else
        CheckARDetails = False
    End If
    Set rsAR = Nothing
End Function

Function CheckAPDetails(xSJVOUCHERNO As String, xACCT_CODE As String, xJOURNAL_DET_ID As Long) As Boolean
    Dim rsAR                                                As ADODB.Recordset
    Set rsAR = New ADODB.Recordset
    rsAR.Open "SELECT VOUCHERNO FROM AMIS_AP WHERE VOUCHERNO='" & xSJVOUCHERNO & "' AND ACCT_CODE='" & xACCT_CODE & "' AND JOURNAL_DET_ID = " & xJOURNAL_DET_ID & "", gconDMIS, adOpenForwardOnly
    If Not rsAR.EOF And Not rsAR.BOF Then
        CheckAPDetails = True
    Else
        CheckAPDetails = False
    End If
    Set rsAR = Nothing
End Function

Function CheckARPaymentDetails(xJType As String, xVOUCHERNO As String, xACCT_CODE As String, xJOURNAL_DET_ID As Long) As Boolean
    Dim rsAR                                                As ADODB.Recordset
    Set rsAR = New ADODB.Recordset
    rsAR.Open "SELECT JTYPE,VOUCHERNO FROM AMIS_DETAIL WHERE JTYPE='" & xJType & "' AND VOUCHERNO = '" & xVOUCHERNO & "' AND ACCT_CODE='" & xACCT_CODE & "' AND JOURNAL_DET_ID = " & xJOURNAL_DET_ID & "", gconDMIS, adOpenForwardOnly
    If Not rsAR.EOF And Not rsAR.BOF Then
        CheckARPaymentDetails = True
    Else
        CheckARPaymentDetails = False
    End If
    Set rsAR = Nothing
End Function

Function CheckAPPaymentDetails(xJType As String, xVOUCHERNO As String, xACCT_CODE As String, xJOURNAL_DET_ID As Long) As Boolean
    Dim rsAR                                                As ADODB.Recordset
    Set rsAR = New ADODB.Recordset
    rsAR.Open "SELECT JTYPE,VOUCHERNO FROM AMIS_DETAILS WHERE JTYPE='" & xJType & "' AND VOUCHERNO = '" & xVOUCHERNO & "' AND ACCT_CODE='" & xACCT_CODE & "' AND JOURNAL_DET_ID = " & xJOURNAL_DET_ID & "", gconDMIS, adOpenForwardOnly
    If Not rsAR.EOF And Not rsAR.BOF Then
        CheckAPPaymentDetails = True
    Else
        CheckAPPaymentDetails = False
    End If
    Set rsAR = Nothing
End Function

Function JOURNALLASTTRANS(XXX As String) As String
    Dim rsJOURNAL                                           As ADODB.Recordset
    Set rsJOURNAL = New ADODB.Recordset
    rsJOURNAL.Open "SELECT CASE WHEN MAX(JDATE) IS NULL THEN CAST(CONVERT(VARCHAR(10),GETDATE(),101) AS SMALLDATETIME) ELSE MAX(JDATE) END AS JDATE FROM AMIS_JOURNAL_HD WHERE JTYPE='" & XXX & "' AND STATUS='N'", gconDMIS, adOpenForwardOnly
    If Not rsJOURNAL.EOF And Not rsJOURNAL.BOF Then
        JOURNALLASTTRANS = Null2String(rsJOURNAL!JDATE)
    Else
        JOURNALLASTTRANS = ""
    End If
    Set rsJOURNAL = Nothing
End Function

Function JOURNALFIRSTTRANS(XXX As String) As String
    Dim rsJOURNAL                                           As ADODB.Recordset
    Set rsJOURNAL = New ADODB.Recordset
    rsJOURNAL.Open "SELECT CASE WHEN MIN(JDATE) IS NULL THEN CAST(CONVERT(VARCHAR(10),GETDATE(),101) AS SMALLDATETIME) ELSE MIN(JDATE) END AS JDATE FROM AMIS_JOURNAL_HD WHERE JTYPE='" & XXX & "' AND STATUS='N'", gconDMIS, adOpenForwardOnly
    If Not rsJOURNAL.EOF And Not rsJOURNAL.BOF Then
        JOURNALFIRSTTRANS = Null2String(rsJOURNAL!JDATE)
    Else
        JOURNALFIRSTTRANS = ""
    End If
    Set rsJOURNAL = Nothing
End Function

Function SetEntityName(CODE As String) As String
    Dim rsEntity                                            As ADODB.Recordset
    Set rsEntity = New ADODB.Recordset
    rsEntity.Open "Select CODE,ACCOUNTNAME from ALL_ENTITY where COMPLET_CODE = " & N2Str2Null(CODE), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsEntity.EOF And Not rsEntity.BOF Then
        SetEntityName = Null2String(rsEntity!ACCOUNTNAME)
    Else
        SetEntityName = ""
    End If
    Set rsEntity = Nothing
End Function

Sub DetailsTrueFalse(XXX As Boolean)
    frmAMISJournalEntry_Details.picDetails.Enabled = XXX
    frmAMISJournalEntry_Details.Picture1.Visible = XXX
    frmAMISJournalEntry_Details.Picture2.Visible = XXX
End Sub

Sub DetailsPaymentTrueFalse(XXX As Boolean)
    frmAMISJournalEntry_DetailPayment.picDetails.Enabled = XXX
    frmAMISJournalEntry_DetailPayment.Picture2.Visible = XXX
End Sub

Function VendorATC(XXX As String) As String
    Dim rsVendorATC                                         As ADODB.Recordset
    Set rsVendorATC = New ADODB.Recordset
    rsVendorATC.Open "SELECT ATC FROM ALL_VENDOR_TABLE WHERE CODE='" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsVendorATC.EOF And Not rsVendorATC.BOF Then
        VendorATC = Null2String(rsVendorATC!ATC)
    End If
    Set rsVendorATC = Nothing
End Function

Function CHECKUNPOSTED(xFROM As Date, xTO As Date) As Boolean
    Set rsCHECKUNPOSTED = New ADODB.Recordset
    rsCHECKUNPOSTED.Open "SELECT * FROM AMIS_JOURNAL_HD WHERE JDATE BETWEEN '" & xFROM & "' AND '" & xTO & "' AND STATUS='N'", gconDMIS, adOpenForwardOnly
    If Not rsCHECKUNPOSTED.EOF And Not rsCHECKUNPOSTED.BOF Then
        CHECKUNPOSTED = True
        UnpostedReportPrinting
    Else
        CHECKUNPOSTED = False
    End If
    Set rsCHECKUNPOSTED = Nothing
End Function

Sub UnpostedReportPrinting()
    Dim xlApp                                               As Excel.Application
    Dim xlBook                                              As Excel.Workbook
    Dim xlSheet1                                            As Excel.Worksheet
    If Len(Dir(AMIS_REPORT_PATH & "UnpostedReport.xlt")) = 0 Then
        MsgBox "Please find excel template for Unposted Report", vbInformation
        Exit Sub
    End If
    Dim i As Integer
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(AMIS_REPORT_PATH & "UnpostedReport.xlt")
    Set xlSheet1 = xlBook.Worksheets(1)

    xlSheet1.Cells(1, "A").Font.Bold = True
    xlSheet1.Cells(1, "A") = COMPANY_NAME
    'xlSheet1.Cells(1, "B").Font.Bold = True

    xlSheet1.Cells(2, "A").Font.Bold = True
    xlSheet1.Cells(2, "A") = COMPANY_ADDRESS
    'xlSheet1.Cells(2, "B").Font.Bold = True
    
    Do While Not rsCHECKUNPOSTED.EOF
        xlSheet1.Cells(7 + i, "A") = Null2String(rsCHECKUNPOSTED!JDATE)
        xlSheet1.Cells(7 + i, "B") = Null2String(rsCHECKUNPOSTED!JTYPE)
        xlSheet1.Cells(7 + i, "C") = rsCHECKUNPOSTED!VOUCHERNO
        i = i + 1
    rsCHECKUNPOSTED.MoveNext
    Loop
    
    xlApp.Visible = True
    Set xlBook = Nothing
    Set xlSheet1 = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = 0
End Sub
