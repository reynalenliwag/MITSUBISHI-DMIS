Attribute VB_Name = "modAMISMain"
Option Explicit
Dim xEntity                                       As String
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
    CHARGE_SALES = "'41'"
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
        MsgBox Err.Description
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
        .StatusBar1.Panels(9).Text = "Server Name: " & SQLSERVERNAME
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
    Dim rsProfile                                 As ADODB.Recordset
    Dim CrystalRpt                                As Crystal.CrystalReport
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
        If ReportName = "GeneralJournal" Then
            If COMPANY_CODE = "HAI" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HGC" Or COMPANY_CODE = "HCI" Then
                GJ_REMARKS_XXX
                CrystalRpt.Formulas(59) = "REMARK = '" & xEntity & "'"
            End If
        End If

        If COMPANY_CODE = "HGC" Then
            ' Update By BTT : 07282008
            If ReportName = "AccountsPayable" Or ReportName = "SalesJournal" Or ReportName = "GeneralJournal" Or ReportName = "CashDisbursement" Or ReportName = "CashReceipts" Then
                If REPRINT_CAPTION = "YES" Then
                    CrystalRpt.Formulas(3) = "Reprint= '" & "REPRINTED" & "'"
                Else
                    CrystalRpt.Formulas(3) = "Reprint= '" & "" & "'"
                End If
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
    Dim GJ_REMARKS                                As ADODB.Recordset
    Dim xENTITY2                                  As String
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
        gconDMIS.Execute ("UPDATE AMIS_JOURNAL_HD SET REMARKS=NULL WHERE VOUCHERNO = '" & frmAMISJournalEntry_GJ.txtVoucherNo.Text & "' AND JTYPE = 'GJ'")
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
    Dim rsProfile                                 As ADODB.Recordset
    Dim CrystalRpt                                As Crystal.CrystalReport
    frmMain.rptMain.Reset
    Set CrystalRpt = frmMain.rptMain
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE WHERE MODULENAME = '" & MODULENAME & "'")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        CrystalRpt.ReportFileName = AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt"
        CrystalRpt.WindowShowPrintSetupBtn = True
        CrystalRpt.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        CrystalRpt.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"

        If WithDate = True Then CrystalRpt.Formulas(2) = "ReportDate = '" & LOGDATE & "'"
        CrystalRpt.Formulas(3) = "FromJDate = '" & V_From & "'"
        CrystalRpt.Formulas(4) = "ToJDate = '" & V_To & "'"
        CrystalRpt.ReportTitle = ReportHeading: CrystalRpt.WindowTitle = ReportHeading
        PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", filter, DMIS_REPORT_Connection, 1
        CrystalRpt.PageZoom 89
    End If
    Screen.MousePointer = 0
End Sub

Public Function CheckIfBookIsOpen(vJtype As String, vAcctngMonth As Integer, vAcctngYear As Integer) As Boolean
    Dim FieldName                                 As String
    If vJtype = "APJ" Or vJtype = "CDJ" Or vJtype = "SJ" Or vJtype = "CRJ" Or vJtype = "GJ" Then
        FieldName = Trim(vJtype & "Month" & vAcctngMonth)

        Dim rsCheckRecord                         As ADODB.Recordset
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

Public Function Null2Bit(XXX As Variant) As Integer
    If Null2Bool(XXX) = True Then
        Null2Bit = 1
    Else
        Null2Bit = 0
    End If
End Function
Function ReturnAccountName(XXX As String) As String
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset
    SQL = "SELECT Description FROM AMIS_ChartAccount where acctcode=" & XXX & ""
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)
    If Not RS.EOF And Not RS.BOF Then
        ReturnAccountName = Null2String(RS!Description)
    End If
    Set RS = Nothing
End Function

Sub FormExistsShow(frmx As Form)
    On Error GoTo ErrorCode
    Dim m_Exists                                  As Boolean
    Dim FRM                                       As Form
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
ErrorCode:
    Err.Clear
End Sub

Sub ReportControlAddColumnHeader(LST As ReportControl, StringHeaders As String)
    Dim ar()                                      As String
    Dim i                                         As Integer


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
    Dim ar()                                      As String
    Dim cWidth                                    As Long
    Dim i                                         As Integer
    Dim scwidth                                   As Long
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
    Dim ar()                                      As String
    Dim cWidth                                    As Long
    Dim i                                         As Integer

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
    Dim i                                         As Long
    Dim ItemDataX                                 As Long
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


