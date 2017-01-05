Attribute VB_Name = "modAMISMain"
Option Explicit
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

Public Sub ShowReport(ReportName As Variant, ReportFolder As Variant, filter As Variant, ReportHeading As String, REPORT_DATE As String, WithDate As Boolean)
    Screen.MousePointer = 11
    Dim rsProfile                                                     As ADODB.Recordset
    Dim CrystalRpt                                                    As Crystal.CrystalReport
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
        
        'UPDATED BY: JUN/ARNOLD------------------------------------------------------------------------
        'DATE UPDATED: 06-11-2009
        If ReportName = "CustomersSubsidiaryLedger" Then
                CrystalRpt.Formulas(55) = "CUST_OPENING = '" & ToDoubleNumber(xBALANCE) & "'"
                CrystalRpt.Formulas(56) = "COB_DATE ='" & Format(BEG_BALANCE_DATE, "mmm dd yyyy") & "'"
        End If
        'UPDATED BY: JUN/ARNOLD------------------------------------------------------------------------


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
            PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", filter, DMIS_REPORT_Connection, 1
        End If
        CrystalRpt.PageZoom 89
    End If
    Screen.MousePointer = 0

End Sub

Public Sub ShowRangeReport(V_From As String, V_To As String, ReportName As Variant, ReportFolder As Variant, filter As Variant, ReportHeading As String, WithDate As Boolean)
    On Error Resume Next
    Screen.MousePointer = 11
    Dim rsProfile                                                     As ADODB.Recordset
    Dim CrystalRpt                                                    As Crystal.CrystalReport
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
Dim FieldName As String
If vJtype = "APJ" Or vJtype = "CDJ" Or vJtype = "SJ" Or vJtype = "CRJ" Or vJtype = "GJ" Then
    FieldName = Trim(vJtype & "Month" & vAcctngMonth)
    
    Dim rsCheckRecord As ADODB.Recordset
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
    Dim SQL                                                           As String
    Dim rs                                                            As New ADODB.Recordset
    SQL = "SELECT Description FROM AMIS_ChartAccount where acctcode=" & XXX & ""
    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)
    If Not rs.EOF And Not rs.BOF Then
        ReturnAccountName = Null2String(rs!Description)
    End If
    Set rs = Nothing
End Function


