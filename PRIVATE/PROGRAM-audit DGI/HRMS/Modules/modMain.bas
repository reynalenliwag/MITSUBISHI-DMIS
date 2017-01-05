Attribute VB_Name = "modHRMSMain"
Option Explicit
Dim rsProfile                                                         As ADODB.Recordset
Private Const REG_DWORD = 4&
Private Const REG_SZ = 1
Private Const HKEY_CURRENT_USER = &H80000001
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias _
                                      "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
                                                       phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
                                       "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
                                                         ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal _
                                                                                                                      cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" _
                                     (ByVal hKey As Long) As Long
Dim SERVER_CONNECT, LOCAL_DB_NAME, LOCAL_DATABASE_NAME, LOCAL_DSN_NAME As String
Attribute LOCAL_DB_NAME.VB_VarUserMemId = 1073741825
Attribute LOCAL_DATABASE_NAME.VB_VarUserMemId = 1073741825
Attribute LOCAL_DSN_NAME.VB_VarUserMemId = 1073741825
Public EMPLOYEE_TAX_BASE                                              As String

Public Sub Main()
    If App.PrevInstance = True Then
        MsgBox "There is open HRMS application", vbInformation
        End
    End If
    
    SERVERNAME = GetSetting("DMIS 2.0", "SETTINGS", "SERVERNAME")
    SQLSERVERNAME = GetSetting("DMIS 2.0", "SETTINGS", "SQLSERVERNAME")
    DATABASE = GetSetting("DMIS 2.0", "SETTINGS", "DATABASE")
    If SQLSERVERNAME = "" Or DATABASE = "" Then
        MsgBox "Application Not Yet Configured. Please Configure Server Setting From ADSA.", vbCritical
        End
        Exit Sub
    End If

    ConnStr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & DATABASE & " " & " ;Data Source=" & SQLSERVERNAME
    DMIS_Connection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & DATABASE & " ;Data Source=" & SQLSERVERNAME
    DMIS_REPORT_Connection = "DSN=" & DATABASE & " ;DSQ=" & SQLSERVERNAME
    DMIS_Audit_Connection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=DMIS_AUDIT ;Data Source=" & SQLSERVERNAME


    MODULENAME = "HRMS"
    frmMain.Show

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
    frmSplash.Show: frmSplash.ZOrder 0
    frmSplash.labCon.Caption = "Connecting to SQL Server... Please wait...": DoEvents
    ApplySecurityValidation = True
    On Error GoTo ConnErr
    Set gconDMIS = New ADODB.Connection
    gconDMIS.ConnectionString = DMIS_Connection
    gconDMIS.Mode = adModeReadWrite
    gconDMIS.CursorLocation = adUseClient
    frmSplash.labCon.Caption = "Connecting to CSMS Database... Please wait..."
    gconDMIS.Open
    SEARCH_TAB = 0
    OpenSQLDb = True
    SetCompanyProfile
    Screen.MousePointer = 0
    frmSplash.Command1.Value = True
    Exit Function

ConnErr:
    ShowVBError
    MsgBoxXP "I can't open a connection!!! You may have to " & vbCrLf & _
             "LOG-IN again to connect to the server to run this program. " & vbCrLf & _
             "If you don't have an account contact your friendly " & vbCrLf & _
             "neighborhood SysAdministrator.", "ERROR", XP_OKOnly, msg_Critical
    End
End Function

Public Sub SetUserSettings()
    Call SetUserPathSettings
    With frmMain
        .StatusBar1.Panels(1).Text = "User: " & LOGNAME
        .StatusBar1.Panels(2).Text = "Level: " & LOGLEVEL
        .StatusBar1.Panels(3).Text = "Date: " & Format(LOGDATE, "long date")
        .StatusBar1.Panels(4).Text = "Login Time: " & LOGTIME
        .StatusBar1.Panels(9).Text = "Server Name: " & SQLSERVERNAME
        If CUTTOFF_CODE = "1" Then
            .StatusBar1.Panels(10).Text = "1st Cut-Off" & "-" & PAY_MONTH & " " & PAY_YEAR
        ElseIf CUTTOFF_CODE = "2" Then
            .StatusBar1.Panels(10).Text = "2nd Cut-Off" & "-" & PAY_MONTH & " " & PAY_YEAR
        End If


    End With
End Sub

Public Sub GetThePayrollCode()
    Dim RSTMP                                                         As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select * from HRMS_PayrollSetup")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        PAYROLLCODE_FROM1 = N2Str2Zero(RSTMP!FROMDATE1)
        PAYROLLCODE_TO1 = N2Str2Zero(RSTMP!TODATE1)
        PAYROLLCODE_FROM2 = N2Str2Zero(RSTMP!FROMDATE2)
        PAYROLLCODE_TO2 = N2Str2Zero(RSTMP!TODATE2)
        EMPLOYEE_TAX_BASE = Null2String(RSTMP!TAXCODE)
        CUTTOFF_CODE = Null2String(RSTMP!NOTEDBY2)
        PAY_MONTH = N2Str2Zero(RSTMP!PERIODMONTH)
        PAY_YEAR = N2Str2Zero(RSTMP!PERIODYEAR)
    End If
    Set RSTMP = Nothing
End Sub

Public Function FindPrevMonth(comboMonth As String) As Integer
    If comboMonth = "January" Then FindPrevMonth = 12
    If comboMonth = "February" Then FindPrevMonth = 1
    If comboMonth = "March" Then FindPrevMonth = 2
    If comboMonth = "April" Then FindPrevMonth = 3
    If comboMonth = "May" Then FindPrevMonth = 4
    If comboMonth = "June" Then FindPrevMonth = 5
    If comboMonth = "July" Then FindPrevMonth = 6
    If comboMonth = "August" Then FindPrevMonth = 7
    If comboMonth = "September" Then FindPrevMonth = 8
    If comboMonth = "October" Then FindPrevMonth = 9
    If comboMonth = "November" Then FindPrevMonth = 10
    If comboMonth = "December" Then FindPrevMonth = 11
End Function


