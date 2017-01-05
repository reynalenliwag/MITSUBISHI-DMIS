Attribute VB_Name = "modCSMIOSMain"
Option Explicit

Public Sub Main()
    If App.PrevInstance = True Then
        MsgBox "There is open CSMS application", vbInformation
        End
    End If
    
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
    frmSplash.Show
    frmSecurity.Show vbModal
    frmMainMenu.Show

 



    'UPdate by AXP-062620071225
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
    End With
End Sub

