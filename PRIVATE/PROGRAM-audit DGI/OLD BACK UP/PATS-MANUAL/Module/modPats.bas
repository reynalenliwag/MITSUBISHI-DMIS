Attribute VB_Name = "modPats"
Option Explicit
'Public gconDMIS As Database
Public rsEMPINFO                                       As ADODB.Recordset
Public rsAttend                                        As ADODB.Recordset
Public rsCard                                          As ADODB.Recordset
Public rsTempSummary                                   As ADODB.Recordset
Public rsdivref                                        As ADODB.Recordset
Public Flag                                            As Integer
Public MO                                              As String
Public Path1                                           As String
Public Path2                                           As String
Public OldPass                                         As String

Public MODULENAME                                      As String
Public ConnStr                                         As String
Public DMIS_Connection                                 As String
Public DMIS_REPORT_Connection                          As String
Public DMIS_Audit_Connection                           As String
Public SERVERNAME                                      As String
Public SQLSERVERNAME                                   As String
Public DATABASE                                        As String
'Public thedivcode As String

Public Sub Main()
    If App.PrevInstance = True Then
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
    DMIS_Audit_Connection = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ADMS_AUDIT ;Data Source=" & SQLSERVERNAME
    MODULENAME = "PATS"
    HRMS_PICTURES_PATH = GetSetting("DMIS 2.0", "REPORTS", "HRMS") & "\images\"
    If OpenSQLDb = True Then
        frmLOGIN.Show
    End If
End Sub

Public Function OpenSQLDb() As Boolean
    Screen.MousePointer = 11
    'HRMS_PICTURES_PATH = "\\DMISSERVER\DMIS 2.0\REPORTS\HRMS\images\"
    Set GCONDMIS = New ADODB.Connection
    GCONDMIS.ConnectionString = DMIS_Connection
    GCONDMIS.Mode = adModeReadWrite
    GCONDMIS.CursorLocation = adUseClient
    GCONDMIS.Open
    OpenSQLDb = True
    Set rsEMPINFO = New ADODB.Recordset
    rsEMPINFO.Open "select * from HRMS_EmpInfo where activeinactive='A'", GCONDMIS, adOpenKeyset
    Set rsAttend = New ADODB.Recordset
    rsAttend.Open "select * from HRMS_Attend", GCONDMIS, adOpenKeyset
    Set rsTempSummary = New ADODB.Recordset
    rsTempSummary.Open "select * from HRMS_TempSummary", GCONDMIS, adOpenKeyset
    Set rsdivref = New ADODB.Recordset
    rsdivref.Open "select * from HRMS_DivRef", GCONDMIS, adOpenKeyset
    If Not (rsEMPINFO.BOF And rsEMPINFO.EOF) Then
        rsEMPINFO.MoveFirst
    End If
    If Not (rsAttend.BOF And rsAttend.EOF) Then
        rsAttend.MoveFirst
    End If
    frmLOGIN.Show
    Screen.MousePointer = 0
    Exit Function
ConnErr:
    MsgBox Err.Description
    MsgBox "I can't open a connection!!! You may have to " & vbCrLf & _
           "LOG-IN again to connect to the server to run this program. " & vbCrLf & _
           "If you don't have an account contact your friendly " & vbCrLf & _
           "neighborhood SysAdministrator.", _
           vbOKOnly + vbCritical, "ERROR"
    End
End Function

Function Date2Month(Value As String)
    MO = "January  February March    April    May      June     July     August   SeptemberOctober  November December "
    Date2Month = Mid$(MO, (Month(Value) - 1) * 9 + 1, 9)
End Function
