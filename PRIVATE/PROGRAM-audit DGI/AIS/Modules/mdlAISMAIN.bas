Attribute VB_Name = "mdlAISMAIN"
Public LOGCODE, LOGNAME, LOGLEVEL, LOGTIME, LOGDATE As String


Public Maxwiz, AccessCNT As Long
Public wizVar, CryptVar, EmpInfoEmpno As Object
Public rsProfile As ADODB.Recordset

Public COMPANY_NAME, COMPANY_ADDRESS, COMPANY_TIN As String
Public PREPARED_BY, CHECKED_BY, APPROVED_BY, ACCOUNT_NO, BANK_NAME, BANK_LOCATION, BANK_MANAGER, SECRETARY, NOTED_BY As String
Public EMPLOYEE_NO       As String
Public PREPARED_BY_DESIGNATION, APPROVED_BY_DESIGNATION As String

Public PROCESS_OPTION    As String
Public IMPNO             As String

Public OVERTIME_CODES    As String
Public OVERTIME_RATE     As Double

Public Sub Main()
'MsgBox Command
'If Command = "K" Then


'Else
    frmMain.Show
    frmSplash.Show
    frmSecurity.Show vbModal
    frmSecurity.ZOrder 1
'End If

    frmAISMAIN2.Show
End Sub

Public Function OpenSQLDb() As Boolean
    Screen.MousePointer = 11
    frmSecurity.Hide
    frmSplash.Show
    frmSplash.labCon.Caption = "Connecting to HRMS Database... Please wait..."
    frmSplash.ZOrder 0
    DoEvents
    ApplySecurityValidation = True
    If ValidPassword(Trim(LOGNAME), Trim(LOGPASS), "SYSTEM") Then
        AIS_REPORT_PATH = "\\SERVER\HMI\REPORTS\AIS\"
        HRMS_PICTURES_PATH = "\\SERVER\HMI\REPORTS\HRMS\images\"
        On Error GoTo ConnErr
        Set gconDMIS = New ADODB.Connection
        gconDMIS.ConnectionString = DMIS_Connection
        gconDMIS.Mode = adModeReadWrite
        gconDMIS.CursorLocation = adUseClient
        gconDMIS.Open
        OpenSQLDb = True
        Dim rsProfile As ADODB.Recordset
        Set rsProfile = New ADODB.Recordset
        Set rsProfile = gconDMIS.Execute("Select * from ALL_Profile Where ModuleName='HRMS'")
        If Not rsProfile.EOF And Not rsProfile.BOF Then
            COMPANY_NAME = Null2String(rsProfile!CompanyName)
            COMPANY_ADDRESS = Null2String(rsProfile!CompanyAddress)
            COMPANY_TIN = Null2String(rsProfile!CompanyTINNo)
            PREPARED_BY = Null2String(rsProfile!PreparedBy)
            CHECKED_BY = Null2String(rsProfile!CheckedBy)
            APPROVED_BY = Null2String(rsProfile!ApprovedBy)
            ACCOUNT_NO = Null2String(rsProfile!AccountNo)
            BANK_NAME = "SECURITY BANK"
            BANK_LOCATION = "NAGA CITY"
            PREPARED_BY_DESIGNATION = "Secretary"
            APPROVED_BY_DESIGNATION = "Operations Manager"
            BANK_MANAGER = Null2String(rsProfile!BankManager)
            SECRETARY = Null2String(rsProfile!SECRETARY)
            NOTED_BY = Null2String(rsProfile!NotedBy1)
        End If
    End If
    Screen.MousePointer = 0
    frmSplash.labCon.Caption = "HRMS Successfully Connected... Initializing Please wait..."
    Unload frmSplash
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



Public Sub SetUserMenuSettings()
    With frmMain
        If LOGLEVEL = wizVar.DecryptAccess("41444D_]jUU") Or LOGLEVEL = .wizEnc1.DecryptAccess("415554ctleze") Then
            '.mnuMaintenance.Enabled = True
            ''''FIX HERE
        Else
            '.mnuMaintenance.Enabled = False
        End If
        .StatusBar1.Panels(1).Text = "User: " & LOGNAME
        .StatusBar1.Panels(2).Text = "Level: " & LOGLEVEL
        .StatusBar1.Panels(3).Text = "Date: " & Format(LOGDATE, "long date")
        .StatusBar1.Panels(4).Text = "Login Time: " & LOGTIME
    End With
    'SetDataEntryVariables
End Sub
