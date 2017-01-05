Attribute VB_Name = "modOSMSMain"
Option Explicit
'==========================================================================================
'FUNCTION / FEATURE :Added New OSMS Main Modules Same Like Every Other Modules
'DATE STARTED       :04272007
'LAST UPDATED       :
'DATABASE UPDATES   :ALL TABLES ARE RENAMED WITH PREFIX OSMS_
'WHO UPDATED        :AXP
'==========================================================================================
'FUNCTION / FEATURE :To Enable One Picture Box in Form And Look Make It Look A Like VB Modal Form , Use Ful in Master and Detail Record
'DATE STARTED       :04272007
'LAST UPDATED       :04272007
'DATABASE UPDATES   :
'WHO UPDATED        :AXP
'UPDATING CODE      :AXP0427200716:46
'==========================================================================================
'FUNCTION / FEATURE :Centering Picture Box aslo used in Master and Detail
'DATE STARTED       :04272007
'LAST UPDATED       :
'DATABASE UPDATES   :
'WHO UPDATED        :AXP
'UPDATING CODE      :AXP0427200716:47
'==========================================================================================
'FUNCTION / FEATURE :Added getsetting Method TO Get the Report Path of The System, Skin Path from Regsitry
'DATE STARTED       :04272007
'LAST UPDATED       :04272007
'DATABASE UPDATES   :
'WHO UPDATED        :AXP
'UPDATING CODE      :AXP0427200716:49
'==========================================================================================
'FUNCTION / FEATURE : Added Column Resize Function which Resizes the Listview to Certain Size Uses on Form Load Generally
'DATE STARTED       :5/4/200715:20
'LAST UPDATED       :5/4/200715:20
'DATABASE UPDATES   :
'WHO UPDATED        :AXP5/4/200715:20
'==========================================================================================
'FUNCTION / FEATURE :AddColumnHeader:Adds COlumn TO To List View Generally Used Together with ResizeColumn Header
'DATE STARTED       :5/4/200715:24
'LAST UPDATED       :5/4/200715:24
'DATABASE UPDATES   :
'WHO UPDATED        :AXP5/4/200715:24
'==========================================================================================

Sub Main()
DMIS_Connection = GetSetting("DMIS 2.0")
    frmMain.Show
    frmMain.ZOrder 1
    frmSplash.Show
    frmSecurity.Show vbModal
    frmSecurity.ZOrder 1
    frmMainMenu.Show
    'Upating Code       :AXP-062620071225
    ReminderModule ""
End Sub



Public Function OpenSqlDb() As Boolean
'AXP0427200716:49
    Screen.MousePointer = 11
    frmSecurity.Hide
    frmSplash.Show
    frmSplash.ZOrder 0
    frmSplash.labCon.Caption = "Connecting to SQL Server... Please wait..."
    DoEvents
    ApplySecurityValidation = True
          
    On Error GoTo ConnErr
    Set gconDMIS = New ADODB.Connection
    gconDMIS.ConnectionString = DMIS_Connection
    frmSplash.labCon.Caption = "Connecting to DMIS Database... Please wait..."
    DoEvents
    gconDMIS.Mode = adModeReadWrite
    gconDMIS.CursorLocation = adUseClient
    gconDMIS.Open

    Dim rsProfile As ADODB.Recordset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        Company_name = Null2String(rsProfile!CompanyName)
        Company_Address = Null2String(rsProfile!Companyaddress)
    End If
    'deGVD.deConnGVD.ConnectionString = GVDDATA_Connection
    OpenSqlDb = True
    Screen.MousePointer = 0
    frmSplash.Command1.Value = True

    Exit Function

ConnErr:
    MsgBox Err.Description
    MsgBox "I can't open a connection!!! You may have to " & vbCrLf & _
           "LOG-IN again to connect to the server to run this program. " & vbCrLf & _
           "If you don't have an account contact your friendly " & vbCrLf & _
           "neighborhood SysAdministrator.", _
           vbOKOnly + vbCritical, "ERROR"
End Function



Public Sub SetUserMenuSettings()
    With frmMain
        If LOGLEVEL = .wizEnc1.DecryptAccess("41444D_]jUU") Or LOGLEVEL = .wizEnc1.DecryptAccess("415554ctleze") Then
            '     .mnuMaintenance.Enabled = True
        Else
            '    .mnuMaintenance.Enabled = False
        End If
        .StatusBar1.Panels(1).Text = "User: " & LOGNAME
        .StatusBar1.Panels(2).Text = "Level: " & LOGLEVEL
        .StatusBar1.Panels(3).Text = "Date: " & Format(LOGDATE, "long date")
        .StatusBar1.Panels(4).Text = "Login Time: " & LOGTIME
    End With
End Sub





Sub ShowHidePictureBox(hwnd As Long, State As Boolean, FRM As Form)
'AXP0427200716:46
    Dim cntl As Control
    For Each cntl In FRM.ControlS
        If TypeOf cntl Is PictureBox Then
            If Not cntl.Container.hwnd = hwnd Then
                If cntl.hwnd = hwnd Then
                    cntl.Enabled = State
                    cntl.Visible = State
                    If State = True Then
                        cntl.ZOrder 0
                    Else
                        cntl.ZOrder 1
                    End If
                Else

                    cntl.Enabled = Not (State)
                    If State = True Then
                        ' cntl.ZOrder 1
                    Else
                        'cntl.ZOrder 0
                    End If
                End If
            End If
        End If
    Next

End Sub



Sub CenterPictureBox(picx As PictureBox, FRM As Form)
'AXP0427200716:47
    picx.Left = (FRM.ScaleWidth - picx.Width) / 2
    picx.Top = (FRM.ScaleHeight - picx.Height) / 2
End Sub





Public Sub AddColumnHeader(StringHeaders As String, lvGrid As ListView)
'AXP5/4/200715:24
    Dim AR() As String
    Dim cWidth As Long
    Dim I As Integer

    AR = Split(StringHeaders, ",")
    cWidth = lvGrid.Width
    lvGrid.ColumnHeaders.Clear
    For I = LBound(AR) To UBound(AR)
        lvGrid.ColumnHeaders.Add , , AR(I)
    Next
    Erase AR
    StringHeaders = vbNullString
End Sub

Public Sub ResizeColumnHeader(grd As Object, SizeArray As String)
'AXP5/4/200715:20
    grd.Visible = False

    Dim AR() As String
    Dim cWidth As Long
    Dim I As Integer
    Dim scwidth As Long
    AR = Split(SizeArray, ",")
    cWidth = grd.Width

    If TypeOf grd Is ListView Then
        For I = LBound(AR) To UBound(AR)
            If I <= grd.ColumnHeaders.Count Then
                scwidth = cWidth * (CDec(AR(I)) / 100)
                grd.ColumnHeaders(I + 1).Width = scwidth
            End If
        Next


    End If

    Erase AR
    grd.Visible = True
End Sub
