VERSION 5.00
Object = "{54AC2DF1-B6CB-406E-BB23-DC06DF6AAD9E}#16.0#0"; "wizCrypto.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmSecurity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log In"
   ClientHeight    =   4200
   ClientLeft      =   4005
   ClientTop       =   3435
   ClientWidth     =   5070
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   Icon            =   "Security.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4200
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2850
      Left            =   0
      Picture         =   "Security.frx":74F2
      ScaleHeight     =   2850
      ScaleWidth      =   2130
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   780
      Width           =   2130
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2970
      Top             =   7035
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Security.frx":A62D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4110
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Security.frx":A8AB
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cancel"
      Top             =   3480
      Width           =   885
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Security.frx":ABE9
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Log-In"
      Top             =   3480
      Width           =   885
   End
   Begin wizCrypto.Crypto Crypto1 
      Height          =   465
      Left            =   240
      TabIndex        =   8
      Top             =   7125
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   820
   End
   Begin VB.TextBox txtUserPass 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   22
      PasswordChar    =   "l"
      TabIndex        =   4
      Top             =   3015
      Width           =   2835
   End
   Begin VB.TextBox txtlevel 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   7125
      Visible         =   0   'False
      Width           =   2685
   End
   Begin MSComctlLib.ListView lstUserName 
      Height          =   1890
      Left            =   2160
      TabIndex        =   0
      Top             =   810
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   3334
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   -2147483633
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Security.frx":AE84
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "USERNAME"
         Object.Width           =   5644
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   -45
      X2              =   6750
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   45
      Picture         =   "Security.frx":AFE6
      Top             =   0
      Width           =   720
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   870
      Left            =   -45
      TabIndex        =   1
      Top             =   -90
      Width           =   7665
      _Version        =   655364
      _ExtentX        =   13520
      _ExtentY        =   1535
      _StockProps     =   14
      Caption         =   "           Security Window"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
   Begin VB.Label labUserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select User..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2205
      TabIndex        =   3
      Top             =   2700
      Width           =   1605
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsALL_vw_RAMS_PAccess                              As ADODB.Recordset
Dim COUNTER                                            As Integer
Dim PicClickCnt                                        As Integer

Const LogOnMessage                                     As String = _
      " The System Could Not Log You On, Make Sure Your User Name is Correct then Type " _
      & " Your Password Again. " _
      & " ** Letter in Passwords Must Be Typed Using Correct Care " _
      & " ** Make Sure That Cap Locks Is Not Accidently On. " _


Sub initMemvars()
    txtUserPass = vbNullString
End Sub

Sub StoreMemVars()
    txtUserPass.Text = vbNullString
    txtlevel.Text = vbNullString
End Sub

Sub MakeConnection()
    Screen.MousePointer = 11
    If OpenSQLDb = False Then

        MsgBoxXP "I can't open a connection!!! You may have to " & vbCrLf & _
                 "LOG-IN again to connect to the (LOCAL) to run this program. " & vbCrLf & _
                 "If you don't have an account contact your friendly " & vbCrLf & _
                 "neighborhood SysAdministrator.", "ERROR", XP_OKOnly, msg_Critical
        UnloadApp
    End If
    If OpenSQLAudit = False Then
        MessagePop InfoWarning, "Error Opening Connection ", "Please Contact Your System Administrator. Error Opening Audit Database "
    End If

    frmSplash.Command1.Value = True
    SetUserSettings
    Screen.MousePointer = 0
End Sub

Sub Lst_ALL_vw_RAMS_PAccessLoadval(TisoyView As ListItems, RecSet As ADODB.Recordset)
    Dim Indx                                           As Long
    Dim i                                              As Long
    TisoyView.Clear
    If Not (RecSet.BOF And RecSet.EOF) Then
        While Not RecSet.EOF
            Indx = TisoyView.Count + 1
            TisoyView.Add Indx, , IIf(IsNull(RecSet(0)), "", wizVar.DecryptAccess(Trim(RecSet(0)))), 1, 1
            For i = 1 To RecSet.Fields.Count - 1
                TisoyView(Indx).ListSubItems.Add i, , IIf(IsNull(RecSet(i)), "", wizVar.DecryptAccess(Trim(RecSet(i)))), 1
            Next i
            RecSet.MoveNext
        Wend
    End If
    Set RecSet = Nothing
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ErrorCode:
    If txtUserPass.Enabled = False Then Exit Sub

    With wizVar
        Set rsALL_vw_RAMS_PAccess = New ADODB.Recordset
        'If CHANGE_USER = True Then
        If COMPANY_CODE = COMPANY_VERSION Then
            rsALL_vw_RAMS_PAccess.Open "select * from ALL_vw_RAMS_PAccess where user_name = '" & labUserName.Caption & "'", gconACCESS, adOpenKeyset
        Else
            rsALL_vw_RAMS_PAccess.Open "select * from ALL_vw_RAMS_PAccess where username = '" & labUserName.Caption & "'", gconACCESS, adOpenKeyset
        End If

        If Not rsALL_vw_RAMS_PAccess.EOF And Not rsALL_vw_RAMS_PAccess.BOF Then

            If txtUserPass.Text <> .DecryptAccess(rsALL_vw_RAMS_PAccess!userPass) Then
                MessagePop InfoWarning, "Warning", LogOnMessage
                GoTo Messages2
            Else
                LOGCODE = Null2String(rsALL_vw_RAMS_PAccess!USERCODE)
                LOGNAME = labUserName.Caption
                LOGPASS = txtUserPass.Text
                LOGLEVEL = Null2String(rsALL_vw_RAMS_PAccess!LOGLEVEL)
                LOGID = Null2String(rsALL_vw_RAMS_PAccess!USERID)
                LOGTIME = Time: LOGDATE = Date: MakeConnection

                On Error Resume Next
                Dim RS                                 As ADODB.Recordset
                Set RS = gconDMIS.Execute("Select getdate() as DateNow, host_name() as PCName")
                If RS!pcname <> "SERVER" Or RS!pcname <> "DMISSERVER" Or RS!pcname <> "MASTER" Or RS!pcname <> "MISSERVER" Then
                    Date = RS!DateNow
                    Time = RS!DateNow
                    LOGTIME = Time
                    LOGDATE = Date
                End If
                Set RS = Nothing
            End If
        Else
            MessagePop InfoWarning, "Warning", LogOnMessage
            GoTo Messages
        End If
    End With

    Exit Sub

Messages:

    If COUNTER <= 3 Then
        COUNTER = COUNTER + 1
        labUserName.Caption = "Select User..."
        txtUserPass.Text = ""
        If lstUserName.ListItems.Count > 0 And lstUserName.Enabled = True Then
            lstUserName.SetFocus
        End If
    Else
        Me.Visible = False
        MessagePop NoEntry, "Access Denied", "Warning: Intruder Alert! ..."
        MsgBox "Critical Error: System will Shutdown!", vbCritical, "Login Failed"
        End
    End If
    Exit Sub

Messages2:
    If COUNTER <= 3 Then
        COUNTER = COUNTER + 1
        txtUserPass.Text = ""
        On Error Resume Next
        txtUserPass.SetFocus


    Else
        Me.Visible = False
        MessagePop NoEntry, "Access Denied", "Warning: Intruder Alert! ..."
        MsgBox "Critical Error: System will Shutdown!", vbCritical, "Login Failed"
        End
    End If

    Exit Sub

ErrorCode:
    MsgBox "I can't open a connection!!! You may have to " & vbCrLf & _
           "LOG-IN again to connect to the server to run this program. " & vbCrLf & _
           "If you don't have an account contact your friendly " & vbCrLf & _
           "neighborhood SysAdministrator.", _
           vbOKOnly + vbCritical, "ERROR"
    End
End Sub

Private Sub cmdCancel_Click()
    End
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn Or vbKeyTab
        If Me.ActiveControl.Name = "cbousername" Then
            If lstUserName.SelectedItem <> "" Then SendKeys "{TAB}"
        End If
        If Me.ActiveControl.Name = "txtUserPass" Then
            If txtUserPass.Text <> "" And labUserName.Caption <> "" Then cmdOk_Click
        End If
    Case vbKeyF4
        If Picture1.Visible = False Then Unload Me
    Case Else
        MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    DrawXPCtl Me
    Dim FILTERUSER                                     As String
    Set wizVar = frmMain.wizEnc1
    Set CryptVar = frmSecurity.Crypto1
    Set oVoice = New SpVoice
    Me.BorderStyle = 0
    DoEvents
    COUNTER = 0: PicClickCnt = 0
    CryptVar.Visible = False
    initMemvars
    txtUserPass.Enabled = False
    Dim RSX                                            As ADODB.Recordset
    Dim UIDX                                           As Long

    If OpenAccessDb = True Then
        lstUserName.Sorted = False
        lstUserName.ListItems.Clear
        GET_COMPANYCODE
        COMPANYCODE_VERSION
        DMIS_VERSION

10      Set rsALL_vw_RAMS_PAccess = New ADODB.Recordset

        FILTERUSER = " USERID IN (SELECT USERID FROM ALL_RAMS_USER_MODULES WHERE MAINMODULENAME='" & App.TITLE & "')"

        On Error GoTo USERERROR
        'If CHANGE_USER = True Then
        If COMPANY_CODE = COMPANY_VERSION Then
            rsALL_vw_RAMS_PAccess.Open "select User_Name from ALL_vw_RAMS_PAccess WHERE " & FILTERUSER & " AND LOCK=0 AND loglevel<>'SDM' Order by User_Name Asc", gconACCESS, adOpenKeyset
        Else
            rsALL_vw_RAMS_PAccess.Open "select Username from ALL_vw_RAMS_PAccess WHERE " & FILTERUSER & " AND LOCK=0 AND loglevel<>'SDM' Order by UserName Asc", gconACCESS, adOpenKeyset
        End If

LISTVIEWFILL:         If Not rsALL_vw_RAMS_PAccess.EOF And Not rsALL_vw_RAMS_PAccess.BOF Then
            Listview_Loadval lstUserName.ListItems, rsALL_vw_RAMS_PAccess
        Else
            Dim LC, UN, PW, LV                         As String
            LC = "NET": UN = "NETSPEED": PW = "616C6Fñ™Æùù": LV = "ADM"
            On Error GoTo VBERROR:
            'If CHANGE_USER = True Then
            If COMPANY_CODE = COMPANY_VERSION Then
                Set RSX = gconACCESS.Execute("Insert into ALL_vw_RAMS_PAccess (usercode, user_name, userpass, loglevel) Values ('" & LC & "', '" & UN & "', '" & PW & "', '" & LV & "') " & vbCrLf & "SELECT @@IDENTITY")
            Else
                Set RSX = gconACCESS.Execute("Insert into ALL_vw_RAMS_PAccess (usercode, username, userpass, loglevel) Values ('" & LC & "', '" & UN & "', '" & PW & "', '" & LV & "') " & vbCrLf & "SELECT @@IDENTITY")
            End If
            Set RSX = RSX.NextRecordset
            If Not RSX Is Nothing Then
                UIDX = RSX.Collect(0)
            End If
            gconACCESS.Execute ("INSERT INTO ALL_RAMS_USER_MODULES(USERID,MAINMODULENAME)VALUES(" & UIDX & ",'PMIS')")
            gconACCESS.Execute ("INSERT INTO ALL_RAMS_USER_MODULES(USERID,MAINMODULENAME)VALUES(" & UIDX & ",'SMIS')")
            gconACCESS.Execute ("INSERT INTO ALL_RAMS_USER_MODULES(USERID,MAINMODULENAME)VALUES(" & UIDX & ",'AMIS')")
            gconACCESS.Execute ("INSERT INTO ALL_RAMS_USER_MODULES(USERID,MAINMODULENAME)VALUES(" & UIDX & ",'PMIS')")
            gconACCESS.Execute ("INSERT INTO ALL_RAMS_USER_MODULES(USERID,MAINMODULENAME)VALUES(" & UIDX & ",'CSMS')")
            gconACCESS.Execute ("INSERT INTO ALL_RAMS_USER_MODULES(USERID,MAINMODULENAME)VALUES(" & UIDX & ",'CMIS')")
            gconACCESS.Execute ("INSERT INTO ALL_RAMS_USER_MODULES(USERID,MAINMODULENAME)VALUES(" & UIDX & ",'HRMS')")

            gconACCESS.Execute ("INSERT INTO ALL_RAMS_USERGROUPS(CODE,GROUPNAME)VALUES('ADM','ADMINISTRATORS')")
            gconACCESS.Execute ("INSERT INTO ALL_RAMS_USERGROUPS(CODE,GROUPNAME)VALUES('SDM','SYSTEM ADMINISTRATORS')")
            gconACCESS.Execute ("INSERT INTO ALL_RAMS_USERGROUPS(CODE,GROUPNAME)VALUES('USR','USERS')")
            GoTo 10
        End If
        StoreMemVars
        CenterMe Screen, Me, 0
    Else
        MessagePop InfoWarning, "Warning", LogOnMessage
    End If
    Screen.MousePointer = 0
    Exit Sub
USERERROR:
    'If CHANGE_USER = True Then
    If COMPANY_CODE = COMPANY_VERSION Then
        rsALL_vw_RAMS_PAccess.Open "select User_name from ALL_vw_RAMS_PAccess WHERE LOCK=0 AND loglevel<>'SDM' Order by User_Name Asc", gconACCESS, adOpenKeyset
    Else
        rsALL_vw_RAMS_PAccess.Open "select Username from ALL_vw_RAMS_PAccess WHERE LOCK=0 AND loglevel<>'SDM' Order by UserName Asc", gconACCESS, adOpenKeyset
    End If
    GoTo LISTVIEWFILL
    Exit Sub
VBERROR:
    ShowVBError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If frmSecurity.Visible = True Then
        If Picture1.Visible = False And AccessCNT = 0 Then
            MakeConnection
        Else

        End If
    End If
End Sub

Private Sub lstUserName_DblClick()
    SendKeys "{TAB}"
End Sub

Private Sub lstUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub lstUserName_LostFocus()
    Dim MINDEX                                         As Integer
    Dim FILTERUSER                                     As String

    On Error Resume Next

    labUserName.Caption = lstUserName.SelectedItem
    Set rsALL_vw_RAMS_PAccess = New ADODB.Recordset
    If Not lstUserName.SelectedItem Is Nothing Then
        MINDEX = lstUserName.SelectedItem.Index
    End If


    Set rsALL_vw_RAMS_PAccess = New ADODB.Recordset
    FILTERUSER = "USERID IN (SELECT USERID FROM ALL_RAMS_USER_MODULES WHERE MAINMODULENAME='" & MODULENAME & "')"
    On Error GoTo USERERROR
    'If CHANGE_USER = True Then
    If COMPANY_CODE = COMPANY_VERSION Then
        rsALL_vw_RAMS_PAccess.Open "SELECT USER_NAME FROM ALL_VW_RAMS_PACCESS WHERE " & FILTERUSER & " AND LOCK=0 AND LOGLEVEL<>'SDM' ORDER BY USER_NAME ASC", gconACCESS, adOpenKeyset
    Else
        rsALL_vw_RAMS_PAccess.Open "SELECT USERNAME FROM ALL_VW_RAMS_PACCESS WHERE " & FILTERUSER & " AND LOCK=0 AND LOGLEVEL<>'SDM' ORDER BY USERNAME ASC", gconACCESS, adOpenKeyset
    End If





FillListview:     If Not rsALL_vw_RAMS_PAccess.EOF And Not rsALL_vw_RAMS_PAccess.BOF Then
        Listview_Loadval lstUserName.ListItems, rsALL_vw_RAMS_PAccess
    End If

    If Not MINDEX = 0 Then
        lstUserName.ListItems(MINDEX).Selected = True
        lstUserName.ListItems(MINDEX).EnsureVisible
    End If
    txtUserPass.Enabled = True
    On Error Resume Next
    txtUserPass.SetFocus
    Exit Sub

USERERROR:
    'If CHANGE_USER = True Then
    If COMPANY_CODE = COMPANY_VERSION Then
        rsALL_vw_RAMS_PAccess.Open "SELECT USER_NAME FROM ALL_VW_RAMS_PACCESS WHERE LOCK=0 AND LOGLEVEL<>'SDM' ORDER BY USER_NAME ASC", gconACCESS, adOpenKeyset
    Else
        rsALL_vw_RAMS_PAccess.Open "SELECT USERNAME FROM ALL_VW_RAMS_PACCESS WHERE LOCK=0 AND LOGLEVEL<>'SDM' ORDER BY USERNAME ASC", gconACCESS, adOpenKeyset
    End If
    GoTo FillListview

End Sub

Private Sub txtUserPass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOk_Click
End Sub

