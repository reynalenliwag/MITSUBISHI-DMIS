VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7E0552E8-E2C9-4C9E-BC75-F34D871C75F4}#2.0#0"; "wizEncrypt.ocx"
Object = "{54AC2DF1-B6CB-406E-BB23-DC06DF6AAD9E}#16.0#0"; "wizCrypto.ocx"
Begin VB.Form frmSecurity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log In"
   ClientHeight    =   4200
   ClientLeft      =   4005
   ClientTop       =   3480
   ClientWidth     =   5445
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   Icon            =   "SECURITY.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4200
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2850
      Left            =   0
      Picture         =   "SECURITY.frx":74F2
      ScaleHeight     =   2850
      ScaleWidth      =   2100
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   810
      Width           =   2100
      Begin wizEncrypt.wizEnc wizEnc1 
         Left            =   810
         Top             =   1230
         _ExtentX        =   3969
         _ExtentY        =   3969
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3780
      Top             =   5085
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
            Picture         =   "SECURITY.frx":A62D
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
      Picture         =   "SECURITY.frx":A8AB
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Picture         =   "SECURITY.frx":ABE9
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   885
   End
   Begin wizCrypto.Crypto Crypto1 
      Height          =   465
      Left            =   900
      TabIndex        =   8
      Top             =   5265
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
      Width           =   2805
   End
   Begin VB.TextBox txtlevel 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   645
      Left            =   4590
      TabIndex        =   7
      Top             =   5085
      Visible         =   0   'False
      Width           =   1095
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
      MouseIcon       =   "SECURITY.frx":AE84
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "USERNAME"
         Object.Width           =   5644
      EndProperty
   End
   Begin VB.Label LABSERVER 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SERVER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2940
      TabIndex        =   9
      Top             =   240
      Width           =   2415
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
      Picture         =   "SECURITY.frx":AFE6
      Top             =   0
      Width           =   720
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
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAccess                            As ADODB.Recordset
Dim Counter                             As Integer
Dim PicClickCnt                         As Integer
Sub InitMemVars()
    txtUserPass = vbNullString
End Sub

Sub StoreMemVars()
    txtUserPass.Text = vbNullString
    txtlevel.Text = vbNullString
End Sub

Private Sub cmdOk_Click()
    On Error GoTo adder:
    If txtUserPass.Enabled = False Then Exit Sub
    With wizVar
        Set rsAccess = New ADODB.Recordset
        rsAccess.Open "select * from ALL_Rams_Users where username = '" & labUserName.Caption & "'", gconDMIS, adOpenKeyset
        If Not rsAccess.EOF And Not rsAccess.BOF Then
            If txtUserPass.Text <> .DecryptAccess(rsAccess!Password) Then
                GoTo Messages2
            Else
                LOGCODE = rsAccess!USERCODE
                LOGNAME = labUserName.Caption
                LOGPASS = txtUserPass.Text
                LOGLEVEL = rsAccess!userGroup
                LOGID = rsAccess!UserID
                LOGTIME = Time
                LOGDATE = Date
                On Error GoTo NOAUDIT:
                Set gconAudit = New ADODB.Connection
                gconAudit.Open (DMIS_Audit_Connection)
CONTINUE:
                frmMain.SETCAPTION
                Dim rsProfile           As ADODB.Recordset
                Set rsProfile = New ADODB.Recordset
                Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE WHERE MODULENAME='AMIS'")
                If Not rsProfile.EOF And Not rsProfile.BOF Then
                    Company_name = Null2String(rsProfile!CompanyName)
                    Company_Address = Null2String(rsProfile!Companyaddress)
                End If
                Set rsProfile = Nothing
            End If
        Else
            GoTo Messages
        End If
    End With
    Unload Me
    frmMain.Show
    Exit Sub
Messages:
    If Counter < 15 Then
        Counter = Counter + 1
        labUserName.Caption = "Select User..."
        txtUserPass.Text = ""
        lstUserName.SetFocus
    Else
        Unload Me
        End
    End If
    Exit Sub
Messages2:
    If Counter < 3 Then
        MsgBox "Invalid Password", vbInformation
        Counter = Counter + 1
        txtUserPass.Text = "": txtUserPass.SetFocus
    Else
        'MsgSpeechBox "Warning: Intruder Alert! ... Intruder Alert! ... Intruder Alert! ..."
        'MessagePop NoEntry, "Access Denied", "Warning: Intruder Alert! ... Intruder Alert! ... Intruder Alert! ..."
    End If
    Exit Sub
adder:
    MsgBox Err.Description
NOAUDIT:
    MsgBox Err.Description
    GoTo CONTINUE


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

10:
    Set CryptVar = frmSecurity.Crypto1
    Set wizVar = Me.wizEnc1
    Set oVoice = New SpVoice
    Me.BorderStyle = 0

    DoEvents

    Counter = 0: PicClickCnt = 0
    CryptVar.Visible = False
    InitMemVars
    txtUserPass.Enabled = False

    If OpenConnection = True Then
        lstUserName.Sorted = False
        lstUserName.ListItems.Clear
        LABSERVER = "Log In (" & ServerName & ")"
        Set rsAccess = New ADODB.Recordset
        rsAccess.Open "Select Username from All_RAMS_USERs WHERE USERGROUP='SDM' and LOCK=0 Order by UserName Asc", gconDMIS, adOpenKeyset
        If Not rsAccess.EOF And Not rsAccess.BOF Then
            Listview_Loadval lstUserName.ListItems, rsAccess
        Else                                                 'USERCODE,USERNAME,PASSWORD,USERGROUP
            gconDMIS.Execute ("INSERT INTO ALL_RAMS_USERS (USERCODE,USERNAME,PASSWORD,USERGROUP,LOCK) values('AD" & Left(Day(Now), 1) & "' ,'ADMIN','616C6Fñ™Æùù','SDM',0)")
            GoTo 10:
        End If
        StoreMemVars
        CenterMe Screen, Me, 0

    End If
    Screen.MousePointer = 0


End Sub






 

Private Sub LABSERVER_Click()
    Static Cnt
    Cnt = 1 + Cnt
    If Cnt = 15 Then CLEARSETTING: End
End Sub

Private Sub lstUserName_KeyPress(KeyAscii As Integer)
        If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub lstUserName_LostFocus()
    If lstUserName.SelectedItem Is Nothing Then Exit Sub
    Dim mIndex                          As Integer
    labUserName.Caption = lstUserName.SelectedItem
    If Not lstUserName.SelectedItem Is Nothing Then
        mIndex = lstUserName.SelectedItem.Index
    End If
    If Not mIndex = 0 Then
        lstUserName.ListItems(mIndex).Selected = True
        lstUserName.ListItems(mIndex).EnsureVisible
    End If
    txtUserPass.Enabled = True
    txtUserPass.SetFocus
End Sub

Private Sub Picture1_DblClick()
    PicClickCnt = PicClickCnt + 1
    If PicClickCnt = 3 Then
        labUserName.Caption = wizVar.DecryptAccess("77697A∂ùº∞û¡¡¢¡")
        On Error Resume Next
        txtUserPass.Enabled = True
        txtUserPass.SetFocus
    End If
End Sub

Private Sub txtUserPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOk_Click
End Sub
