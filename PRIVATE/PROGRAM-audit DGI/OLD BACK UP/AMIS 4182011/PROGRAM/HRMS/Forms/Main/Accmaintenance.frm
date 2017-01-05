VERSION 5.00
Begin VB.Form frmAccMaintenance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password File Maintenance"
   ClientHeight    =   3315
   ClientLeft      =   5415
   ClientTop       =   2385
   ClientWidth     =   5790
   ForeColor       =   &H00D8E9EC&
   Icon            =   "Accmaintenance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3315
   ScaleWidth      =   5790
   Begin VB.PictureBox picMaintenance 
      BorderStyle     =   0  'None
      Height          =   2205
      Left            =   120
      ScaleHeight     =   2205
      ScaleWidth      =   5925
      TabIndex        =   5
      Top             =   90
      Width           =   5925
      Begin VB.TextBox txtUserName 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   1
         Text            =   "txtUserName"
         Top             =   480
         Width           =   3645
      End
      Begin VB.TextBox txtPassWord 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   22
         PasswordChar    =   "l"
         TabIndex        =   2
         Top             =   900
         Width           =   3645
      End
      Begin VB.ComboBox cboLevel 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         ItemData        =   "Accmaintenance.frx":030A
         Left            =   1560
         List            =   "Accmaintenance.frx":030C
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "cboLevel"
         Top             =   1740
         Width           =   3675
      End
      Begin VB.TextBox txtConfirm 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   22
         PasswordChar    =   "l"
         TabIndex        =   3
         Top             =   1320
         Width           =   3645
      End
      Begin VB.TextBox txtUserCode 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "txtUserName"
         Top             =   60
         Width           =   585
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   4080
         TabIndex        =   6
         Text            =   "txtLevel"
         Top             =   1410
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   12
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   11
         Top             =   930
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   10
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   9
         Top             =   1350
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "User Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label labPassCode 
         BackColor       =   &H8000000D&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4830
         TabIndex        =   7
         Top             =   570
         Width           =   135
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4230
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   16
      Top             =   2295
      Width           =   1440
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   720
         MouseIcon       =   "Accmaintenance.frx":030E
         MousePointer    =   99  'Custom
         Picture         =   "Accmaintenance.frx":0460
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -30
         MouseIcon       =   "Accmaintenance.frx":079E
         MousePointer    =   99  'Custom
         Picture         =   "Accmaintenance.frx":08F0
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   90
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   13
      Top             =   2295
      Width           =   5580
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4860
         MouseIcon       =   "Accmaintenance.frx":0C40
         MousePointer    =   99  'Custom
         Picture         =   "Accmaintenance.frx":0D92
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4110
         MouseIcon       =   "Accmaintenance.frx":10F8
         MousePointer    =   99  'Custom
         Picture         =   "Accmaintenance.frx":124A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmAccMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPAccess                                As ADODB.Recordset
Dim AddorEdit                                As String

Sub FillCboLevel()
    cboLevel.Clear
    cboLevel.AddItem "GUEST"
    cboLevel.AddItem "ADMIN"
    cboLevel.AddItem "ACCOUNTANT"
    cboLevel.AddItem "BILLING"
    cboLevel.AddItem "CASHIER"
    cboLevel.AddItem "SUPERVISOR"
    cboLevel.AddItem "MANAGER"
End Sub



Private Sub cmdCANCEL_Click()
    picMaintenance.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    StoreMemVars
End Sub



Private Sub cmdEdit_Click()
    FillCboLevel
    cboLevel.Text = wizVar.DecryptAccess(Null2String(rsPAccess!LOGLEVEL))
    picMaintenance.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim findStr                              As String
    findStr = InputSpeechBox("Please enter Username to find...", txtUserName.Text)
    If findStr <> "" Then
        On Error GoTo ErrorCode
        rsPAccess.Bookmark = rsFind(rsPAccess.Clone, "username", findStr).Bookmark
    End If
    StoreMemVars
    Exit Sub

ErrorCode:
    If Err.Number = 3021 Then
        ShowCantFind findStr
        Resume Next
    End If
End Sub





Private Sub cmdPrint_Click()
    ShowAdviceMsg
End Sub

Private Sub cmdSave_Click()
    If txtUserCode.Text = "" Then
        ShowIsRequiredMsg "User Code"
        txtUserCode.SetFocus
        Exit Sub
    End If
    If txtUserName.Text = "" Then
        ShowIsRequiredMsg "User Name"
        txtUserName.SetFocus
        Exit Sub
    End If
    If AddorEdit = "ADD" Then
        Dim rsPAccessDup                     As ADODB.Recordset
        Set rsPAccessDup = New ADODB.Recordset
        Set rsPAccessDup = rsPAccess.Clone
        rsPAccessDup.Find "username = '" & txtUserName.Text & "'"
        If Not rsPAccessDup.EOF Then
            ShowAlreadyExistMsg "User Name"
            txtUserName.SetFocus
            Exit Sub
        End If
    End If
    If txtPassWord.Text = "" Then
        ShowIsRequiredMsg "Password"
        txtPassWord.SetFocus
        Exit Sub
    End If
    If txtConfirm.Text = "" Then
        ShowIsRequiredMsg "Confirm Password"
        txtConfirm.SetFocus
        Exit Sub
    End If
    If txtConfirm.Text <> txtPassWord.Text Then
        MsgSpeechBox "Passwords do not match!"
        txtPassWord.SetFocus
        Exit Sub
    End If
    If cboLevel.Text = "" Then
        ShowIsRequiredMsg "Level"
        cboLevel.SetFocus
        Exit Sub
    End If
    If Len(txtUserCode.Text) <> 3 Then
        MsgSpeechBox "Length of User Code must be 3"
        Exit Sub
    End If
    With wizVar
        If AddorEdit = "ADD" Then
            gconACCESS.Execute "Insert into PAccess " & _
                               "(usercode,username, userpass, logLevel) Values ('" & .EncryptAccess(txtUserCode.Text) & "','" & .EncryptAccess(txtUserName.Text) & "', '" & .EncryptAccess(txtPassWord.Text) & "', '" & .EncryptAccess(cboLevel.Text) & "')"
            ShowSuccessFullyAdded
            
        Else
            gconACCESS.Execute "Update PAccess set " & _
                               "usercode = '" & .EncryptAccess(txtUserCode.Text) & "', " & _
                               "username = '" & .EncryptAccess(txtUserName.Text) & "', " & _
                               "userpass = '" & .EncryptAccess(txtPassWord.Text) & "', " & _
                               "logLevel = '" & .EncryptAccess(cboLevel.Text) & "' "
                               
            gconDMIS.Execute ("Update ALL_RAMS_USERS SET PASSWORD='" & .EncryptAccess(txtPassWord) & "'")
            
            ShowSuccessFullyUpdated
        End If
      
        cmdCancel.Value = True
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    If OpenAccessDb = True Then
        Set rsPAccess = New ADODB.Recordset
        rsPAccess.Open "select * from PAccess WHERE USERNAME=" & N2Str2Null(wizVar.EncryptAccess(LOGNAME)), gconACCESS, adOpenKeyset, adLockReadOnly

        StoreMemVars
    Else
        ShowVBError
        MsgSpeechBox "I can't open a connection!!! Cryptofile or Datafile for Access is missing or Invalid " & vbCrLf & _
                     "Contact your friendly neighborhood SysAdministrator."
    End If
    '
    '
    DrawXPCtl Me
End Sub

Sub StoreMemVars()
    With wizVar
        If Not (rsPAccess.EOF Or rsPAccess.BOF) Then
            labPassCode.Caption = rsPAccess!ID
            txtUserCode.Text = .DecryptAccess(Null2String(rsPAccess!usercode))
            txtUserName.Text = .DecryptAccess(Null2String(rsPAccess!Username))
            txtPassWord.Text = .DecryptAccess(Null2String(rsPAccess!userPass))
            txtConfirm.Text = .DecryptAccess(Null2String(rsPAccess!userPass))
            txtLevel.Text = .DecryptAccess(Null2String(rsPAccess!LOGLEVEL))
            cboLevel.Text = .DecryptAccess(Null2String(rsPAccess!LOGLEVEL))
        Else
            AddorEdit = "ADD"
            picMaintenance.Enabled = True
            InitMemVars
            Picture1.Visible = False
            Picture2.Visible = True
        End If
    End With
End Sub

Sub InitMemVars()
    txtUserCode.Text = ""
    txtUserName.Text = ""
    txtPassWord.Text = ""
    txtConfirm.Text = ""
    FillCboLevel
    txtLevel.Text = ""
End Sub

Sub rsRefresh()
    Set rsPAccess = New ADODB.Recordset
    rsPAccess.Open "select * from paccess order by username", gconACCESS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

