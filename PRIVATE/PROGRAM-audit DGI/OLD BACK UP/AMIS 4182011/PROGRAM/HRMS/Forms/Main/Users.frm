VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmusers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " User Administration"
   ClientHeight    =   3870
   ClientLeft      =   1560
   ClientTop       =   720
   ClientWidth     =   8895
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Users.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   8895
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1545
      Left            =   0
      ScaleHeight     =   1545
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   0
      Width           =   4395
      Begin VB.TextBox txtpass2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1485
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   750
         Width           =   2745
      End
      Begin VB.TextBox txtpass1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1485
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   405
         Width           =   2745
      End
      Begin VB.TextBox txtUser 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1485
         MaxLength       =   20
         TabIndex        =   5
         Top             =   60
         Width           =   2745
      End
      Begin VB.ComboBox cboGroups 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1485
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   1125
         Width           =   2805
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   540
         TabIndex        =   11
         Top             =   1140
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   30
         TabIndex        =   10
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   630
         TabIndex        =   9
         Top             =   495
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   645
         TabIndex        =   8
         Top             =   150
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   4440
      ScaleHeight     =   3195
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   0
      Width           =   4395
      Begin MSComctlLib.ListView lvwUsers 
         Height          =   2865
         Left            =   30
         TabIndex        =   1
         Top             =   270
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   5054
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Username"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "UserGroup"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "  "
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Users List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   90
         TabIndex        =   2
         Top             =   30
         Width           =   2475
      End
   End
   Begin MSForms.CommandButton cmdEdit 
      Height          =   450
      Left            =   4500
      TabIndex        =   18
      Top             =   2625
      Visible         =   0   'False
      Width           =   825
      Caption         =   "Edit"
      PicturePosition =   327683
      Size            =   "1455;794"
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdAddMod 
      Height          =   495
      Left            =   2250
      TabIndex        =   17
      Top             =   3240
      Width           =   1560
      Caption         =   "Add Modules"
      PicturePosition =   327683
      Size            =   "2752;873"
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdNew 
      Height          =   495
      Left            =   615
      TabIndex        =   16
      Top             =   3240
      Width           =   1560
      Caption         =   "New User"
      PicturePosition =   327683
      Size            =   "2752;873"
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdSave 
      Height          =   495
      Left            =   3870
      TabIndex        =   15
      Top             =   3240
      Width           =   1560
      VariousPropertyBits=   25
      Caption         =   "Save"
      PicturePosition =   327683
      Size            =   "2752;873"
      FontName        =   "Arial"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdDelete 
      Height          =   495
      Left            =   7140
      TabIndex        =   14
      Top             =   3240
      Width           =   1560
      Caption         =   "Delete"
      PicturePosition =   327683
      Size            =   "2752;873"
      Picture         =   "Users.frx":08CA
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   495
      Left            =   5505
      TabIndex        =   13
      Top             =   3240
      Width           =   1560
      Caption         =   "Refresh"
      PicturePosition =   327683
      Size            =   "2752;873"
      Picture         =   "Users.frx":0C05
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lblDisplayUser 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1500
      TabIndex        =   12
      Top             =   1620
      Width           =   2775
   End
End
Attribute VB_Name = "frmusers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs                                       As ADODB.Recordset
Dim CurID                                    As Long
Dim isNew                                    As Boolean

Private Sub DisplayUsers()
    Set rs = gconDMIS.Execute("Select username, usergroup, userid from HRMS_USERS order by username")
    If Not (rs.BOF And rs.EOF) Then
        Listview_Loadval lvwUsers.ListItems, rs
        lvwUsers.Enabled = True

    Else
        lvwUsers.Enabled = False
        lvwUsers.ListItems.Clear
    End If
    Set rs = Nothing
End Sub

Private Sub cboGroups_Change()
    cboGroups_Click
End Sub


Private Sub cboGroups_Click()
    On Error GoTo errorlbl
    lblDisplayUser.Caption = gconDMIS.Execute("Select groupname from HRMS_USERGROUPS Where Code='" & ReplaceQuote(cboGroups.Text) & "'").Collect(0)
    Exit Sub
errorlbl:
    Err.Clear
End Sub

Private Sub cmdAddMod_Click()
    If lvwUsers.SelectedItem Is Nothing Then
        MessagePop RecNotFound, "No Record", "There are No Record", 1000
        Exit Sub
    End If
    frmUserModules.UserID = Me.lvwUsers.SelectedItem.SubItems(2)
    frmUserModules.Username = Me.lvwUsers.SelectedItem
    frmUserModules.Show
End Sub

Private Sub cmdDELETE_Click()
    If Me.lvwUsers.ListItems.Count = 0 Then Exit Sub
    If MsgBox("Are you sure you want to remove user " & Me.lvwUsers.SelectedItem & "?", vbExclamation + vbYesNo, "Remove User") = vbYes Then
        If Me.lvwUsers.ListItems.Count = 1 Then
            MsgBox "Sorry, can't remove selected user.", vbCritical, "Access denied!"
        Else
            gconDMIS.Execute ("Delete from HRMS_USERS where username = '" & Me.lvwUsers.SelectedItem & "' ")
            gconACCESS.Execute ("Delete from PAccess where username = '" & wizVar.DecryptAccess(Me.lvwUsers.SelectedItem) & "' ")
            Call Combo_Loadval(cboGroups, gconDMIS.Execute("Select code,groupname from HRMS_USERGROUPS"))
            'Call populateCbo("Select code, groupname from HRMS_USERGROUPS", gconDMIS, Me.cboGroups)
            Call DisplayUsers
            ResetEntry
        End If
    End If
End Sub

Private Sub cmdNew_Click()
    isNew = True
    Me.txtUser.SetFocus
    Me.cmdSAVE.Enabled = True
    Me.cmdEdit.Enabled = False
    Me.cmdDELETE.Enabled = False
    Me.cmdAddMod.Enabled = False
    Me.cmdNEW.Enabled = False
    Me.txtUser = vbNullString
    Me.txtpass1 = vbNullString
    Me.txtpass2 = vbNullString

End Sub

Private Sub cmdRefresh_Click()
    Call ResetEntry
End Sub

Private Sub cmdSAVE_Click()
    Call SaveUser
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        KeyCode = 0
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 0
    'Call populateCbo("Select code, groupname from HRMS_USERGROUPS", gconDMIS, cboGroups)

    Call Combo_Loadval(cboGroups, gconDMIS.Execute("Select code,groupname from HRMS_USERGROUPS"))
    Call DisplayUsers
End Sub

Private Sub ResetEntry()

    Me.txtUser = vbNullString
    Me.txtpass1 = vbNullString
    Me.txtpass2 = vbNullString
    Me.cboGroups.ListIndex = -1
    Me.cmdNEW.Enabled = True
    Me.cmdNEW.Caption = "New User"
    Me.cmdNEW.SetFocus
    Me.cmdSAVE.Enabled = False
    Me.cmdAddMod.Enabled = True
    Me.cmdEdit.Enabled = True
    Me.cmdDELETE.Enabled = True
End Sub

Private Sub SaveUser()
    'On Error GoTo errlab

    Dim pass                                 As String

    If Trim(Me.txtUser) = vbNullString Then
        MsgBox "Username is required!", vbExclamation, "Warning"
        Me.txtUser.SetFocus
        Exit Sub
    ElseIf Trim(Me.txtpass1) = vbNullString Then
        MsgBox "Password is required!", vbExclamation, "Warning"
        Me.txtpass1.SetFocus
        Exit Sub
    ElseIf Trim(Me.txtpass2) = vbNullString Then
        MsgBox "Please verify your password!", vbExclamation, "Warning"
        Me.txtpass2.SetFocus
        Exit Sub
    ElseIf Trim(Me.txtpass1) <> Trim(Me.txtpass2) Then
        MsgBox "Please check your password!", vbExclamation, "Warning"
        Me.txtpass2.SetFocus
        Exit Sub
    ElseIf Trim(Me.cboGroups.Text) = vbNullString Then
        MsgBox "Please select USer group!", vbExclamation, "Warning"
        Me.cboGroups.SetFocus
        Exit Sub
    End If
    '=== users table ====
    pass = wizVar.EncryptAccess(Trim(Me.txtpass1))
    If isNew Then
        If Username_Exists(Trim(Me.txtUser)) Then
            Me.txtUser.SetFocus
            Exit Sub
        End If
        gconDMIS.Execute "Insert into HRMS_USERS ([username],[password],usergroup)" _
                       & " values ('" & Trim(Me.txtUser) & "'," & N2Str2Null(pass) & ",'" & Me.cboGroups.Text & "')"
        gconACCESS.Execute "Insert into PAccess (usercode,[username],[userpass],LOGLEVEL)" _
                         & " values ('" & wizVar.EncryptAccess(Left(Trim(Me.txtUser), 3)) & "','" & wizVar.EncryptAccess(Trim(Me.txtUser)) & "'," & N2Str2Null(pass) & ",'" & wizVar.EncryptAccess(Me.cboGroups.Text) & "')"
        MsgBox "User " & Trim(Me.txtUser) & " successfully added!", vbInformation, "New User"
    Else
        gconDMIS.Execute ("update HRMS_USERS set [username] = " & N2Str2Null(Trim(Me.txtUser)) & ", [password] = " & _
                          N2Str2Null(wizVar.EncryptAccess(Trim(Me.txtpass1))) & ",usergroup = " & N2Str2Null(Me.cboGroups.Text) & " where userid = " & CurID)
        gconACCESS.Execute ("Update PAccess set [username] = " & N2Str2Null(wizVar.EncryptAccess(Trim(Me.txtUser))) & ", [userpass] = " & _
                            N2Str2Null(wizVar.EncryptAccess(Trim(Me.txtpass1))) & ",LOGLEVEL = " & N2Str2Null(wizVar.EncryptAccess(Me.cboGroups.Text)) & " where username = '" & wizVar.DecryptAccess(Me.txtUser) & "'")
        isNew = True
    End If
    Call ResetEntry
    Call Combo_Loadval(cboGroups, gconDMIS.Execute("Select code,groupname from HRMS_USERGROUPS"))
    'Call populateCbo("Select code, groupname from HRMS_USERGROUPS", gconDMIS, Me.cboGroups)
    Call DisplayUsers

    Exit Sub
    'errlab:
    '  MsgBox "Theres a problem encountered... pls try again.", vbInformation, "Remarks"
End Sub

Private Sub Display_Edit(UserID As Long)
    Set rs = gconDMIS.Execute("Select * from HRMS_USERS where userid = " & UserID & "")
    If Not (rs.BOF And rs.EOF) Then
        Me.txtUser = rs!Username
        Me.txtpass1 = wizVar.DecryptAccess(rs!Password)
        Me.txtpass2 = Me.txtpass1
        CurID = Me.lvwUsers.SelectedItem.SubItems(2)
        'Me.cboGroups.Value = rs!usergroup
        Me.cboGroups.Text = rs!usergroup
        cmdSAVE.Enabled = True
        isNew = False
    End If
    Set rs = Nothing
End Sub

Private Function Username_Exists(Username As String) As Boolean
    Set rs = gconDMIS.Execute("Select username from HRMS_USERS where username = '" & Username & "' ")
    If Not (rs.BOF And rs.EOF) Then
        MsgBox "Username " & Username & "  already exists!, please try another.", vbCritical, "Alert"
        Username_Exists = True

        Set rs = Nothing
        Exit Function
    End If
    Set rs = Nothing
    Username_Exists = False
End Function

Private Sub lvwUsers_DblClick()
    Call Display_Edit(Me.lvwUsers.SelectedItem.SubItems(2))
End Sub

Private Sub txtpass1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtpass2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


