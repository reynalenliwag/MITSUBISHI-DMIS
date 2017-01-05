VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Object = "{54AC2DF1-B6CB-406E-BB23-DC06DF6AAD9E}#16.0#0"; "WIZCRYPTO.OCX"
Begin VB.Form frmSecurity 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Security Window"
   ClientHeight    =   2025
   ClientLeft      =   2880
   ClientTop       =   2205
   ClientWidth     =   5250
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   Icon            =   "Security.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2025
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin wizCrypto.Crypto Crypto1 
      Height          =   465
      Left            =   210
      TabIndex        =   5
      Top             =   2280
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   820
   End
   Begin wizButton.cmd cmdOkey 
      Height          =   345
      Left            =   4200
      TabIndex        =   3
      Top             =   1050
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   609
      TX              =   "&Okey"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "Security.frx":0442
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1680
      MaxLength       =   22
      TabIndex        =   0
      Top             =   1080
      Width           =   1665
   End
   Begin VB.TextBox txtUserPass 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   22
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   1665
   End
   Begin VB.TextBox txtlevel 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   285
      Left            =   5850
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   150
   End
   Begin wizButton.cmd cmdCancel 
      Height          =   345
      Left            =   4200
      TabIndex        =   4
      Top             =   1440
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   609
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "Security.frx":045E
   End
   Begin VB.Image Image1 
      Height          =   2160
      Left            =   -90
      Picture         =   "Security.frx":047A
      Top             =   -60
      Width           =   5400
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPAccess As ADODB.Recordset
Dim Counter As Integer

Sub initMemvars()
txtUserName = "": txtUserPass = ""
End Sub

Sub StoreMemvars()
If rsPAccess.EOF() And rsPAccess.BOF() Then
   Dim mysql, LC, UN, PW, LV As String
   LC = "41444D^X^":   UN = "61646D££±í¶±§ù´ö°°í":   PW = "696365Zòò¢ö£®•©¨ó°®ñöúú":   LV = "41444D_]jUU"
   gconAccess.Execute "Insert into PAccess (usercode, username, userpass, loglevel) Values ('" & LC & "', '" & UN & "', '" & PW & "', '" & LV & "')"
End If
txtUserName.Text = "": txtUserPass.Text = "": txtlevel.Text = ""
End Sub

Private Sub cmdCancel_Click()
UnloadApp
End Sub

Private Sub cmdOkey_Click()
On Error Resume Next
With wizVar
     If CryptVar.CheckAccount(txtUserName.Text, txtUserPass.Text) = True Then
        LOGCODE = .DecryptAccess("57495AÅbÅ"):       LOGNAME = txtUserName.Text
        LOGLEVEL = .DecryptAccess("415554ctleze"):   LOGTIME = Time
        LOGDATE = Date
        CryptVar.Visible = True: SetCrypto
     Else
        Set rsPAccess = New ADODB.Recordset
            rsPAccess.Open "select * from PAccess where username = '" & .EncryptAccess(txtUserName.Text) & "'", gconAccess, adOpenKeyset
        If Not rsPAccess.EOF And Not rsPAccess.BOF Then
           If txtUserPass.Text <> .DecryptAccess(rsPAccess!userPass) Then
              MsgSpeechBox "Access Denied: Invalid Password!"
              GoTo Messages2
           Else
              LOGCODE = .DecryptAccess(rsPAccess!usercode):  LOGNAME = txtUserName.Text: LOGPASS = txtUserPass.Text
              LOGLEVEL = .DecryptAccess(rsPAccess!LOGLEVEL): LOGTIME = Time
              LOGDATE = Date
           End If
        Else
           MsgSpeechBox "Access Denied: User Name not found!"
           GoTo Messages
        End If
        MakeConnection
     End If
End With
Exit Sub
      
Messages:
If Counter < 3 Then
   Counter = Counter + 1
   txtUserName.Text = "":   txtUserPass.Text = ""
   txtUserName.SetFocus
Else
   Me.Visible = False
   MsgSpeechBox "Warning: Intruder Alert! ... Intruder Alert! ... Intruder Alert! ..."
   UnloadApp
End If
Exit Sub

Messages2:
If Counter < 3 Then
   Counter = Counter + 1
   txtUserPass.Text = "":   txtUserPass.SetFocus
Else
   Me.Visible = False
   MsgSpeechBox "Warning: Intruder Alert! ... Intruder Alert! ... Intruder Alert! ..."
   UnloadApp
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
       Case vbKeyReturn Or vbKeyTab
            If Me.ActiveControl.Name = "txtUserName" Then
               If txtUserName.Text <> "" Then SendKeys "{TAB}"
            End If
            If Me.ActiveControl.Name = "txtUserPass" Then
               If txtUserPass.Text <> "" And txtUserName.Text <> "" Then cmdOkey_Click
            End If
       Case vbKeyF4
            If txtUserName.Visible = False Then Unload Me
       Case Else
            MoveKeyPress KeyCode
End Select
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
Set wizVar = frmMain.wizEnc1
Set CryptVar = frmSecurity.Crypto1
Set oVoice = New SpVoice
Me.BorderStyle = 0
MakeTransparent Me.hWnd, 200
DoEvents
Counter = 0:
CryptVar.Visible = False
initMemvars
If OpenAccessDb = True Then
   Set rsPAccess = New ADODB.Recordset
       rsPAccess.Open "select * from PAccess", gconAccess, adOpenKeyset
   StoreMemvars
   CenterMe frmMain, Me, 1
Else
   MsgBoxXP "I can't open a connection!!! Cryptofile or Datafile for Access is missing or Invalid " & vbCrLf & _
            "Contact your friendly neighborhood SysAdministrator.", "ERROR", XP_OKOnly, msg_Critical
End If
Screen.MousePointer = 0
End Sub

Sub MakeConnection()
Screen.MousePointer = 11
If OpenSQLDb = False Then
   MsgSpeech "I can't open a connection!!! You may have to LOG-IN again to connect to the server to run this program. If you don't have an account contact your friendly neighborhood System Administreytor."
   MsgBoxXP "I can't open a connection!!! You may have to " & vbCrLf & _
            "LOG-IN again to connect to the server to run this program. " & vbCrLf & _
            "If you don't have an account contact your friendly " & vbCrLf & _
            "neighborhood SysAdministrator.", "ERROR", XP_OKOnly, msg_Critical
   UnloadApp
End If
frmSplash.Command1.Value = True
SetUserMenuSettings
frmMain.Show
'MainForm.Show
Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If frmSecurity.Visible = True Then
   If txtUserName.Visible = False And AccessCNT = 0 Then
      MakeConnection
   Else
      On Error Resume Next
      UnloadApp
   End If
End If
End Sub

Sub SetCrypto()
If AccessCNT = 0 Then
   MakeOpaque Me.hWnd
   CryptVar.Top = 0: CryptVar.Left = 0: CryptVar.SetCryptoSys
   Me.Width = CryptVar.Width: Me.Height = CryptVar.Height - 300
   txtUserName.Visible = False: txtUserPass.Visible = False
   cmdOkey.Visible = False: cmdCancel.Visible = False
   Me.BorderStyle = 1
   CenterMe frmMain, Me, 1
End If
End Sub
