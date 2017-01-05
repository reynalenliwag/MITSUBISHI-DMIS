VERSION 5.00
Object = "{54AC2DF1-B6CB-406E-BB23-DC06DF6AAD9E}#12.0#0"; "wizCrypto.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{7E0552E8-E2C9-4C9E-BC75-F34D871C75F4}#2.0#0"; "wizEncrypt.ocx"
Begin VB.Form frmSecurity 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Security Window"
   ClientHeight    =   2055
   ClientLeft      =   2880
   ClientTop       =   2205
   ClientWidth     =   5295
   FillColor       =   &H8000000D&
   ForeColor       =   &H00000000&
   Icon            =   "Security.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2055
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin wizButton.cmd cmdOkey 
      Height          =   345
      Left            =   4110
      TabIndex        =   4
      Top             =   1020
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
   Begin wizCrypto.Crypto Crypto1 
      Height          =   465
      Left            =   1170
      TabIndex        =   3
      Top             =   2670
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   820
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
      Top             =   1110
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
      Top             =   1470
      Width           =   1665
   End
   Begin VB.TextBox txtlevel 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   285
      Left            =   5880
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   150
   End
   Begin wizButton.cmd cmdCancel 
      Height          =   345
      Left            =   4110
      TabIndex        =   5
      Top             =   1380
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   -60
      Picture         =   "Security.frx":047A
      ScaleHeight     =   2085
      ScaleWidth      =   5355
      TabIndex        =   6
      Top             =   -30
      Width           =   5355
      Begin wizEncrypt.wizEnc wizEnc1 
         Left            =   1410
         Top             =   2310
         _ExtentX        =   3969
         _ExtentY        =   3969
      End
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
   LC = "41444D^X^":   UN = "61646D££±í¶±§ù´ö°°í":   PW = "696365û¶ß®õôû•ô•¢òûñû":   LV = "41444D_]jUU"
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
        Me.BorderStyle = 1
        CryptVar.Visible = True: SetCrypto
     Else
        Set rsPAccess = New ADODB.Recordset
            rsPAccess.Open "select * from PAccess where username = '" & .EncryptAccess(txtUserName.Text) & "'", gconAccess, adOpenForwardOnly, adLockReadOnly
        If Not rsPAccess.EOF And Not rsPAccess.BOF Then
           If txtUserPass.Text <> .DecryptAccess(rsPAccess!userPass) Then
              MsgSpeechBox "Access Denied: Invalid Password!"
              GoTo Messages2
           Else
              LOGCODE = .DecryptAccess(rsPAccess!usercode):  LOGNAME = txtUserName.Text
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
Set wizVar = frmSecurity.wizEnc1
Set CryptVar = frmSecurity.Crypto1
Set oVoice = New SpVoice
MakeTransparent Me.hwnd, 200
DoEvents
Counter = 0:
CryptVar.Visible = False
initMemvars
If OpenAccessDb = True Then
   Set rsPAccess = New ADODB.Recordset
       rsPAccess.Open "select * from PAccess", gconAccess, adOpenForwardOnly, adLockReadOnly
   StoreMemvars
End If
DrawXPCtl Me
Screen.MousePointer = 0
End Sub

Sub MakeConnection()
Screen.MousePointer = 11
On Error Resume Next
deAccess.deConnAccess.Close
If OpenSQLDb = True Then
   gconLOGIN.Execute "insert into USERS " & _
                     "(logcode,logname,logintime,logdate) " & _
                     " values (" & N2Str2Null(LOGCODE) & ", " & N2Str2Null(LOGNAME) & ", " & N2Str2Null(LOGTIME) & ", " & N2Str2Null(LOGDATE) & ")"
Else
   MsgSpeechBox "I can't open a connection!!! You may have to LOG-IN again to connect to the server to run this program. If you don't have an account contact your friendly neighborhood SysAdministrator."
   MsgBoxXP "I can't open a connection!!! You may have to " & vbCrLf & _
            "LOG-IN again to connect to the server to run this program. " & vbCrLf & _
            "If you don't have an account contact your friendly " & vbCrLf & _
            "neighborhood SysAdministrator.", "ERROR", XP_OKOnly, msg_Critical
   UnloadApp
End If
Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If frmSecurity.Visible = True Then
   If txtUserName.Visible = False And AccessCNT = 0 Then
      MakeConnection
   Else
      On Error Resume Next
      End
   End If
End If
UnloadForm Me
End Sub

Sub SetCrypto()
If AccessCNT = 0 Then
   MakeOpaque Me.hwnd
   CryptVar.Top = 0: CryptVar.Left = 0: CryptVar.SetCryptoSys
   Me.Width = CryptVar.Width: Me.Height = CryptVar.Height - 300
   txtUserName.Visible = False: txtUserPass.Visible = False
   cmdOkey.Visible = False: CmdCancel.Visible = False
   Me.BorderStyle = 1
End If
End Sub
