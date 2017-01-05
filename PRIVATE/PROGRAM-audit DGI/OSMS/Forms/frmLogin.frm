VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log-in"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   ForeColor       =   &H8000000F&
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3735
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   420
      Width           =   2265
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   1290
      MaxLength       =   10
      TabIndex        =   0
      Top             =   60
      Width           =   2265
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   2040
      TabIndex        =   3
      Top             =   810
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   405
      Left            =   600
      TabIndex        =   2
      Top             =   810
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   450
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   1455
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsUser As ADODB.Recordset

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()
    Set rsUser = New Recordset
        rsUser.Open "Select * from [user] where username = '" & txtUserName.Text & "'", gconOSMS
        If Not rsUser.EOF And Not rsUser.BOF Then
           If txtPassWord.Text = rsUser!UserPass Then
              Unload Me
           Else
              MsgBoxXP "Invalid Password!", "Invalid Password", XP_OKOnly, msg_Exclamation
              txtPassWord.SetFocus
           End If
        Else
            MsgBoxXP "Invalid Username!", "Invalid Username", XP_OKOnly, msg_Exclamation
            txtUserName.SetFocus
        End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
    CenterMe Me
End Sub
