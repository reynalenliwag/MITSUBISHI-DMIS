VERSION 5.00
Begin VB.Form frmPassMaintenance 
   Caption         =   "Change Password"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form9"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Okey"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtConfirm 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtNew 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtOld 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Confirm Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "New Password:"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Old Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmPassMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If txtOld.Text <> OldPass Then
   MsgBox "Old Password is invalid.", vbInformation
   txtOld.SetFocus
   Exit Sub
End If

If txtNew.Text = "" Then
   MsgBox "New Password must have a value.", vbInformation
   txtNew.SetFocus
   Exit Sub
End If

If txtConfirm.Text = "" Then
   MsgBox "Confirm Password must have a value.", vbInformation
   txtConfirm.SetFocus
   Exit Sub
End If

If txtConfirm.Text <> txtNew Then
   MsgBox "Password do not match.", vbInformation
   txtNew.SetFocus
   Exit Sub
End If

'Dim mysql As String
'mysql = "update HRMS_DivRef set word = '" & rewrite(txtNew.Text, True) & "' where divcode = '" & thedivcode & "'"
'gconDMIS.Execute mysql, dbFailOnError

MsgBox "Password changed.", vbInformation
OldPass = txtNew.Text
Command2.Value = True
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub Form_Load()
'CenterMe Me, Me, 0
Label1.Caption = frmLOGIN.Caption
txtOld.Text = ""
txtNew.Text = ""
txtConfirm.Text = ""
End Sub

