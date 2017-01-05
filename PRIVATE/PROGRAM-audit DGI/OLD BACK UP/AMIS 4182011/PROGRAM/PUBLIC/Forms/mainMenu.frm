VERSION 5.00
Begin VB.Form frmRAMS_MainMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Browse Module"
      Height          =   555
      Left            =   240
      TabIndex        =   2
      Top             =   1590
      Width           =   1905
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Modules"
      Height          =   555
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   1905
   End
   Begin VB.CommandButton cmdUserMaintain 
      Caption         =   "User Maintainence"
      Height          =   555
      Left            =   240
      TabIndex        =   1
      Top             =   900
      Width           =   1905
   End
End
Attribute VB_Name = "frmRAMS_MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdUserMaintain_Click()
    frmusers.Show
End Sub

Private Sub Command1_Click()
frmRAMS_AEModule.Show

End Sub

Private Sub Command2_Click()
    frmRAMS_ModulesSheet.Show
End Sub

Sub Form_Load()
    CenterMe frmMain, Me, 1
End Sub
