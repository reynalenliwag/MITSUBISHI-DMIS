VERSION 5.00
Begin VB.Form frmMaintenance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maintenance"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2715
   Icon            =   "Maintenance.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Back"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Employees"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Password"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmSecurity.Show vbModal
End Sub

Private Sub Command2_Click()
frmEmployees.Show vbModal
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterMe Me, Me, 0
End Sub
