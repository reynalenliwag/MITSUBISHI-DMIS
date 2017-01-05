VERSION 5.00
Begin VB.Form frmOSMSWindowAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About OSMS"
   ClientHeight    =   4200
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5820
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":030A
   ScaleHeight     =   2898.915
   ScaleMode       =   0  'User
   ScaleWidth      =   5465.281
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   180
      Picture         =   "frmAbout.frx":53B3E
      Top             =   1560
      Width           =   1500
   End
End
Attribute VB_Name = "frmOSMSWindowAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
End Sub
