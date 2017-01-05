VERSION 5.00
Begin VB.Form frmHRMSAbout 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "About the Author"
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6045
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H8000000D&
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "About.frx":030A
   ScaleHeight     =   3870
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmHRMSAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Unload Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
End Sub
