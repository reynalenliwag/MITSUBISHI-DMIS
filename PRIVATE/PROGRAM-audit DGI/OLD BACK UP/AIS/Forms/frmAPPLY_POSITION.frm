VERSION 5.00
Begin VB.Form frmAISAPPLY_POSITION 
   Caption         =   "Form1"
   ClientHeight    =   720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1905
   LinkTopic       =   "Form1"
   ScaleHeight     =   720
   ScaleWidth      =   1905
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmAISAPPLY_POSITION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
End Sub
