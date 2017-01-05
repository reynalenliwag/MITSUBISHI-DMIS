VERSION 5.00
Begin VB.Form frmAISTMPFILES 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TEMPORARY FILES"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   7065
End
Attribute VB_Name = "frmAISTMPFILES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
End Sub

