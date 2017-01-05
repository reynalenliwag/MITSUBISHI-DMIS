VERSION 5.00
Begin VB.Form frmPMISPQIRForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PQIR Form - Data Input"
   ClientHeight    =   14670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11625
   Icon            =   "PQIRForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   14670
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   15000
      Left            =   -300
      Picture         =   "PQIRForm.frx":08CA
      ScaleHeight     =   15000
      ScaleWidth      =   12225
      TabIndex        =   0
      Top             =   -330
      Width           =   12225
   End
End
Attribute VB_Name = "frmPMISPQIRForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
CenterMe Me, Me, 0
End Sub
