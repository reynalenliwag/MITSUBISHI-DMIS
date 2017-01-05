VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCSMSAddPms_NEO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preventive Maintenance Service Schedule"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11580
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMSAddPms_NEO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   11580
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5325
      Left            =   60
      ScaleHeight     =   5295
      ScaleWidth      =   11385
      TabIndex        =   1
      Top             =   1860
      Width           =   11415
      Begin MSComctlLib.ListView ListView1 
         Height          =   5085
         Left            =   30
         TabIndex        =   2
         Top             =   60
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   8969
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   60
      ScaleHeight     =   1545
      ScaleWidth      =   11385
      TabIndex        =   0
      Top             =   180
      Width           =   11415
   End
End
Attribute VB_Name = "frmCSMSAddPms_NEO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub rsRefresh()

End Sub

Sub StoreMemVars()

End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
End Sub

