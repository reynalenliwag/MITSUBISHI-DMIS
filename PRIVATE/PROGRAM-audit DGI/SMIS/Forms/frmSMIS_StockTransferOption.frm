VERSION 5.00
Begin VB.Form frmSMIS_StockTransferOption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK TRANSFER OPTION"
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4740
   Icon            =   "frmSMIS_StockTransferOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   4740
   Begin VB.CommandButton cmdTransferOut 
      Caption         =   "STOCK TRANSFER OUT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   4755
   End
   Begin VB.CommandButton CmdTransferIn 
      Caption         =   "STOCK TRANSFER IN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4755
   End
End
Attribute VB_Name = "frmSMIS_StockTransferOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdTransferIN_Click()
    If Module_Access(LOGID, "STOCK TRANSFER IN", "TRANSACTION") = False Then Exit Sub
    frmSMIS_Trans_MRR2.Show
End Sub

Private Sub cmdTransferOut_Click()
    If Module_Access(LOGID, "STOCK TRANSFER OUT", "TRANSACTION") = False Then Exit Sub
    frmSMIS_Trans_MRR1.Show
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Screen.MousePointer = 0
End Sub
