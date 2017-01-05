VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCRIS_GridCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Search"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "GridMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   9510
   Begin XtremeReportControl.ReportControl ReportControl2 
      Height          =   6195
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   9345
      _Version        =   655364
      _ExtentX        =   16484
      _ExtentY        =   10927
      _StockProps     =   64
      BorderStyle     =   2
      ShowFooter      =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3900
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmCRIS_GridCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
