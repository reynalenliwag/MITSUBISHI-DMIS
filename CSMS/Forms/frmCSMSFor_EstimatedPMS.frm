VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMS_EstimatedPMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "For Follow Up (PMS)"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4035
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMSFor_EstimatedPMS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4035
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   825
      Left            =   3150
      MouseIcon       =   "frmCSMSFor_EstimatedPMS.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSFor_EstimatedPMS.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   660
      Width           =   795
   End
   Begin MSComCtl2.DTPicker txtDate 
      Height          =   375
      Left            =   1590
      TabIndex        =   0
      Top             =   180
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   107020289
      CurrentDate     =   39622
   End
   Begin Crystal.CrystalReport rptFOLLOW 
      Left            =   60
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Technician Efficiency Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   825
      Left            =   2370
      MouseIcon       =   "frmCSMSFor_EstimatedPMS.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSFor_EstimatedPMS.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   660
      Width           =   795
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Estimated to Return"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   30
      TabIndex        =   3
      Top             =   180
      Width           =   1500
   End
End
Attribute VB_Name = "frmCSMS_EstimatedPMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlApp                                              As Excel.Application
Dim xlBook                                             As Excel.Workbook
Dim xlSheet                                            As Excel.Worksheet

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "FOR FOLLOW UP.xls")
    Set xlSheet = xlBook.Worksheets(1)

    xlSheet.Cells(4, "A") = COMPANY_NAME
    xlSheet.Cells(5, "A") = COMPANY_ADDRESS
    xlSheet.Cells(34, "F") = GENERAL_MANAGER



    xlApp.Visible = True
    Set xlApp = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
End Sub

