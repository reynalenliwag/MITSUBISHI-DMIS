VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMSPrintYTDProcessing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Year to Date"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3135
   ForeColor       =   &H00D8E9EC&
   Icon            =   "PrintYTDProcessing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1605
   ScaleWidth      =   3135
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2160
      MouseIcon       =   "PrintYTDProcessing.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "PrintYTDProcessing.frx":0594
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   660
      Width           =   765
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1410
      MouseIcon       =   "PrintYTDProcessing.frx":09DF
      MousePointer    =   99  'Custom
      Picture         =   "PrintYTDProcessing.frx":0B31
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   660
      Width           =   765
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   1965
   End
   Begin Crystal.CrystalReport rptPrintYTD 
      Left            =   2085
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   825
   End
End
Attribute VB_Name = "frmHRMSPrintYTDProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200711:48
Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:

    'If Function_Access(LOGID, "Acess_Print", "YEAR TO DATE REPORT") = False Then Exit Sub
    Screen.MousePointer = 11
    rptPrintYTD.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptPrintYTD.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptPrintYTD.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    rptPrintYTD.Formulas(3) = "PREPARED_BY = '" & PREPARED_BY & "'"
    rptPrintYTD.Formulas(4) = "CHECKED_BY = '" & CHECKED_BY & "'"
    rptPrintYTD.Formulas(5) = "APPROVED_BY = '" & APPROVED_BY & "'"
    rptPrintYTD.Formulas(6) = "PrintedBy = '" & LOGNAME & "'"

    PrintSQLReport rptPrintYTD, HRMS_REPORT_PATH & "ytddetails.rpt", "val({ytdDetails.yeer}) = " & cboyear.Text, DMIS_REPORT_Connection, 1
    LogAudit "G", "YEAR TO DATE", cboyear
    Screen.MousePointer = 0

    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    DrawXPCtl Me
    'FillcboYear cboyear
    fillcombo_up cboyear
    cboyear.Text = YEAR(LOGDATE)
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

