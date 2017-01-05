VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_LTOStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LTO STATUS"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Report_ltostatus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1710
   ScaleWidth      =   5010
   Begin VB.ComboBox cboLTO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Report_ltostatus.frx":030A
      Left            =   600
      List            =   "Report_ltostatus.frx":0323
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   3795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Vehicles Group Listing"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   90
      MouseIcon       =   "Report_ltostatus.frx":03D1
      MousePointer    =   99  'Custom
      Picture         =   "Report_ltostatus.frx":0523
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Customer Directory by Customer Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   60
      MouseIcon       =   "Report_ltostatus.frx":09C2
      MousePointer    =   99  'Custom
      Picture         =   "Report_ltostatus.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1740
      Width           =   4575
   End
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
      Left            =   2475
      MouseIcon       =   "Report_ltostatus.frx":0FB3
      MousePointer    =   99  'Custom
      Picture         =   "Report_ltostatus.frx":1105
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   780
      Width           =   885
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
      Left            =   1605
      MouseIcon       =   "Report_ltostatus.frx":1550
      MousePointer    =   99  'Custom
      Picture         =   "Report_ltostatus.frx":16A2
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   780
      Width           =   885
   End
   Begin Crystal.CrystalReport rptReleased 
      Left            =   4050
      Top             =   1050
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "MMPC Monthly Inventory Control"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "SELECT LTO STATUS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   540
      TabIndex        =   5
      Top             =   60
      Width           =   2835
   End
End
Attribute VB_Name = "frmSMIS_Report_LTOStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    Dim FILTER                                                        As String
    On Error GoTo ErrorCode:


    Screen.MousePointer = 11

    rptReleased.WindowTitle = "LTO STATUS REPORT"
    rptReleased.Reset
    rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    LogAudit "V", "LTO STATUS REPORT", "FOR " & cboLTO
    If cboLTO.Text = "ALL" Then
        PrintSQLReport rptReleased, SMIS_REPORT_PATH & "LTOSTAUS.rpt", "", DMIS_REPORT_Connection, 1
        'UPDATED BY: JUN
        'DATE UPDATED: 09032008 5:00
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "LTO STATUS", "", "", "", cboLTO, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Else
        PrintSQLReport rptReleased, SMIS_REPORT_PATH & "LTOSTAUS.rpt", "{VEHICLE.LTOSTATUS}=" & N2Str2Null(cboLTO), DMIS_REPORT_Connection, 1
        'UPDATED BY: JUN
        'DATE UPDATED: 09032008 5:00
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "LTO STATUS", "", "", "", cboLTO, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

    End If

    Screen.MousePointer = 0


    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
             
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (LTO STATUS)"
            Call frmALL_AuditInquiry.DisplayHistory("", "LTO STATUS", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
End Sub

