VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form warrantyrep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Warranty Report"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CSMS_WARANTYREP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1995
   ScaleWidth      =   4635
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
      Height          =   795
      Left            =   2310
      MouseIcon       =   "CSMS_WARANTYREP.frx":27A2
      MousePointer    =   99  'Custom
      Picture         =   "CSMS_WARANTYREP.frx":28F4
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Close Window"
      Top             =   1140
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1290
      TabIndex        =   1
      Top             =   780
      Width           =   225
   End
   Begin MSComCtl2.DTPicker txtfrom 
      Height          =   345
      Left            =   90
      TabIndex        =   0
      Top             =   330
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   609
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   39546
   End
   Begin Crystal.CrystalReport warrantyrep 
      Left            =   180
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker txtto 
      Height          =   345
      Left            =   2430
      TabIndex        =   4
      Top             =   330
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   609
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   39546
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
      Height          =   795
      Left            =   1590
      MouseIcon       =   "CSMS_WARANTYREP.frx":2D3F
      MousePointer    =   99  'Custom
      Picture         =   "CSMS_WARANTYREP.frx":2E91
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   60
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2430
      TabIndex        =   6
      Top             =   60
      Width           =   555
   End
   Begin VB.Label Label3 
      Caption         =   "Summary Only"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1620
      TabIndex        =   5
      Top             =   810
      Width           =   1665
   End
End
Attribute VB_Name = "warrantyrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error Resume Next
    If Function_Access(LOGID, "Acess_PRINT", "SERVICE REPORT") = False Then Exit Sub

    Screen.MousePointer = 11

    warrantyrep.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    warrantyrep.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    warrantyrep.Formulas(2) = "printedby = '" & LOGNAME & "'"

    If Check1.Value = 1 Then
        warrantyrep.WindowTitle = "Warranty Summary Report"
        warrantyrep.ReportTitle = "Warranty Summary Report"
        PrintSQLReport warrantyrep, CSMS_REPORT_PATH & "Warrantysumrepor.rpt", "{REPOR.DTE_COMP} >= DATESERIAL(" & Year(txtFROM.Value) & "," & Month(txtFROM.Value) & "," & Day(txtFROM.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1
        LogAudit "G", "WARRANTY SUMMARY - REPORT"
    Else
        warrantyrep.WindowTitle = "Warranty Report"
        warrantyrep.ReportTitle = "Warranty Report Detailed"
        PrintSQLReport warrantyrep, CSMS_REPORT_PATH & "Warrantyrepor.rpt", "{REPOR.DTE_COMP} >= DATESERIAL(" & Year(txtFROM.Value) & "," & Month(txtFROM.Value) & "," & Day(txtFROM.Value) & ") AND {REPOR.DTE_COMP} <= DATE(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1
        LogAudit "V", "WARRANTY - REPORT"
    End If

    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtFROM.Value = firstDay(LOGDATE)
    txtTO.Value = LOGDATE
End Sub

