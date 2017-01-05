VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSMIS_Report_LeadSource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LeadSource Report"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3360
      Top             =   1950
   End
   Begin Crystal.CrystalReport rptLeadSource 
      Left            =   2880
      Top             =   1650
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   1950
      MouseIcon       =   "frmSMIS_Report_LeadSource.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmSMIS_Report_LeadSource.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   1530
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
      Left            =   1050
      MouseIcon       =   "frmSMIS_Report_LeadSource.frx":059D
      MousePointer    =   99  'Custom
      Picture         =   "frmSMIS_Report_LeadSource.frx":06EF
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   1530
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "Option Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   3765
      Begin VB.OptionButton optSAE 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   270
         Value           =   -1  'True
         Width           =   225
      End
      Begin VB.OptionButton optSummaryReport 
         Caption         =   "Summary Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   150
         TabIndex        =   4
         Top             =   900
         Width           =   3405
      End
      Begin VB.ComboBox cboSA 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frmSMIS_Report_LeadSource.frx":0B8E
         Left            =   390
         List            =   "frmSMIS_Report_LeadSource.frx":0B95
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   3225
      End
   End
   Begin Crystal.CrystalReport rptHitRatio 
      Left            =   3480
      Top             =   2940
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Monthly Sales Executive Listing"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   765
      Left            =   150
      TabIndex        =   6
      Top             =   1530
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   1349
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmSMIS_Report_LeadSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fltr                                                              As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    'fltr = "(Lcase({Cris_Prospects.SAE})='" & LCase(cboSA.Text) & "')"
    fltr = "(Lcase({Cris_Prospects.SAE})<= 'C')"

    With rptLeadSource
        .Formulas(0) = "CompanyName='" & COMPANY_NAME & "'"
        .Formulas(1) = "CompanyAddress='" & COMPANY_ADDRESS & "'"
    End With

    If Me.optSAE.Value = True Then
        If Me.cboSA.Text = "" Then
            MsgBox "Please Select Sales Agent", vbOKOnly + vbInformation
            Exit Sub
        ElseIf Me.cboSA.Text = "ALL" Then
            PrintSQLReport rptLeadSource, SMIS_REPORT_PATH & "\VS\LeadSource.rpt", "", DMIS_Connection, 1
        Else
            PrintSQLReport rptLeadSource, SMIS_REPORT_PATH & "\VS\LeadSource.rpt", fltr, DMIS_Connection, 1
        End If

    ElseIf Me.optSummaryReport.Value = True Then
        PrintSQLReport rptLeadSource, SMIS_REPORT_PATH & "\VS\LeadSource.rpt", "", DMIS_Connection, 1
    End If

End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    FillCombo "Select ID, Name from SMIS_vw_Srep Order By 2 ", 0, 1, cboSA
    With Me.cboSA
        .AddItem "ALL"
        .ListIndex = 0
    End With

End Sub

Private Sub optSAE_Click()
    Me.cboSA.Enabled = True
End Sub

Private Sub optSummaryReport_Click()
    Me.cboSA.Enabled = False
End Sub

Private Sub Timer1_Timer()
    Dim ctr                                                           As Integer
    ctr = ctr + 1


    Me.ProgressBar1.Value = ctr

    If ctr = 100 Then
        Exit Sub
        Me.Timer1.Enabled = False
    End If

End Sub

