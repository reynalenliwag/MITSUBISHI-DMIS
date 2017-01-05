VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_HitRatio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HIT RATIO REPORT"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3915
   Icon            =   "frmSMIS_Reports_HitRatio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3540
   ScaleWidth      =   3915
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
      TabIndex        =   9
      Top             =   810
      Width           =   3405
   End
   Begin VB.OptionButton optSAE 
      Height          =   375
      Left            =   150
      TabIndex        =   8
      Top             =   420
      Width           =   225
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
      Height          =   1245
      Left            =   60
      TabIndex        =   6
      Top             =   90
      Width           =   3765
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
         ItemData        =   "frmSMIS_Reports_HitRatio.frx":08CA
         Left            =   390
         List            =   "frmSMIS_Reports_HitRatio.frx":08D1
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   270
         Width           =   3225
      End
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
      ForeColor       =   &H00973640&
      Height          =   465
      ItemData        =   "frmSMIS_Reports_HitRatio.frx":08DA
      Left            =   2340
      List            =   "frmSMIS_Reports_HitRatio.frx":08DC
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1950
      Width           =   1485
   End
   Begin VB.ComboBox cboMonth 
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
      ForeColor       =   &H00973640&
      Height          =   465
      ItemData        =   "frmSMIS_Reports_HitRatio.frx":08DE
      Left            =   60
      List            =   "frmSMIS_Reports_HitRatio.frx":08E0
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1950
      Width           =   2205
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
      Left            =   900
      MouseIcon       =   "frmSMIS_Reports_HitRatio.frx":08E2
      MousePointer    =   99  'Custom
      Picture         =   "frmSMIS_Reports_HitRatio.frx":0A34
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   2580
      Width           =   885
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
      Left            =   1800
      MouseIcon       =   "frmSMIS_Reports_HitRatio.frx":0ED3
      MousePointer    =   99  'Custom
      Picture         =   "frmSMIS_Reports_HitRatio.frx":1025
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Window"
      Top             =   2580
      Width           =   885
   End
   Begin Crystal.CrystalReport rptHitRatio 
      Left            =   3540
      Top             =   3030
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
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
      Left            =   60
      TabIndex        =   3
      Top             =   1530
      Width           =   1665
   End
   Begin VB.Label Label1 
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
      Left            =   2280
      TabIndex        =   2
      Top             =   1500
      Width           =   825
   End
End
Attribute VB_Name = "frmSMIS_Report_HitRatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FilterSAE                                                         As String
Dim Filtered                                                          As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    FilterSAE = "({CRIS_Prospects.SAE})='" & cboSA.Text & "' And month({CRIS_Prospects.LogInitialInquiry})<=" & What_month(Me.cboMonth.Text) & ""
    Filtered = "(month({CRIS_Prospects.LogInitialInquiry})=" & What_month(Me.cboMonth.Text) & " And year({CRIS_Prospects.LogInitialInquiry})=" & cboYear.Text & ")"
    With rptHitRatio
        .Formulas(0) = "CompanyName='" & COMPANY_NAME & "'"
        .Formulas(1) = "CompanyAddress='" & COMPANY_ADDRESS & "'"
    End With

    If Me.optSAE.Value = True Then
        If Me.cboSA.Text = "" Then
            MsgBox "Please Select Sales Agent", vbInformation + vbOKOnly
            Exit Sub
        End If
        If Me.cboSA.Text = "ALL" Then
            PrintSQLReport rptHitRatio, SMIS_REPORT_PATH & "VS\HitRatio.rpt", "", DMIS_REPORT_Connection, 1
        Else
            PrintSQLReport rptHitRatio, SMIS_REPORT_PATH & "VS\HitRatio.rpt", FilterSAE, DMIS_REPORT_Connection, 1
        End If
    ElseIf Me.optSummaryReport.Value = True Then
        PrintSQLReport rptHitRatio, SMIS_REPORT_PATH & "VS\HitRatio.rpt", Filtered, DMIS_REPORT_Connection, 1
    End If
    
    'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
     If optSAE.Value = True Then
       Call NEW_LogAudit("V", "HIT RATIO", "", "", "", cboSA & " " & cboMonth & " " & cboYear, "", "")
     Else
      Call NEW_LogAudit("V", "HIT RATIO", "", "", "", "SUMMARY -" & " " & cboMonth & " " & cboYear, "", "")
     End If
    'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

    
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (HIT RATIO)"
            Call frmALL_AuditInquiry.DisplayHistory("", "HIT RATIO", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    fillcbomonth cboMonth
    fillcbomoreyear cboYear
    FillCombo "Select ID, Name from SMIS_vw_Srep Order By 2 ", 0, 1, cboSA
    cboSA.AddItem "ALL"
    cboMonth = MonthName(Month(LOGDATE))
    cboYear = Year(LOGDATE)
    CenterMe frmMain, Me, 1
    optSAE.Value = True
    Me.cboSA.ListIndex = 0
End Sub

Private Sub optSAE_Click()
    cboSA.Enabled = True
End Sub

Private Sub optSummaryReport_Click()
    cboSA.Enabled = False
End Sub

