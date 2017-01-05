VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMSPHMonthly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PhilHealth Monthly Remittance"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4860
   ForeColor       =   &H00D8E9EC&
   Icon            =   "PHMonthly.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1725
   ScaleWidth      =   4860
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
      Left            =   2430
      MouseIcon       =   "PHMonthly.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "PHMonthly.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   810
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
      Left            =   1560
      MouseIcon       =   "PHMonthly.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "PHMonthly.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   810
      Width           =   885
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
      Left            =   3420
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   135
      Width           =   1365
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
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   2355
   End
   Begin Crystal.CrystalReport rptPHMonthly 
      Left            =   3540
      Top             =   1065
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
      Left            =   2550
      TabIndex        =   2
      Top             =   165
      Width           =   825
   End
End
Attribute VB_Name = "frmHRMSPHMonthly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    rptPHMonthly.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptPHMonthly.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptPHMonthly.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    rptPHMonthly.Formulas(3) = "PREPARED_BY = '" & PREPARED_BY & "'"
    rptPHMonthly.Formulas(4) = "CHECKED_BY = '" & CHECKED_BY & "'"
    rptPHMonthly.Formulas(5) = "APPROVED_BY = '" & APPROVED_BY & "'"
    rptPHMonthly.Formulas(9) = "PRINTEDBY = '" & LOGNAME & "'"
    
    Select Case Me.Caption
        Case "PHILHEALTH MONTHLY REMITTANCE":
            'If Function_Access(LOGID, "Acess_Print", "REPORT PHILHEALTH MONTHLY REMITTANCE") = False Then Exit Sub
            PrintSQLReport rptPHMonthly, HRMS_REPORT_PATH & "Phremit.rpt", "{HRMS_PAYROLL.PAY_YEAR} = " & cboyear.Text & " and {HRMS_PAYROLL.Pay_month} = " & What_month(cboMOnth), DMIS_REPORT_Connection, 1
            LogAudit "V", "PHILHEALTH MONTHLY REMITTANCE" & cboMOnth & ", " & cboyear
        Case "REPORT SSS MONTHLY REMITTANCE":
            'If Function_Access(LOGID, "Acess_Print", "REPORT SSS MONTHLY REMITTANCE") = False Then Exit Sub
            PrintSQLReport rptPHMonthly, HRMS_REPORT_PATH & "SSSremit.rpt", "{HRMS_PAYROLL.PAY_YEAR} = " & cboyear.Text & " and {HRMS_PAYROLL.PAY_MONTH} = " & What_month(cboMOnth), DMIS_REPORT_Connection, 1
            LogAudit "V", "PRINT SSS MONTHLY REPORT", cboMOnth & ", " & cboyear
        Case "REPORT PAG-IBIG MONTHLY REMITTANCE":
            'If Function_Access(LOGID, "Acess_Print", "REPORT PAGIBIG MONTHLY REMITTANCE") = False Then Exit Sub
            PrintSQLReport rptPHMonthly, HRMS_REPORT_PATH & "Pagibigremit.rpt", "{pagibigdet.PAY_YEAR} = " & cboyear.Text & " and {pagibigdet.PAY_MONTH} = " & What_month(cboMOnth), DMIS_REPORT_Connection, 1
            LogAudit "V", "PAB-IBIG MONTHLY REMITTANCE", cboMOnth & ", " & cboyear
        Case "REPORT WITHHOLDING TAX MONTHLY REMITTANCE":
            'If Function_Access(LOGID, "Acess_Print", "REPORT TAX MONTHLY REMITTANCE") = False Then Exit Sub
            PrintSQLReport rptPHMonthly, HRMS_REPORT_PATH & "TAXremit_.rpt", "{HRMS_PAYROLL.PAY_YEAR} = " & cboyear.Text & " and {HRMS_PAYROLL.PAY_MONTH} = " & What_month(cboMOnth), DMIS_REPORT_Connection, 1
            LogAudit "V", "PRINT TAX MONTHLY REMITTANCE", cboMOnth & ", " & cboyear
    End Select
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMOnth
    cboMOnth.Text = The_month(MONTH(LOGDATE))
    'FillcboYear cboyear
    fillcombo_up cboyear
    cboyear.Text = YEAR(LOGDATE)
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

