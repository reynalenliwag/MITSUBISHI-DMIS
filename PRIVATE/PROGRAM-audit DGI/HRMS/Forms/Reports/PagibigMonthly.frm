VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMSPagibigMonthly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pag-Ibig Monthly Remittance"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4875
   ForeColor       =   &H00D8E9EC&
   Icon            =   "PagibigMonthly.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4875
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
      Left            =   2355
      MouseIcon       =   "PagibigMonthly.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "PagibigMonthly.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   750
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
      Left            =   1485
      MouseIcon       =   "PagibigMonthly.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "PagibigMonthly.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   750
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
      Left            =   3450
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
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
      Top             =   180
      Width           =   2355
   End
   Begin Crystal.CrystalReport rptPagibigMonthly 
      Left            =   3480
      Top             =   1110
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
      Left            =   2550
      TabIndex        =   2
      Top             =   210
      Width           =   825
   End
End
Attribute VB_Name = "frmHRMSPagibigMonthly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPagibigdet                             As ADODB.Recordset
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "Acess_Print", "REPORT PAG-IBIG MONTHLY REMITTANCE") = False Then Exit Sub
    Screen.MousePointer = 11
    rptPagibigMonthly.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptPagibigMonthly.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptPagibigMonthly.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    rptPagibigMonthly.Formulas(3) = "PREPARED_BY = '" & PREPARED_BY & "'"
    rptPagibigMonthly.Formulas(4) = "CHECKED_BY = '" & CHECKED_BY & "'"
    rptPagibigMonthly.Formulas(5) = "APPROVED_BY = '" & APPROVED_BY & "'"
    rptPagibigMonthly.Formulas(6) = "PayrollCode = '" & PAYROLL_CODE & "'"
    rptPagibigMonthly.Formulas(9) = "Printedby = '" & LOGNAME & "'"
    rptPagibigMonthly.Formulas(10) = "Yeer = '" & cboyear & "'"
    rptPagibigMonthly.Formulas(11) = "Month = '" & cboMOnth & "'"
    'remarked 03262008:1330HRS
    '=====================================================
    'PrintSQLReport rptPagibigMonthly, HRMS_REPORT_PATH & "Pagibigremit.rpt", "year({Pagibigdet.deyt}) =" & cboYear.Text & " and month({Pagibigdet.deyt}) =" & What_month(cboMonth), DMIS_REPORT_CONNECTION, 1
    '=====================================================
    'updated 03262008:1330HRS
    '=====================================================
    PrintSQLReport rptPagibigMonthly, HRMS_REPORT_PATH & "Pagibigremit.rpt", "{Pagibigdet.PAY_YEAR} = " & cboyear.Text & " and {Pagibigdet.Pay_month} = " & What_month(cboMOnth) & " and ({Pagibigdet.Cut_off} = '1' OR {Pagibigdet.Cut_off} = '2' )", DMIS_REPORT_Connection, 1
    '=====================================================
    LogAudit "V", "PAB-IBIG MONTHLY REMITTANCE", Left(cboMOnth, 3) & " " & PAYROLLCODE_FROM1 & "-" & PAYROLLCODE_TO1
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
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    
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
