VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMSSSSLoanQuarter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SSS Loan Quarterly Report"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4725
   ForeColor       =   &H00D8E9EC&
   Icon            =   "SSSLoanQuarter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4725
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
      Left            =   2310
      MouseIcon       =   "SSSLoanQuarter.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "SSSLoanQuarter.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   1260
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
      Left            =   1440
      MouseIcon       =   "SSSLoanQuarter.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "SSSLoanQuarter.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   1260
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
      TabIndex        =   2
      Top             =   660
      Width           =   1245
   End
   Begin VB.ComboBox cboQuarter 
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
      Top             =   660
      Width           =   2355
   End
   Begin VB.ComboBox cboLoanType 
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
      TabIndex        =   1
      Top             =   120
      Width           =   4635
   End
   Begin Crystal.CrystalReport rptSSSLoanQuarter 
      Left            =   3480
      Top             =   1440
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
      TabIndex        =   3
      Top             =   660
      Width           =   825
   End
End
Attribute VB_Name = "frmHRMSSSSLoanQuarter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLoanmasDet                                                      As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200711:50
Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Print", "SSS LOAN QUARTERLY REPORT") = False Then Exit Sub
    If cboQuarter.Text = "First Quarter" Then
        Set rsLoanmasDet = New ADODB.Recordset
        rsLoanmasDet.Open "select * from HRMS_LoanMasDet where year(deyt) = " & cboYear.Text & " and month(deyt) >= " & 1 & " AND month(deyt) <= " & 3, gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf cboQuarter.Text = "Second Quarter" Then
        Set rsLoanmasDet = New ADODB.Recordset
        rsLoanmasDet.Open "select * from HRMS_LoanMasDet where year(deyt) = " & cboYear.Text & " and month(deyt) >= " & 4 & " AND month(deyt) <= " & 6, gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf cboQuarter.Text = "Third Quarter" Then
        Set rsLoanmasDet = New ADODB.Recordset
        rsLoanmasDet.Open "select * from HRMS_LoanMasDet where year(deyt) = " & cboYear.Text & " and month(deyt) >= " & 7 & " AND month(deyt) <= " & 9, gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        Set rsLoanmasDet = New ADODB.Recordset
        rsLoanmasDet.Open "select * from HRMS_LoanMasDet where year(deyt) = " & cboYear.Text & " and month(deyt) >= " & 10 & " AND month(deyt) <= " & 12, gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not rsLoanmasDet.EOF And Not rsLoanmasDet.BOF Then
        Dim FILTER, RepName                                           As String

        If cboLoanType.Text = "Salary Loan" Then
            RepName = HRMS_REPORT_PATH & "SSSSalLoanQuart.rpt"
        Else
            RepName = HRMS_REPORT_PATH & "SSSCalLoanQuart.rpt"
        End If

        If cboQuarter.Text = "First Quarter" Then FILTER = "year({Loanmasdet.deyt}) = " & cboYear.Text & " and month({Loanmasdet.deyt}) >= " & 1 & " AND month({Loanmasdet.deyt}) <= " & 3
        If cboQuarter.Text = "Second Quarter" Then FILTER = "year({Loanmasdet.deyt}) = " & cboYear.Text & " and month({Loanmasdet.deyt}) >= " & 4 & " AND month({Loanmasdet.deyt}) <= " & 6
        If cboQuarter.Text = "Third Quarter" Then FILTER = "year({Loanmasdet.deyt}) = " & cboYear.Text & " and month({Loanmasdet.deyt}) >= " & 7 & " AND month({Loanmasdet.deyt}) <= " & 9
        If cboQuarter.Text = "Fourth Quarter" Then FILTER = "year({Loanmasdet.deyt}) = " & cboYear.Text & " and month({Loanmasdet.deyt}) >= " & 10 & " AND month({Loanmasdet.deyt}) <= " & 12
        Screen.MousePointer = 11
        rptSSSLoanQuarter.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
        rptSSSLoanQuarter.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
        rptSSSLoanQuarter.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
        rptSSSLoanQuarter.Formulas(3) = "PREPARED_BY = '" & PREPARED_BY & "'"
        rptSSSLoanQuarter.Formulas(4) = "CHECKED_BY = '" & CHECKED_BY & "'"
        rptSSSLoanQuarter.Formulas(5) = "APPROVED_BY = '" & APPROVED_BY & "'"
        rptSSSLoanQuarter.Formulas(6) = "HEADING = '" & "SSS LOAN QUARTERLY REPORT" & "'"

        PrintSQLReport rptSSSLoanQuarter, RepName, FILTER, DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0

        Call LogAudit("V", "PRINT LOAN QUARTERLY REPORT", cboLoanType & "-" & cboQuarter & "-" & cboYear)
    End If

    Exit Sub

Errorcode:
    Call ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    cboQuarter.Clear
    cboQuarter.AddItem "First Quarter"
    cboQuarter.AddItem "Second Quarter"
    cboQuarter.AddItem "Third Quarter"
    cboQuarter.AddItem "Fourth Quarter"
    cboLoanType.Clear
    cboLoanType.AddItem "Salary Loan"
    cboLoanType.AddItem "Calamity Loan"
    cboYear.Text = YEAR(LOGDATE)
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

