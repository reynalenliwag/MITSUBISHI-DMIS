VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMSPagibigLoanMonthly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pag-Ibig Loan Monthly Remittance"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4845
   ForeColor       =   &H00D8E9EC&
   Icon            =   "PagibigLoanMonthly.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4845
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
      Left            =   2385
      MouseIcon       =   "PagibigLoanMonthly.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "PagibigLoanMonthly.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   840
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
      Left            =   1515
      MouseIcon       =   "PagibigLoanMonthly.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "PagibigLoanMonthly.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   840
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
      Top             =   225
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
      Top             =   225
      Width           =   2355
   End
   Begin Crystal.CrystalReport rptPagibigLoanMonthly 
      Left            =   3450
      Top             =   1185
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
      Top             =   255
      Width           =   825
   End
End
Attribute VB_Name = "frmHRMSPagibigLoanMonthly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLoanmasDet                                                      As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200711:47
Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:

    'If Function_Access(LOGID, "Acess_Print", "PAG-IBIG LOAN MONTHLY REMITTANCE") = False Then Exit Sub
    Screen.MousePointer = 11
    rptPagibigLoanMonthly.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptPagibigLoanMonthly.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptPagibigLoanMonthly.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    rptPagibigLoanMonthly.Formulas(3) = "PREPARED_BY = '" & PREPARED_BY & "'"
    rptPagibigLoanMonthly.Formulas(4) = "CHECKED_BY = '" & CHECKED_BY & "'"
    rptPagibigLoanMonthly.Formulas(5) = "APPROVED_BY = '" & APPROVED_BY & "'"
    PrintSQLReport rptPagibigLoanMonthly, HRMS_REPORT_PATH & "PagibigLoanremit.rpt", "year({Loanmasdet.deyt}) = " & cboYear.Text & " and month({Loanmasdet.deyt}) = " & What_month(cboMonth), DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0

    LogAudit "V", "PRINT PAGIBIG MONTHL LOAN", LOGNAME
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


    fillcbomonth cboMonth
    cboMonth.Text = The_month(Month(LOGDATE))
    FillcboYear cboYear
    cboYear.Text = YEAR(LOGDATE)
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

