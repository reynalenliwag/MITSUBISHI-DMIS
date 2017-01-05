VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMSLoansMonthly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Monthly Remittance"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4875
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00D8E9EC&
   Icon            =   "SSSLoanMonthly.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2490
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
      Left            =   2535
      MouseIcon       =   "SSSLoanMonthly.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "SSSLoanMonthly.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   1605
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
      Left            =   1650
      MouseIcon       =   "SSSLoanMonthly.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "SSSLoanMonthly.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   1605
      Width           =   885
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1110
      Width           =   1605
   End
   Begin VB.ComboBox cboLoanType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   210
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   450
      Width           =   4515
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1140
      Width           =   2715
   End
   Begin Crystal.CrystalReport rptSSSLoanMonthly 
      Left            =   8280
      Top             =   180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   180
      TabIndex        =   7
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Select Your Loan Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   180
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   3090
      TabIndex        =   3
      Top             =   870
      Width           =   510
   End
End
Attribute VB_Name = "frmHRMSLoansMonthly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub FillLoanType()
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim LENLOAN                                                       As String
    Dim SPACES                                                        As String

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_LoanCode Order By Description ASc")
    cboLoanType.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            If Len(RSTMP!CODE) = 3 Then SPACES = " "
            If Len(RSTMP!CODE) = 2 Then SPACES = "  "
            If Len(RSTMP!CODE) = 1 Then SPACES = "   "
            cboLoanType.AddItem Null2String(RSTMP!Description) & "-" & SPACES & RSTMP!CODE

            RSTMP.MoveNext
        Loop

    End If
    cboLoanType.AddItem "ALL", 0
    cboLoanType.ListIndex = 0

    Set RSTMP = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:
    Dim FILTER                                                        As String
    Dim LOANCODE                                                      As String
    LOANCODE = Trim(Right(cboLoanType, 4))
    Screen.MousePointer = 11
    If cboLoanType = "ALL" Then
        FILTER = "year({Loanmasdet.deyt}) = " & cboyear.Text & " and month({Loanmasdet.deyt}) = " & What_month(cboMOnth)
    Else
        FILTER = "{LoanMasDet.LoanType} = '" & LOANCODE & "' AND year({Loanmasdet.deyt}) = " & cboyear.Text & " and month({Loanmasdet.deyt}) = " & What_month(cboMOnth)
    End If



    rptSSSLoanMonthly.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptSSSLoanMonthly.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptSSSLoanMonthly.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    rptSSSLoanMonthly.Formulas(3) = "PREPARED_BY = '" & PREPARED_BY & "'"
    rptSSSLoanMonthly.Formulas(4) = "CHECKED_BY = '" & CHECKED_BY & "'"
    rptSSSLoanMonthly.Formulas(5) = "APPROVED_BY = '" & APPROVED_BY & "'"
    rptSSSLoanMonthly.Formulas(6) = "PrintedBy    = '" & LOGNAME & "'"
    PrintSQLReport rptSSSLoanMonthly, HRMS_REPORT_PATH & "Loanremit.rpt", FILTER, DMIS_REPORT_Connection, 1
    LogAudit "V", "PRINT LOAN MONTHLY REMITTANCE", cboLoanType.Text & " - " & cboMOnth.Text & " - " & cboyear.Text
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
    'FillcboYear cboyear
    fillcombo_up cboyear
    FillLoanType
    cboMOnth.Text = The_month(MONTH(LOGDATE))
    cboyear.Text = YEAR(LOGDATE)
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

