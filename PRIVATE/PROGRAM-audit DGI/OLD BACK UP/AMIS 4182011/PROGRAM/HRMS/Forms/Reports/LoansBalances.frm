VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMSLoansBalances 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Balance Reports"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4905
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
   Icon            =   "LoansBalances.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4905
   Begin VB.ComboBox cboCUTOFF 
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check1"
      Height          =   255
      Left            =   30
      TabIndex        =   7
      Top             =   690
      Width           =   225
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   30
      TabIndex        =   6
      Top             =   210
      Width           =   225
   End
   Begin VB.ComboBox cboMonth 
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1581
      Width           =   1695
   End
   Begin VB.ComboBox cboName 
      Height          =   360
      Left            =   510
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   180
      Width           =   4245
   End
   Begin VB.ComboBox cboYear 
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2100
      Width           =   1695
   End
   Begin VB.ComboBox cbotype 
      Height          =   360
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   630
      Width           =   4275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2475
      MouseIcon       =   "LoansBalances.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "LoansBalances.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   2655
      Width           =   885
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1605
      MouseIcon       =   "LoansBalances.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "LoansBalances.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   2655
      Width           =   885
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
End
Attribute VB_Name = "frmHRMSLoansBalances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub FillLoanType()
    Dim RSTMP                                                         As New ADODB.Recordset
    Set RSTMP = New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT DESCRIPTION + '     ' + CODE FROM HRMS_LOANCODE ORDER BY CODE")
    If Not RSTMP.EOF And Not RSTMP.BOF Then
        Combo_Loadval cboTYPE, RSTMP
    End If
    cboTYPE.AddItem "ALL", 0
    cboTYPE.ListIndex = 0
    Set RSTMP = Nothing
End Sub

Sub FillYear()
    cboYear.Clear
    cboYear.AddItem (YEAR(Date) - 8)
    cboYear.AddItem (YEAR(Date) - 7)
    cboYear.AddItem (YEAR(Date) - 6)
    cboYear.AddItem (YEAR(Date) - 5)
    cboYear.AddItem (YEAR(Date) - 4)
    cboYear.AddItem (YEAR(Date) - 3)
    cboYear.AddItem (YEAR(Date) - 2)
    cboYear.AddItem (YEAR(Date) - 1)
    cboYear.AddItem YEAR(Date)
    cboYear.AddItem (YEAR(Date) + 1)
    cboYear.AddItem (YEAR(Date) + 2)
    cboYear.AddItem (YEAR(Date) + 3)
    cboYear.AddItem (YEAR(Date) + 4)
    cboYear.AddItem (YEAR(Date) + 5)
    cboYear.AddItem (YEAR(Date) + 6)
    cboYear.AddItem (YEAR(Date) + 7)
    cboYear.AddItem (YEAR(Date) + 8)
    cboYear.AddItem "ALL", 0
    cboYear.ListIndex = 0
End Sub

Sub FillEmpNo()
    cboName.Clear
    Dim rsEmpInfo                                                     As ADODB.Recordset
    Set rsEmpInfo = New ADODB.Recordset
    Set rsEmpInfo = gconDMIS.Execute("Select empno + '        '  + lastname + ', ' + firstname from HRMS_Empinfo order by lastname ")
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Combo_Loadval cboName, rsEmpInfo
    End If
    cboName.AddItem "ALL", 0
    cboName.ListIndex = 0
    Set rsEmpInfo = Nothing
End Sub

Sub FillMonth()
    cboMOnth.Clear
    cboMOnth.AddItem MonthName(1)
    cboMOnth.AddItem MonthName(2)
    cboMOnth.AddItem MonthName(3)
    cboMOnth.AddItem MonthName(4)
    cboMOnth.AddItem MonthName(5)
    cboMOnth.AddItem MonthName(6)
    cboMOnth.AddItem MonthName(7)
    cboMOnth.AddItem MonthName(8)
    cboMOnth.AddItem MonthName(9)
    cboMOnth.AddItem MonthName(10)
    cboMOnth.AddItem MonthName(11)
    cboMOnth.AddItem MonthName(12)
    cboMOnth.AddItem "ALL", 0
    cboMOnth.ListIndex = 0
End Sub

Sub FillCUTOFF()
    cboCUTOFF.AddItem "1st Cut-Off"
    cboCUTOFF.AddItem "2nd Cut-Off"
    cboCUTOFF.AddItem "ALL", 0
    cboCUTOFF.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:
    Dim FILTER                                                        As String
    Dim FILTERDET1                                                    As String
    Dim FILTERDET2                                                    As String
    Dim FILTERDET3                                                    As String
    Dim FILTERDET4                                                    As String
    Dim FILTERDET5                                                    As String
    Dim LOANCODE                                                      As String
    Dim strCUTOFF                                                     As String
    FILTER = ""

    LOANCODE = Trim(Right(cboTYPE, 4))
    Screen.MousePointer = 11

    If cboCUTOFF = "1st Cut-Off" Then
        strCUTOFF = "1"
    ElseIf cboCUTOFF = "2nd Cut-Off" Then
        strCUTOFF = "2"
    End If

    If COMPANY_CODE = "HAI" Then
        If Check1.Value = 1 Then
            FILTERDET1 = "{Loanmasdet.empno} = '" & Trim(Left(cboName.Text, 11)) & "'"
        Else
            FILTERDET1 = ""
        End If
    Else
        If Check1.Value = 1 Then
            FILTERDET1 = "{Loanmasdet.empno} = '" & Trim(Left(cboName.Text, 6)) & "'"
        Else
            FILTERDET1 = ""
        End If
    End If

    If Check2.Value = 1 Then
        FILTERDET2 = "RIGHT({Loanmasdet.loantype},4) = '" & LOANCODE & "'"
    Else
        FILTERDET2 = ""
    End If
    
    If Check3.Value = 1 Then
        FILTERDET3 = "{Loanmasdet.PAY_MONTH} = " & What_month(cboMOnth.Text) & " "
    Else
        FILTERDET3 = ""
    End If
    
    If Check4.Value = 1 Then
        FILTERDET4 = "{Loanmasdet.PAY_YEAR} = " & cboYear.Text & " "
    Else
        FILTERDET4 = ""
    End If
    If Check5.Value = 1 Then
        FILTERDET5 = "{Loanmasdet.CUT_OFF} = " & N2Str2Null(strCUTOFF) & " "
    Else
        FILTERDET5 = ""
    End If

    If cboName.Text = "ALL" Then
        FILTERDET1 = ""
    End If
    
    If cboTYPE.Text = "ALL" Then
        FILTERDET2 = ""
    End If
    
    If cboMOnth.Text = "ALL" Then
        FILTERDET3 = ""
    End If
    
    If cboYear.Text = "ALL" Then
        FILTERDET4 = ""
    End If
    
    If cboCUTOFF.Text = "ALL" Then
        FILTERDET5 = ""
    End If

    If Len(FILTERDET1) > 0 Then
        FILTER = FILTERDET1
    ElseIf Len(FILTERDET2) > 0 Then
        FILTER = FILTERDET2
    ElseIf Len(FILTERDET3) > 0 Then
        FILTER = FILTERDET3
    ElseIf Len(FILTERDET4) > 0 Then
        FILTER = FILTERDET4
    ElseIf Len(FILTERDET5) > 0 Then
        FILTER = FILTERDET5
    End If

    If Len(FILTERDET2) > 0 Then
        FILTER = FILTER & " AND " & FILTERDET2
    End If
    If Len(FILTERDET3) > 0 Then
        FILTER = FILTER & " AND " & FILTERDET3
    End If
    If Len(FILTERDET4) > 0 Then
        FILTER = FILTER & " AND " & FILTERDET4
    End If
    If Len(FILTERDET5) > 0 Then
        FILTER = FILTER & " AND " & FILTERDET5
    End If

    rptSSSLoanMonthly.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptSSSLoanMonthly.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptSSSLoanMonthly.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    rptSSSLoanMonthly.Formulas(3) = "PREPARED_BY = '" & PREPARED_BY & "'"
    rptSSSLoanMonthly.Formulas(4) = "CHECKED_BY = '" & CHECKED_BY & "'"
    rptSSSLoanMonthly.Formulas(5) = "APPROVED_BY = '" & APPROVED_BY & "'"
    rptSSSLoanMonthly.Formulas(6) = "PrintedBy    = '" & LOGNAME & "'"
    PrintSQLReport rptSSSLoanMonthly, HRMS_REPORT_PATH & "LOAN BALANCES DETAIL.RPT", FILTER, DMIS_REPORT_Connection, 1
    'LogAudit "V", "PRINT LOAN MONTHLY REMITTANCE", cboType.Text & " - " & cboMonth.Text & " - " & cboYear.Text

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
    DrawXPCtl Me
    FillLoanType
    FillYear
    FillEmpNo
    FillMonth
    FillCUTOFF
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub Option1_Click()
    FillLoanType
End Sub

