VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMSPayJournal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Journal"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   3720
   ForeColor       =   &H8000000F&
   Icon            =   "PayJournal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3720
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
      Left            =   1605
      MouseIcon       =   "PayJournal.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "PayJournal.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Close Window"
      Top             =   1680
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
      Left            =   735
      MouseIcon       =   "PayJournal.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "PayJournal.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print Report"
      Top             =   1680
      Width           =   885
   End
   Begin VB.CheckBox chkContractual 
      Caption         =   "Print for Contractual Employees"
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
      Left            =   60
      TabIndex        =   3
      Top             =   915
      Width           =   3345
   End
   Begin VB.CheckBox chkAllowanceBase 
      Caption         =   "Print for Allowance Base Employees"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      TabIndex        =   4
      Top             =   1200
      Width           =   3345
   End
   Begin VB.CheckBox chkProbReg 
      Caption         =   "Print for Probationary/Regular Employees"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      TabIndex        =   2
      Top             =   555
      Value           =   1  'Checked
      Width           =   3405
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
      Left            =   300
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   165
      Width           =   1845
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
      Left            =   2190
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   165
      Width           =   885
   End
   Begin Crystal.CrystalReport rptPrintPay 
      Left            =   2700
      Top             =   1965
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Payroll Journal"
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
End
Attribute VB_Name = "frmHRMSPayJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200711:47
Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:

    'If Function_Access(LOGID, "Acess_Print", "PAYROLL JOURNAL") = False Then Exit Sub
    Dim MM, YY                                                        As String
    MM = What_month(cboMonth)
    YY = cboYear.Text
    Screen.MousePointer = 11

    If chkProbReg.Value = 1 Then
        If LOGLEVEL = "ADM" Then
            rptPrintPay.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
            rptPrintPay.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
            rptPrintPay.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
            rptPrintPay.Formulas(3) = "PREPARED_BY = '" & PREPARED_BY & "'"
            rptPrintPay.Formulas(4) = "CHECKED_BY = '" & CHECKED_BY & "'"
            rptPrintPay.Formulas(5) = "APPROVED_BY = '" & APPROVED_BY & "'"

            PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "payrolljournal.rpt", "month({payroll.paydatefrom}) = " & MM & " AND year({payroll.paydatefrom}) = " & YY, DMIS_REPORT_Connection, 1
            LogAudit "V", "PAYROLL JOURNAL", cboMonth & ", " & cboYear & " REGULAR-ADM"
        Else
            rptPrintPay.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
            rptPrintPay.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
            rptPrintPay.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
            rptPrintPay.Formulas(3) = "PREPARED_BY = '" & PREPARED_BY & "'"
            rptPrintPay.Formulas(4) = "CHECKED_BY = '" & CHECKED_BY & "'"
            rptPrintPay.Formulas(5) = "APPROVED_BY = '" & APPROVED_BY & "'"

            PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "payrolljournalE.rpt", "month({payroll.paydatefrom}) = " & MM & " AND year({payroll.paydatefrom}) = " & YY, DMIS_REPORT_Connection, 1
            LogAudit "V", "PAYROLL JOURNAL", cboMonth & ", " & cboYear & " REGULAR"
        End If
    End If
    If chkContractual.Value = 1 Then
        PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "payrolljournalC.rpt", "month({payroll.paydatefrom}) = " & MM & " AND year({payroll.paydatefrom}) = " & YY, DMIS_REPORT_Connection, 1
        LogAudit "V", "PAYROLL JOURNAL", cboMonth & ", " & cboYear & " CONTRACTUAL"
    End If
    If chkAllowanceBase.Value = 1 Then
        PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "payrolljournalA.rpt", "month({payroll.paydatefrom}) = " & MM & " AND year({payroll.paydatefrom}) = " & YY, DMIS_REPORT_Connection, 1
        LogAudit "V", "PAYROLL JOURNAL", cboMonth & ", " & cboYear & " ALLOWANCE"
    End If

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


    fillcbomonth cboMonth
    FillcboYear cboYear
    cboYear.Text = YEAR(LOGDATE)
    cboMonth.Text = The_month(Month(LOGDATE))
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmHRMSPayJournal = Nothing
End Sub

