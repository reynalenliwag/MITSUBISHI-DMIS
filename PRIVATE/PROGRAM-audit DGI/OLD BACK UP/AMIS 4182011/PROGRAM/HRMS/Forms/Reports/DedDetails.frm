VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMSDedDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deduction Details"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2820
   ForeColor       =   &H00D8E9EC&
   Icon            =   "DedDetails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   2820
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
      Left            =   1410
      MouseIcon       =   "DedDetails.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "DedDetails.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Left            =   540
      MouseIcon       =   "DedDetails.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "DedDetails.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print Report"
      Top             =   1530
      Width           =   885
   End
   Begin VB.CheckBox chkInclude 
      Caption         =   "Include Managers Sheet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   30
      TabIndex        =   3
      Top             =   975
      Width           =   2655
   End
   Begin VB.CheckBox chkPreview 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   30
      TabIndex        =   4
      Top             =   1215
      Width           =   1275
   End
   Begin VB.ComboBox cboQuensina 
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
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   2775
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
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   555
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   555
      Width           =   885
   End
   Begin Crystal.CrystalReport rptPrintPay 
      Left            =   1800
      Top             =   1575
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
End
Attribute VB_Name = "frmHRMSDedDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPAYROLL                                                         As ADODB.Recordset
Dim FromDate, ToDate                                                  As String
Attribute ToDate.VB_VarUserMemId = 1073938433

Function FindNextMonth() As String
    If cboMonth.Text = "January" Then FindNextMonth = "2"
    If cboMonth.Text = "Febuary" Then FindNextMonth = "3"
    If cboMonth.Text = "March" Then FindNextMonth = "4"
    If cboMonth.Text = "April" Then FindNextMonth = "5"
    If cboMonth.Text = "May" Then FindNextMonth = "6"
    If cboMonth.Text = "June" Then FindNextMonth = "7"
    If cboMonth.Text = "July" Then FindNextMonth = "8"
    If cboMonth.Text = "August" Then FindNextMonth = "9"
    If cboMonth.Text = "September" Then FindNextMonth = "10"
    If cboMonth.Text = "October" Then FindNextMonth = "11"
    If cboMonth.Text = "November" Then FindNextMonth = "12"
    If cboMonth.Text = "December" Then FindNextMonth = "1"
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    'If Function_Access(LOGID, "Acess_Print", "DEDUCTION DETAILS") = False Then Exit Sub
    Dim MM, ddFROM, YY                                                As String
    Dim vYEAR                                                         As String
    MM = What_month(cboMonth)
    YY = cboYear.Text

    If cboQuensina.Text = "1st Cut-Off" Then
        FromDate = DateSerial(YY, MM, PAYROLLCODE_FROM1)
        ToDate = DateSerial(YY, MM, PAYROLLCODE_TO1)
    Else
        FromDate = DateSerial(YY, MM, PAYROLLCODE_FROM2)

        If cboMonth.Text = "December" Then vYEAR = CDbl(cboYear) + 1 Else vYEAR = CDbl(cboYear)
        ToDate = FindNextMonth & "/" & PAYROLLCODE_TO2 & "/" & vYEAR
    End If

    Screen.MousePointer = 11
    rptPrintPay.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptPrintPay.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptPrintPay.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    rptPrintPay.Formulas(3) = "PREPARED_BY = '" & PREPARED_BY & "'"
    rptPrintPay.Formulas(4) = "CHECKED_BY = '" & CHECKED_BY & "'"
    rptPrintPay.Formulas(5) = "APPROVED_BY = '" & APPROVED_BY & "'"
    rptPrintPay.Formulas(6) = "PrintedBy = '" & LOGNAME & "'"

    If chkPreview.Value = 1 Then
        PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "DedDetails.rpt", "month({payroll.paydatefrom}) = " & Month(FromDate) & " AND day({payroll.paydatefrom}) = " & Day(FromDate) & " AND year({payroll.paydatefrom}) = " & YEAR(FromDate) & " AND month({payroll.paydateto}) = " & Month(ToDate) & " AND day({payroll.paydateto}) = " & Day(ToDate) & " AND year({payroll.paydateto}) = " & YEAR(ToDate) & " AND {empinfo.emplevel} = 'E'", DMIS_REPORT_Connection, 1
        If chkInclude.Value = 1 Then PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "DedDetails.rpt", "month({payroll.paydatefrom}) = " & Month(FromDate) & " AND day({payroll.paydatefrom}) = " & Day(FromDate) & " AND year({payroll.paydatefrom}) = " & YEAR(FromDate) & " AND month({payroll.paydateto}) = " & Month(ToDate) & " AND day({payroll.paydateto}) = " & Day(ToDate) & " AND year({payroll.paydateto}) = " & YEAR(ToDate) & " AND {empinfo.emplevel} = 'M'", DMIS_REPORT_Connection, 1
    Else
        PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "DedDetails.rpt", "month({payroll.paydatefrom}) = " & Month(FromDate) & " AND day({payroll.paydatefrom}) = " & Day(FromDate) & " AND year({payroll.paydatefrom}) = " & YEAR(FromDate) & " AND month({payroll.paydateto}) = " & Month(ToDate) & " AND day({payroll.paydateto}) = " & Day(ToDate) & " AND year({payroll.paydateto}) = " & YEAR(ToDate) & " AND {empinfo.emplevel} = 'E'", DMIS_REPORT_Connection, 0
        If chkInclude.Value = 1 Then PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "DedDetails.rpt", "month({payroll.paydatefrom}) = " & Month(FromDate) & " AND day({payroll.paydatefrom}) = " & Day(FromDate) & " AND year({payroll.paydatefrom}) = " & YEAR(FromDate) & " AND month({payroll.paydateto}) = " & Month(ToDate) & " AND day({payroll.paydateto}) = " & Day(ToDate) & " AND year({payroll.paydateto}) = " & YEAR(ToDate) & " AND {empinfo.emplevel} = 'M'", DMIS_REPORT_Connection, 0
    End If

    LogAudit "V", "PRINT DEDUCTION DETAILS", cboQuensina & " " & cboMonth & ", " & cboYear
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    cboQuensina.AddItem "1st Cut-Off"
    cboQuensina.AddItem "2nd Cut-Off"
    fillcbomonth cboMonth
    FillcboYear cboYear
    If Day(LOGDATE) > 15 Then
        cboQuensina.Text = "2nd Cut-Off"
    Else
        cboQuensina.Text = "1st Cut-Off"
    End If
    cboYear.Text = YEAR(LOGDATE)
    cboMonth.Text = The_month(Month(LOGDATE))
    chkPreview.Value = 1
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

