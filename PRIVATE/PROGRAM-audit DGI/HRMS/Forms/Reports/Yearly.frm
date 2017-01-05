VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmHRMSYearly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report By Year"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3135
   ForeColor       =   &H00D8E9EC&
   Icon            =   "Yearly.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1830
   ScaleWidth      =   3135
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
      Left            =   2280
      MouseIcon       =   "Yearly.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "Yearly.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   960
      Width           =   795
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
      Left            =   1500
      MouseIcon       =   "Yearly.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "Yearly.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   960
      Width           =   795
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
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   390
      Width           =   2025
   End
   Begin Crystal.CrystalReport rptYearly 
      Left            =   1185
      Top             =   1230
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
   Begin XtremeShortcutBar.ShortcutCaption cap 
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3255
      _Version        =   655364
      _ExtentX        =   5741
      _ExtentY        =   609
      _StockProps     =   14
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      ForeColor       =   8388608
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
      Left            =   120
      TabIndex        =   1
      Top             =   390
      Width           =   825
   End
End
Attribute VB_Name = "frmHRMSYearly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL As String
Dim REPORT As String
Dim TAGGED As Integer

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    'FillcboYear cboyear
    fillcombo_up cboyear
    cboyear.Text = YEAR(LOGDATE)
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    rptYearly.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptYearly.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    'rptYearly.Reset
    
    If FORMYEARLYREQUEST = "ALTERMINATED" Then
        If Function_Access(LOGID, "Acess_Print", "ALPHALIST REPORT") = False Then Exit Sub
        SQL = "SELECT * FROM HRMS_EMPINFO WHERE YEAR(RESIGNED) = '" & cboyear.Text & "' AND (EMPLEVEL = 'E' OR EMPLEVEL = 'M') ORDER BY LASTNAME"
        REPORT = "ALPHATERMINATED.XLT"
        TAGGED = 1
        Call PrintALPHAExcel(SQL, REPORT, TAGGED)
        Call LogAudit("V", "PRINT YEARLY REPORT", cboyear)
    ElseIf FORMYEARLYREQUEST = "ALWITHEMP" Then
        If Function_Access(LOGID, "Acess_Print", "ALPHALIST REPORT") = False Then Exit Sub
        SQL = "SELECT * FROM HRMS_EMPINFO WHERE YEAR(DATEHIRED) = '" & cboyear.Text & "' AND (EMPLEVEL = 'E' OR EMPLEVEL = 'M') AND PREVIOUSCOMPANY IS NOT NULL AND RESIGNED IS NULL ORDER BY LASTNAME"
        REPORT = "ALPHAWITHEMP.XLT"
        TAGGED = 2
        Call PrintALPHAExcel(SQL, REPORT, TAGGED)
        Call LogAudit("V", "PRINT YEARLY REPORT", cboyear)
    ElseIf FORMYEARLYREQUEST = "ALWITHNOEMP" Then
        If Function_Access(LOGID, "Acess_Print", "ALPHALIST REPORT") = False Then Exit Sub
        SQL = "SELECT * FROM HRMS_EMPINFO WHERE (EMPLEVEL = 'E' OR EMPLEVEL = 'M') AND YEAR(DATEHIRED) <= " & cboyear.Text & " AND RESIGNED IS NULL AND (PREVIOUSCOMPANY IS NULL OR YEAR(DATEHIRED) <> '" & cboyear.Text & "') ORDER BY LASTNAME"
        REPORT = "ALPHAWITHNOEMP.XLT"
        TAGGED = 3
        Call PrintALPHAExcel(SQL, REPORT, TAGGED)
        Call LogAudit("V", "PRINT YEARLY REPORT", cboyear)
    ElseIf FORMYEARLYREQUEST = "SCHEDSSS" Then
        If Function_Access(LOGID, "Acess_Print", "REPORT SCHEDULE OF SSS PREMIUM CONTRIBUTION") = False Then Exit Sub
        PrintSQLReport rptYearly, HRMS_REPORT_PATH & "schedSSS.rpt", "{PAYROLL.PAY_YEAR} = " & cboyear.Text, DMIS_REPORT_Connection, 1
        Call LogAudit("V", "SCHEDULE SSS REPORT", cboyear)
    ElseIf FORMYEARLYREQUEST = "SCHEDPHIC" Then
        If Function_Access(LOGID, "Acess_Print", "REPORT SCHEDULE OF PHILHEALTH PREMIUM CONTRIBUTION") = False Then Exit Sub
        PrintSQLReport rptYearly, HRMS_REPORT_PATH & "schedPHIC.rpt", "{PAYROLL.PAY_YEAR} = " & cboyear.Text, DMIS_REPORT_Connection, 1
        Call LogAudit("V", "SCHEDULE PHILHEALTH REPORT", cboyear)
    ElseIf FORMYEARLYREQUEST = "SCHEDPAGIBIG" Then
        If Function_Access(LOGID, "Acess_Print", "REPORT SCHEDULE OF PAGIBIG PREMIUM CONTRIBUTION") = False Then Exit Sub
        PrintSQLReport rptYearly, HRMS_REPORT_PATH & "schedPagIbig.rpt", "{PAYROLL.PAY_YEAR} = " & cboyear.Text, DMIS_REPORT_Connection, 1
        Call LogAudit("V", "SCHEDULE PAGIBIG REPORT", cboyear)
    ElseIf FORMYEARLYREQUEST = "SCHEDTAX" Then
        If LOGLEVEL = "ADM" Then
            If Function_Access(LOGID, "Acess_Print", "REPORT SCHEDULE OF TAX WITHHELD") = False Then Exit Sub
            PrintSQLReport rptYearly, HRMS_REPORT_PATH & "schedTax.rpt", "{PAYROLL.PAY_YEAR} = " & cboyear.Text, DMIS_REPORT_Connection, 1
            Call LogAudit("V", "SCHEDULE TAX ADMINSTATOR REPORT", cboyear)
        Else
            If Function_Access(LOGID, "Acess_Print", "REPORT SCHEDULE OF TAX WITHHELD") = False Then Exit Sub
            PrintSQLReport rptYearly, HRMS_REPORT_PATH & "schedTax.rpt", "{Payroll.Emplevel} = 'E' and " & "{HRMS_PAYROLL.PAY_YEAR} = " & cboyear.Text, DMIS_REPORT_Connection, 1
            Call LogAudit("V", "SCHEDULE TAX REPORT", cboyear)
        End If
    ElseIf FORMYEARLYREQUEST = "SCHEDOVERTIME" Then
        If Function_Access(LOGID, "Acess_Print", "REPORT SCHEDULE OF OVERTIME PAY") = False Then Exit Sub
        PrintSQLReport rptYearly, HRMS_REPORT_PATH & "schedOvertime.rpt", "{PAYROLL.PAY_YEAR} = " & cboyear.Text, DMIS_REPORT_Connection, 1
        Call LogAudit("V", "SCHEDULE OVERTIME REPORT", cboyear)
    ElseIf FORMYEARLYREQUEST = "SCHEDPAYROLL" Then
        If LOGLEVEL = "ADM" Then
            If Function_Access(LOGID, "Acess_Print", "REPORT SCHEDULE OF PAYROLL") = False Then Exit Sub
            PrintSQLReport rptYearly, HRMS_REPORT_PATH & "schedPayroll.rpt", "{PAYROLL.PAY_YEAR} = " & cboyear.Text, DMIS_REPORT_Connection, 1
            Call LogAudit("V", "SCHEDULE PAYROLL ADMIN REPORT", cboyear)
        Else
            If Function_Access(LOGID, "Acess_Print", "REPORT SCHEDULE OF PAYROLL") = False Then Exit Sub
            PrintSQLReport rptYearly, HRMS_REPORT_PATH & "schedPayroll.rpt", "{Payroll.Emplevel} = 'E' and " & "{HRMS_PAYROLL.PAY_YEAR} = " & cboyear.Text, DMIS_REPORT_Connection, 1
            Call LogAudit("V", "SCHEDULE PAYROLL REPORT", cboyear)
        End If
    
    ElseIf FORMYEARLYREQUEST = "YEARLYSCHEDPAYROLL" Then
            
            If Function_Access(LOGID, "Acess_Print", "REPORTS YEARLY INDIVIDUAL") = False Then Exit Sub
            PrintSQLReport rptYearly, HRMS_REPORT_PATH & "SchedPayroll_year.rpt", "{PAYROLL.PAY_YEAR} = " & cboyear.Text, DMIS_REPORT_Connection, 1
                
    ElseIf FORMYEARLYREQUEST = "SCHEDCOMMISSION" Then
        If Function_Access(LOGID, "Acess_Print", "REPORT SCHEDULE OF COMMISSION") = False Then Exit Sub
        PrintSQLReport rptYearly, HRMS_REPORT_PATH & "schedCommission.rpt", "{HRMS_PAYROLL.PAY_YEAR}  = " & cboyear.Text & " AND {@TotalCont} > 0", DMIS_REPORT_Connection, 1
        Call LogAudit("V", "SCHEDULE COMMISSION REPORT", cboyear)
    ElseIf FORMYEARLYREQUEST = "SCHEDCOMMISSIONTAX" Then
        If Function_Access(LOGID, "Acess_Print", "REPORT SCHEDULE OF COMMISSION TAX") = False Then Exit Sub
        PrintSQLReport rptYearly, HRMS_REPORT_PATH & "schedCommissionTax.rpt", "{HRMS_PAYROLL.PAY_YEAR}  = " & cboyear.Text & " AND {@TotalCont} > 0", DMIS_REPORT_Connection, 1
        Call LogAudit("V", "SCHEDULE COMMISSION TAX REPORT", cboyear)
    ElseIf FORMYEARLYREQUEST = "SCHEDTAXDUEREFUND" Then
        If Function_Access(LOGID, "Acess_Print", "REPORT SCHEDULE OF TAX DUE") = False Then Exit Sub
        PrintSQLReport rptYearly, HRMS_REPORT_PATH & "refund.rpt", "{PAYROLL.PAY_YEAR}  = '" & cboyear.Text & "'", DMIS_REPORT_Connection, 1
        Call LogAudit("V", "SCHEDULE TAX DUE REFUND REPORT", cboyear)
'    ElseIf FORMYEARLYREQUEST = "LEAVESUMMARY" Then
'        If Function_Access(LOGID, "Acess_Print", "REPORT SCHEDULE OF COMMISSION TAX") = False Then Exit Sub
'        'PrintSQLReport rptYearly, HRMS_REPORT_PATH & "Leave Summary.rpt", "", DMIS_REPORT_Connection, 1
'        PrintSQLReport rptYearly, HRMS_REPORT_PATH & "Leave_Summary.rpt", "", DMIS_REPORT_Connection, 1
'        Call LogAudit("V", "SCHEDULE TAX DUE REFUND REPORT", cboyear)
    ElseIf FORMYEARLYREQUEST = "SCHEDDEDUCTION" Then
        If Function_Access(LOGID, "Acess_Print", "REPORT SCHEDULE OF DEDUCTION") = False Then Exit Sub
        Call Print_Excel_SCHEDDEDUCTION(cboyear)
        Call LogAudit("V", "SCHEDULE DEDUCTION REPORT", cboyear)
    End If
    Screen.MousePointer = 0
End Sub

Function GetSumGross(EMPNO As String, EMPLEVEL As String, YEAR As Integer) As Double
    Dim rsPAYROLL As ADODB.Recordset
    Set rsPAYROLL = New ADODB.Recordset
    Set rsPAYROLL = gconDMIS.Execute("SELECT (SUM(RATE) - SUM(UNDERTIME) - SUM(ABSENT) + SUM(OVERTIME) + SUM(TAXABLEADJ)) as SUMGROSS FROM HRMS_PAYROLL WHERE EMPNO = '" & EMPNO & "' AND EMPLEVEL = '" & EMPLEVEL & "' AND PAY_YEAR = " & YEAR)
    
    GetSumGross = 0
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        GetSumGross = Round(N2Str2Zero(rsPAYROLL!SUMGROSS), 2)
    End If
    Set rsPAYROLL = Nothing
End Function

Function GetSumEmployerContribution(EMPNO As String, EMPLEVEL As String, YEAR As Integer) As Double
    Dim EC As Double
    Dim rsEC As ADODB.Recordset
    Set rsEC = New ADODB.Recordset

'original
'    Set rsEC = gconDMIS.Execute("SELECT SSSE FROM HRMS_PAYROLL WHERE EMPNO = '" & EMPNO & "' AND EMPLEVEL = '" & EMPLEVEL & "' AND PAY_YEAR = " & YEAR)
'    EC = 0
'    If Not rsEC.EOF And Not rsEC.BOF Then
'        rsEC.MoveFirst
'        While Not rsEC.EOF
'            If N2Str2Zero(rsEC!SSSE) >= 500# Then
'                EC = EC + 30
'            ElseIf N2Str2Zero(rsEC!SSSE) < 500# And N2Str2Zero(rsEC!SSSE) > 0 Then
'                EC = EC + 10
'            Else
'                EC = EC + 0
'            End If
'            rsEC.MoveNext
'        Wend
'    End If
'
'    Dim rsPAYROLL                                                     As ADODB.Recordset
'    Set rsPAYROLL = New ADODB.Recordset
'    Set rsPAYROLL = gconDMIS.Execute("SELECT (SUM(SSSR) + SUM(PAGIBIG) + SUM(PHILHEALTHR)) as SUMEMPLOYERCONTRI FROM HRMS_PAYROLL WHERE EMPNO = '" & EMPNO & "' AND EMPLEVEL = '" & EMPLEVEL & "' AND PAY_YEAR = " & YEAR)
'
'    GetSumEmployerContribution = 0
'    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
'        GetSumEmployerContribution = Round(N2Str2Zero(rsPAYROLL!SUMEMPLOYERCONTRI), 2)
'    End If
'
'    GetSumEmployerContribution = GetSumEmployerContribution + EC
'    GetSumEmployerContribution = Round(GetSumEmployerContribution, 2)
'
'    Set rsPAYROLL = Nothing
    
    
    Set rsEC = gconDMIS.Execute("SELECT SSSE FROM HRMS_PAYROLL WHERE EMPNO = '" & EMPNO & "' AND EMPLEVEL = '" & EMPLEVEL & "' AND PAY_YEAR = " & YEAR)
    EC = 0
    If Not rsEC.EOF And Not rsEC.BOF Then
        rsEC.MoveFirst
        While Not rsEC.EOF
            If N2Str2Zero(rsEC!SSSE) >= 500# Then
                EC = EC + 30
            ElseIf N2Str2Zero(rsEC!SSSE) < 500# And N2Str2Zero(rsEC!SSSE) > 0 Then
                EC = EC + 10
            Else
                EC = EC + 0
            End If
            rsEC.MoveNext
        Wend
    End If
    
    Dim rsPAYROLL                                                     As ADODB.Recordset
    Set rsPAYROLL = New ADODB.Recordset
    Set rsPAYROLL = gconDMIS.Execute("SELECT (SUM(SSSE) + SUM(PAGIBIG) + SUM(PHILHEALTHE)) as SUMEMPLOYERCONTRI FROM HRMS_PAYROLL WHERE EMPNO = '" & EMPNO & "' AND EMPLEVEL = '" & EMPLEVEL & "' AND PAY_YEAR = " & YEAR)

    GetSumEmployerContribution = 0
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        GetSumEmployerContribution = Round(N2Str2Zero(rsPAYROLL!SUMEMPLOYERCONTRI), 2)
    End If
    
    GetSumEmployerContribution = GetSumEmployerContribution + EC
    GetSumEmployerContribution = Round(GetSumEmployerContribution, 2)
    
    Set rsPAYROLL = Nothing
       
End Function

Function GetSumPremiumContri(EMPNO As String, EMPLEVEL As String, YEAR As Integer) As Double
    Dim rsPAYROLL                                                     As ADODB.Recordset
    Set rsPAYROLL = New ADODB.Recordset
    Set rsPAYROLL = gconDMIS.Execute("SELECT (SUM(SSSE) + SUM(PAGIBIG) + SUM(PHILHEALTHE)) as SUMPREMIUMCONTRI FROM HRMS_PAYROLL WHERE EMPNO = '" & EMPNO & "' AND EMPLEVEL = '" & EMPLEVEL & "' AND PAY_YEAR = " & YEAR)

    GetSumPremiumContri = 0
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        GetSumPremiumContri = Round(N2Str2Zero(rsPAYROLL!SUMPREMIUMCONTRI), 2)
    End If
    Set rsPAYROLL = Nothing
End Function

Function GetSum13thMonthNonTaxable(EMPNO As String, EMPLEVEL As String, YEAR As Integer) As Double
    Dim SALARY As Double
    Dim MONTHS_ENTERED As Double
    SALARY = 0
    
    GetSum13thMonthNonTaxable = 0
      
    Dim rsPAYROLL                                                     As ADODB.Recordset
    Set rsPAYROLL = New ADODB.Recordset
    Set rsPAYROLL = gconDMIS.Execute("SELECT CUT_OFF FROM HRMS_PAYROLL WHERE EMPNO = '" & EMPNO & "' AND EMPLEVEL = '" & EMPLEVEL & "' AND PAY_YEAR = " & YEAR)

    MONTHS_ENTERED = 0
    
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        MONTHS_ENTERED = rsPAYROLL.RecordCount / 2
    End If
    
    Dim rsEmpInfo As ADODB.Recordset
    Set rsEmpInfo = New ADODB.Recordset
    Set rsEmpInfo = gconDMIS.Execute("SELECT BASICSALARY, PAYROLLTYPE FROM HRMS_EMPINFO WHERE EMPNO = '" & EMPNO & "'")
    
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        SALARY = N2Str2Zero(rsEmpInfo!BASICSALARY)
        If Null2String(rsEmpInfo!payrolltype) = "Monthly Base" Then
            MONTHS_ENTERED = MONTHS_ENTERED * 2
        End If
    End If
    
    GetSum13thMonthNonTaxable = SALARY * MONTHS_ENTERED
    
    GetSum13thMonthNonTaxable = Round(GetSum13thMonthNonTaxable / 12, 2)
    Set rsPAYROLL = Nothing
End Function

Function GetSumTaxJanNov(EMPNO As String, EMPLEVEL As String, YEAR As Integer) As Double
    Dim rsPAYROLL                                                     As ADODB.Recordset
    Set rsPAYROLL = New ADODB.Recordset
    Set rsPAYROLL = gconDMIS.Execute("SELECT SUM(TAX) as SUMTAXJANNOV FROM HRMS_PAYROLL WHERE EMPNO = '" & EMPNO & "' AND EMPLEVEL = '" & EMPLEVEL & "' AND PAY_YEAR = " & YEAR & " AND (PAY_MONTH <= 11)")

    GetSumTaxJanNov = 0
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        GetSumTaxJanNov = Round(N2Str2Zero(rsPAYROLL!SUMTAXJANNOV), 2)
    End If
    Set rsPAYROLL = Nothing
End Function

Function GetSumTaxJanDec(EMPNO As String, EMPLEVEL As String, YEAR As Integer) As Double
    Dim rsPAYROLL                                                     As ADODB.Recordset
    Set rsPAYROLL = New ADODB.Recordset
    Set rsPAYROLL = gconDMIS.Execute("SELECT SUM(TAX) as SUMTAXJANDEC FROM HRMS_PAYROLL WHERE EMPNO = '" & EMPNO & "' AND EMPLEVEL = '" & EMPLEVEL & "' AND PAY_YEAR = " & YEAR)

    GetSumTaxJanDec = 0
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        GetSumTaxJanDec = Round(N2Str2Zero(rsPAYROLL!SUMTAXJANDEC), 2)
    End If
    Set rsPAYROLL = Nothing
End Function

Function Personal_Ex2(STATUS As String) As Double
    Personal_Ex2 = 0
    If STATUS = "Z" Then
        Personal_Ex2 = 0
    ElseIf STATUS = "ME" Or STATUS = "S" Then
        Personal_Ex2 = 50000
    Else
        Personal_Ex2 = 50000
        If Mid(STATUS, 3, 1) = "1" Then
            Personal_Ex2 = Personal_Ex2 + 25000#
        ElseIf Mid(STATUS, 3, 1) = "2" Then
            Personal_Ex2 = Personal_Ex2 + 25000# * 2#
        ElseIf Mid(STATUS, 3, 1) = "3" Then
            Personal_Ex2 = Personal_Ex2 + 25000# * 3#
        ElseIf Mid(STATUS, 3, 1) = "4" Then
            Personal_Ex2 = Personal_Ex2 + 25000# * 4#
        End If
    End If
    Personal_Ex2 = Round(Personal_Ex2, 2)
End Function

Function ComputeTaxDue(AMOUNT As Double, Optional YearHired As Integer) As Double
    ComputeTaxDue = 0
    
    If AMOUNT > 0 And AMOUNT <= 10000 Then
        ComputeTaxDue = AMOUNT * 0.05
    ElseIf AMOUNT > 10000 And AMOUNT <= 30000 Then
        ComputeTaxDue = 500# + (AMOUNT - 10000#) * 0.1
    ElseIf AMOUNT > 30000 And AMOUNT <= 70000 Then
        ComputeTaxDue = 2500# + (AMOUNT - 30000#) * 0.15
    ElseIf AMOUNT > 70000 And AMOUNT <= 140000 Then
        ComputeTaxDue = 8500# + (AMOUNT - 70000#) * 0.2
    ElseIf AMOUNT > 140000 And AMOUNT <= 250000 Then
        ComputeTaxDue = 22500# + (AMOUNT - 140000#) * 0.25
    ElseIf AMOUNT > 25000 And AMOUNT <= 500000 Then
        ComputeTaxDue = 50000# + (AMOUNT - 250000#) * 0.3
    ElseIf AMOUNT > 500000 Then
        If YearHired >= 2001 Then
            ComputeTaxDue = 125000# + (AMOUNT - 500000#) * 0.32
        ElseIf YearHired < 2000 Then
            ComputeTaxDue = 125000# + (AMOUNT - 500000#) * 0.33
        Else
            ComputeTaxDue = 125000# + (AMOUNT - 500000#) * 0.34
        End If
    End If
    
    ComputeTaxDue = ComputeTaxDue
    ComputeTaxDue = Round(ComputeTaxDue, 2)
End Function

Sub PrintALPHAExcel(STATEMENT As String, REPORT As String, CHOICE As Integer)
    Dim xlApp                           As Excel.Application
    Dim xlsheet                         As Excel.Worksheet
    Dim xlbook                          As Excel.Workbook
    
    Dim PHIC_NO As String
    Dim TIN_NO As String
    Dim SSS_NO  As String
    Dim COMP_NAME  As String
    Dim COMP_ADDRESS  As String
    Dim COMP_TELEPHONE As String
    Dim PREPARED_BY As String
    Dim CHECKED_BY As String
    Dim APPROVED_BY As String

    PHIC_NO = ""
    TIN_NO = ""
    SSS_NO = ""
    COMP_NAME = ""
    COMP_ADDRESS = ""
    COMP_TELEPHONE = ""
    PREPARED_BY = ""
    CHECKED_BY = ""
    APPROVED_BY = ""
    
    Dim RS_HEADER                       As ADODB.Recordset
    Set RS_HEADER = gconDMIS.Execute("SELECT * FROM ALL_PROFILE WHERE MODULENAME='HRMS'")
    If Not (RS_HEADER.EOF And Not RS_HEADER.BOF) Then
        PHIC_NO = Null2String(RS_HEADER!CompanyPHICNo)
        TIN_NO = Null2String(RS_HEADER!companytinno)
        SSS_NO = Null2String(RS_HEADER!companysssno)
        COMP_NAME = Null2String(RS_HEADER!CompanyName)
        COMP_ADDRESS = Null2String(RS_HEADER!Companyaddress)
        COMP_TELEPHONE = Null2String(RS_HEADER!Companyaddress)
        PREPARED_BY = Null2String(RS_HEADER!PREPAREDBY)
        CHECKED_BY = Null2String(RS_HEADER!CHECKEDBY)
        APPROVED_BY = Null2String(RS_HEADER!APPROVEDBY)
    End If
    
    Set xlApp = New Excel.Application
    Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & REPORT)
    Set xlsheet = xlbook.Worksheets(1)
    
    xlsheet.Cells(1, "A") = "" & COMP_NAME & ""
    xlsheet.Cells(2, "A") = "" & COMP_ADDRESS & ""
    xlsheet.Cells(3, "A") = "T.I.N. " & TIN_NO & ""
    
    Dim I As Integer
    Dim j As Integer
    I = 0
    
    Dim rsEmpInfo As ADODB.Recordset
    Set rsEmpInfo = New ADODB.Recordset
    Set rsEmpInfo = gconDMIS.Execute(STATEMENT)
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        While Not rsEmpInfo.EOF
            xlsheet.Cells(14 + I, "A") = I + 1
            If Null2String(rsEmpInfo!tinno) = "" Then
                xlsheet.Cells(14 + I, "B").AddComment ("Please input the Valid Tin Number")
            Else
                xlsheet.Cells(14 + I, "B") = Null2String(rsEmpInfo!tinno)
            End If
            xlsheet.Cells(14 + I, "C") = Null2String(rsEmpInfo!lastname)
            xlsheet.Cells(14 + I, "D") = Null2String(rsEmpInfo!FIRSTNAME)
            xlsheet.Cells(14 + I, "E") = Left(Null2String(rsEmpInfo!MIDDLENAME), 1)
            If TAGGED <> 2 Then
                xlsheet.Cells(14 + I, "F") = GetSum13thMonthNonTaxable(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), cboyear)
                If xlsheet.Cells(14 + I, "F") > 30000 Then
                    xlsheet.Cells(14 + I, "F") = 30000
                End If
                
                xlsheet.Cells(14 + I, "G") = GetSumEmployerContribution(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), cboyear)
                
                xlsheet.Cells(14 + I, "H") = 0
                If (GetSum13thMonthNonTaxable(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), cboyear)) > 30000 Then
                    xlsheet.Cells(14 + I, "I") = (GetSum13thMonthNonTaxable(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), cboyear)) - 30000
                Else
                    xlsheet.Cells(14 + I, "I") = 0
                End If
                xlsheet.Cells(14 + I, "J") = GetSumGross(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), cboyear) - GetSumPremiumContri(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), cboyear)
                xlsheet.Cells(14 + I, "K") = Personal_Ex2(Null2String(rsEmpInfo!EXSTATUS))
                xlsheet.Cells(14 + I, "L") = ComputeTaxDue(xlsheet.Cells(14 + I, "I") + xlsheet.Cells(14 + I, "J") - xlsheet.Cells(14 + I, "K"), YEAR(Null2String(rsEmpInfo!DateHired)))
                xlsheet.Cells(14 + I, "M") = GetSumTaxJanNov(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), cboyear)
                xlsheet.Cells(14 + I, "N") = GetSumTaxJanDec(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), cboyear)
                xlsheet.Cells(14 + I, "O") = xlsheet.Cells(14 + I, "L") - xlsheet.Cells(14 + I, "M")
                If xlsheet.Cells(14 + I, "O") < 0 Then
                    xlsheet.Cells(14 + I, "O") = 0
                End If
                xlsheet.Cells(14 + I, "P") = xlsheet.Cells(14 + I, "M") - xlsheet.Cells(14 + I, "L")
                If xlsheet.Cells(14 + I, "P") < 0 Then
                    xlsheet.Cells(14 + I, "P") = 0
                End If
                xlsheet.Cells(14 + I, "Q") = 0
            Else
                xlsheet.Cells(14 + I, "F") = Get_NONTAX13THMONTH_PREVEMP(Null2String(rsEmpInfo!EMPNO), cboyear)
                xlsheet.Cells(14 + I, "G") = Get_NONTAXPREMIUM_PREVEMP(Null2String(rsEmpInfo!EMPNO), cboyear)
                xlsheet.Cells(14 + I, "H") = Get_NONTAXSALARIES_PREVEMP(Null2String(rsEmpInfo!EMPNO), cboyear)
                xlsheet.Cells(14 + I, "I") = Get_TAX13THMONTH_PREVEMP(Null2String(rsEmpInfo!EMPNO), cboyear)
                xlsheet.Cells(14 + I, "J") = Get_TAXSALARIES_PREVEMP(Null2String(rsEmpInfo!EMPNO), cboyear)
                xlsheet.Cells(14 + I, "K") = xlsheet.Cells(14 + I, "I") + xlsheet.Cells(14 + I, "J")
                xlsheet.Cells(14 + I, "L") = GetSum13thMonthNonTaxable(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), cboyear)
                If xlsheet.Cells(14 + I, "L") > 30000 Then
                    xlsheet.Cells(14 + I, "L") = 30000
                End If
                xlsheet.Cells(14 + I, "M") = GetSumEmployerContribution(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), cboyear)
                xlsheet.Cells(14 + I, "N") = 0
                If (GetSum13thMonthNonTaxable(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), cboyear)) > 30000 Then
                    xlsheet.Cells(14 + I, "O") = (GetSum13thMonthNonTaxable(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), cboyear)) - 30000
                Else
                    xlsheet.Cells(14 + I, "O") = 0
                End If
                xlsheet.Cells(14 + I, "P") = GetSumGross(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), cboyear) - GetSumPremiumContri(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), cboyear)
                xlsheet.Cells(14 + I, "Q") = xlsheet.Cells(14 + I, "K") + xlsheet.Cells(14 + I, "O") + xlsheet.Cells(14 + I, "P")
                xlsheet.Cells(14 + I, "R") = Personal_Ex2(Null2String(rsEmpInfo!EXSTATUS))
                xlsheet.Cells(14 + I, "S") = 0
                xlsheet.Cells(14 + I, "T") = ComputeTaxDue(xlsheet.Cells(14 + I, "Q") - xlsheet.Cells(14 + I, "R"))
                xlsheet.Cells(14 + I, "U") = Get_TAXWITHHELD_PREVEMP(Null2String(rsEmpInfo!EMPNO), cboyear)
                xlsheet.Cells(14 + I, "V") = GetSumTaxJanDec(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), cboyear)
                xlsheet.Cells(14 + I, "W") = xlsheet.Cells(14 + I, "T") - (xlsheet.Cells(14 + I, "U") + xlsheet.Cells(14 + I, "V"))
                If xlsheet.Cells(14 + I, "W") < 0 Then
                    xlsheet.Cells(14 + I, "W") = 0
                End If
                xlsheet.Cells(14 + I, "X") = (xlsheet.Cells(14 + I, "U") + xlsheet.Cells(14 + I, "V")) - xlsheet.Cells(14 + I, "T")
                If xlsheet.Cells(14 + I, "X") < 0 Then
                    xlsheet.Cells(14 + I, "X") = 0
                End If
                xlsheet.Cells(14 + I, "Y") = 0
            End If
            I = I + 1
            rsEmpInfo.MoveNext
        Wend
    End If
    
    If TAGGED <> 2 Then
        xlsheet.Cells(14 + I + 1, "A") = "(1)"
        xlsheet.Cells(14 + I + 1, "B") = "(2)"
        xlsheet.Cells(14 + I + 1, "C") = "(3a)"
        xlsheet.Cells(14 + I + 1, "D") = "(3b)"
        xlsheet.Cells(14 + I + 1, "E") = "(3c)"
        xlsheet.Cells(14 + I + 1, "F") = "'(4a)"
        xlsheet.Cells(14 + I + 1, "G") = "'(4b)"
        xlsheet.Cells(14 + I + 1, "H") = "'(4c)"
        xlsheet.Cells(14 + I + 1, "I") = "'(4d)"
        xlsheet.Cells(14 + I + 1, "J") = "'(4e)"
        xlsheet.Cells(14 + I + 1, "L") = "'(7)"
        xlsheet.Cells(14 + I + 1, "M") = "'(8)"
        xlsheet.Cells(14 + I + 1, "O") = "'(9a)=(7)-(8)"
        xlsheet.Cells(14 + I + 1, "P") = "'(9b)=(8)-(7)"
    Else
        xlsheet.Cells(14 + I + 1, "A") = "(1)"
        xlsheet.Cells(14 + I + 1, "B") = "(2)"
        xlsheet.Cells(14 + I + 1, "C") = "(3a)"
        xlsheet.Cells(14 + I + 1, "D") = "(3b)"
        xlsheet.Cells(14 + I + 1, "E") = "(3c)"
        xlsheet.Cells(14 + I + 1, "F") = "'(4a)"
        xlsheet.Cells(14 + I + 1, "G") = "'(4b)"
        xlsheet.Cells(14 + I + 1, "H") = "'(4c)"
        xlsheet.Cells(14 + I + 1, "I") = "'(4d)"
        xlsheet.Cells(14 + I + 1, "J") = "'(4e)"
        xlsheet.Cells(14 + I + 1, "K") = "'(4f = 4d + 4e)"
        xlsheet.Cells(14 + I + 1, "L") = "'(4g)"
        xlsheet.Cells(14 + I + 1, "M") = "'(4h)"
        xlsheet.Cells(14 + I + 1, "N") = "'(4i)"
        xlsheet.Cells(14 + I + 1, "O") = "'(4j)"
        xlsheet.Cells(14 + I + 1, "P") = "'(4k)"
        xlsheet.Cells(14 + I + 1, "Q") = "'(4l = 4f + 4j + 4k)"
        xlsheet.Cells(14 + I + 1, "R") = "'(5)"
        xlsheet.Cells(14 + I + 1, "S") = "'(6)"
        xlsheet.Cells(14 + I + 1, "T") = "'(7)"
        xlsheet.Cells(14 + I + 1, "U") = "(8a)"
        xlsheet.Cells(14 + I + 1, "V") = "(8b)"
        xlsheet.Cells(14 + I + 1, "W") = "'(9a) = (7)-(8a + 8b)"
        xlsheet.Cells(14 + I + 1, "X") = "'(9b)=(8a + 8b)-(7)"
        xlsheet.Cells(14 + I + 1, "Y") = "'(10)=(8b + 9a) or (8b - 9b)"
    End If
    
    xlsheet.Cells(14 + I + 2, "F").Formula = "=SUM(F14:" & "F" & 13 + I & ")"
    xlsheet.Cells(14 + I + 2, "G").Formula = "=SUM(G14:" & "G" & 13 + I & ")"
    xlsheet.Cells(14 + I + 2, "H").Formula = "=SUM(H14:" & "H" & 13 + I & ")"
    xlsheet.Cells(14 + I + 2, "I").Formula = "=SUM(I14:" & "I" & 13 + I & ")"
    xlsheet.Cells(14 + I + 2, "J").Formula = "=SUM(J14:" & "J" & 13 + I & ")"
    xlsheet.Cells(14 + I + 2, "K").Formula = "=SUM(K14:" & "K" & 13 + I & ")"
    xlsheet.Cells(14 + I + 2, "L").Formula = "=SUM(L14:" & "L" & 13 + I & ")"
    xlsheet.Cells(14 + I + 2, "M").Formula = "=SUM(M14:" & "M" & 13 + I & ")"
    xlsheet.Cells(14 + I + 2, "N").Formula = "=SUM(N14:" & "N" & 13 + I & ")"
    xlsheet.Cells(14 + I + 2, "O").Formula = "=SUM(O14:" & "O" & 13 + I & ")"
    xlsheet.Cells(14 + I + 2, "P").Formula = "=SUM(P14:" & "P" & 13 + I & ")"
    xlsheet.Cells(14 + I + 2, "Q").Formula = "=SUM(Q14:" & "Q" & 13 + I & ")"
    
    If TAGGED = 2 Then
        xlsheet.Cells(14 + I + 2, "R").Formula = "=SUM(R14:" & "R" & 13 + I & ")"
        xlsheet.Cells(14 + I + 2, "S").Formula = "=SUM(S14:" & "S" & 13 + I & ")"
        xlsheet.Cells(14 + I + 2, "T").Formula = "=SUM(T14:" & "T" & 13 + I & ")"
        xlsheet.Cells(14 + I + 2, "U").Formula = "=SUM(U14:" & "U" & 13 + I & ")"
        xlsheet.Cells(14 + I + 2, "V").Formula = "=SUM(V14:" & "V" & 13 + I & ")"
        xlsheet.Cells(14 + I + 2, "W").Formula = "=SUM(W14:" & "W" & 13 + I & ")"
        xlsheet.Cells(14 + I + 2, "X").Formula = "=SUM(X14:" & "X" & 13 + I & ")"
        xlsheet.Cells(14 + I + 2, "Y").Formula = "=SUM(Y14:" & "Y" & 13 + I & ")"
    End If
    
    xlsheet.Cells(14 + I + 2, "B") = "TOTAL"
    xlsheet.Cells(14 + I + 7, "C") = "Prepared by:"
    xlsheet.Cells(14 + I + 7, "H") = "Certified Correct by:"
    xlsheet.Cells(14 + I + 12, "C") = "Admin. Manager"
    xlsheet.Cells(14 + I + 12, "H") = "Asst. Gen. Manager"
    
    For I = 1 To xlsheet.Comments.count
        xlsheet.Comments(I).Visible = True
    Next

    xlApp.Visible = True
    Set xlsheet = Nothing
    Set xlbook = Nothing
    Set xlApp = Nothing
    
    Set rsEmpInfo = Nothing
    Set RS_HEADER = Nothing
End Sub

Function Get_NONTAX13THMONTH_PREVEMP(EMPNO As String, YEAR As Integer) As Double
    Get_NONTAX13THMONTH_PREVEMP = 0
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT NONTAX13THMONTH FROM HRMS_PREVEMP WHERE PREVEMPYEAR ='" & YEAR & "' AND EMPNO = '" & EMPNO & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        Get_NONTAX13THMONTH_PREVEMP = N2Str2Zero(rsTemp!NONTax13thMonth)
    End If
    Get_NONTAX13THMONTH_PREVEMP = Round(Get_NONTAX13THMONTH_PREVEMP, 2)
    Set rsTemp = Nothing
End Function

Function Get_NONTAXPREMIUM_PREVEMP(EMPNO As String, YEAR As Integer) As Double
    Get_NONTAXPREMIUM_PREVEMP = 0
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT NONTAXPREMIUM FROM HRMS_PREVEMP WHERE PREVEMPYEAR ='" & YEAR & "' AND EMPNO = '" & EMPNO & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        Get_NONTAXPREMIUM_PREVEMP = N2Str2Zero(rsTemp!NONTaxPremium)
    End If
    Get_NONTAXPREMIUM_PREVEMP = Round(Get_NONTAXPREMIUM_PREVEMP, 2)
    Set rsTemp = Nothing
End Function

Function Get_NONTAXSALARIES_PREVEMP(EMPNO As String, YEAR As Integer) As Double
    Get_NONTAXSALARIES_PREVEMP = 0
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT NONTAXSALARIES FROM HRMS_PREVEMP WHERE PREVEMPYEAR ='" & YEAR & "' AND EMPNO = '" & EMPNO & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        Get_NONTAXSALARIES_PREVEMP = N2Str2Zero(rsTemp!NONTaxSalaries)
    End If
    Get_NONTAXSALARIES_PREVEMP = Round(Get_NONTAXSALARIES_PREVEMP, 2)
    Set rsTemp = Nothing
End Function

Function Get_TAX13THMONTH_PREVEMP(EMPNO As String, YEAR As Integer) As Double
    Get_TAX13THMONTH_PREVEMP = 0
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT TAX13THMONTH FROM HRMS_PREVEMP WHERE PREVEMPYEAR ='" & YEAR & "' AND EMPNO = '" & EMPNO & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        Get_TAX13THMONTH_PREVEMP = N2Str2Zero(rsTemp!Tax13thMonth)
    End If
    Get_TAX13THMONTH_PREVEMP = Round(Get_TAX13THMONTH_PREVEMP, 2)
    Set rsTemp = Nothing
End Function

Function Get_TAXSALARIES_PREVEMP(EMPNO As String, YEAR As Integer) As Double
Get_TAXSALARIES_PREVEMP = 0
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT TAXSALARIES FROM HRMS_PREVEMP WHERE PREVEMPYEAR ='" & YEAR & "' AND EMPNO = '" & EMPNO & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        Get_TAXSALARIES_PREVEMP = N2Str2Zero(rsTemp!TaxSalaries)
    End If
    Get_TAXSALARIES_PREVEMP = Round(Get_TAXSALARIES_PREVEMP, 2)
    Set rsTemp = Nothing
End Function

Function Get_TAXWITHHELD_PREVEMP(EMPNO As String, YEAR As Integer) As Double
    Get_TAXWITHHELD_PREVEMP = 0
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT TAXWITHHELD FROM HRMS_PREVEMP WHERE PREVEMPYEAR ='" & YEAR & "' AND EMPNO = '" & EMPNO & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        Get_TAXWITHHELD_PREVEMP = N2Str2Zero(rsTemp!TaxWithheld)
    End If
    Get_TAXWITHHELD_PREVEMP = Round(Get_TAXWITHHELD_PREVEMP, 2)
    Set rsTemp = Nothing
End Function

Function Get_TOTALTAXABLE_PREVEMP(EMPNO As String, YEAR As Integer) As Double
    Get_TOTALTAXABLE_PREVEMP = 0
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT TOTALTAXABLE FROM HRMS_PREVEMP WHERE PREVEMPYEAR ='" & YEAR & "' AND EMPNO = '" & EMPNO & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        Get_TOTALTAXABLE_PREVEMP = N2Str2Zero(rsTemp!TOTALTAXABLE)
    End If
    Get_TOTALTAXABLE_PREVEMP = Round(Get_TOTALTAXABLE_PREVEMP, 2)
    Set rsTemp = Nothing
End Function

Sub Print_Excel_SCHEDDEDUCTION(YEAR As Integer)
    Dim I                               As Integer
    I = 0
    
    Dim xlApp                           As Excel.Application
    Dim xlsheet                         As Excel.Worksheet
    Dim xlbook                          As Excel.Workbook
    
    Set xlApp = New Excel.Application
    Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & "NEW.XLT")
    Set xlsheet = xlbook.Worksheets(1)
    
    Dim rsHeader As ADODB.Recordset
    Set rsHeader = New ADODB.Recordset
    Set rsHeader = gconDMIS.Execute("SELECT * FROM ALL_PROFILE WHERE MODULENAME = 'HRMS'")
    
    If Not rsHeader.EOF And Not rsHeader.BOF Then
        xlsheet.Cells(2, "B") = Null2String(rsHeader!CompanyName)
        xlsheet.Cells(3, "B") = Null2String(rsHeader!Companyaddress)
    End If
    xlsheet.Cells(4, "B") = "PAY YEAR " & YEAR
    xlsheet.Cells(5, "B") = "DEDUCTION DETAIL REPORT"
    
    Dim rsPayroll_Det As ADODB.Recordset
    Set rsPayroll_Det = New ADODB.Recordset
    Set rsPayroll_Det = gconDMIS.Execute("SELECT DISTINCT DET_CODE FROM HRMS_PAYROLL_DET WHERE PAY_YEAR = '" & YEAR & "'")
    
    Dim rsPAYROLL As ADODB.Recordset
    Set rsPAYROLL = New ADODB.Recordset
    Set rsPAYROLL = gconDMIS.Execute("SELECT EMPNO, LASTNAME + ', ' + FIRSTNAME AS FULLNAME FROM HRMS_EMPINFO WHERE EMPNO IN (SELECT DISTINCT EMPNO FROM HRMS_PAYROLL WHERE PAY_YEAR = '" & YEAR & "') ORDER BY LASTNAME + ', ' + FIRSTNAME")
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        rsPAYROLL.MoveFirst
        While Not rsPAYROLL.EOF
            xlsheet.Cells((I * 16) + 7, GetLetter(1)) = Null2String(rsPAYROLL!FULLNAME)
            xlsheet.Cells((I * 16) + 8, GetLetter(1)) = "MONTH\DEDUCTION CODE"
            Dim COLUMN_DET As Integer
            COLUMN_DET = 1
            If Not rsPayroll_Det.EOF Or Not rsPayroll_Det.BOF Then
                rsPayroll_Det.MoveFirst
                While Not rsPayroll_Det.EOF
                    xlsheet.Cells((I * 16) + 8, GetLetter(1 + COLUMN_DET)) = Null2String(rsPayroll_Det!DET_CODE)
                    xlsheet.Cells((I * 16) + 9, GetLetter(1 + COLUMN_DET)) = GET_AMOUNT(xlsheet.Cells((I * 16) + 8, GetLetter(1 + COLUMN_DET)), 1, cboyear, Null2String(rsPAYROLL!EMPNO))
                    xlsheet.Cells((I * 16) + 10, GetLetter(1 + COLUMN_DET)) = GET_AMOUNT(xlsheet.Cells((I * 16) + 8, GetLetter(1 + COLUMN_DET)), 2, cboyear, Null2String(rsPAYROLL!EMPNO))
                    xlsheet.Cells((I * 16) + 11, GetLetter(1 + COLUMN_DET)) = GET_AMOUNT(xlsheet.Cells((I * 16) + 8, GetLetter(1 + COLUMN_DET)), 3, cboyear, Null2String(rsPAYROLL!EMPNO))
                    xlsheet.Cells((I * 16) + 12, GetLetter(1 + COLUMN_DET)) = GET_AMOUNT(xlsheet.Cells((I * 16) + 8, GetLetter(1 + COLUMN_DET)), 4, cboyear, Null2String(rsPAYROLL!EMPNO))
                    xlsheet.Cells((I * 16) + 13, GetLetter(1 + COLUMN_DET)) = GET_AMOUNT(xlsheet.Cells((I * 16) + 8, GetLetter(1 + COLUMN_DET)), 5, cboyear, Null2String(rsPAYROLL!EMPNO))
                    xlsheet.Cells((I * 16) + 14, GetLetter(1 + COLUMN_DET)) = GET_AMOUNT(xlsheet.Cells((I * 16) + 8, GetLetter(1 + COLUMN_DET)), 6, cboyear, Null2String(rsPAYROLL!EMPNO))
                    xlsheet.Cells((I * 16) + 15, GetLetter(1 + COLUMN_DET)) = GET_AMOUNT(xlsheet.Cells((I * 16) + 8, GetLetter(1 + COLUMN_DET)), 7, cboyear, Null2String(rsPAYROLL!EMPNO))
                    xlsheet.Cells((I * 16) + 16, GetLetter(1 + COLUMN_DET)) = GET_AMOUNT(xlsheet.Cells((I * 16) + 8, GetLetter(1 + COLUMN_DET)), 8, cboyear, Null2String(rsPAYROLL!EMPNO))
                    xlsheet.Cells((I * 16) + 17, GetLetter(1 + COLUMN_DET)) = GET_AMOUNT(xlsheet.Cells((I * 16) + 8, GetLetter(1 + COLUMN_DET)), 9, cboyear, Null2String(rsPAYROLL!EMPNO))
                    xlsheet.Cells((I * 16) + 18, GetLetter(1 + COLUMN_DET)) = GET_AMOUNT(xlsheet.Cells((I * 16) + 8, GetLetter(1 + COLUMN_DET)), 10, cboyear, Null2String(rsPAYROLL!EMPNO))
                    xlsheet.Cells((I * 16) + 19, GetLetter(1 + COLUMN_DET)) = GET_AMOUNT(xlsheet.Cells((I * 16) + 8, GetLetter(1 + COLUMN_DET)), 11, cboyear, Null2String(rsPAYROLL!EMPNO))
                    xlsheet.Cells((I * 16) + 20, GetLetter(1 + COLUMN_DET)) = GET_AMOUNT(xlsheet.Cells((I * 16) + 8, GetLetter(1 + COLUMN_DET)), 12, cboyear, Null2String(rsPAYROLL!EMPNO))
                    xlsheet.Cells((I * 16) + 21, GetLetter(1 + COLUMN_DET)).Formula = "= SUM(" & GetLetter(1 + COLUMN_DET) & ((I * 16) + 9) & ":" & GetLetter(1 + COLUMN_DET) & ((I * 16) + 20) & ")"
                    
                    COLUMN_DET = COLUMN_DET + 1
                    rsPayroll_Det.MoveNext
                Wend
            End If
            xlsheet.Cells((I * 16) + 9, "A") = Left(MonthName(1), 3)
            xlsheet.Cells((I * 16) + 10, "A") = Left(MonthName(2), 3)
            xlsheet.Cells((I * 16) + 11, "A") = Left(MonthName(3), 3)
            xlsheet.Cells((I * 16) + 12, "A") = Left(MonthName(4), 3)
            xlsheet.Cells((I * 16) + 13, "A") = Left(MonthName(5), 3)
            xlsheet.Cells((I * 16) + 14, "A") = Left(MonthName(6), 3)
            xlsheet.Cells((I * 16) + 15, "A") = Left(MonthName(7), 3)
            xlsheet.Cells((I * 16) + 16, "A") = Left(MonthName(8), 3)
            xlsheet.Cells((I * 16) + 17, "A") = Left(MonthName(9), 3)
            xlsheet.Cells((I * 16) + 18, "A") = Left(MonthName(10), 3)
            xlsheet.Cells((I * 16) + 19, "A") = Left(MonthName(11), 3)
            xlsheet.Cells((I * 16) + 20, "A") = Left(MonthName(12), 3)
            xlsheet.Cells((I * 16) + 21, "A") = "TOTAL"
            I = I + 1
            rsPAYROLL.MoveNext
        Wend
    End If
        
    xlApp.Visible = True
    Set xlsheet = Nothing
    Set xlbook = Nothing
    Set xlApp = Nothing
    
    Set rsPAYROLL = Nothing
    Set rsPayroll_Det = Nothing
End Sub

Function GET_NAME(EMPNO As String) As String
    GET_NAME = ""
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT LASTNAME + ', ' + FIRSTNAME AS FULLNAME FROM HRMS_EMPINFO WHERE EMPNO = '" & EMPNO & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        GET_NAME = Null2String(rsTemp!FULLNAME)
    End If
    Set rsTemp = Nothing
End Function

Function GetLetter(NUMBER As Integer) As String
    GetLetter = ""
    If NUMBER = 1 Then
        GetLetter = "A"
    ElseIf NUMBER = 2 Then
        GetLetter = "B"
    ElseIf NUMBER = 3 Then
        GetLetter = "C"
    ElseIf NUMBER = 4 Then
        GetLetter = "D"
    ElseIf NUMBER = 5 Then
        GetLetter = "E"
    ElseIf NUMBER = 6 Then
        GetLetter = "F"
    ElseIf NUMBER = 7 Then
        GetLetter = "G"
    ElseIf NUMBER = 8 Then
        GetLetter = "H"
    ElseIf NUMBER = 9 Then
        GetLetter = "I"
    ElseIf NUMBER = 10 Then
        GetLetter = "J"
    ElseIf NUMBER = 11 Then
        GetLetter = "K"
    ElseIf NUMBER = 12 Then
        GetLetter = "L"
    ElseIf NUMBER = 13 Then
        GetLetter = "M"
    ElseIf NUMBER = 14 Then
        GetLetter = "N"
    ElseIf NUMBER = 15 Then
        GetLetter = "O"
    ElseIf NUMBER = 16 Then
        GetLetter = "P"
    ElseIf NUMBER = 17 Then
        GetLetter = "Q"
    ElseIf NUMBER = 18 Then
        GetLetter = "R"
    ElseIf NUMBER = 19 Then
        GetLetter = "S"
    ElseIf NUMBER = 20 Then
        GetLetter = "T"
    ElseIf NUMBER = 21 Then
        GetLetter = "U"
    ElseIf NUMBER = 22 Then
        GetLetter = "V"
    ElseIf NUMBER = 23 Then
        GetLetter = "W"
    ElseIf NUMBER = 24 Then
        GetLetter = "X"
    ElseIf NUMBER = 25 Then
        GetLetter = "Y"
    ElseIf NUMBER = 26 Then
        GetLetter = "Z"
    End If
End Function

Function GET_AMOUNT(DED_CODE As String, MONTH As Integer, YEAR As Integer, EMPNO As String) As Double
    GET_AMOUNT = 0
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT SUM(DET_AMOUNT) AS AMOUNT FROM HRMS_PAYROLL_DET WHERE EMPNO = '" & EMPNO & "' AND PAY_MONTH = '" & MONTH & "' AND PAY_YEAR = '" & YEAR & "' AND DET_CODE = '" & DED_CODE & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        GET_AMOUNT = N2Str2Zero(rsTemp!AMOUNT)
    End If
    
    GET_AMOUNT = Round(GET_AMOUNT, 2)
    Set rsTemp = Nothing
End Function

