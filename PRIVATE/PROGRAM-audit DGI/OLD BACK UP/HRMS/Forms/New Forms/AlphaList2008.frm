VERSION 5.00
Begin VB.Form frmHRMS_AlphaList2008 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alpha List Report 2008"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   Icon            =   "AlphaList2008.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   4785
   Begin VB.CommandButton cmdWithNoPRevJulDec 
      Caption         =   "Alphalist of Employees With No Previous Employer (Jul-Dec)"
      Height          =   525
      Left            =   120
      TabIndex        =   5
      Top             =   2700
      Width           =   4545
   End
   Begin VB.CommandButton cmdWithNoPrevJanJun 
      Caption         =   "Alphalist of Employees With No Previous Employer (Jan-Jun)"
      Height          =   525
      Left            =   120
      TabIndex        =   4
      Top             =   2190
      Width           =   4545
   End
   Begin VB.CommandButton cmdWithPrevJulDec 
      Caption         =   "Alphalist of Employees With Previous Employer (Jul-Dec)"
      Height          =   525
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4545
   End
   Begin VB.CommandButton cmdWithPrevJanJun 
      Caption         =   "Alphalist of Employees With Previous Employer (Jan-Jun)"
      Height          =   525
      Left            =   120
      TabIndex        =   2
      Top             =   1170
      Width           =   4545
   End
   Begin VB.CommandButton cmdTerminatedJulDec 
      Caption         =   "Alphalist of Terminated Employees (Jul-Dec)"
      Height          =   525
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   4545
   End
   Begin VB.CommandButton cmdTerminatedJanJun 
      Caption         =   "Alphalist of Terminated Employees (Jan-Jun)"
      Height          =   525
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   4545
   End
End
Attribute VB_Name = "frmHRMS_AlphaList2008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SQL As String
Dim REPORT As String
Dim MONTHFROM As Integer
Dim MONTHTO As Integer
Dim TAGGED As Integer

Sub PrintALPHAExcel(STATEMENT As String, REPORT As String, PERIODFROM As Integer, PERIODTO As Integer)
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
    
    Dim DISPLAY_CAPTION1 As String
    Dim DISPLAY_CAPTION2 As String
    
    If PERIODFROM = 1 Then
        DISPLAY_CAPTION1 = "(Jan.-May.)"
        DISPLAY_CAPTION2 = "(Jan.-June)"
    Else
        DISPLAY_CAPTION1 = "(Jul.-Dec.)"
        DISPLAY_CAPTION2 = "(Jul.-Nov)"
    End If
    
    xlsheet.Cells(9, "L") = DISPLAY_CAPTION1
    xlsheet.Cells(10, "M") = DISPLAY_CAPTION1
    xlsheet.Cells(10, "N") = DISPLAY_CAPTION2
    
    Dim rsEmpInfo As ADODB.Recordset
    Set rsEmpInfo = New ADODB.Recordset
    Set rsEmpInfo = gconDMIS.Execute(STATEMENT)
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        While Not rsEmpInfo.EOF
        
            xlsheet.Cells(14 + I, "A") = I + 1
            xlsheet.Cells(14 + I, "B") = Null2String(rsEmpInfo!tinno)
            xlsheet.Cells(14 + I, "C") = Null2String(rsEmpInfo!lastname)
            xlsheet.Cells(14 + I, "D") = Null2String(rsEmpInfo!FIRSTNAME)
            xlsheet.Cells(14 + I, "E") = Left(Null2String(rsEmpInfo!MIDDLENAME), 1)
            xlsheet.Cells(14 + I, "F") = GetSum13thMonthNonTaxable(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), 2008, PERIODFROM, PERIODTO)
            xlsheet.Cells(14 + I, "G") = GetSumEmployerContribution(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), 2008, PERIODFROM, PERIODTO)
            xlsheet.Cells(14 + I, "H") = 0
            xlsheet.Cells(14 + I, "I").Formula = "=IF(F" & 14 + I & "> (30000/2), F" & 14 + I & "- (30000/2),0)"
            xlsheet.Cells(14 + I, "J") = GetSumGross(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), 2008, PERIODFROM, PERIODTO) - GetSumPremiumContri(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), 2008, PERIODFROM, PERIODTO)
            xlsheet.Cells(14 + I, "K") = Personal_Ex2(Null2String(rsEmpInfo!ExStatus), PERIODFROM) / 2
            
            xlsheet.Cells(14 + I, "L") = ComputeTaxDue(GetSumGross(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), 2008, PERIODFROM, PERIODTO) - (Personal_Ex2(Null2String(rsEmpInfo!ExStatus), PERIODFROM) / 2) - GetSumPremiumContri(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), 2008, PERIODFROM, PERIODTO))
                                           
            xlsheet.Cells(14 + I, "M") = GetSumTaxJanNov(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), 2008, PERIODFROM, PERIODTO - 1)
            xlsheet.Cells(14 + I, "N") = GetSumTaxJanDec(Null2String(rsEmpInfo!EMPNO), Null2String(rsEmpInfo!EMPLEVEL), 2008, PERIODFROM, PERIODTO)
            xlsheet.Cells(14 + I, "O").Formula = "=L" & 14 + I & "-" & "M" & 14 + I
            xlsheet.Cells(14 + I, "P").Formula = "=M" & 14 + I & "-" & "L" & 14 + I
            If TAGGED = 1 Then
                xlsheet.Cells(14 + I, "R") = GetResignedDate(Null2String(rsEmpInfo!EMPNO))
            End If
            I = I + 1
            rsEmpInfo.MoveNext
        Wend
    End If
    
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
    
    xlsheet.Cells(14 + I + 2, "B") = "TOTAL"
    xlsheet.Cells(14 + I + 7, "C") = "Prepared by:"
    xlsheet.Cells(14 + I + 7, "H") = "Certified Correct by:"
    xlsheet.Cells(14 + I + 12, "C") = "Admin. Manager"
    xlsheet.Cells(14 + I + 12, "H") = "Asst. Gen. Manager"
    
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
    
    
    xlApp.Visible = True
    Set xlsheet = Nothing
    Set xlbook = Nothing
    Set xlApp = Nothing
    
    Set rsEmpInfo = Nothing
    Set RS_HEADER = Nothing
    TAGGED = 0
End Sub
    
 Function GetSumPremiumContri(EMPNO As String, EMPLEVEL As String, YEAR As Integer, FROM_MONTH As Integer, TO_MONTH As Integer) As Double
    Dim rsPAYROLL As ADODB.Recordset
    Set rsPAYROLL = New ADODB.Recordset
    Set rsPAYROLL = gconDMIS.Execute("SELECT (SUM(SSSE) + SUM(PAGIBIG) + SUM(PHILHEALTHE)) as SUMPREMIUMCONTRI FROM HRMS_PAYROLL WHERE EMPNO = '" & EMPNO & "' AND EMPLEVEL = '" & EMPLEVEL & "' AND PAY_YEAR = " & YEAR & " AND (PAY_MONTH >= " & FROM_MONTH & " AND PAY_MONTH <= " & TO_MONTH & ")")
    
    GetSumPremiumContri = 0
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        GetSumPremiumContri = Round(N2Str2Zero(rsPAYROLL!SUMPREMIUMCONTRI), 2)
    End If
    Set rsPAYROLL = Nothing
End Function

Function GetSum13thMonthNonTaxable(EMPNO As String, EMPLEVEL As String, YEAR As Integer, FROM_MONTH As Integer, TO_MONTH As Integer) As Double
    Dim CUTOFFS_ENTERED As Double
    Dim SALARY As Double
    Dim rsEmpInfo As ADODB.Recordset
    Set rsEmpInfo = New ADODB.Recordset
    Set rsEmpInfo = gconDMIS.Execute("SELECT BASICSALARY FROM HRMS_EMPINFO WHERE EMPNO = '" & EMPNO & "'")
    
    Dim rsPAYROLL As ADODB.Recordset
    Set rsPAYROLL = New ADODB.Recordset
    Set rsPAYROLL = gconDMIS.Execute("SELECT CUT_OFF FROM HRMS_PAYROLL WHERE EMPNO = '" & EMPNO & "' AND PAY_YEAR = '2008' AND PAY_MONTH BETWEEN " & FROM_MONTH & " AND " & TO_MONTH)
    
    GetSum13thMonthNonTaxable = 0
    SALARY = 0
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        SALARY = N2Str2Zero(rsEmpInfo!BASICSALARY)
    End If
    CUTOFFS_ENTERED = 0
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        CUTOFFS_ENTERED = N2Str2Zero(rsPAYROLL.RecordCount) / 2
    End If
    GetSum13thMonthNonTaxable = (SALARY * CUTOFFS_ENTERED) / 6
    GetSum13thMonthNonTaxable = Round(GetSum13thMonthNonTaxable, 2)
    Set rsPAYROLL = Nothing
    Set rsEmpInfo = Nothing
End Function

Function GetSumGross(EMPNO As String, EMPLEVEL As String, YEAR As Integer, FROM_MONTH As Integer, TO_MONTH As Integer) As Double
    Dim rsPAYROLL As ADODB.Recordset
    Set rsPAYROLL = New ADODB.Recordset
    Set rsPAYROLL = gconDMIS.Execute("SELECT (SUM(RATE) - SUM(UNDERTIME) - SUM(ABSENT) + SUM(OVERTIME) + SUM(TAXABLEADJ)) as SUMGROSS FROM HRMS_PAYROLL WHERE EMPNO = '" & EMPNO & "' AND EMPLEVEL = '" & EMPLEVEL & "' AND PAY_YEAR = " & YEAR & " AND (PAY_MONTH >= " & FROM_MONTH & " AND PAY_MONTH <= " & TO_MONTH & ")")
    
    GetSumGross = 0
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        GetSumGross = Round(N2Str2Zero(rsPAYROLL!SUMGROSS), 2)
    End If
    Set rsPAYROLL = Nothing
End Function

Function GetSumTaxJanNov(EMPNO As String, EMPLEVEL As String, YEAR As Integer, FROM_MONTH As Integer, TO_MONTH As Integer) As Double
    Dim rsPAYROLL As ADODB.Recordset
    Set rsPAYROLL = New ADODB.Recordset
    Set rsPAYROLL = gconDMIS.Execute("SELECT SUM(TAX) as SUMTAXJANNOV FROM HRMS_PAYROLL WHERE EMPNO = '" & EMPNO & "' AND EMPLEVEL = '" & EMPLEVEL & "' AND PAY_YEAR = " & YEAR & " AND PAY_MONTH >= '" & FROM_MONTH & "' AND PAY_MONTH <= '" & TO_MONTH & "'")
    
    GetSumTaxJanNov = 0
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        GetSumTaxJanNov = Round(N2Str2Zero(rsPAYROLL!SUMTAXJANNOV), 2)
    End If
    Set rsPAYROLL = Nothing
End Function

Function GetSumTaxJanDec(EMPNO As String, EMPLEVEL As String, YEAR As Integer, FROM_MONTH As Integer, TO_MONTH As Integer) As Double
    Dim rsPAYROLL As ADODB.Recordset
    Set rsPAYROLL = New ADODB.Recordset
    Set rsPAYROLL = gconDMIS.Execute("SELECT SUM(TAX) as SUMTAXJANDEC FROM HRMS_PAYROLL WHERE EMPNO = '" & EMPNO & "' AND EMPLEVEL = '" & EMPLEVEL & "' AND PAY_YEAR = " & YEAR & " AND PAY_MONTH >= '" & FROM_MONTH & "' AND PAY_MONTH <= '" & TO_MONTH & "'")
    
    GetSumTaxJanDec = 0
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        GetSumTaxJanDec = Round(N2Str2Zero(rsPAYROLL!SUMTAXJANDEC), 2)
    End If
    Set rsPAYROLL = Nothing
End Function

Function ComputeTaxDue(AMOUNT As Double) As Double
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
         ComputeTaxDue = 125000# + (AMOUNT - 500000#) * 0.34
    
    End If
    ComputeTaxDue = ComputeTaxDue
    ComputeTaxDue = Round(ComputeTaxDue, 2)
    
End Function

Function Personal_Ex2(STATUS As String, FROM_MONTH As Integer) As Double
    Personal_Ex2 = 0
    If FROM_MONTH = 1 Then
        Personal_Ex2 = Personal_EX(STATUS)
    ElseIf FROM_MONTH = 7 Then
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
    End If
    Personal_Ex2 = Round(Personal_Ex2, 2)
End Function

Function GetSumEmployerContribution(EMPNO As String, EMPLEVEL As String, YEAR As Integer, FROM_MONTH As Integer, TO_MONTH As Integer) As Double
    Dim EC As Double
    Dim rsEC As ADODB.Recordset
    Set rsEC = New ADODB.Recordset
    Set rsEC = gconDMIS.Execute("SELECT SSSE FROM HRMS_PAYROLL WHERE EMPNO = '" & EMPNO & "' AND EMPLEVEL = '" & EMPLEVEL & "' AND PAY_YEAR = " & YEAR & " AND (PAY_MONTH >= " & FROM_MONTH & " AND PAY_MONTH <= " & TO_MONTH & ")")
    
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
    
    Dim rsPAYROLL As ADODB.Recordset
    Set rsPAYROLL = New ADODB.Recordset
    Set rsPAYROLL = gconDMIS.Execute("SELECT (SUM(SSSR) + SUM(PHILHEALTHR) + SUM(PAGIBIG)) as SUMEMPLOYERCONTRIBUTION FROM HRMS_PAYROLL WHERE EMPNO = '" & EMPNO & "' AND EMPLEVEL = '" & EMPLEVEL & "' AND PAY_YEAR = " & YEAR & " AND (PAY_MONTH >= " & FROM_MONTH & " AND PAY_MONTH <= " & TO_MONTH & ")")
    
    GetSumEmployerContribution = 0
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        GetSumEmployerContribution = Round(N2Str2Zero(rsPAYROLL!SUMEMPLOYERCONTRIBUTION), 2)
    End If
    
    GetSumEmployerContribution = Round(GetSumEmployerContribution + EC, 2)
    Set rsPAYROLL = Nothing
End Function

Function GetResignedDate(EMPNO As String) As String
    Dim rsEmpInfo As ADODB.Recordset
    Set rsEmpInfo = New ADODB.Recordset
    Set rsEmpInfo = gconDMIS.Execute("SELECT RESIGNED FROM HRMS_EMPINFO WHERE EMPNO = '" & EMPNO & "'")
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        GetResignedDate = Format(Null2String(rsEmpInfo!RESIGNED), "SHORT DATE")
    End If
    Set rsEmpInfo = Nothing
End Function
Private Sub cmdTerminatedJanJun_Click()
    SQL = "SELECT * FROM HRMS_EMPINFO WHERE YEAR(RESIGNED) = '2008'" & _
          " AND (EMPLEVEL = 'E' OR EMPLEVEL = 'M') ORDER BY LASTNAME"
    REPORT = "ALPHATERMINATED.XLT"
    MONTHFROM = 1
    MONTHTO = 6
    
    Call PrintALPHAExcel(SQL, REPORT, MONTHFROM, MONTHTO)
    SQL = ""
    MONTHFROM = 0
    MONTHTO = 0
    TAGGED = 1
End Sub

Private Sub cmdTerminatedJulDec_Click()
    SQL = "SELECT * FROM HRMS_EMPINFO WHERE YEAR(RESIGNED) = '2008'" & _
          " AND (EMPLEVEL = 'E' OR EMPLEVEL = 'M') AND MONTH(RESIGNED) > '6' ORDER BY LASTNAME"
    REPORT = "ALPHATERMINATED.XLT"
    MONTHFROM = 7
    MONTHTO = 12
    
    Call PrintALPHAExcel(SQL, REPORT, MONTHFROM, MONTHTO)
    SQL = ""
    MONTHFROM = 0
    MONTHTO = 0
    TAGGED = 1
End Sub

Private Sub cmdWithNoPrevJanJun_Click()
    SQL = "SELECT * FROM HRMS_EMPINFO WHERE (EMPLEVEL = 'E' OR EMPLEVEL = 'M') AND YEAR(DATEHIRED) <= '2008' AND RESIGNED IS NULL AND (PREVIOUSCOMPANY IS NULL OR YEAR(DATEHIRED) <> '2008') ORDER BY LASTNAME"
    REPORT = "ALPHAWITHNOEMP.XLT"
    MONTHFROM = 1
    MONTHTO = 6
    
    Call PrintALPHAExcel(SQL, REPORT, MONTHFROM, MONTHTO)
    SQL = ""
    MONTHFROM = 0
    MONTHTO = 0
End Sub

Private Sub cmdWithNoPRevJulDec_Click()
    SQL = "SELECT * FROM HRMS_EMPINFO WHERE (EMPLEVEL = 'E' OR EMPLEVEL = 'M') AND YEAR(DATEHIRED) <= '2008' AND RESIGNED IS NULL AND (PREVIOUSCOMPANY IS NULL OR YEAR(DATEHIRED) <> '2008') ORDER BY LASTNAME"
    REPORT = "ALPHAWITHNOEMP.XLT"
    MONTHFROM = 7
    MONTHTO = 12
    
    Call PrintALPHAExcel(SQL, REPORT, MONTHFROM, MONTHTO)
    SQL = ""
    MONTHFROM = 0
    MONTHTO = 0
End Sub

Private Sub cmdWithPrevJanJun_Click()
    SQL = "SELECT * FROM HRMS_EMPINFO WHERE YEAR(DATEHIRED) = '2008' AND (EMPLEVEL = 'E' OR EMPLEVEL = 'M') AND PREVIOUSCOMPANY IS NOT NULL AND YEAR(RESIGNED) <> '2008' AND MONTH(DATEHIRED) < '7' ORDER BY LASTNAME"
    REPORT = "ALPHAWITHEMP.XLT"
    MONTHFROM = 1
    MONTHTO = 6
    
    Call PrintALPHAExcel(SQL, REPORT, MONTHFROM, MONTHTO)
    SQL = ""
    MONTHFROM = 0
    MONTHTO = 0
End Sub

Private Sub cmdWithPrevJulDec_Click()
    SQL = "SELECT * FROM HRMS_EMPINFO WHERE YEAR(DATEHIRED) = '2008' AND (EMPLEVEL = 'E' OR EMPLEVEL = 'M') AND PREVIOUSCOMPANY IS NOT NULL AND YEAR(RESIGNED) <> '2008' ORDER BY LASTNAME"
    REPORT = "ALPHAWITHEMP.XLT"
    MONTHFROM = 7
    MONTHTO = 12
    
    Call PrintALPHAExcel(SQL, REPORT, MONTHFROM, MONTHTO)
    SQL = ""
    MONTHFROM = 0
    MONTHTO = 0
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    DrawXPCtl Me
End Sub
