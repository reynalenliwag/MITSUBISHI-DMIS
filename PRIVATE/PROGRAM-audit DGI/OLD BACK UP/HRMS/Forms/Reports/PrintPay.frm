VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMSPrintPayroll 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Payroll"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3585
   ForeColor       =   &H00D8E9EC&
   Icon            =   "PrintPay.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3855
   ScaleWidth      =   3585
   Begin VB.CheckBox Check1 
      Caption         =   "Individual Payroll Summary"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   390
      TabIndex        =   12
      Top             =   1500
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.CheckBox chkManager 
      Caption         =   "Print for Managers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   90
      TabIndex        =   11
      Top             =   1800
      Width           =   3345
   End
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
      Height          =   795
      Left            =   1710
      MouseIcon       =   "PrintPay.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "PrintPay.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Close Window"
      Top             =   3000
      Width           =   855
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
      Height          =   795
      Left            =   870
      MouseIcon       =   "PrintPay.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "PrintPay.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print Report"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CheckBox chkProbReg 
      Caption         =   "Print for Probationary/Regular Employees"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   90
      TabIndex        =   5
      Top             =   2040
      Width           =   3345
   End
   Begin VB.CheckBox chkAllowanceBase 
      Caption         =   "Print for Allowance Base Employees"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   90
      TabIndex        =   7
      Top             =   2580
      Width           =   3345
   End
   Begin VB.CheckBox chkContractual 
      Caption         =   "Print for Contractual Employees"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   90
      TabIndex        =   6
      Top             =   2310
      Width           =   3345
   End
   Begin VB.CheckBox chkInclude 
      Caption         =   "Print Payroll Sheet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   390
      TabIndex        =   3
      Top             =   960
      Width           =   2445
   End
   Begin VB.CheckBox chkPreview 
      BackColor       =   &H00D8E9EC&
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2190
      TabIndex        =   8
      Top             =   4980
      Width           =   1275
   End
   Begin VB.CheckBox chkPaySlip 
      Caption         =   "Print Payslip Only"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   390
      TabIndex        =   4
      Top             =   1230
      Width           =   2445
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
      Left            =   390
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2895
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
      Left            =   390
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   540
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   540
      Width           =   1005
   End
   Begin Crystal.CrystalReport rptPrintPay 
      Left            =   2760
      Top             =   2970
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
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      Height          =   285
      Left            =   2670
      TabIndex        =   13
      Top             =   3570
      Visible         =   0   'False
      Width           =   435
   End
End
Attribute VB_Name = "frmHRMSPrintPayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function GetEmloyeeName(EMPNO As String) As String
    GetEmloyeeName = ""
    Dim rsTemp                                       As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE EMPNO = '" & EMPNO & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        GetEmloyeeName = Null2String(rsTemp!lastname) & ", " & Null2String(rsTemp!FIRSTNAME)
    End If
    Set rsTemp = Nothing
End Function

Function GetDescription(CODE As String) As String
    GetDescription = ""
    Dim rsTemp                                       As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT DET_DESC FROM HRMS_PAYROLL_DET WHERE DET_CODE = '" & CODE & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        GetDescription = Null2String(rsTemp!DET_DESC)
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

Sub PrintPayrollExcel(LEVEL As String)
    Dim xlApp                                        As Excel.Application
    Dim xlsheet                                      As Excel.Worksheet
    Dim xlbook                                       As Excel.Workbook
    Dim count1                                       As Integer
    count1 = 0
    Dim count2                                       As Integer
    count2 = 0
    Dim I                                            As Integer
    Dim codes(50)                                    As String
    Dim fld                                          As Field
    I = 1

    Dim matt                                         As String
    If cboQuensina.Text = "1st Cut-Off" Then
        matt = "1"
    ElseIf cboQuensina.Text = "2nd Cut-Off" Then
        matt = "2"
    Else
        MsgBox "SELECT CUT-OFF!"
    End If

    Set xlApp = New Excel.Application
    Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & "NEW.XLT")
    Set xlsheet = xlbook.Worksheets(1)


    Dim Y                                            As Integer
    Y = 0

    xlsheet.Cells(1, 1) = "EMPNO"
    xlsheet.Cells(1, 2) = "NAME"
    xlsheet.Cells(1, 3) = "RATE"
    xlsheet.Cells(1, 4) = "OT"
    xlsheet.Cells(1, 5) = "TAXABLE ADJ"
    xlsheet.Cells(1, 6) = "NON-TAXABLE ADJ"

    Dim rsCodes                                      As ADODB.Recordset
    Set rsCodes = New ADODB.Recordset
    Set rsCodes = gconDMIS.Execute("SELECT DISTINCT DET_CODE FROM HRMS_PAYROLL_DET WHERE CUT_OFF = " & matt & " AND PAY_MONTH = " & What_month(cboMONTH) & " AND PAY_YEAR = " & cboyear)
    If Not rsCodes.EOF And Not rsCodes.BOF Then
        rsCodes.MoveFirst
        While Not rsCodes.EOF
            xlsheet.Cells(1, Y + 7) = rsCodes!DET_CODE
            Y = Y + 1
            rsCodes.MoveNext
        Wend
    End If

    xlsheet.Cells(1, 7 + Y) = "SSSE"
    xlsheet.Cells(1, 8 + Y) = "PHICE"
    xlsheet.Cells(1, 9 + Y) = "PAGIBIGE"
    xlsheet.Cells(1, 10 + Y) = "TAX"
    xlsheet.Cells(1, 11 + Y) = "ALLOWANCE"
    xlsheet.Cells(1, 12 + Y) = "NET"

    Dim empnoint                                     As Integer
    empnoint = 0

    Dim rsTemp                                       As ADODB.Recordset
    Dim RSTOTAL                                      As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_PAYROLL WHERE CUT_OFF = " & matt & " AND PAY_MONTH = " & What_month(cboMONTH) & " AND PAY_YEAR = " & cboyear & " AND EMPLEVEL = '" & LEVEL & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        rsTemp.MoveFirst
        While Not rsTemp.EOF
            xlsheet.Cells(2 + empnoint, 1) = Null2String(rsTemp!EMPNO)
            xlsheet.Cells(2 + empnoint, 2) = GetEmloyeeName(Null2String(rsTemp!EMPNO))
            xlsheet.Cells(2 + empnoint, 3) = N2Str2Zero(rsTemp!Rate)
            xlsheet.Cells(2 + empnoint, 4) = N2Str2Zero(rsTemp!OVERTIME)
            xlsheet.Cells(2 + empnoint, 5) = N2Str2Zero(rsTemp!TAXABLEADJ)
            xlsheet.Cells(2 + empnoint, 6) = N2Str2Zero(rsTemp!NONTAXABLEADJ)
            xlsheet.Cells(2 + empnoint, 7 + Y) = N2Str2Zero(rsTemp!SSSE)
            
            xlsheet.Cells(2 + empnoint, 8 + Y) = N2Str2Zero(rsTemp!PHILHEALTHE)
            xlsheet.Cells(2 + empnoint, 9 + Y) = N2Str2Zero(rsTemp!PAGIBIG)
            xlsheet.Cells(2 + empnoint, 10 + Y) = N2Str2Zero(rsTemp!TAX)
            xlsheet.Cells(2 + empnoint, 11 + Y) = N2Str2Zero(rsTemp!ALLOWANCE)
            xlsheet.Cells(2 + empnoint, 12 + Y) = N2Str2Zero(rsTemp!NETPAY) + N2Str2Zero(rsTemp!ALLOWANCE)

            Set RSTOTAL = gconDMIS.Execute("SELECT EMPNO, DET_CODE, SUM(DET_AMOUNT) AMT FROM HRMS_PAYROLL_DET WHERE CUT_OFF = " & matt & " AND PAY_MONTH = " & What_month(cboMONTH) & " AND PAY_YEAR =" & cboyear & " AND EMPNO = '" & rsTemp!EMPNO & "' GROUP BY DET_CODE, EMPNO")
            While Not RSTOTAL.EOF
                For j = 7 To Y + 6
                    If xlsheet.Cells(1, j) = RSTOTAL!DET_CODE Then
                        xlsheet.Cells(2 + empnoint, j) = N2Str2Zero(RSTOTAL!amt)
                    End If
                Next
                RSTOTAL.MoveNext
            Wend

            empnoint = empnoint + 1
            rsTemp.MoveNext
        Wend
    End If

    Dim matthew                                      As Integer
    For matthew = 3 To Y + 12
        xlsheet.Cells(empnoint + 2, matthew).Formula = "=SUM(" & GetLetter(matthew) & "1:" & GetLetter(matthew) & empnoint + 1 & ")"
    Next

    xlsheet.Cells(empnoint + 3, 2) = cboQuensina.Text & " " & cboMONTH & " " & cboyear
    xlsheet.Cells(empnoint + 4, 1) = "LEGEND"
    For j = 7 To Y + 6
        xlsheet.Cells(empnoint + 5 + (j - 7), 1) = xlsheet.Cells(1, j)
        xlsheet.Cells(empnoint + 5 + (j - 7), 2) = GetDescription(xlsheet.Cells(1, j))
    Next

    Set RSTOTAL = Nothing
    Set rsTemp = Nothing

    xlApp.Visible = True
    Set xlsheet = Nothing
    Set xlbook = Nothing
    Set xlApp = Nothing
End Sub
Sub PrintHARIPayroll(LEVEL As String)
    Dim xlApp                                        As Excel.Application
    Dim xlsheet                                      As Excel.Worksheet
    Dim xlbook                                       As Excel.Workbook
    Dim count1                                       As Integer
    count1 = 0
    Dim count2                                       As Integer
    count2 = 0
    Dim I                                            As Integer
    Dim codes(50)                                    As String
    Dim fld                                          As Field
    I = 1

    Dim matt                                         As String
    If cboQuensina.Text = "1st Cut-Off" Then
        matt = "1"
    ElseIf cboQuensina.Text = "2nd Cut-Off" Then
        matt = "2"
    Else
        MsgBox "SELECT CUT-OFF!"
    End If

    Set xlApp = New Excel.Application
    Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & "payroll sheet.xlt")
    Set xlsheet = xlbook.Worksheets(1)


    Dim Y                                            As Integer
    Y = 0

    xlsheet.Cells(1, "B") = COMPANY_NAME
    xlsheet.Cells(2, "B") = COMPANY_ADDRESS
    xlsheet.Cells(3, "B") = "Payroll Sheet for " & cboQuensina & " " & cboMONTH & " " & cboyear
    
    Dim RG As Excel.Range
    Dim rsPAYROLL                                    As ADODB.Recordset
    Dim RSPAYROLLDET                                 As ADODB.Recordset
    Dim RSPAYROLL_PAGIBIG                            As ADODB.Recordset
    Dim RSPAYROLL_CUSTOMERDEPOSIT                    As ADODB.Recordset
    Dim RSPAYROLL_ARE                                As ADODB.Recordset
    Dim rsDepGroup  As ADODB.Recordset
    Set rsPAYROLL = New ADODB.Recordset
    Dim RSPAYROLL_TRANSPO As ADODB.Recordset
    Dim RSPAYROLL_MEAL As ADODB.Recordset
    Dim RSMONTHLYRATE As ADODB.Recordset
    Dim DEP_CNTR As Integer
    
    Set rsDepGroup = gconDMIS.Execute("SELECT DEPTNAME,DEPTCODE FROM HRMS_DEPARTMENT ORDER BY 1 ASC")
    Y = 7
    
    DEP_CNTR = 7
    While Not rsDepGroup.EOF
        
        Set rsPAYROLL = gconDMIS.Execute("SELECT HRMS_EMPINFO.DEPTCODE, HRMS_EMPINFO.EMPNO,HRMS_EMPINFO.LASTNAME + ' ,' + HRMS_EMPINFO.FIRSTNAME + ' .' + LEFT(HRMS_EMPINFO.MIDDLENAME,1) as FULLNAME, HRMS_PAYROLL.*  FROM HRMS_PAYROLL INNER JOIN HRMS_EMPINFO ON HRMS_EMPINFO.EMPNO=HRMS_PAYROLL.EMPNO  WHERE CUT_OFF = " & matt & " AND PAY_MONTH = " & What_month(cboMONTH) & " AND PAY_YEAR = " & cboyear & " AND HRMS_EMPINFO.DEPTCODE=" & N2Str2Null(rsDepGroup!DEPTCODE) & " AND HRMS_EMPINFO.ACTIVEINACTIVE <> 'I' ORDER BY HRMS_EMPINFO.LASTNAME  asc")
    If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
        Set RG = xlsheet.Range(xlsheet.Cells(Y, "A"), xlsheet.Cells(Y, "T"))
            RG.Merge
            RG.Interior.Color = &HC0C0C0
            xlsheet.Cells(Y, "A") = Null2String(rsDepGroup!DEPTNAME)
            Y = Y + 1
            rsPAYROLL.MoveFirst
            DEP_CNTR = Y
        While Not rsPAYROLL.EOF
            
            Set RSPAYROLLDET = gconDMIS.Execute("SELECT  ISNULL(SUM(DET_AMOUNT),0) AS SSSLOAN FROM HRMS_PAYROLL_DET WHERE TRANTYPE='L' AND DET_CODE IN('CSAL','SSAL') AND EMPNO ='" & rsPAYROLL!EMPNO & "' AND CUT_OFF = " & matt & " AND PAY_MONTH = " & What_month(cboMONTH) & " AND PAY_YEAR = " & cboyear)
            Set RSPAYROLL_PAGIBIG = gconDMIS.Execute("SELECT  ISNULL(SUM(DET_AMOUNT),0) AS PAGIBIG FROM HRMS_PAYROLL_DET WHERE TRANTYPE='L' AND DET_CODE IN ('HMDF','OPML','PSAL') AND EMPNO ='" & rsPAYROLL!EMPNO & "' AND CUT_OFF = " & matt & " AND PAY_MONTH = " & What_month(cboMONTH) & " AND PAY_YEAR = " & cboyear)
            Set RSPAYROLL_CUSTOMERDEPOSIT = gconDMIS.Execute("SELECT  ISNULL(SUM(DET_AMOUNT),0) AS CUSTOMERDEPOSIT FROM HRMS_PAYROLL_DET WHERE TRANTYPE='L' AND DET_CODE NOT IN ('CSAL','SSAL','HMDF','OPML','PSAL','ARE')  AND EMPNO ='" & rsPAYROLL!EMPNO & "' AND CUT_OFF = " & matt & " AND PAY_MONTH = " & What_month(cboMONTH) & " AND PAY_YEAR = " & cboyear)
            Set RSPAYROLL_ARE = gconDMIS.Execute("SELECT  ISNULL(SUM(DET_AMOUNT),0) AS ARELOAN FROM HRMS_PAYROLL_DET WHERE ((TRANTYPE='L'    AND DET_CODE   IN ('ARE' )  )  OR (TRANTYPE ='D' AND DET_CODE IN('TL','UN', 'UL','AR', 'CA','CI', 'OA','VC'))) AND EMPNO ='" & rsPAYROLL!EMPNO & "' AND CUT_OFF = " & matt & " AND PAY_MONTH = " & What_month(cboMONTH) & " AND PAY_YEAR = " & cboyear)
            Set RSPAYROLL_MEAL = gconDMIS.Execute("SELECT SUM(AMOUNT) AS MEALALLOWANCE  FROM HRMS_ADJUSTMENT  WHERE CUT_OFF = " & matt & " AND PAY_MONTH = " & What_month(cboMONTH) & " AND PAY_YEAR = " & cboyear & " AND PARTICULAR='003' AND EMPNO=" & N2Str2Null(rsPAYROLL!EMPNO))
            Set RSPAYROLL_TRANSPO = gconDMIS.Execute("SELECT SUM(AMOUNT) AS TRANSPO  FROM HRMS_ADJUSTMENT  WHERE CUT_OFF = " & matt & " AND PAY_MONTH = " & What_month(cboMONTH) & " AND PAY_YEAR = " & cboyear & " AND PARTICULAR='004' AND EMPNO=" & N2Str2Null(rsPAYROLL!EMPNO))
            Set RSMONTHLYRATE = gconDMIS.Execute("SELECT MONTHLYRATE  FROM HRMS_PAYROLL WHERE CUT_OFF = " & matt & " AND PAY_MONTH = " & What_month(cboMONTH) & " AND PAY_YEAR = " & cboyear & " AND EMPNO=" & N2Str2Null(rsPAYROLL!EMPNO))
            xlsheet.Cells(Y, "A") = rsPAYROLL!EMPNO
            xlsheet.Cells(Y, "B") = rsPAYROLL!FULLNAME
            xlsheet.Cells(Y, "C") = RSMONTHLYRATE!MONTHLYRATE
            xlsheet.Cells(Y, "D") = rsPAYROLL!OVERTIME
            xlsheet.Cells(Y, "E") = rsPAYROLL!ALLOWANCE
            xlsheet.Cells(Y, "F") = rsPAYROLL!TAX
            xlsheet.Cells(Y, "G") = RSPAYROLL_MEAL!MEALALLOWANCE
            xlsheet.Cells(Y, "H") = RSPAYROLL_TRANSPO!TRANSPO
            xlsheet.Cells(Y, "I") = rsPAYROLL!SSSE
            xlsheet.Cells(Y, "J") = rsPAYROLL!SSSR
            If NumericVal(rsPAYROLL!SSSE) > 0 Then
                If NumericVal(rsPAYROLL!SSSE) < 500 Then
                    xlsheet.Cells(Y, "K") = 10
                Else
                    xlsheet.Cells(Y, "K") = 30
                End If
            End If
            
            xlsheet.Cells(Y, "L") = rsPAYROLL!PHILHEALTHR
            xlsheet.Cells(Y, "M") = rsPAYROLL!PHILHEALTHE
            xlsheet.Cells(Y, "N") = rsPAYROLL!PAGIBIGR
            xlsheet.Cells(Y, "O") = rsPAYROLL!PAGIBIG
            xlsheet.Cells(Y, "P") = RSPAYROLLDET!SSSLOAN
            xlsheet.Cells(Y, "Q") = RSPAYROLL_PAGIBIG!PAGIBIG
            xlsheet.Cells(Y, "R") = RSPAYROLL_CUSTOMERDEPOSIT!CUSTOMERDEPOSIT
            xlsheet.Cells(Y, "S") = RSPAYROLL_ARE!ARELOAN
            xlsheet.Cells(Y, "U") = rsPAYROLL!NETPAY
            Y = Y + 1
            rsPAYROLL.MoveNext
        Wend
                xlsheet.Cells(Y, "B") = UCase("TOTAL FOR " & Null2String(rsDepGroup!DEPTNAME))
            Set RG = xlsheet.Range(xlsheet.Cells(Y, "C"), xlsheet.Cells(Y, "T"))
                RG.Formula = "=SUM(C" & DEP_CNTR & ":C" & Y - 1 & ")"
                RG.Font.Bold = True
               ' RG.Borders(xlEdgeBottom) = 2
                Set RG = Nothing
                Y = Y + 1
    End If
    rsDepGroup.MoveNext
Wend
    Set RSTOTAL = Nothing
    Set rsTemp = Nothing
    xlApp.Visible = True
    Set xlsheet = Nothing
    Set xlbook = Nothing
    Set xlApp = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    Dim vYEAR                                        As Integer
    Dim vMONTH                                       As Integer
    Dim VCUT_OFF                                     As String
    If cboQuensina.Text = "1st Cut-Off" Then
        VCUT_OFF = "1"
    Else
        VCUT_OFF = "2"
    End If
    vYEAR = cboyear.Text
    vMONTH = What_month(cboMONTH.Text)
    rptPrintPay.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptPrintPay.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptPrintPay.WindowShowSearchBtn = True

    If chkInclude.Value = 1 Then
        If chkProbReg.Value = 1 Then

            If COMPANY_CODE = "HARI" Then
                PrintHARIPayroll ("E")
            Else
                Call PrintPayrollExcel("E")
            End If
             PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "payroll.rpt", "{payroll.CUT_OFF} = '" & VCUT_OFF & "' AND {payroll.PAY_MONTH} = " & vMONTH & " AND {payroll.PAY_YEAR} = " & vYEAR & " AND {empinfo.emplevel} = 'E' AND {empinfo.ACTIVEINACTIVE} <> 'I'", DMIS_REPORT_Connection, 1
        End If
        If chkContractual.Value = 1 Then
            PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "payroll.rpt", "{payroll.CUT_OFF} = '" & VCUT_OFF & "' AND {payroll.PAY_MONTH} = " & vMONTH & " AND {payroll.PAY_YEAR} = " & vYEAR & " AND {empinfo.emplevel} = 'C' AND {empinfo.ACTIVEINACTIVE} <> 'I'", DMIS_REPORT_Connection, 1
        End If
        If chkAllowanceBase.Value = 1 Then
            PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "payroll.rpt", "{payroll.CUT_OFF} = '" & VCUT_OFF & "' AND {payroll.PAY_MONTH} = " & vMONTH & " AND {payroll.PAY_YEAR} = " & vYEAR & " AND {empinfo.emplevel} = 'A' AND {empinfo.ACTIVEINACTIVE} <> 'I'", DMIS_REPORT_Connection, 1
        End If
        If chkManager.Value = 1 Then
            'Call PrintPayrollExcel("M")
            PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "payroll.rpt", "{payroll.CUT_OFF} = '" & VCUT_OFF & "' AND {payroll.PAY_MONTH} = " & vMONTH & " AND {payroll.PAY_YEAR} = " & vYEAR & " AND {empinfo.emplevel} = 'M' AND {empinfo.ACTIVEINACTIVE} <> 'I'", DMIS_REPORT_Connection, 1
        End If
    End If

    If chkPaySlip.Value = 1 Then
        If chkProbReg.Value = 1 Then
            rptPrintPay.WindowTitle = "Regular/Probationary Payslip"
            PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "payslip.rpt", "{payroll.CUT_OFF} = '" & VCUT_OFF & "' AND {payroll.PAY_MONTH} = " & vMONTH & " AND {payroll.PAY_YEAR} = " & vYEAR & " AND {empinfo.emplevel} = 'E' AND {empinfo.ACTIVEINACTIVE} <> 'I'", DMIS_REPORT_Connection, 1
        End If
        If chkContractual.Value = 1 Then
            rptPrintPay.WindowTitle = "Contractual Payslip"
            PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "payslip.rpt", "{payroll.CUT_OFF} = '" & VCUT_OFF & "' AND {payroll.PAY_MONTH} = " & vMONTH & " AND {payroll.PAY_YEAR} = " & vYEAR & " AND {empinfo.emplevel} = 'A' AND {empinfo.ACTIVEINACTIVE} <> 'I'", DMIS_REPORT_Connection, 1
        End If
        If chkAllowanceBase.Value = 1 Then
            rptPrintPay.WindowTitle = "Allowance Base Payslip"
            PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "payslip.rpt", "{payroll.CUT_OFF} = '" & VCUT_OFF & "' AND {payroll.PAY_MONTH} = " & vMONTH & " AND {payroll.PAY_YEAR} = " & vYEAR & " AND {empinfo.emplevel} = 'C' AND {empinfo.ACTIVEINACTIVE} <> 'I'", DMIS_REPORT_Connection, 1
        End If
        If chkManager.Value = 1 Then
            rptPrintPay.WindowTitle = "Managers Payslip"
            PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "payslip.rpt", "{payroll.CUT_OFF} = '" & VCUT_OFF & "' AND {payroll.PAY_MONTH} = " & vMONTH & " AND {payroll.PAY_YEAR} = " & vYEAR & " AND {empinfo.emplevel} = 'M' AND {empinfo.ACTIVEINACTIVE} <> 'I'", DMIS_REPORT_Connection, 1
        End If
    End If

    If Check1.Value = 1 Then
        If chkProbReg.Value = 1 Then

            PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "PayrollSheet_Individual.rpt", "{payroll.PAY_YEAR} = " & vYEAR & " AND {empinfo.emplevel} = 'E' AND {empinfo.ACTIVEINACTIVE} <> 'I'", DMIS_REPORT_Connection, 1
        End If
        If chkContractual.Value = 1 Then

            PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "PayrollSheet_Individual.rpt", "{payroll.PAY_YEAR} = " & vYEAR & " AND {empinfo.emplevel} = 'A' AND {empinfo.ACTIVEINACTIVE} <> 'I'", DMIS_REPORT_Connection, 1
        End If
        If chkAllowanceBase.Value = 1 Then
            PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "PayrollSheet_Individual.rpt", "{payroll.PAY_YEAR} = " & vYEAR & " AND {empinfo.emplevel} = 'C' AND {empinfo.ACTIVEINACTIVE} <> 'I'", DMIS_REPORT_Connection, 1
        End If
        If chkManager.Value = 1 Then
            PrintSQLReport rptPrintPay, HRMS_REPORT_PATH & "PayrollSheet_Individual.rpt", "{payroll.PAY_YEAR} = " & vYEAR & " AND {empinfo.emplevel} = 'M' AND {empinfo.ACTIVEINACTIVE} <> 'I'", DMIS_REPORT_Connection, 1
        End If
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'DrawXPCtl Me
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    If LOGLEVEL <> "ADM" Then chkInclude.Enabled = False
    Dim rsCutoff                                     As ADODB.Recordset
    Set rsCutoff = New ADODB.Recordset
    Set rsCutoff = gconDMIS.Execute("SELECT PERIODMONTH,PERIODYEAR,NOTEDBY2 FROM HRMS_PAYROLLSETUP")
    If Not (rsCutoff.EOF And rsCutoff.BOF) Then
        If NumericVal(rsCutoff!NOTEDBY2) = 1 Then
            cboQuensina.Clear
            cboQuensina.AddItem "1st Cut-Off"
            cboQuensina.Text = "1st Cut-Off"
        ElseIf NumericVal(rsCutoff!NOTEDBY2) = 2 Then
            cboQuensina.Clear
            cboQuensina.AddItem "2nd Cut-Off"
            cboQuensina.Text = "2nd Cut-Off"
        Else
            MsgBox "Cut-off not set"
        End If
        cboMONTH.Clear
        cboMONTH.AddItem MonthName(Null2String(rsCutoff!PERIODMONTH))
        cboMONTH.Text = MonthName(Null2String(rsCutoff!PERIODMONTH))
        cboyear.Clear
        cboyear.AddItem Null2String(rsCutoff!PERIODYEAR)
        cboyear.Text = Null2String(rsCutoff!PERIODYEAR)
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

