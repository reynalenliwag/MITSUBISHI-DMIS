VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMSPRINT_AllowanceComputation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Allowance Computation"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AllowanceComputation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3120
   Begin VB.Timer tme_Load 
      Interval        =   200
      Left            =   300
      Top             =   2220
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   540
      Width           =   885
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
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   540
      Width           =   1845
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
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2775
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
      Left            =   2130
      MouseIcon       =   "AllowanceComputation.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "AllowanceComputation.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Window"
      Top             =   1020
      Width           =   855
   End
   Begin Crystal.CrystalReport rptBreak 
      Left            =   180
      Top             =   990
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
      Left            =   1290
      MouseIcon       =   "AllowanceComputation.frx":0B27
      MousePointer    =   99  'Custom
      Picture         =   "AllowanceComputation.frx":0C79
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label lblLoad 
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   180
      TabIndex        =   5
      Top             =   1590
      Visible         =   0   'False
      Width           =   945
   End
End
Attribute VB_Name = "frmHRMSPRINT_AllowanceComputation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

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

Function FindPrevMonth() As String
    If cboMonth.Text = "January" Then FindPrevMonth = "12"
    If cboMonth.Text = "Febuary" Then FindPrevMonth = "1"
    If cboMonth.Text = "March" Then FindPrevMonth = "2"
    If cboMonth.Text = "April" Then FindPrevMonth = "3"
    If cboMonth.Text = "May" Then FindPrevMonth = "4"
    If cboMonth.Text = "June" Then FindPrevMonth = "5"
    If cboMonth.Text = "July" Then FindPrevMonth = "6"
    If cboMonth.Text = "August" Then FindPrevMonth = "7"
    If cboMonth.Text = "September" Then FindPrevMonth = "8"
    If cboMonth.Text = "October" Then FindPrevMonth = "9"
    If cboMonth.Text = "November" Then FindPrevMonth = "10"
    If cboMonth.Text = "December" Then FindPrevMonth = "11"
End Function

Private Sub cmdPrint_Click()
    Dim matt As Integer
    If cboQuensina.Text = "1st Cut-Off" Then
        matt = 1
    ElseIf cboQuensina.Text = "2nd Cut-Off" Then
        matt = 2
    End If
    Select Case Me.Caption
        Case "Print OverTime BreakDown":
            'If Function_Access(LOGID, "ACESS_PRINT", "REPORT PRINT DEDUCTION BREAKDOWN") = False Then Exit Sub
           rptBreak.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
           rptBreak.Formulas(1) = "PrintedBy = '" & LOGNAME & "'"
           rptBreak.Formulas(4) = "companyAddress = '" & COMPANY_ADDRESS & "'"
           PrintSQLReport rptBreak, HRMS_REPORT_PATH & "OverTime BreakDown.rpt", "{hrms_overtime.cut_off} = '" & matt & "' AND {hrms_overtime.pay_month} = " & What_month(cboMonth) & " and {hrms_overtime.pay_year} = " & cboYear & "", DMIS_REPORT_Connection, 1
           LogAudit "V", "PRINT OVERTIME BREAKDOWN", cboQuensina & "-" & cboMonth & ", " & cboYear
        
        Case "Print Commission BreakDown":
            'If Function_Access(LOGID, "ACESS_PRINT", "REPORT PRINT COMMISSION BREAKDOWN") = False Then Exit Sub
            rptBreak.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
            rptBreak.Formulas(1) = "PrintedBy = '" & LOGNAME & "'"
            rptBreak.Formulas(2) = "CUT-OFF = '" & cboQuensina.Text & "'"
            'rptBreak.Formulas(3) = "RangeDate = '" & MonthName(month(FromDate)) & " " & Day(FromDate) & "," & _
            'year(FromDate) & " - " & MonthName(month(ToDate)) & " " & Day(ToDate) & "," & year(ToDate) & "'"
            rptBreak.Formulas(4) = "companyAddress = '" & COMPANY_ADDRESS & "'"
            
            Execute_Commission_BreakDown
            
            PrintSQLReport rptBreak, HRMS_REPORT_PATH & "Commission BreakDown.rpt", "", DMIS_REPORT_Connection, 1
            
            LogAudit "V", "PRINT COMMISSION BREAKDOWN", cboQuensina & "-" & cboMonth & ", " & cboYear
            
        Case "Print Deduction BreakDown":
            'If Function_Access(LOGID, "ACESS_PRINT", "REPORT PRINT DEDUCTION BREAKDOWN") = False Then Exit Sub
            rptBreak.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
            rptBreak.Formulas(1) = "PrintedBy = '" & LOGNAME & "'"
            rptBreak.Formulas(4) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptBreak, HRMS_REPORT_PATH & "Deduction BreakDown.rpt", "{hrms_deductions.cut_off} = '" & matt & "' AND {hrms_deductions.pay_month} = " & What_month(cboMonth) & " and {hrms_deductions.pay_year} = " & cboYear & "", DMIS_REPORT_Connection, 1
            LogAudit "V", "PRINT DEDUCTION BREAKDOWN", cboQuensina & "-" & cboMonth & ", " & cboYear
            
        Case "Print Adjustment BreakDown":
            'If Function_Access(LOGID, "ACESS_PRINT", "REPORT PRINT AJUSTMENT BREAKDOWN") = False Then Exit Sub
            rptBreak.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
            rptBreak.Formulas(1) = "PrintedBy = '" & LOGNAME & "'"
            'rptBreak.Formulas(2) = "CUT-OFF = '" & cboQuensina.Text & "'"
            'rptBreak.Formulas(3) = "RangeDate = '" & MonthName(month(FromDate)) & " " & Day(FromDate) & "," & _
            '    year(FromDate) & " - " & MonthName(month(ToDate)) & " " & Day(ToDate) & "," & year(ToDate) & "'"
            rptBreak.Formulas(4) = "companyAddress = '" & COMPANY_ADDRESS & "'"
            
            Execute_Adjustment_BreakDown
            
            PrintSQLReport rptBreak, HRMS_REPORT_PATH & "Adjustment BreakDown.rpt", "", DMIS_REPORT_Connection, 1
            LogAudit "V", "PRINT ADJUSTMENT BREAKDOWN", cboQuensina & "-" & cboMonth & ", " & cboYear
    End Select
End Sub

Function GetDateToPrint()
    If cboQuensina.Text = "1st Cut-Off" Then
        GetDateToPrint = FindPrevMonth
    End If
End Function

Sub Execute_Commission_BreakDown()
    Dim rsTmp As New ADODB.Recordset
    Dim rsCOMM As New ADODB.Recordset
    Dim MM, YY, FromDate, ToDate             As String
    
    lblLoad.Visible = True: DoEvents
    MM = What_month(cboMonth): YY = cboYear.Text
    
    If cboQuensina.Text = "2nd Cut-Off" Then
        FromDate = DateSerial(YY, MM, PAYROLLCODE_FROM2)
        ToDate = DateSerial(YY, MM, PAYROLLCODE_TO2)
    End If
    
    If cboQuensina.Text = "1st Cut-Off" Then
        If PAYROLL_CODE = 1 Then
            FromDate = DateSerial(YY, MM, PAYROLLCODE_FROM1)
            ToDate = DateSerial(YY, MM, PAYROLLCODE_TO1)
        Else
            If cboMonth.Text = "January" Then
                vYEAR = CDbl(cboYear) - 1
            Else
                vYEAR = CDbl(cboYear)
            End If
            
            FromDate = DateSerial(vYEAR, FindPrevMonth, PAYROLLCODE_FROM1)
            ToDate = DateSerial(YY, MM, PAYROLLCODE_TO1)
        End If
    Else
        If PAYROLL_CODE = 1 Then
            FromDate = DateSerial(YY, MM, PAYROLLCODE_FROM2)
            ToDate = DateSerial(YY, MM, PAYROLLCODE_TO2)
        Else
            FromDate = DateSerial(YY, MM, PAYROLLCODE_FROM2)
            ToDate = DateSerial(YY, MM, PAYROLLCODE_TO2)
        End If
    End If
        
    gconDMIS.Execute ("Delete From HRMS_Commission_BD")
    Set rsTmp = gconDMIS.Execute("Select * From HRMS_Commission Where Deyt Between '" & FromDate & _
        "' And '" & ToDate & "'")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set rsCOMM = gconDMIS.Execute("Select * From HRMS_Commission_BD Where Empno = '" & rsTmp!EMPNO & "'")
            If Not (rsCOMM.BOF And rsCOMM.EOF) Then
                gconDMIS.Execute ("Update HRMS_Commission_BD Set Amount = " & rsCOMM!AMOUNT + rsTmp!AMOUNT & _
                    ",TaxAmount = " & rsCOMM!TaxAMount + rsTmp!tax & " Where EMpno = '" & rsTmp!EMPNO & "'")
            Else
                gconDMIS.Execute ("Insert Into HRMS_Commission_BD Values('" & rsTmp!EMPNO & "'," & rsTmp!AMOUNT & _
                    "," & rsTmp!tax & ")")
            End If
            
            rsTmp.MoveNext
        Loop
    End If
    
    lblLoad.Visible = False
    Set rsTmp = Nothing
End Sub

Sub Execute_Adjustment_BreakDown()
    Dim rsTmp As New ADODB.Recordset
    Dim rsDB As New ADODB.Recordset
    Dim TAXABLE As Currency
    Dim NONTAXABLE As Currency
    Dim MM, YY, FromDate, ToDate             As String
    
    lblLoad.Visible = True: DoEvents
    MM = What_month(cboMonth): YY = cboYear.Text
    
    If cboQuensina.Text = "2nd Cut-Off" Then
        FromDate = DateSerial(YY, MM, PAYROLLCODE_FROM2)
        ToDate = DateSerial(YY, MM, PAYROLLCODE_TO2)
    End If
    
    If cboQuensina.Text = "1st Cut-Off" Then
        If PAYROLL_CODE = 1 Then
            FromDate = DateSerial(YY, MM, PAYROLLCODE_FROM1)
            ToDate = DateSerial(YY, MM, PAYROLLCODE_TO1)
        Else
            If cboMonth.Text = "January" Then
                vYEAR = CDbl(cboYear) - 1
            Else
                vYEAR = CDbl(cboYear)
            End If
            
            FromDate = DateSerial(vYEAR, FindPrevMonth, PAYROLLCODE_FROM1)
            ToDate = DateSerial(YY, MM, PAYROLLCODE_TO1)
        End If
    Else
        If PAYROLL_CODE = 1 Then
            FromDate = DateSerial(YY, MM, PAYROLLCODE_FROM2)
            ToDate = DateSerial(YY, MM, PAYROLLCODE_TO2)
        Else
            FromDate = DateSerial(YY, MM, PAYROLLCODE_FROM2)
            ToDate = DateSerial(YY, MM, PAYROLLCODE_TO2)
        End If
    End If
    
    gconDMIS.Execute ("Delete From HRMS_Adjustment_BD")
    Set rsTmp = gconDMIS.Execute("Select * from HRMS_Adjustment Where Deyt Between '" & FromDate & _
        "' And '" & ToDate & "' Order By Empno ASC")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            TAXABLE = 0
            NONTAXABLE = 0
            
            If rsTmp!Type = "T" Then
                TAXABLE = rsTmp!AMOUNT
            Else
                NONTAXABLE = rsTmp!AMOUNT
            End If
            
            Set rsDB = gconDMIS.Execute("Select EmpNO From HRMS_Adjustment_BD Where Empno = '" & rsTmp!EMPNO & "'")
            If Not (rsDB.BOF And rsDB.EOF) Then
                gconDMIS.Execute ("Update HRMS_Adjustment_BD Set Taxable = " & rsDB!TAXABLE + TAXABLE & _
                    ",NonTaxable = " & rsDB!NONTAXABLE + NONTAXABLE & " Where Empno = '" & rsTmp!EMPNO & "'")
            Else
                gconDMIS.Execute ("Insert Into HRMS_Adjustment_BD Values('" & rsTmp!EMPNO & _
                    "'," & TAXABLE & _
                    "," & NONTAXABLE & ")")
            End If
    
            rsTmp.MoveNext
        Loop
    End If
    
    lblLoad.Visible = False
    Set rsTmp = Nothing
End Sub

Sub Execute_Deduction_BreakDown()
    Dim rsTmp As New ADODB.Recordset
    Dim rsTMP1 As New ADODB.Recordset
    Dim DAYT1 As String: Dim DAYT2 As String
    Dim DateFrom As String: Dim DateTo As String
        
    Dim DLATE As Currency
    Dim DHDABSENT As Currency
    Dim DWDABSENT As Currency
    Dim DOTHERS As Currency
    
    Dim AMOUNT As Currency
    Dim MM, YY, FromDate, ToDate             As String
    
    lblLoad.Visible = True: DoEvents
    
    MM = What_month(cboMonth): YY = cboYear.Text
   
    If cboQuensina.Text = "1st Cut-Off" Then
        FromDate = DateSerial(YY, 1, 21)
        ToDate = DateSerial(YY, 2, 5)
    Else
        FromDate = DateSerial(YY, 2, 6)
        ToDate = DateSerial(YY, 2, 20)
    End If
    
    gconDMIS.Execute ("Delete from HRMS_Deduction_BD")
    Set rsTmp = gconDMIS.Execute("Select * From HRMS_Deductions Where Deyt between '" & FromDate & "' and '" & ToDate & "'")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            If rsTmp!PARTICULAR = "LT" Then
                DLATE = rsTmp!AMOUNT
            ElseIf rsTmp!PARTICULAR = "HD" Then
                DHDABSENT = rsTmp!AMOUNT
            ElseIf rsTmp!PARTICULAR = "WD" Then
                DWDABSENT = rsTmp!AMOUNT
            Else
                DOTHERS = rsTmp!AMOUNT
            End If
            
            Set rsTMP1 = gconDMIS.Execute("Select * From HRMS_Deduction_BD Where Empno = '" & rsTmp!EMPNO & "'")
            If Not (rsTMP1.BOF And rsTMP1.EOF) Then
                gconDMIS.Execute ("Update HRMS_Deduction_BD Set Late = " & rsTMP1!Late + DLATE & _
                ",HDAbsent = " & rsTMP1!HDAbsent + DHDABSENT & _
                ",WDAbsent = " & rsTMP1!WDAbsent + DWDABSENT & _
                ",Other = " & rsTMP1!Other + DOTHERS & _
                " Where Empno = '" & rsTmp!EMPNO & "'")
            Else
                gconDMIS.Execute ("Insert Into HRMS_Deduction_BD Values('" & rsTmp!EMPNO & _
                    "'," & DLATE & "," & DHDABSENT & "," & DWDABSENT & "," & DOTHERS & ")")
            End If
                   
            DLATE = 0
            DHDABSENT = 0
            DWDABSENT = 0
            DOTHERS = 0
            
            rsTmp.MoveNext
        Loop
    End If
    
    lblLoad.Visible = False
    
    Set rsTmp = Nothing
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    
    cboQuensina.AddItem "1st Cut-Off"
    cboQuensina.AddItem "2nd Cut-Off"
    cboQuensina.ListIndex = 0
     
    fillcbomonth cboMonth
    FillcboYear cboYear
    
    cboMonth.Text = MonthName(Month(Now))
    cboYear.Text = Year(Now)
End Sub

Private Sub tme_Load_Timer()
    If lblLoad.ForeColor = vbBlue Then
        lblLoad.ForeColor = vbRed
    Else
        lblLoad.ForeColor = vbBlue
    End If
End Sub
