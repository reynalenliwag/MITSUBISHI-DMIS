VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMSActualManningReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actual Manning"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3195
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMSActualManningReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2085
   ScaleWidth      =   3195
   Begin VB.CheckBox Option1 
      Caption         =   "Summary"
      Height          =   195
      Left            =   990
      TabIndex        =   6
      Top             =   900
      Width           =   1695
   End
   Begin VB.ComboBox cboMOnth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   510
      Width           =   2055
   End
   Begin Crystal.CrystalReport rptActualManningReport 
      Left            =   240
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Actual Manning Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   795
      Left            =   2400
      MouseIcon       =   "frmCSMSActualManningReport.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSActualManningReport.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Close Window"
      Top             =   1200
      Width           =   645
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   885
      Left            =   3480
      MouseIcon       =   "frmCSMSActualManningReport.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSActualManningReport.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   2820
      Width           =   675
   End
   Begin Crystal.CrystalReport RPT1 
      Left            =   870
      Top             =   1350
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Actual Manning Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   795
      Left            =   1770
      MouseIcon       =   "frmCSMSActualManningReport.frx":1C10
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSActualManningReport.frx":1D62
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Print Report"
      Top             =   1200
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   375
      TabIndex        =   4
      Top             =   240
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   585
      Width           =   735
   End
End
Attribute VB_Name = "frmCSMSActualManningReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlApp                                              As Excel.Application
Dim xlBook                                             As Excel.Workbook
Dim xlSheet                                            As Excel.Worksheet

Function CheckMonthName(MNAME As String) As Integer
    If MNAME = "January" Then CheckMonthName = 1

End Function

Function GetMyDate(mMONTH As Integer)
    If mMONTH = 1 Then GetMyDate = 1 & "/" & Day(lastDay(DateSerial(cboYear, 1, 1))) & "/" & cboYear
    If mMONTH = 2 Then GetMyDate = 2 & "/" & Day(lastDay(DateSerial(cboYear, 2, 1))) & "/" & cboYear
    If mMONTH = 3 Then GetMyDate = 3 & "/" & Day(lastDay(DateSerial(cboYear, 3, 1))) & "/" & cboYear
    If mMONTH = 4 Then GetMyDate = 4 & "/" & Day(lastDay(DateSerial(cboYear, 4, 1))) & "/" & cboYear
    If mMONTH = 5 Then GetMyDate = 5 & "/" & Day(lastDay(DateSerial(cboYear, 5, 1))) & "/" & cboYear
    If mMONTH = 6 Then GetMyDate = 6 & "/" & Day(lastDay(DateSerial(cboYear, 6, 1))) & "/" & cboYear
    If mMONTH = 7 Then GetMyDate = 7 & "/" & Day(lastDay(DateSerial(cboYear, 7, 1))) & "/" & cboYear
    If mMONTH = 8 Then GetMyDate = 8 & "/" & Day(lastDay(DateSerial(cboYear, 8, 1))) & "/" & cboYear
    If mMONTH = 9 Then GetMyDate = 9 & "/" & Day(lastDay(DateSerial(cboYear, 9, 1))) & "/" & cboYear
    If mMONTH = 10 Then GetMyDate = 10 & "/" & Day(lastDay(DateSerial(cboYear, 10, 1))) & "/" & cboYear
    If mMONTH = 11 Then GetMyDate = 11 & "/" & Day(lastDay(DateSerial(cboYear, 11, 1))) & "/" & cboYear
    If mMONTH = 12 Then GetMyDate = 12 & "/" & Day(lastDay(DateSerial(cboYear, 12, 1))) & "/" & cboYear
End Function

Function GetDeptName()
    If COMPANY_CODE = "HAI" Then GetDeptName = "DEP-050"
    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCI" Or COMPANY_CODE = "HAS" Then GetDeptName = "SERVICE"
    If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HSB" Then GetDeptName = ""
    If COMPANY_CODE = "HPC" Then GetDeptName = "0005"
End Function

Sub cmdPrint_Click()
    '    If Function_Access(LOGID, "Acess_Print", "ACTUAL MANNING REPORT") = False Then Exit Sub

    'On Error GoTo Errorcode
    Dim DEPT_NAME                                      As String
    DEPT_NAME = GetDeptName

    If Option1.Value = 1 Then
        Screen.MousePointer = 11

        RPT1.Formulas(1) = "COMPANYNAME = '" & COMPANY_NAME & "'"
        RPT1.Formulas(2) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
        RPT1.Formulas(3) = "printedby = '" & LOGNAME & "'"
        RPT1.Formulas(4) = "MONTHNAME = '" & cboMonth & "'"
        RPT1.Formulas(5) = "YEARREPORT = '" & cboYear & "'"

        PrintSQLReport RPT1, CSMS_REPORT_PATH & "Actual_Manning_Report_SUM.rpt", "{HRMS_EMPINFO.DEPTCODE} = '" & DEPT_NAME & "'", CSMS_REPORT_CONNECTION, 1
        LogAudit "V", "ACTUAL MANNING REPORT_SUMMARY", cboYear

        Screen.MousePointer = 0
        Exit Sub
    End If
    Screen.MousePointer = 11

    Dim rsActualManning                                As ADODB.Recordset
    Set rsActualManning = New ADODB.Recordset
    Set rsActualManning = gconDMIS.Execute("Select * from HRMS_EmpInfo where Year(datehired) = '" & cboYear.Text & "' and deptcode = '" & DEPT_NAME & "'")
    If Not rsActualManning.EOF And Not rsActualManning.BOF Then
        Dim Prob_Count_Jan                             As Integer
        Dim Prob_Count_Feb                             As Integer
        Dim Prob_Count_Mar                             As Integer
        Dim Prob_Count_Apr                             As Integer
        Dim Prob_Count_May                             As Integer
        Dim Prob_Count_June                            As Integer
        Dim Prob_Count_July                            As Integer
        Dim Prob_Count_August                          As Integer
        Dim Prob_Count_Sept                            As Integer
        Dim Prob_Count_Oct                             As Integer
        Dim Prob_Count_Nov                             As Integer
        Dim Prob_Count_Dec                             As Integer
        Dim MonthHired_Prob                            As Integer

        Prob_Count_Jan = 0: Prob_Count_Apr = 0: Prob_Count_August = 0: Prob_Count_Dec = 0
        Prob_Count_Feb = 0: Prob_Count_July = 0: Prob_Count_June = 0: Prob_Count_Mar = 0
        Prob_Count_May = 0: Prob_Count_Nov = 0: Prob_Count_Oct = 0: Prob_Count_Sept = 0
        Dim MMNAME                                     As Integer
        Dim mDEYT                                      As String
        Dim tDEYT                                      As String
        Dim d1                                         As String


        'COMPUTE FOR THE BEGGINING


        Dim BEG_CNT                                    As Integer
        Dim rstmp                                      As New ADODB.Recordset
        Set rstmp = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE YEAR(DATEHIRED) < " & cboYear & "  AND EMPLEVEL = '" & "E" & "' AND DEPTCODE = '" & DEPT_NAME & "'")
        If Not (rstmp.BOF And rstmp.EOF) Then
            Do While Not rstmp.EOF
                BEG_CNT = BEG_CNT + 1
                rstmp.MoveNext
            Loop
        End If
        Set rstmp = Nothing
        Set rstmp = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE YEAR(RESIGNED) < " & cboYear & " AND EMPLEVEL = '" & "E" & "' AND DEPTCODE = '" & DEPT_NAME & "'")
        If Not (rstmp.BOF And rstmp.EOF) Then
            Do While Not rstmp.EOF
                BEG_CNT = BEG_CNT - 1
                rstmp.MoveNext
            Loop
        End If
        Set rstmp = Nothing
        'COMPUTE FOR THE BEGGINING

        MMNAME = 1

        Dim rsProbationary                             As ADODB.Recordset
        Do While MMNAME <= What_month(cboMonth)
            d1 = GetMyDate(MMNAME)

            Set rsProbationary = New ADODB.Recordset
            Set rsProbationary = gconDMIS.Execute("Select * from HRMS_EmpInfo where EMPLEVEL = '" & "E" & "' AND month(datehired) = " & What_month(MonthName(Month(d1))) & " and year(DATEHIRED) = " & cboYear & " AND DEPTCODE = '" & DEPT_NAME & "'")
            If Not rsProbationary.BOF And Not rsProbationary.EOF Then
                Do While Not rsProbationary.EOF
                    MonthHired_Prob = MMNAME
                    Select Case MonthHired_Prob
                        Case 1: Prob_Count_Jan = Prob_Count_Jan + 1
                        Case 2: Prob_Count_Feb = Prob_Count_Feb + 1
                        Case 3: Prob_Count_Mar = Prob_Count_Mar + 1
                        Case 4: Prob_Count_Apr = Prob_Count_Apr + 1
                        Case 5: Prob_Count_May = Prob_Count_May + 1
                        Case 6: Prob_Count_June = Prob_Count_June + 1
                        Case 7: Prob_Count_July = Prob_Count_July + 1
                        Case 8: Prob_Count_August = Prob_Count_August + 1
                        Case 9: Prob_Count_Sept = Prob_Count_Sept + 1
                        Case 10: Prob_Count_Oct = Prob_Count_Oct + 1
                        Case 11: Prob_Count_Nov = Prob_Count_Nov + 1
                        Case 12: Prob_Count_Dec = Prob_Count_Dec + 1
                    End Select
                    rsProbationary.MoveNext
                Loop
            End If
            MMNAME = MMNAME + 1
            Set rsProbationary = Nothing
        Loop

        Dim OJT_Count_Jan                              As Integer
        Dim OJT_Count_Feb                              As Integer
        Dim OJT_Count_Mar                              As Integer
        Dim OJT_Count_Apr                              As Integer
        Dim OJT_Count_May                              As Integer
        Dim OJT_Count_June                             As Integer
        Dim OJT_Count_July                             As Integer
        Dim OJT_Count_August                           As Integer
        Dim OJT_Count_Sept                             As Integer
        Dim OJT_Count_Oct                              As Integer
        Dim OJT_Count_Nov                              As Integer
        Dim OJT_Count_Dec                              As Integer
        Dim MonthHired_OJT                             As Integer
        '
        OJT_Count_Jan = 0: OJT_Count_Apr = 0: OJT_Count_August = 0: OJT_Count_Dec = 0
        OJT_Count_Feb = 0: OJT_Count_July = 0: OJT_Count_June = 0: OJT_Count_Mar = 0
        OJT_Count_May = 0: OJT_Count_Nov = 0: OJT_Count_Oct = 0: OJT_Count_Sept = 0

        MMNAME = 1
        Dim rsOJT                                      As ADODB.Recordset
        Do While MMNAME <= What_month(cboMonth)
            d1 = GetMyDate(MMNAME)
            Set rsOJT = New ADODB.Recordset
            Set rsOJT = gconDMIS.Execute("Select * from HRMS_EmpInfo where EMPLEVEL = '" & "A" & "' AND month(datehired) = " & What_month(MonthName(Month(d1))) & " and year(DATEHIRED) = " & cboYear & " AND DEPTCODE = '" & DEPT_NAME & "'")

            If Not rsOJT.BOF And Not rsOJT.EOF Then
                rsOJT.MoveFirst
                Do While Not rsOJT.EOF
                    MonthHired_OJT = MMNAME
                    Select Case MonthHired_OJT
                        Case 1: OJT_Count_Jan = OJT_Count_Jan + 1
                        Case 2: OJT_Count_Feb = OJT_Count_Feb + 1
                        Case 3: OJT_Count_Mar = OJT_Count_Mar + 1
                        Case 4: OJT_Count_Apr = OJT_Count_Apr + 1
                        Case 5: OJT_Count_May = OJT_Count_May + 1
                        Case 6: OJT_Count_June = OJT_Count_June + 1
                        Case 7: OJT_Count_July = OJT_Count_July + 1
                        Case 8: OJT_Count_August = OJT_Count_August + 1
                        Case 9: OJT_Count_Sept = OJT_Count_Sept + 1
                        Case 10: OJT_Count_Oct = OJT_Count_Oct + 1
                        Case 11: OJT_Count_Nov = OJT_Count_Nov + 1
                        Case 12: OJT_Count_Dec = OJT_Count_Dec + 1
                    End Select
                    rsOJT.MoveNext
                Loop
            End If
            Set rsOJT = Nothing
            MMNAME = MMNAME + 1
        Loop
        '
        Dim Contractual_Count_Jan                      As Integer
        Dim Contractual_Count_Feb                      As Integer
        Dim Contractual_Count_Mar                      As Integer
        Dim Contractual_Count_Apr                      As Integer
        Dim Contractual_Count_May                      As Integer
        Dim Contractual_Count_June                     As Integer
        Dim Contractual_Count_July                     As Integer
        Dim Contractual_Count_August                   As Integer
        Dim Contractual_Count_Sept                     As Integer
        Dim Contractual_Count_Oct                      As Integer
        Dim Contractual_Count_Nov                      As Integer
        Dim Contractual_Count_Dec                      As Integer
        Dim MonthHired_Cont                            As Integer
        '
        Contractual_Count_Jan = 0: Contractual_Count_Apr = 0: Contractual_Count_August = 0: Contractual_Count_Dec = 0
        Contractual_Count_Feb = 0: Contractual_Count_July = 0: Contractual_Count_June = 0: Contractual_Count_Mar = 0
        Contractual_Count_May = 0: Contractual_Count_Nov = 0: Contractual_Count_Oct = 0: Contractual_Count_Sept = 0

        Dim rsContractual                              As ADODB.Recordset
        MMNAME = 1

        Do While MMNAME <= What_month(cboMonth)
            d1 = GetMyDate(MMNAME)
            Set rsContractual = New ADODB.Recordset
            Set rsContractual = gconDMIS.Execute("Select * from HRMS_EmpInfo where EMPLEVEL = '" & "C" & "' AND month(datehired) = " & What_month(MonthName(Month(d1))) & " and year(DATEHIRED) = " & cboYear & " AND DEPTCODE = '" & DEPT_NAME & "'")
            If Not rsContractual.BOF And Not rsContractual.EOF Then
                rsContractual.MoveFirst
                Do While Not rsContractual.EOF
                    MonthHired_Cont = MMNAME
                    Select Case MonthHired_Cont
                        Case 1: Contractual_Count_Jan = Contractual_Count_Jan + 1
                        Case 2: Contractual_Count_Feb = Contractual_Count_Feb + 1
                        Case 3: Contractual_Count_Mar = Contractual_Count_Mar + 1
                        Case 4: Contractual_Count_Apr = Contractual_Count_Apr + 1
                        Case 5: Contractual_Count_May = Contractual_Count_May + 1
                        Case 6: Contractual_Count_June = Contractual_Count_June + 1
                        Case 7: Contractual_Count_July = Contractual_Count_July + 1
                        Case 8: Contractual_Count_August = Contractual_Count_August + 1
                        Case 9: Contractual_Count_Sept = Contractual_Count_Sept + 1
                        Case 10: Contractual_Count_Oct = Contractual_Count_Oct + 1
                        Case 11: Contractual_Count_Nov = Contractual_Count_Nov + 1
                        Case 12: Contractual_Count_Dec = Contractual_Count_Dec + 1
                    End Select
                    rsContractual.MoveNext
                Loop
            End If
            MMNAME = MMNAME + 1
            Set rsContractual = Nothing
        Loop

        Dim Finished_Contract_Count_Jan                As Integer
        Dim Finished_Contract_Count_Feb                As Integer
        Dim Finished_Contract_Count_Mar                As Integer
        Dim Finished_Contract_Count_Apr                As Integer
        Dim Finished_Contract_Count_May                As Integer
        Dim Finished_Contract_Count_June               As Integer
        Dim Finished_Contract_Count_July               As Integer
        Dim Finished_Contract_Count_August             As Integer
        Dim Finished_Contract_Count_Sept               As Integer
        Dim Finished_Contract_Count_Oct                As Integer
        Dim Finished_Contract_Count_Nov                As Integer
        Dim Finished_Contract_Count_Dec                As Integer
        Dim Month_Finished_Contract                    As Integer
        '
        Finished_Contract_Count_Jan = 0: Finished_Contract_Count_Apr = 0: Finished_Contract_Count_August = 0: Finished_Contract_Count_Dec = 0
        Finished_Contract_Count_Feb = 0: Finished_Contract_Count_July = 0: Finished_Contract_Count_June = 0: Finished_Contract_Count_Mar = 0
        Finished_Contract_Count_May = 0: Finished_Contract_Count_Nov = 0: Finished_Contract_Count_Oct = 0: Finished_Contract_Count_Sept = 0

        Dim rsFinished_Contract                        As ADODB.Recordset
        MMNAME = 1

        Do While MMNAME <= What_month(cboMonth)
            d1 = GetMyDate(MMNAME)
            Set rsFinished_Contract = New ADODB.Recordset
            Set rsFinished_Contract = gconDMIS.Execute("Select * from HRMS_EmpInfo where EMPLEVEL = 'C' AND month(RESIGNED) = " & What_month(MonthName(Month(d1))) & " and year(RESIGNED) = " & cboYear & " AND DEPTCODE = '" & DEPT_NAME & "'")

            If Not rsFinished_Contract.BOF And Not rsFinished_Contract.EOF Then
                rsFinished_Contract.MoveFirst
                Do While Not rsFinished_Contract.EOF
                    Month_Finished_Contract = MMNAME
                    Select Case Month_Finished_Contract
                        Case 1: Finished_Contract_Count_Jan = Finished_Contract_Count_Jan + 1
                        Case 2: Finished_Contract_Count_Feb = Finished_Contract_Count_Feb + 1
                        Case 3: Finished_Contract_Count_Mar = Finished_Contract_Count_Mar + 1
                        Case 4: Finished_Contract_Count_Apr = Finished_Contract_Count_Apr + 1
                        Case 5: Finished_Contract_Count_May = Finished_Contract_Count_May + 1
                        Case 6: Finished_Contract_Count_June = Finished_Contract_Count_June + 1
                        Case 7: Finished_Contract_Count_July = Finished_Contract_Count_July + 1
                        Case 8: Finished_Contract_Count_August = Finished_Contract_Count_August + 1
                        Case 9: Finished_Contract_Count_Sept = Finished_Contract_Count_Sept + 1
                        Case 10: Finished_Contract_Count_Oct = Finished_Contract_Count_Oct + 1
                        Case 11: Finished_Contract_Count_Nov = Finished_Contract_Count_Nov + 1
                        Case 12: Finished_Contract_Count_Dec = Finished_Contract_Count_Dec + 1
                    End Select
                    rsFinished_Contract.MoveNext
                Loop
            End If
            Set rsFinished_Contract = Nothing
            MMNAME = MMNAME + 1
        Loop
        '
        Dim Completed_Training_Count_Jan               As Integer
        Dim Completed_Training_Count_Feb               As Integer
        Dim Completed_Training_Count_Mar               As Integer
        Dim Completed_Training_Count_Apr               As Integer
        Dim Completed_Training_Count_May               As Integer
        Dim Completed_Training_Count_June              As Integer
        Dim Completed_Training_Count_July              As Integer
        Dim Completed_Training_Count_August            As Integer
        Dim Completed_Training_Count_Sept              As Integer
        Dim Completed_Training_Count_Oct               As Integer
        Dim Completed_Training_Count_Nov               As Integer
        Dim Completed_Training_Count_Dec               As Integer
        Dim Month_Completed_Training                   As Integer
        '
        Completed_Training_Count_Jan = 0: Completed_Training_Count_Apr = 0: Completed_Training_Count_August = 0: Completed_Training_Count_Dec = 0
        Completed_Training_Count_Feb = 0: Completed_Training_Count_July = 0: Completed_Training_Count_June = 0: Completed_Training_Count_Mar = 0
        Completed_Training_Count_May = 0: Completed_Training_Count_Nov = 0: Completed_Training_Count_Oct = 0: Completed_Training_Count_Sept = 0

        MMNAME = 1
        Dim rsCompleted_Training                       As ADODB.Recordset
        Do While MMNAME <= What_month(cboMonth)
            d1 = GetMyDate(MMNAME)
            Set rsCompleted_Training = New ADODB.Recordset
            Set rsCompleted_Training = gconDMIS.Execute("Select * from HRMS_EmpInfo where EMPLEVEL = 'A' AND month(RESIGNED) = " & What_month(MonthName(Month(d1))) & " and year(RESIGNED) = " & cboYear & " AND DEPTCODE = '" & DEPT_NAME & "'")
            If Not rsCompleted_Training.BOF And Not rsCompleted_Training.EOF Then
                rsCompleted_Training.MoveFirst
                Do While Not rsCompleted_Training.EOF
                    Month_Completed_Training = MMNAME
                    Select Case Month_Completed_Training
                        Case 1: Completed_Training_Count_Jan = Completed_Training_Count_Jan + 1
                        Case 2: Completed_Training_Count_Feb = Completed_Training_Count_Feb + 1
                        Case 3: Completed_Training_Count_Mar = Completed_Training_Count_Mar + 1
                        Case 4: Completed_Training_Count_Apr = Completed_Training_Count_Apr + 1
                        Case 5: Completed_Training_Count_May = Completed_Training_Count_May + 1
                        Case 6: Completed_Training_Count_June = Completed_Training_Count_June + 1
                        Case 7: Completed_Training_Count_July = Completed_Training_Count_July + 1
                        Case 8: Completed_Training_Count_August = Completed_Training_Count_August + 1
                        Case 9: Completed_Training_Count_Sept = Completed_Training_Count_Sept + 1
                        Case 10: Completed_Training_Count_Oct = Completed_Training_Count_Oct + 1
                        Case 11: Completed_Training_Count_Nov = Completed_Training_Count_Nov + 1
                        Case 12: Completed_Training_Count_Dec = Completed_Training_Count_Dec + 1
                    End Select
                    rsCompleted_Training.MoveNext
                Loop
            End If
            MMNAME = MMNAME + 1
            Set rsCompleted_Training = Nothing
        Loop
        '
        Dim NO_OF_SEPAREATION_FOR_Jan                  As Integer
        Dim NO_OF_SEPAREATION_FOR_Feb                  As Integer
        Dim NO_OF_SEPAREATION_FOR_Mar                  As Integer
        Dim NO_OF_SEPAREATION_FOR_Apr                  As Integer
        Dim NO_OF_SEPAREATION_FOR_May                  As Integer
        Dim NO_OF_SEPAREATION_FOR_June                 As Integer
        Dim NO_OF_SEPAREATION_FOR_July                 As Integer
        Dim NO_OF_SEPAREATION_FOR_August               As Integer
        Dim NO_OF_SEPAREATION_FOR_Sept                 As Integer
        Dim NO_OF_SEPAREATION_FOR_Oct                  As Integer
        Dim NO_OF_SEPAREATION_FOR_Nov                  As Integer
        Dim NO_OF_SEPAREATION_FOR_Dec                  As Integer
        Dim NO_OF_SEPAREATION_FOR                      As Integer
        '
        NO_OF_SEPAREATION_FOR_Jan = 0: NO_OF_SEPAREATION_FOR_Apr = 0: NO_OF_SEPAREATION_FOR_August = 0: NO_OF_SEPAREATION_FOR_Dec = 0
        NO_OF_SEPAREATION_FOR_Feb = 0: NO_OF_SEPAREATION_FOR_July = 0: NO_OF_SEPAREATION_FOR_June = 0: NO_OF_SEPAREATION_FOR_Mar = 0
        NO_OF_SEPAREATION_FOR_May = 0: NO_OF_SEPAREATION_FOR_Nov = 0: NO_OF_SEPAREATION_FOR_Oct = 0: NO_OF_SEPAREATION_FOR_Sept = 0

        MMNAME = 1

        Dim rsNO_OF_SEPAREATION                        As ADODB.Recordset

        Do While MMNAME <= What_month(cboMonth)
            d1 = GetMyDate(MMNAME)
            Set rsNO_OF_SEPAREATION = New ADODB.Recordset
            Set rsNO_OF_SEPAREATION = gconDMIS.Execute("Select * from HRMS_EmpInfo where emplevel = 'E' AND month(RESIGNED) = " & What_month(MonthName(Month(d1))) & " and year(RESIGNED) = " & cboYear & " AND DEPTCODE = '" & DEPT_NAME & "'")
            If Not rsNO_OF_SEPAREATION.BOF And Not rsNO_OF_SEPAREATION.EOF Then
                rsNO_OF_SEPAREATION.MoveFirst
                Do While Not rsNO_OF_SEPAREATION.EOF
                    NO_OF_SEPAREATION_FOR = MMNAME
                    Select Case NO_OF_SEPAREATION_FOR
                        Case 1: NO_OF_SEPAREATION_FOR_Jan = NO_OF_SEPAREATION_FOR_Jan + 1
                        Case 2: NO_OF_SEPAREATION_FOR_Feb = NO_OF_SEPAREATION_FOR_Feb + 1
                        Case 3: NO_OF_SEPAREATION_FOR_Mar = NO_OF_SEPAREATION_FOR_Mar + 1
                        Case 4: NO_OF_SEPAREATION_FOR_Apr = NO_OF_SEPAREATION_FOR_Apr + 1
                        Case 5: NO_OF_SEPAREATION_FOR_May = NO_OF_SEPAREATION_FOR_May + 1
                        Case 6: NO_OF_SEPAREATION_FOR_June = NO_OF_SEPAREATION_FOR_June + 1
                        Case 7: NO_OF_SEPAREATION_FOR_July = NO_OF_SEPAREATION_FOR_July + 1
                        Case 8: NO_OF_SEPAREATION_FOR_August = NO_OF_SEPAREATION_FOR_August + 1
                        Case 9: NO_OF_SEPAREATION_FOR_Sept = NO_OF_SEPAREATION_FOR_Sept + 1
                        Case 10: NO_OF_SEPAREATION_FOR_Oct = NO_OF_SEPAREATION_FOR_Oct + 1
                        Case 11: NO_OF_SEPAREATION_FOR_Nov = NO_OF_SEPAREATION_FOR_Nov + 1
                        Case 12: NO_OF_SEPAREATION_FOR_Dec = NO_OF_SEPAREATION_FOR_Dec + 1
                    End Select
                    rsNO_OF_SEPAREATION.MoveNext
                Loop
            End If
            MMNAME = MMNAME + 1
            Set rsNO_OF_SEPAREATION = Nothing
        Loop

        rptActualManningReport.Formulas(79) = "BEG = " & BEG_CNT & ""
        rptActualManningReport.Formulas(15) = "Probationary_For_Jan = " & Prob_Count_Jan & ""
        rptActualManningReport.Formulas(16) = "Probationary_For_Feb = " & Prob_Count_Feb & ""
        rptActualManningReport.Formulas(17) = "Probationary_For_Mar = " & Prob_Count_Mar & ""
        rptActualManningReport.Formulas(18) = "Probationary_For_Apr = " & Prob_Count_Apr & ""
        rptActualManningReport.Formulas(19) = "Probationary_For_May = " & Prob_Count_May & ""
        rptActualManningReport.Formulas(20) = "Probationary_For_June = " & Prob_Count_June & ""
        rptActualManningReport.Formulas(21) = "Probationary_For_July = " & Prob_Count_July & ""
        rptActualManningReport.Formulas(22) = "Probationary_For_Aug = " & Prob_Count_August & ""
        rptActualManningReport.Formulas(23) = "Probationary_For_Sept = " & Prob_Count_Sept & ""
        rptActualManningReport.Formulas(24) = "Probationary_For_Oct = " & Prob_Count_Oct & ""
        rptActualManningReport.Formulas(25) = "Probationary_For_Nov = " & Prob_Count_Nov & ""
        rptActualManningReport.Formulas(26) = "Probationary_For_Dec = " & Prob_Count_Dec & ""

        rptActualManningReport.Formulas(27) = "OJT_For_Jan = " & OJT_Count_Jan & ""
        rptActualManningReport.Formulas(28) = "OJT_For_Feb = " & OJT_Count_Feb & ""
        rptActualManningReport.Formulas(29) = "OJT_For_Mar = " & OJT_Count_Mar & ""
        rptActualManningReport.Formulas(30) = "OJT_For_Apr = " & OJT_Count_Apr & ""
        rptActualManningReport.Formulas(31) = "OJT_For_May = " & OJT_Count_May & ""
        rptActualManningReport.Formulas(32) = "OJT_For_June = " & OJT_Count_June & ""
        rptActualManningReport.Formulas(33) = "OJT_For_July = " & OJT_Count_July & ""
        rptActualManningReport.Formulas(34) = "OJT_For_Aug = " & OJT_Count_August & ""
        rptActualManningReport.Formulas(35) = "OJT_For_Sept = " & OJT_Count_Sept & ""
        rptActualManningReport.Formulas(36) = "OJT_For_Oct = " & OJT_Count_Oct & ""
        rptActualManningReport.Formulas(37) = "OJT_For_Nov = " & OJT_Count_Nov & ""
        rptActualManningReport.Formulas(38) = "OJT_For_Dec = " & OJT_Count_Dec & ""


        rptActualManningReport.Formulas(2) = "YearReport = '" & cboYear.Text & "'"
        rptActualManningReport.Formulas(3) = "Contractual_For_Jan = " & Contractual_Count_Jan & ""
        rptActualManningReport.Formulas(4) = "Contractual_For_Feb = " & Contractual_Count_Feb & ""
        rptActualManningReport.Formulas(5) = "Contractual_For_Mar = " & Contractual_Count_Mar & ""
        rptActualManningReport.Formulas(6) = "Contractual_For_Apr = " & Contractual_Count_Apr & ""
        rptActualManningReport.Formulas(7) = "Contractual_For_May = " & Contractual_Count_May & ""
        rptActualManningReport.Formulas(8) = "Contractual_For_June = " & Contractual_Count_June & ""
        rptActualManningReport.Formulas(9) = "Contractual_For_July = " & Contractual_Count_July & ""
        rptActualManningReport.Formulas(10) = "Contractual_For_Aug = " & Contractual_Count_August & ""
        rptActualManningReport.Formulas(11) = "Contractual_For_Sept = " & Contractual_Count_Sept & ""
        rptActualManningReport.Formulas(12) = "Contractual_For_Oct = " & Contractual_Count_Oct & ""
        rptActualManningReport.Formulas(13) = "Contractual_For_Nov = " & Contractual_Count_Nov & ""
        rptActualManningReport.Formulas(14) = "Contractual_For_Dec = " & Contractual_Count_Dec & ""


        rptActualManningReport.Formulas(39) = "Finished_Contract_For_Jan = " & Finished_Contract_Count_Jan & ""
        rptActualManningReport.Formulas(40) = "Finished_Contract_For_Feb = " & Finished_Contract_Count_Feb & ""
        rptActualManningReport.Formulas(41) = "Finished_Contract_For_Mar = " & Finished_Contract_Count_Mar & ""
        rptActualManningReport.Formulas(42) = "Finished_Contract_For_Apr = " & Finished_Contract_Count_Apr & ""
        rptActualManningReport.Formulas(43) = "Finished_Contract_For_May = " & Finished_Contract_Count_May & ""
        rptActualManningReport.Formulas(44) = "Finished_Contract_For_June = " & Finished_Contract_Count_June & ""
        rptActualManningReport.Formulas(45) = "Finished_Contract_For_July = " & Finished_Contract_Count_July & ""
        rptActualManningReport.Formulas(46) = "Finished_Contract_For_Aug = " & Finished_Contract_Count_August & ""
        rptActualManningReport.Formulas(47) = "Finished_Contract_For_Sept = " & Finished_Contract_Count_Sept & ""
        rptActualManningReport.Formulas(48) = "Finished_Contract_For_Oct = " & Finished_Contract_Count_Oct & ""
        rptActualManningReport.Formulas(49) = "Finished_Contract_For_Nov = " & Finished_Contract_Count_Nov & ""
        rptActualManningReport.Formulas(50) = "Finished_Contract_For_Dec = " & Finished_Contract_Count_Dec & ""

        rptActualManningReport.Formulas(51) = "OJT_Completed_For_Jan = " & Completed_Training_Count_Jan & ""
        rptActualManningReport.Formulas(52) = "OJT_Completed_For_Feb = " & Completed_Training_Count_Feb & ""
        rptActualManningReport.Formulas(53) = "OJT_Completed_For_Mar = " & Completed_Training_Count_Mar & ""
        rptActualManningReport.Formulas(54) = "OJT_Completed_For_Apr = " & Completed_Training_Count_Apr & ""
        rptActualManningReport.Formulas(55) = "OJT_Completed_For_May = " & Completed_Training_Count_May & ""
        rptActualManningReport.Formulas(56) = "OJT_Completed_For_June = " & Completed_Training_Count_June & ""
        rptActualManningReport.Formulas(57) = "OJT_Completed_For_July = " & Completed_Training_Count_July & ""
        rptActualManningReport.Formulas(58) = "OJT_Completed_For_Aug = " & Completed_Training_Count_August & ""
        rptActualManningReport.Formulas(59) = "OJT_Completed_For_Sept = " & Completed_Training_Count_Sept & ""
        rptActualManningReport.Formulas(60) = "OJT_Completed_For_Oct = " & Completed_Training_Count_Oct & ""
        rptActualManningReport.Formulas(61) = "OJT_Completed_For_Nov = " & Completed_Training_Count_Nov & ""
        rptActualManningReport.Formulas(62) = "OJT_Completed_For_Dec = " & Completed_Training_Count_Dec & ""

        rptActualManningReport.Formulas(63) = "NO_OF_SEPARATIONS_FOR_JAN = " & NO_OF_SEPAREATION_FOR_Jan & ""
        rptActualManningReport.Formulas(64) = "NO_OF_SEPARATIONS_FOR_FEB = " & NO_OF_SEPAREATION_FOR_Feb & ""
        rptActualManningReport.Formulas(65) = "NO_OF_SEPARATIONS_FOR_MAR = " & NO_OF_SEPAREATION_FOR_Mar & ""
        rptActualManningReport.Formulas(66) = "NO_OF_SEPARATIONS_FOR_APR = " & NO_OF_SEPAREATION_FOR_Apr & ""
        rptActualManningReport.Formulas(67) = "NO_OF_SEPARATIONS_FOR_MAY = " & NO_OF_SEPAREATION_FOR_May & ""
        rptActualManningReport.Formulas(68) = "NO_OF_SEPARATIONS_FOR_JUNE = " & NO_OF_SEPAREATION_FOR_June & ""
        rptActualManningReport.Formulas(69) = "NO_OF_SEPARATIONS_FOR_JULY = " & NO_OF_SEPAREATION_FOR_July & ""
        rptActualManningReport.Formulas(70) = "NO_OF_SEPARATIONS_FOR_AUG = " & NO_OF_SEPAREATION_FOR_August & ""
        rptActualManningReport.Formulas(71) = "NO_OF_SEPARATIONS_FOR_SEPT = " & NO_OF_SEPAREATION_FOR_Sept & ""
        rptActualManningReport.Formulas(72) = "NO_OF_SEPARATIONS_FOR_OCT = " & NO_OF_SEPAREATION_FOR_Oct & ""
        rptActualManningReport.Formulas(73) = "NO_OF_SEPARATIONS_FOR_NOV = " & NO_OF_SEPAREATION_FOR_Nov & ""
        rptActualManningReport.Formulas(74) = "NO_OF_SEPARATIONS_FOR_DEC = " & NO_OF_SEPAREATION_FOR_Dec & ""

        If What_month(cboMonth.Text) = 1 Then rptActualManningReport.Formulas(80) = "T_JAN = " & 0 & ""
        If What_month(cboMonth.Text) = 2 Then rptActualManningReport.Formulas(80) = "T_FEB = " & 0 & ""
        If What_month(cboMonth.Text) = 3 Then rptActualManningReport.Formulas(80) = "T_MAR = " & 0 & ""
        If What_month(cboMonth.Text) = 4 Then rptActualManningReport.Formulas(80) = "T_APR = " & 0 & ""
        If What_month(cboMonth.Text) = 5 Then rptActualManningReport.Formulas(80) = "T_MAY = " & 0 & ""
        If What_month(cboMonth.Text) = 6 Then rptActualManningReport.Formulas(80) = "T_JUN = " & 0 & ""
        If What_month(cboMonth.Text) = 7 Then rptActualManningReport.Formulas(80) = "T_JUL = " & 0 & ""
        If What_month(cboMonth.Text) = 8 Then rptActualManningReport.Formulas(80) = "T_AUG = " & 0 & ""
        If What_month(cboMonth.Text) = 9 Then rptActualManningReport.Formulas(80) = "T_SEP = " & 0 & ""
        If What_month(cboMonth.Text) = 10 Then rptActualManningReport.Formulas(80) = "T_OCT = " & 0 & ""
        If What_month(cboMonth.Text) = 11 Then rptActualManningReport.Formulas(80) = "T_NOV = " & 0 & ""


        '        'JUN 02/05/2008
        rptActualManningReport.Formulas(75) = "COMPANYNAME = '" & COMPANY_NAME & "'"
        rptActualManningReport.Formulas(76) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
        rptActualManningReport.Formulas(77) = "printedby = '" & LOGNAME & "'"
        rptActualManningReport.Formulas(78) = "MONTHNAME = '" & cboMonth & "'"

        PrintSQLReport rptActualManningReport, CSMS_REPORT_PATH & "Actual_Manning_Report.rpt", "Year({HRMS_EmpInfo.datehired}) = " & cboYear.Text, CSMS_REPORT_CONNECTION, 1
        LogAudit "V", "ACTUAL MANNING REPORT", cboYear

        Screen.MousePointer = 0
    Else
        ShowNoRecord
        Screen.MousePointer = 0
        Exit Sub
    End If
    Exit Sub
Errorcode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    'On Error GoTo Errorcode

    Dim DEPT_NAME                                      As String
    Dim rsHRMS                                         As New ADODB.Recordset
    Dim cnt                                            As Integer

    DEPT_NAME = GetDeptName

    '    If MsgBox("Print in Excel", vbQuestion + vbYesNo, "CSMS") = vbNo Then
    '        Call cmdPrint_Click
    '        Exit Sub
    '    End If
    If Option1.Value = 1 Then
        Screen.MousePointer = 11

        cnt = 15

        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "Actual Manning Report Summary.xls")
        Set xlSheet = xlBook.Worksheets(1)

        Set rsHRMS = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE DEPTCODE = '" & DEPT_NAME & "' ORDER BY LASTNAME")
        If Not (rsHRMS.BOF And rsHRMS.EOF) Then
            xlSheet.Cells(9, "A") = "Dealer :" & COMPANY_NAME
            xlSheet.Cells(11, "A") = "Month of : " & MonthName(Month(Date))
            xlSheet.Cells(9, "P") = Date
            xlSheet.Cells(35, "K") = GENERAL_MANAGER

            Do While Not rsHRMS.EOF
                xlSheet.Cells(cnt, "B") = Null2String(rsHRMS!lastname) & ", " & Null2String(rsHRMS!Firstname) & " " & Left(Null2String(rsHRMS!MIDDLENAME), 1) & "."
                xlSheet.Cells(cnt, "C") = Null2String(rsHRMS!Position)

                If Not Null2String(rsHRMS!DATEHIRED) = "" Then
                    xlSheet.Cells(cnt, "D") = Month(Null2String(rsHRMS!DATEHIRED))
                    xlSheet.Cells(cnt, "E") = Day(Null2String(rsHRMS!DATEHIRED))
                    xlSheet.Cells(cnt, "F") = Year(Null2String(rsHRMS!DATEHIRED))

                    If Null2String(rsHRMS!EMPLEVEL) = "E" Or Null2String(rsHRMS!EMPLEVEL) = "M" Then
                        If Month(rsHRMS!DATEHIRED) = 1 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "G") = "H"
                        If Month(rsHRMS!DATEHIRED) = 2 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "H") = "H"
                        If Month(rsHRMS!DATEHIRED) = 3 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "I") = "H"
                        If Month(rsHRMS!DATEHIRED) = 4 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "J") = "H"
                        If Month(rsHRMS!DATEHIRED) = 5 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "K") = "H"
                        If Month(rsHRMS!DATEHIRED) = 6 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "L") = "H"
                        If Month(rsHRMS!DATEHIRED) = 7 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "M") = "H"
                        If Month(rsHRMS!DATEHIRED) = 8 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "N") = "H"
                        If Month(rsHRMS!DATEHIRED) = 9 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "O") = "H"
                        If Month(rsHRMS!DATEHIRED) = 10 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "P") = "H"
                        If Month(rsHRMS!DATEHIRED) = 11 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "Q") = "H"
                        If Month(rsHRMS!DATEHIRED) = 12 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "R") = "H"

                        If Not Null2String(rsHRMS!RESIGNED) = "" Then
                            If Month(rsHRMS!RESIGNED) = 1 And Year(rsHRMS!RESIGNED) = Year(Date) Then xlSheet.Cells(cnt, "G") = "R"
                            If Month(rsHRMS!RESIGNED) = 2 And Year(rsHRMS!RESIGNED) = Year(Date) Then xlSheet.Cells(cnt, "H") = "R"
                            If Month(rsHRMS!RESIGNED) = 3 And Year(rsHRMS!RESIGNED) = Year(Date) Then xlSheet.Cells(cnt, "I") = "R"
                            If Month(rsHRMS!RESIGNED) = 4 And Year(rsHRMS!RESIGNED) = Year(Date) Then xlSheet.Cells(cnt, "J") = "R"
                            If Month(rsHRMS!RESIGNED) = 5 And Year(rsHRMS!RESIGNED) = Year(Date) Then xlSheet.Cells(cnt, "K") = "R"
                            If Month(rsHRMS!RESIGNED) = 6 And Year(rsHRMS!RESIGNED) = Year(Date) Then xlSheet.Cells(cnt, "L") = "R"
                            If Month(rsHRMS!RESIGNED) = 7 And Year(rsHRMS!RESIGNED) = Year(Date) Then xlSheet.Cells(cnt, "M") = "R"
                            If Month(rsHRMS!RESIGNED) = 8 And Year(rsHRMS!RESIGNED) = Year(Date) Then xlSheet.Cells(cnt, "N") = "R"
                            If Month(rsHRMS!RESIGNED) = 9 And Year(rsHRMS!RESIGNED) = Year(Date) Then xlSheet.Cells(cnt, "O") = "R"
                            If Month(rsHRMS!RESIGNED) = 10 And Year(rsHRMS!RESIGNED) = Year(Date) Then xlSheet.Cells(cnt, "P") = "R"
                            If Month(rsHRMS!RESIGNED) = 11 And Year(rsHRMS!RESIGNED) = Year(Date) Then xlSheet.Cells(cnt, "Q") = "R"
                            If Month(rsHRMS!RESIGNED) = 12 And Year(rsHRMS!RESIGNED) = Year(Date) Then xlSheet.Cells(cnt, "R") = "R"
                        End If
                    End If
                    If Null2String(rsHRMS!EMPLEVEL) = "C" Then
                        If Month(rsHRMS!DATEHIRED) = 1 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "G") = "C"
                        If Month(rsHRMS!DATEHIRED) = 2 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "H") = "C"
                        If Month(rsHRMS!DATEHIRED) = 3 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "I") = "C"
                        If Month(rsHRMS!DATEHIRED) = 4 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "J") = "C"
                        If Month(rsHRMS!DATEHIRED) = 5 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "K") = "C"
                        If Month(rsHRMS!DATEHIRED) = 6 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "L") = "C"
                        If Month(rsHRMS!DATEHIRED) = 7 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "M") = "C"
                        If Month(rsHRMS!DATEHIRED) = 8 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "N") = "C"
                        If Month(rsHRMS!DATEHIRED) = 9 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "O") = "C"
                        If Month(rsHRMS!DATEHIRED) = 10 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "P") = "C"
                        If Month(rsHRMS!DATEHIRED) = 11 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "Q") = "C"
                        If Month(rsHRMS!DATEHIRED) = 12 And Year(rsHRMS!DATEHIRED) = Year(Date) Then xlSheet.Cells(cnt, "R") = "C"

                        If Not Null2String(rsHRMS!RESIGNED) = "" Then
                            If Month(rsHRMS!RESIGNED) = 1 Then xlSheet.Cells(cnt, "G") = "R"
                            If Month(rsHRMS!RESIGNED) = 2 Then xlSheet.Cells(cnt, "H") = "R"
                            If Month(rsHRMS!RESIGNED) = 3 Then xlSheet.Cells(cnt, "I") = "R"
                            If Month(rsHRMS!RESIGNED) = 4 Then xlSheet.Cells(cnt, "J") = "R"
                            If Month(rsHRMS!RESIGNED) = 5 Then xlSheet.Cells(cnt, "K") = "R"
                            If Month(rsHRMS!RESIGNED) = 6 Then xlSheet.Cells(cnt, "L") = "R"
                            If Month(rsHRMS!RESIGNED) = 7 Then xlSheet.Cells(cnt, "M") = "R"
                            If Month(rsHRMS!RESIGNED) = 8 Then xlSheet.Cells(cnt, "N") = "R"
                            If Month(rsHRMS!RESIGNED) = 9 Then xlSheet.Cells(cnt, "O") = "R"
                            If Month(rsHRMS!RESIGNED) = 10 Then xlSheet.Cells(cnt, "P") = "R"
                            If Month(rsHRMS!RESIGNED) = 11 Then xlSheet.Cells(cnt, "Q") = "R"
                            If Month(rsHRMS!RESIGNED) = 12 Then xlSheet.Cells(cnt, "R") = "R"
                        End If
                    End If
                End If

                cnt = cnt + 1
                If cnt > 29 Then
                    cnt = 15
                    xlApp.Visible = True
                    Set xlApp = Nothing

                    Set xlApp = CreateObject("Excel.Application")
                    Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "Actual Manning Report Summary.xls")
                    Set xlSheet = xlBook.Worksheets(1)

                    xlSheet.Cells(9, "A") = "Dealer : " & COMPANY_NAME
                    xlSheet.Cells(11, "A") = "Month of : " & MonthName(Month(Date))
                    xlSheet.Cells(9, "P") = Date
                    xlSheet.Cells(35, "K") = GENERAL_MANAGER
                End If

                rsHRMS.MoveNext
            Loop
        End If

        xlApp.Visible = True
        Set xlApp = Nothing

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "ACTUAL MANNING REPORT", "", "", "", "SUMMARY", "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        'LogAudit "V", "ACTUAL MANNING REPORT SUMMARY", cboYear

        Screen.MousePointer = 0
        Exit Sub
    Else
        Screen.MousePointer = 11
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "Actual Manning Report.xls")
        Set xlSheet = xlBook.Worksheets(1)

        'COMPUTE FOR THE BEGGINING EXISTING EMPLOYEE LAST YEAR
        Dim BEG_CNT                                    As Integer
        Dim rstmp                                      As New ADODB.Recordset
        Set rstmp = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE YEAR(DATEHIRED) < " & cboYear & "  AND EMPLEVEL = '" & "E" & "' AND DEPTCODE = '" & DEPT_NAME & "'")
        If Not (rstmp.BOF And rstmp.EOF) Then
            Do While Not rstmp.EOF
                BEG_CNT = BEG_CNT + 1
                rstmp.MoveNext
            Loop
        End If
        Set rstmp = Nothing
        Set rstmp = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE YEAR(RESIGNED) < " & cboYear & " AND EMPLEVEL = '" & "E" & "' AND DEPTCODE = '" & DEPT_NAME & "'")
        If Not (rstmp.BOF And rstmp.EOF) Then
            Do While Not rstmp.EOF
                BEG_CNT = BEG_CNT - 1
                rstmp.MoveNext
            Loop
        End If
        Set rstmp = Nothing

        xlSheet.Cells(14, "C") = BEG_CNT
        xlSheet.Cells(8, "A") = "Dealer : " & COMPANY_NAME
        xlSheet.Cells(10, "A") = "Month Of " & cboMonth
        xlSheet.Cells(10, "K") = Date
        xlSheet.Cells(27, "J") = GENERAL_MANAGER
        'COMPUTE FOR THE BEGGINING

        Dim MMNAME                                     As Integer
        Dim mDEYT                                      As String
        Dim tDEYT                                      As String
        Dim d1                                         As String
        Dim rsActualManning                            As ADODB.Recordset

        Set rsActualManning = New ADODB.Recordset
        Set rsActualManning = gconDMIS.Execute("Select * from HRMS_EmpInfo where Year(datehired) <= '" & cboYear.Text & "' and deptcode = '" & DEPT_NAME & "'")
        If Not (rsActualManning.BOF And rsActualManning.EOF) Then
            'HIRED PROBI =============================================================================================================
            MMNAME = 1

            Dim rsProbationary                         As ADODB.Recordset
            Dim Prob_Count_Jan                         As Integer
            Dim Prob_Count_Feb                         As Integer
            Dim Prob_Count_Mar                         As Integer
            Dim Prob_Count_Apr                         As Integer
            Dim Prob_Count_May                         As Integer
            Dim Prob_Count_June                        As Integer
            Dim Prob_Count_July                        As Integer
            Dim Prob_Count_August                      As Integer
            Dim Prob_Count_Sept                        As Integer
            Dim Prob_Count_Oct                         As Integer
            Dim Prob_Count_Nov                         As Integer
            Dim Prob_Count_Dec                         As Integer
            Dim MonthHired_Prob                        As Integer

            Do While MMNAME <= What_month(cboMonth)
                d1 = GetMyDate(MMNAME)

                Set rsProbationary = New ADODB.Recordset
                Set rsProbationary = gconDMIS.Execute("Select * from HRMS_EmpInfo where EMPLEVEL = '" & "E" & "' AND month(datehired) = " & What_month(MonthName(Month(d1))) & " and year(DATEHIRED) = " & cboYear & " AND DEPTCODE = '" & DEPT_NAME & "'")
                If Not rsProbationary.BOF And Not rsProbationary.EOF Then
                    Do While Not rsProbationary.EOF
                        MonthHired_Prob = MMNAME
                        Select Case MonthHired_Prob
                            Case 1: Prob_Count_Jan = Prob_Count_Jan + 1
                            Case 2: Prob_Count_Feb = Prob_Count_Feb + 1
                            Case 3: Prob_Count_Mar = Prob_Count_Mar + 1
                            Case 4: Prob_Count_Apr = Prob_Count_Apr + 1
                            Case 5: Prob_Count_May = Prob_Count_May + 1
                            Case 6: Prob_Count_June = Prob_Count_June + 1
                            Case 7: Prob_Count_July = Prob_Count_July + 1
                            Case 8: Prob_Count_August = Prob_Count_August + 1
                            Case 9: Prob_Count_Sept = Prob_Count_Sept + 1
                            Case 10: Prob_Count_Oct = Prob_Count_Oct + 1
                            Case 11: Prob_Count_Nov = Prob_Count_Nov + 1
                            Case 12: Prob_Count_Dec = Prob_Count_Dec + 1
                        End Select
                        rsProbationary.MoveNext
                    Loop
                End If
                MMNAME = MMNAME + 1
                Set rsProbationary = Nothing
            Loop

            xlSheet.Cells(15, "C") = Prob_Count_Jan
            xlSheet.Cells(15, "D") = Prob_Count_Feb
            xlSheet.Cells(15, "E") = Prob_Count_Mar
            xlSheet.Cells(15, "F") = Prob_Count_Apr
            xlSheet.Cells(15, "G") = Prob_Count_May
            xlSheet.Cells(15, "H") = Prob_Count_June
            xlSheet.Cells(15, "I") = Prob_Count_July
            xlSheet.Cells(15, "J") = Prob_Count_August
            xlSheet.Cells(15, "K") = Prob_Count_Sept
            xlSheet.Cells(15, "L") = Prob_Count_Oct
            xlSheet.Cells(15, "M") = Prob_Count_Nov
            xlSheet.Cells(15, "N") = Prob_Count_Dec
            'HIRED PROBI =============================================================================================================

            'HIRED OJT =============================================================================================================
            Dim OJT_Count_Jan                          As Integer
            Dim OJT_Count_Feb                          As Integer
            Dim OJT_Count_Mar                          As Integer
            Dim OJT_Count_Apr                          As Integer
            Dim OJT_Count_May                          As Integer
            Dim OJT_Count_June                         As Integer
            Dim OJT_Count_July                         As Integer
            Dim OJT_Count_August                       As Integer
            Dim OJT_Count_Sept                         As Integer
            Dim OJT_Count_Oct                          As Integer
            Dim OJT_Count_Nov                          As Integer
            Dim OJT_Count_Dec                          As Integer
            Dim MonthHired_OJT                         As Integer

            MMNAME = 1
            Dim rsOJT                                  As ADODB.Recordset
            Do While MMNAME <= What_month(cboMonth)
                d1 = GetMyDate(MMNAME)
                Set rsOJT = New ADODB.Recordset
                Set rsOJT = gconDMIS.Execute("Select * from HRMS_EmpInfo where EMPLEVEL = '" & "A" & "' AND month(datehired) = " & What_month(MonthName(Month(d1))) & " and year(DATEHIRED) = " & cboYear & " AND DEPTCODE = '" & DEPT_NAME & "'")

                If Not rsOJT.BOF And Not rsOJT.EOF Then
                    rsOJT.MoveFirst
                    Do While Not rsOJT.EOF
                        MonthHired_OJT = MMNAME
                        Select Case MonthHired_OJT
                            Case 1: OJT_Count_Jan = OJT_Count_Jan + 1
                            Case 2: OJT_Count_Feb = OJT_Count_Feb + 1
                            Case 3: OJT_Count_Mar = OJT_Count_Mar + 1
                            Case 4: OJT_Count_Apr = OJT_Count_Apr + 1
                            Case 5: OJT_Count_May = OJT_Count_May + 1
                            Case 6: OJT_Count_June = OJT_Count_June + 1
                            Case 7: OJT_Count_July = OJT_Count_July + 1
                            Case 8: OJT_Count_August = OJT_Count_August + 1
                            Case 9: OJT_Count_Sept = OJT_Count_Sept + 1
                            Case 10: OJT_Count_Oct = OJT_Count_Oct + 1
                            Case 11: OJT_Count_Nov = OJT_Count_Nov + 1
                            Case 12: OJT_Count_Dec = OJT_Count_Dec + 1
                        End Select
                        rsOJT.MoveNext
                    Loop
                End If
                Set rsOJT = Nothing
                MMNAME = MMNAME + 1
            Loop
            xlSheet.Cells(16, "C") = OJT_Count_Jan
            xlSheet.Cells(16, "D") = OJT_Count_Feb
            xlSheet.Cells(16, "E") = OJT_Count_Mar
            xlSheet.Cells(16, "F") = OJT_Count_Apr
            xlSheet.Cells(16, "G") = OJT_Count_May
            xlSheet.Cells(16, "H") = OJT_Count_June
            xlSheet.Cells(16, "I") = OJT_Count_July
            xlSheet.Cells(16, "J") = OJT_Count_August
            xlSheet.Cells(16, "K") = OJT_Count_Sept
            xlSheet.Cells(16, "L") = OJT_Count_Oct
            xlSheet.Cells(16, "M") = OJT_Count_Nov
            xlSheet.Cells(16, "N") = OJT_Count_Dec
            'HIRED OJT =============================================================================================================


            'HIRED CONTRACT =============================================================================================================
            Dim Contractual_Count_Jan                  As Integer
            Dim Contractual_Count_Feb                  As Integer
            Dim Contractual_Count_Mar                  As Integer
            Dim Contractual_Count_Apr                  As Integer
            Dim Contractual_Count_May                  As Integer
            Dim Contractual_Count_June                 As Integer
            Dim Contractual_Count_July                 As Integer
            Dim Contractual_Count_August               As Integer
            Dim Contractual_Count_Sept                 As Integer
            Dim Contractual_Count_Oct                  As Integer
            Dim Contractual_Count_Nov                  As Integer
            Dim Contractual_Count_Dec                  As Integer
            Dim MonthHired_Cont                        As Integer
            '
            Dim rsContractual                          As ADODB.Recordset
            MMNAME = 1

            Do While MMNAME <= What_month(cboMonth)
                d1 = GetMyDate(MMNAME)
                Set rsContractual = New ADODB.Recordset
                Set rsContractual = gconDMIS.Execute("Select * from HRMS_EmpInfo where EMPLEVEL = '" & "C" & "' AND month(datehired) = " & What_month(MonthName(Month(d1))) & " and year(DATEHIRED) = " & cboYear & " AND DEPTCODE = '" & DEPT_NAME & "'")
                If Not rsContractual.BOF And Not rsContractual.EOF Then
                    rsContractual.MoveFirst
                    Do While Not rsContractual.EOF
                        MonthHired_Cont = MMNAME
                        Select Case MonthHired_Cont
                            Case 1: Contractual_Count_Jan = Contractual_Count_Jan + 1
                            Case 2: Contractual_Count_Feb = Contractual_Count_Feb + 1
                            Case 3: Contractual_Count_Mar = Contractual_Count_Mar + 1
                            Case 4: Contractual_Count_Apr = Contractual_Count_Apr + 1
                            Case 5: Contractual_Count_May = Contractual_Count_May + 1
                            Case 6: Contractual_Count_June = Contractual_Count_June + 1
                            Case 7: Contractual_Count_July = Contractual_Count_July + 1
                            Case 8: Contractual_Count_August = Contractual_Count_August + 1
                            Case 9: Contractual_Count_Sept = Contractual_Count_Sept + 1
                            Case 10: Contractual_Count_Oct = Contractual_Count_Oct + 1
                            Case 11: Contractual_Count_Nov = Contractual_Count_Nov + 1
                            Case 12: Contractual_Count_Dec = Contractual_Count_Dec + 1
                        End Select
                        rsContractual.MoveNext
                    Loop
                End If
                MMNAME = MMNAME + 1
                Set rsContractual = Nothing
            Loop
            xlSheet.Cells(17, "C") = Contractual_Count_Jan
            xlSheet.Cells(17, "D") = Contractual_Count_Feb
            xlSheet.Cells(17, "E") = Contractual_Count_Mar
            xlSheet.Cells(17, "F") = Contractual_Count_Apr
            xlSheet.Cells(17, "G") = Contractual_Count_May
            xlSheet.Cells(17, "H") = Contractual_Count_June
            xlSheet.Cells(17, "I") = Contractual_Count_July
            xlSheet.Cells(17, "J") = Contractual_Count_August
            xlSheet.Cells(17, "K") = Contractual_Count_Sept
            xlSheet.Cells(17, "L") = Contractual_Count_Oct
            xlSheet.Cells(17, "M") = Contractual_Count_Nov
            xlSheet.Cells(17, "N") = Contractual_Count_Dec
            'HIRED CONTRACT =============================================================================================================

            'RESIGNED PROBI =============================================================================================================
            Dim NO_OF_SEPAREATION_FOR_Jan              As Integer
            Dim NO_OF_SEPAREATION_FOR_Feb              As Integer
            Dim NO_OF_SEPAREATION_FOR_Mar              As Integer
            Dim NO_OF_SEPAREATION_FOR_Apr              As Integer
            Dim NO_OF_SEPAREATION_FOR_May              As Integer
            Dim NO_OF_SEPAREATION_FOR_June             As Integer
            Dim NO_OF_SEPAREATION_FOR_July             As Integer
            Dim NO_OF_SEPAREATION_FOR_August           As Integer
            Dim NO_OF_SEPAREATION_FOR_Sept             As Integer
            Dim NO_OF_SEPAREATION_FOR_Oct              As Integer
            Dim NO_OF_SEPAREATION_FOR_Nov              As Integer
            Dim NO_OF_SEPAREATION_FOR_Dec              As Integer
            Dim NO_OF_SEPAREATION_FOR                  As Integer
            '
            MMNAME = 1

            Dim rsNO_OF_SEPAREATION                    As ADODB.Recordset

            Do While MMNAME <= What_month(cboMonth)
                d1 = GetMyDate(MMNAME)
                Set rsNO_OF_SEPAREATION = New ADODB.Recordset
                Set rsNO_OF_SEPAREATION = gconDMIS.Execute("Select * from HRMS_EmpInfo where emplevel = 'E' AND month(RESIGNED) = " & What_month(MonthName(Month(d1))) & " and year(RESIGNED) = " & cboYear & " AND DEPTCODE = '" & DEPT_NAME & "'")
                If Not rsNO_OF_SEPAREATION.BOF And Not rsNO_OF_SEPAREATION.EOF Then
                    rsNO_OF_SEPAREATION.MoveFirst
                    Do While Not rsNO_OF_SEPAREATION.EOF
                        NO_OF_SEPAREATION_FOR = MMNAME
                        Select Case NO_OF_SEPAREATION_FOR
                            Case 1: NO_OF_SEPAREATION_FOR_Jan = NO_OF_SEPAREATION_FOR_Jan + 1
                            Case 2: NO_OF_SEPAREATION_FOR_Feb = NO_OF_SEPAREATION_FOR_Feb + 1
                            Case 3: NO_OF_SEPAREATION_FOR_Mar = NO_OF_SEPAREATION_FOR_Mar + 1
                            Case 4: NO_OF_SEPAREATION_FOR_Apr = NO_OF_SEPAREATION_FOR_Apr + 1
                            Case 5: NO_OF_SEPAREATION_FOR_May = NO_OF_SEPAREATION_FOR_May + 1
                            Case 6: NO_OF_SEPAREATION_FOR_June = NO_OF_SEPAREATION_FOR_June + 1
                            Case 7: NO_OF_SEPAREATION_FOR_July = NO_OF_SEPAREATION_FOR_July + 1
                            Case 8: NO_OF_SEPAREATION_FOR_August = NO_OF_SEPAREATION_FOR_August + 1
                            Case 9: NO_OF_SEPAREATION_FOR_Sept = NO_OF_SEPAREATION_FOR_Sept + 1
                            Case 10: NO_OF_SEPAREATION_FOR_Oct = NO_OF_SEPAREATION_FOR_Oct + 1
                            Case 11: NO_OF_SEPAREATION_FOR_Nov = NO_OF_SEPAREATION_FOR_Nov + 1
                            Case 12: NO_OF_SEPAREATION_FOR_Dec = NO_OF_SEPAREATION_FOR_Dec + 1
                        End Select
                        rsNO_OF_SEPAREATION.MoveNext
                    Loop
                End If
                MMNAME = MMNAME + 1
                Set rsNO_OF_SEPAREATION = Nothing
            Loop
            xlSheet.Cells(19, "C") = NO_OF_SEPAREATION_FOR_Jan
            xlSheet.Cells(19, "D") = NO_OF_SEPAREATION_FOR_Feb
            xlSheet.Cells(19, "E") = NO_OF_SEPAREATION_FOR_Mar
            xlSheet.Cells(19, "F") = NO_OF_SEPAREATION_FOR_Apr
            xlSheet.Cells(19, "G") = NO_OF_SEPAREATION_FOR_May
            xlSheet.Cells(19, "H") = NO_OF_SEPAREATION_FOR_June
            xlSheet.Cells(19, "I") = NO_OF_SEPAREATION_FOR_July
            xlSheet.Cells(19, "J") = NO_OF_SEPAREATION_FOR_August
            xlSheet.Cells(19, "K") = NO_OF_SEPAREATION_FOR_Sept
            xlSheet.Cells(19, "L") = NO_OF_SEPAREATION_FOR_Oct
            xlSheet.Cells(19, "M") = NO_OF_SEPAREATION_FOR_Nov
            xlSheet.Cells(19, "N") = NO_OF_SEPAREATION_FOR_Dec
            'RESIGNED PROBI =============================================================================================================

            'FINISHED OJT =============================================================================================================
            Dim Completed_Training_Count_Jan           As Integer
            Dim Completed_Training_Count_Feb           As Integer
            Dim Completed_Training_Count_Mar           As Integer
            Dim Completed_Training_Count_Apr           As Integer
            Dim Completed_Training_Count_May           As Integer
            Dim Completed_Training_Count_June          As Integer
            Dim Completed_Training_Count_July          As Integer
            Dim Completed_Training_Count_August        As Integer
            Dim Completed_Training_Count_Sept          As Integer
            Dim Completed_Training_Count_Oct           As Integer
            Dim Completed_Training_Count_Nov           As Integer
            Dim Completed_Training_Count_Dec           As Integer
            Dim Month_Completed_Training               As Integer

            MMNAME = 1
            Dim rsCompleted_Training                   As ADODB.Recordset
            Do While MMNAME <= What_month(cboMonth)
                d1 = GetMyDate(MMNAME)
                Set rsCompleted_Training = New ADODB.Recordset
                Set rsCompleted_Training = gconDMIS.Execute("Select * from HRMS_EmpInfo where EMPLEVEL = 'A' AND month(RESIGNED) = " & What_month(MonthName(Month(d1))) & " and year(RESIGNED) = " & cboYear & " AND DEPTCODE = '" & DEPT_NAME & "'")
                If Not rsCompleted_Training.BOF And Not rsCompleted_Training.EOF Then
                    rsCompleted_Training.MoveFirst
                    Do While Not rsCompleted_Training.EOF
                        Month_Completed_Training = MMNAME
                        Select Case Month_Completed_Training
                            Case 1: Completed_Training_Count_Jan = Completed_Training_Count_Jan + 1
                            Case 2: Completed_Training_Count_Feb = Completed_Training_Count_Feb + 1
                            Case 3: Completed_Training_Count_Mar = Completed_Training_Count_Mar + 1
                            Case 4: Completed_Training_Count_Apr = Completed_Training_Count_Apr + 1
                            Case 5: Completed_Training_Count_May = Completed_Training_Count_May + 1
                            Case 6: Completed_Training_Count_June = Completed_Training_Count_June + 1
                            Case 7: Completed_Training_Count_July = Completed_Training_Count_July + 1
                            Case 8: Completed_Training_Count_August = Completed_Training_Count_August + 1
                            Case 9: Completed_Training_Count_Sept = Completed_Training_Count_Sept + 1
                            Case 10: Completed_Training_Count_Oct = Completed_Training_Count_Oct + 1
                            Case 11: Completed_Training_Count_Nov = Completed_Training_Count_Nov + 1
                            Case 12: Completed_Training_Count_Dec = Completed_Training_Count_Dec + 1
                        End Select
                        rsCompleted_Training.MoveNext
                    Loop
                End If
                MMNAME = MMNAME + 1
                Set rsCompleted_Training = Nothing
            Loop
            xlSheet.Cells(20, "C") = Completed_Training_Count_Jan
            xlSheet.Cells(20, "D") = Completed_Training_Count_Feb
            xlSheet.Cells(20, "E") = Completed_Training_Count_Mar
            xlSheet.Cells(20, "F") = Completed_Training_Count_Apr
            xlSheet.Cells(20, "G") = Completed_Training_Count_May
            xlSheet.Cells(20, "H") = Completed_Training_Count_June
            xlSheet.Cells(20, "I") = Completed_Training_Count_July
            xlSheet.Cells(20, "J") = Completed_Training_Count_August
            xlSheet.Cells(20, "K") = Completed_Training_Count_Sept
            xlSheet.Cells(20, "L") = Completed_Training_Count_Oct
            xlSheet.Cells(20, "M") = Completed_Training_Count_Nov
            xlSheet.Cells(20, "N") = Completed_Training_Count_Dec
            'FINISHED OJT =============================================================================================================

            'FINISHED CONTRACT =============================================================================================================
            Dim Finished_Contract_Count_Jan            As Integer
            Dim Finished_Contract_Count_Feb            As Integer
            Dim Finished_Contract_Count_Mar            As Integer
            Dim Finished_Contract_Count_Apr            As Integer
            Dim Finished_Contract_Count_May            As Integer
            Dim Finished_Contract_Count_June           As Integer
            Dim Finished_Contract_Count_July           As Integer
            Dim Finished_Contract_Count_August         As Integer
            Dim Finished_Contract_Count_Sept           As Integer
            Dim Finished_Contract_Count_Oct            As Integer
            Dim Finished_Contract_Count_Nov            As Integer
            Dim Finished_Contract_Count_Dec            As Integer
            Dim Month_Finished_Contract                As Integer

            Dim rsFinished_Contract                    As ADODB.Recordset
            MMNAME = 1

            Do While MMNAME <= What_month(cboMonth)
                d1 = GetMyDate(MMNAME)
                Set rsFinished_Contract = New ADODB.Recordset
                Set rsFinished_Contract = gconDMIS.Execute("Select * from HRMS_EmpInfo where EMPLEVEL = 'C' AND month(RESIGNED) = " & What_month(MonthName(Month(d1))) & " and year(RESIGNED) = " & cboYear & " AND DEPTCODE = '" & DEPT_NAME & "'")

                If Not rsFinished_Contract.BOF And Not rsFinished_Contract.EOF Then
                    rsFinished_Contract.MoveFirst
                    Do While Not rsFinished_Contract.EOF
                        Month_Finished_Contract = MMNAME
                        Select Case Month_Finished_Contract
                            Case 1: Finished_Contract_Count_Jan = Finished_Contract_Count_Jan + 1
                            Case 2: Finished_Contract_Count_Feb = Finished_Contract_Count_Feb + 1
                            Case 3: Finished_Contract_Count_Mar = Finished_Contract_Count_Mar + 1
                            Case 4: Finished_Contract_Count_Apr = Finished_Contract_Count_Apr + 1
                            Case 5: Finished_Contract_Count_May = Finished_Contract_Count_May + 1
                            Case 6: Finished_Contract_Count_June = Finished_Contract_Count_June + 1
                            Case 7: Finished_Contract_Count_July = Finished_Contract_Count_July + 1
                            Case 8: Finished_Contract_Count_August = Finished_Contract_Count_August + 1
                            Case 9: Finished_Contract_Count_Sept = Finished_Contract_Count_Sept + 1
                            Case 10: Finished_Contract_Count_Oct = Finished_Contract_Count_Oct + 1
                            Case 11: Finished_Contract_Count_Nov = Finished_Contract_Count_Nov + 1
                            Case 12: Finished_Contract_Count_Dec = Finished_Contract_Count_Dec + 1
                        End Select
                        rsFinished_Contract.MoveNext
                    Loop
                End If
                Set rsFinished_Contract = Nothing
                MMNAME = MMNAME + 1
            Loop
            xlSheet.Cells(21, "C") = Finished_Contract_Count_Jan
            xlSheet.Cells(21, "D") = Finished_Contract_Count_Feb
            xlSheet.Cells(21, "E") = Finished_Contract_Count_Mar
            xlSheet.Cells(21, "F") = Finished_Contract_Count_Apr
            xlSheet.Cells(21, "G") = Finished_Contract_Count_May
            xlSheet.Cells(21, "H") = Finished_Contract_Count_June
            xlSheet.Cells(21, "I") = Finished_Contract_Count_July
            xlSheet.Cells(21, "J") = Finished_Contract_Count_August
            xlSheet.Cells(21, "K") = Finished_Contract_Count_Sept
            xlSheet.Cells(21, "L") = Finished_Contract_Count_Oct
            xlSheet.Cells(21, "M") = Finished_Contract_Count_Nov
            xlSheet.Cells(21, "N") = Finished_Contract_Count_Dec
            'FINISHED CONTRACT =============================================================================================================

            'CLEANERS =============================================================================================================
            If What_month(cboMonth.Text) = 1 Then xlSheet.Cells(14, "D") = ""
            If What_month(cboMonth.Text) = 2 Then xlSheet.Cells(14, "E") = ""
            If What_month(cboMonth.Text) = 3 Then xlSheet.Cells(14, "F") = ""
            If What_month(cboMonth.Text) = 4 Then xlSheet.Cells(14, "G") = ""
            If What_month(cboMonth.Text) = 5 Then xlSheet.Cells(14, "H") = ""
            If What_month(cboMonth.Text) = 6 Then xlSheet.Cells(14, "I") = ""
            If What_month(cboMonth.Text) = 7 Then xlSheet.Cells(14, "J") = ""
            If What_month(cboMonth.Text) = 8 Then xlSheet.Cells(14, "K") = ""
            If What_month(cboMonth.Text) = 9 Then xlSheet.Cells(14, "L") = ""
            If What_month(cboMonth.Text) = 10 Then xlSheet.Cells(14, "M") = ""
            If What_month(cboMonth.Text) = 11 Then xlSheet.Cells(14, "N") = ""
            'CLEANERS =============================================================================================================

            xlApp.Visible = True
            Set xlApp = Nothing
            'LogAudit "V", "ACTUAL MANNING REPORT", cboYear
            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "ACTUAL MANNING REPORT", "", "", "", "DETAILS", "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        Else
            ShowNoRecord
        End If

        Screen.MousePointer = 0
        Exit Sub
    End If

    Exit Sub
Errorcode:
    ShowVBError
    Screen.MousePointer = 0

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (ACTUAL MANNING REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "ACTUAL MANNING REPORT", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Screen.MousePointer = 0
    fillcbomonth cboMonth
    FillCboMoreYear cboYear

    cboMonth.Text = MonthName(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
End Sub

