VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMSMonthlyTimeControlAnalysis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "M. Time Control Analysis"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3885
   Icon            =   "frmCSMSMonthlyTimeControlAnalysis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1875
   ScaleWidth      =   3885
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
      Left            =   930
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select year from the list"
      Top             =   510
      Width           =   2835
   End
   Begin VB.ComboBox cboMonth 
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
      Left            =   930
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select month from the list"
      Top             =   120
      Width           =   2835
   End
   Begin Crystal.CrystalReport rptCSMSMonthlyTimeControlAnalysis 
      Left            =   4260
      Top             =   1500
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Monthly Time Control Analysis Report"
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
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      MouseIcon       =   "frmCSMSMonthlyTimeControlAnalysis.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSMonthlyTimeControlAnalysis.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   930
      Width           =   765
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4410
      MouseIcon       =   "frmCSMSMonthlyTimeControlAnalysis.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSMonthlyTimeControlAnalysis.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   1170
      Width           =   765
   End
   Begin VB.CommandButton Command1 
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
      Height          =   855
      Left            =   2250
      MouseIcon       =   "frmCSMSMonthlyTimeControlAnalysis.frx":1C10
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSMonthlyTimeControlAnalysis.frx":1D62
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   930
      Width           =   765
   End
   Begin VB.Label lblNOTES 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "COMPUTING MTCA DATA please wait..."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   510
      Left            =   60
      TabIndex        =   7
      Top             =   1110
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   405
      TabIndex        =   6
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   210
      Width           =   585
   End
End
Attribute VB_Name = "frmCSMSMonthlyTimeControlAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlApp                                              As Excel.Application
Dim xlBook                                             As Excel.Workbook
Dim xlSheet                                            As Excel.Worksheet

Function ComputeAttendedHours(vEMPNO As String) As Double()
    Dim rstmp                                          As New ADODB.Recordset
    Dim VTMP(1)                                        As Double
    Dim VTIME                                          As Double
    Dim VAVAI                                          As Double
    Dim nTIME                                          As Double
    Dim nAVAI                                          As Double

    Set rstmp = gconDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE " & _
        " EMPNO = '" & vEMPNO & _
        "' AND MONTH(DATETODAY) = " & What_month(cboMonth) & _
        " AND YEAR(DATETODAY) = " & cboYear & _
        " ORDER BY DATETODAY")
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            nTIME = 0: nAVAI = 0
            If COMPANY_CODE = "HMH" Or COMPANY_CODE = "HAS" Then
                If Not Null2String(rstmp!INAM) = "" Then
                    If Not Null2String(rstmp!outpm) = "" Then
                        VTIME = VTIME + DateDiff("N", rstmp!INAM, rstmp!outpm)
                        nTIME = nTIME + DateDiff("N", rstmp!INAM, rstmp!outpm)
                    End If
                End If
            Else
                If Null2String(rstmp!Shift) = "SHIFT1" Then
                    If Not Null2String(rstmp!INAM) = "" Then
                        If Not Null2String(rstmp!OUTAM) = "" Then
                            VTIME = VTIME + DateDiff("N", rstmp!INAM, rstmp!OUTAM)
                        End If
                    End If
                Else


                End If
            End If

            'MINUS THE LUNCH BREAK 1 hour
            If nTIME > 450 Then
                VTIME = VTIME - 60
            End If

            nAVAI = 7.5
            If (nTIME / 60) > 7.5 Then
                nAVAI = nAVAI + ((nTIME - (nAVAI * 60)) / 60)
                VAVAI = VAVAI + nAVAI
            Else
                VAVAI = VAVAI + 7.5
            End If


            rstmp.MoveNext
        Loop
    End If

    If Not VTIME < 0 Then
        VTMP(0) = VTIME / 60
        VTMP(1) = VAVAI
        ComputeAttendedHours = VTMP
    Else
        VTMP(0) = 0
        VTMP(1) = VAVAI
    End If

    ComputeAttendedHours = VTMP
    Set rstmp = Nothing
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode
    Screen.MousePointer = 11

    Dim rsMonthlyTimeControlAnalysis                   As ADODB.Recordset
    Set rsMonthlyTimeControlAnalysis = New ADODB.Recordset
    Set rsMonthlyTimeControlAnalysis = gconDMIS.Execute("SELECT * from CSMS_vw_RO_Det_For_TechPerformance where Month(SAVEDATE) = '" & What_month(cboMonth.Text) & "' AND Year(SAVEDATE) =" & cboYear.Text)
    If Not rsMonthlyTimeControlAnalysis.BOF And Not rsMonthlyTimeControlAnalysis.EOF Then
        rptCSMSMonthlyTimeControlAnalysis.Formulas(8) = "MonthReport = '" & cboMonth.Text & "'"
        rptCSMSMonthlyTimeControlAnalysis.Formulas(9) = "YearReport = '" & cboYear.Text & "'"

        'JUN 02/05/2005
        rptCSMSMonthlyTimeControlAnalysis.Formulas(0) = "Company Name = '" & COMPANY_NAME & "'"
        rptCSMSMonthlyTimeControlAnalysis.Formulas(1) = "Company Address = '" & COMPANY_ADDRESS & "'"
        rptCSMSMonthlyTimeControlAnalysis.Formulas(2) = "Printedby = '" & LOGNAME & "'"

        'COMMENT BY : MJP 07092008 1:50 PM
        'PrintSQLReport rptCSMSMonthlyTimeControlAnalysis, CSMS_REPORT_PATH & "MonthlyTimeControlAnalysis.rpt", "Month({CSMS_vw_RO_Det_For_TechPerformance.SAVEDATE}) = " & What_month(cboMonth.Text) & " AND Year({CSMS_vw_RO_Det_For_TechPerformance.SAVEDATE}) = " & cboYear.Text, CSMS_REPORT_CONNECTION, 1

        PrintSQLReport rptCSMSMonthlyTimeControlAnalysis, CSMS_REPORT_PATH & "MonthlyTimeControlAnalysis.rpt", "Month({CSMS_REPOR.DTE_COMP}) = " & What_month(cboMonth.Text) & " AND Year({CSMS_REPOR.DTE_COMP}) = " & cboYear.Text, CSMS_REPORT_CONNECTION, 1

        'NEW LOG AUDIT +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '   Call NEW_LogAudit("V", "MONTHLY TIME CONTROL ANALYSIS", "", "", "", "", "", "")
        'NEW LOG AUDIT +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        LogAudit "V", "MONTHLY TIME CONTROL ANALYSIS - REPORTS ", cboMonth & cboYear
    Else
        ShowNoRecord
    End If

    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub Command1_Click()
    Screen.MousePointer = 11
    Dim cnt                                            As Integer
    Dim rsREPOR                                        As New ADODB.Recordset
    Dim rsDet                                          As New ADODB.Recordset
    Dim RSTECH                                         As New ADODB.Recordset
    Dim FLAT_TIME                                      As Double
    Dim PROD_TIME                                      As Double
    Dim ATTEND_HR                                       As Double
    Dim AVAILA_HR                                       As Double
    
    Set xlApp = New Excel.Application
    Set RSTECH = gconDMIS.Execute("SELECT   * FROM CSMS_VW_TECHNICIAN ORDER BY TECH_NAME")
    'Set rsTECH = gconDMIS.Execute("SELECT  * FROM CSMS_VW_TECHNICIAN WHERE EMPNO = '55046'")
    If Not (RSTECH.BOF And RSTECH.EOF) Then
        lblNOTES.Visible = True
        If COMPANY_CODE = "HMH" Then
            cnt = 15
            Set xlApp = CreateObject("Excel.Application")
            Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "Monthly Time Control Analysis.xlt")
            Set xlSheet = xlBook.Worksheets(1)
    
            xlSheet.Cells(8, "A") = "For the Month of " & cboMonth
            xlSheet.Cells(10, "A") = COMPANY_NAME
            xlSheet.Cells(34, "D") = GENERAL_MANAGER
    
            Do While Not RSTECH.EOF
                xlSheet.Cells(cnt, "A") = Null2String(RSTECH!EMPNO)
                xlSheet.Cells(cnt, "B") = Null2String(RSTECH!TECH_NAME)
    
                Set rsREPOR = gconDMIS.Execute("SELECT CSMS_Repor.REP_OR, CSMS_Repor.TRANSTYPE, CSMS_Ro_Det.DET_HRS, CSMS_Ro_Det.HRSWRK, " & _
                        " CSMS_Repor.DTE_COMP FROM CSMS_Repor INNER JOIN " & _
                        " dbo.CSMS_Ro_Det ON dbo.CSMS_Repor.REP_OR = dbo.CSMS_Ro_Det.REP_OR AND " & _
                        " dbo.CSMS_Repor.TRANSTYPE = dbo.CSMS_Ro_Det.TRANSTYPE " & _
                        " Where (Month(dbo.CSMS_Repor.DTE_COMP) = " & What_month(cboMonth) & ") And (Year(dbo.CSMS_Repor.DTE_COMP) = " & cboYear & ") AND LTRIM(RTRIM(CSMS_RO_DET.TECHCODE)) = '" & LTrim(RTrim(RSTECH!Technician)) & "'")
                If Not (rsREPOR.BOF And rsREPOR.EOF) Then
                    Do While Not rsREPOR.EOF
                        FLAT_TIME = FLAT_TIME + NumericVal(rsREPOR!DET_HRS)
                        PROD_TIME = PROD_TIME + NumericVal(rsREPOR!HRSWRK)
    
                        rsREPOR.MoveNext
                    Loop
                End If
    
                ATTEND_HR = 0
                AVAILA_HR = 0
                xlSheet.Cells(cnt, "C") = FLAT_TIME
                xlSheet.Cells(cnt, "D") = PROD_TIME
    
                ATTEND_HR = ComputeAttendedHours(RSTECH!EMPNO)(0)
                AVAILA_HR = ComputeAttendedHours(RSTECH!EMPNO)(1)
    
                xlSheet.Cells(cnt, "I") = ATTEND_HR
                xlSheet.Cells(cnt, "J") = AVAILA_HR
                
'                If AVAILA_HR = 0 Then
'                    xlSheet.Cells(cnt, "K") = 0
'                Else
'                    xlSheet.Cells(cnt, "K") = ATTEND_HR / AVAILA_HR
'                End If
'
'                If PROD_TIME = 0 Then
'                    xlSheet.Cells(cnt, "E") = 0
'                Else
'                    xlSheet.Cells(cnt, "E") = Round(FLAT_TIME / PROD_TIME, 2)
'                End If
    
                FLAT_TIME = 0
                PROD_TIME = 0
    
                cnt = cnt + 1
                If cnt > 31 Then
'                    xlApp.Visible = True
'                    Set xlApp = Nothing
    
                    xlSheet.Cells(8, "A") = "For the Month of " & cboMonth
                    xlSheet.Cells(10, "B") = COMPANY_CODE
                    xlSheet.Cells(34, "D") = GENERAL_MANAGER
                    cnt = 15
                End If
    
                RSTECH.MoveNext
            Loop
    
            xlApp.Visible = True
            Set xlApp = Nothing
            lblNOTES.Visible = False
        Else
            cnt = 15
            Set xlApp = CreateObject("Excel.Application")
            Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "Monthly Time Control Analysis.xls")
            Set xlSheet = xlBook.Worksheets(1)
    
            xlSheet.Cells(8, "A") = "For the Month of " & cboMonth
            xlSheet.Cells(10, "B") = COMPANY_NAME
            xlSheet.Cells(34, "D") = GENERAL_MANAGER
    
            Do While Not RSTECH.EOF
                xlSheet.Cells(cnt, "A") = Null2String(RSTECH!EMPNO)
                xlSheet.Cells(cnt, "B") = Null2String(RSTECH!TECH_NAME)
    
                Set rsREPOR = gconDMIS.Execute("SELECT CSMS_Repor.REP_OR, CSMS_Repor.TRANSTYPE, CSMS_Ro_Det.DET_HRS, CSMS_Ro_Det.HRSWRK, " & _
                        " CSMS_Repor.DTE_COMP FROM CSMS_Repor INNER JOIN " & _
                        " dbo.CSMS_Ro_Det ON dbo.CSMS_Repor.REP_OR = dbo.CSMS_Ro_Det.REP_OR AND " & _
                        " dbo.CSMS_Repor.TRANSTYPE = dbo.CSMS_Ro_Det.TRANSTYPE " & _
                        " Where (Month(dbo.CSMS_Repor.DTE_COMP) = " & What_month(cboMonth) & ") And (Year(dbo.CSMS_Repor.DTE_COMP) = " & cboYear & ") AND LTRIM(RTRIM(CSMS_RO_DET.TECHCODE)) = '" & LTrim(RTrim(RSTECH!Technician)) & "'")
                If Not (rsREPOR.BOF And rsREPOR.EOF) Then
                    Do While Not rsREPOR.EOF
                        FLAT_TIME = FLAT_TIME + NumericVal(rsREPOR!DET_HRS)
                        PROD_TIME = PROD_TIME + NumericVal(rsREPOR!HRSWRK)
    
                        rsREPOR.MoveNext
                    Loop
                Else
                End If
    
   
                ATTEND_HR = 0
                AVAILA_HR = 0
                xlSheet.Cells(cnt, "C") = FLAT_TIME
                xlSheet.Cells(cnt, "D") = PROD_TIME
    
                ATTEND_HR = ComputeAttendedHours(RSTECH!EMPNO)(0)
                AVAILA_HR = ComputeAttendedHours(RSTECH!EMPNO)(1)
    
                xlSheet.Cells(cnt, "F") = ATTEND_HR
                xlSheet.Cells(cnt, "G") = AVAILA_HR
                If AVAILA_HR = 0 Then
                    xlSheet.Cells(cnt, "H") = 0
                Else
                    xlSheet.Cells(cnt, "H") = ATTEND_HR / AVAILA_HR
                End If
    
                If PROD_TIME = 0 Then
                    xlSheet.Cells(cnt, "E") = 0
                Else
                    xlSheet.Cells(cnt, "E") = Round(FLAT_TIME / PROD_TIME, 2)
                End If
    
                FLAT_TIME = 0
                PROD_TIME = 0
    
                cnt = cnt + 1
                If cnt > 31 Then
'                    xlApp.Visible = True
'                    Set xlApp = Nothing
    
                    xlSheet.Cells(8, "A") = "For the Month of " & cboMonth
                    xlSheet.Cells(10, "B") = COMPANY_CODE
                    xlSheet.Cells(34, "D") = GENERAL_MANAGER
                    cnt = 15
                End If
    
                RSTECH.MoveNext
            Loop
            
            xlApp.Visible = True
            Set xlApp = Nothing
            lblNOTES.Visible = False
        End If

        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "MONTHLY TIME CONTROL ANALYSIS", "", "", "", cboMonth & " " & cboYear, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
    Else
        ShowNoRecord
        Exit Sub
    End If

    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (MONTHLY TIME CONTROL ANALYSIS)"
            Call frmALL_AuditInquiry.DisplayHistory("", "MONTHLY TIME CONTROL ANALYSIS", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Screen.MousePointer = 0
    fillcbomonth cboMonth
    FillCboMoreYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
End Sub

Private Sub Timer1_Timer()
    If lblNOTES.ForeColor = vbRed Then
        lblNOTES.ForeColor = vbBlack
    Else
        lblNOTES.ForeColor = vbRed
    End If
End Sub

