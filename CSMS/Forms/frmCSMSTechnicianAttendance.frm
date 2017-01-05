VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMSTechnicianAttendance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tech Attendance"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3060
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMSTechnicianAttendance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1785
   ScaleWidth      =   3060
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   930
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select month from the list"
      Top             =   120
      Width           =   2085
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   930
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select year from the list"
      Top             =   510
      Width           =   2085
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   750
      Left            =   2265
      MouseIcon       =   "frmCSMSTechnicianAttendance.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSTechnicianAttendance.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   945
      Width           =   735
   End
   Begin Crystal.CrystalReport rptTechnicianAttendance 
      Left            =   120
      Top             =   1065
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Technician Attendance Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   750
      Left            =   1545
      MouseIcon       =   "frmCSMSTechnicianAttendance.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSTechnicianAttendance.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   945
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   180
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   570
      Width           =   735
   End
End
Attribute VB_Name = "frmCSMSTechnicianAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlApp                                              As Excel.Application
Dim xlBook                                             As Excel.Workbook
Dim xlSheet                                            As Excel.Worksheet

Sub PrintAttendance()
    Dim fMONTH                                         As Integer
    Dim sMONTH                                         As Integer
    Dim tMONTH                                         As Integer

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "Attendance Sheet.xls")
    Set xlSheet = xlBook.Worksheets(1)

    xlSheet.Cells(6, "A") = "DEALER :  " & Null2String(COMPANY_NAME)
    xlSheet.Cells(6, "P") = Date
    xlSheet.Cells(38, "H") = GENERAL_MANAGER

    If What_month(cboMonth) < 4 Then                  'JAN - MAR
        fMONTH = 1: sMONTH = 2: tMONTH = 3
    ElseIf What_month(cboMonth) < 7 And What_month(cboMonth) > 3 Then    'APR - JUN
        fMONTH = 4: sMONTH = 5: tMONTH = 6
    ElseIf What_month(cboMonth) < 10 And What_month(cboMonth) > 6 Then    'JUL - SEP
        fMONTH = 7: sMONTH = 8: tMONTH = 9
    Else                                              'OCT - DEC
        fMONTH = 10: sMONTH = 11: tMONTH = 12
    End If

    xlSheet.Cells(11, "C") = MonthName(fMONTH)
    xlSheet.Cells(11, "H") = MonthName(sMONTH)
    xlSheet.Cells(11, "M") = MonthName(tMONTH)

    Call ComputeAttendance(fMONTH, sMONTH, tMONTH)

    xlApp.Visible = True
    Set xlApp = Nothing
End Sub

Sub ComputeAttendance(FMON As Integer, SMON As Integer, TMON As Integer)
    Dim rsSA                                           As New ADODB.Recordset
    Dim RSTECH                                         As New ADODB.Recordset
    Dim RSATTEND                                       As New ADODB.Recordset
    Dim i                                              As Integer
    Dim OF_WORK_DAY                                    As Integer
    Dim DAY_PRESENT                                    As Integer
    Dim DAY_ABSENT                                     As Integer
    Dim MIN_LATE                                       As Double
    Dim MIN_ATTEND                                     As Double

    i = 16
    Set rsSA = gconDMIS.Execute("SELECT * FROM CSMS_VW_EMPNO ORDER BY NAYM")
    If Not (rsSA.BOF And rsSA.EOF) Then
        Do While Not rsSA.EOF
            xlSheet.Cells(i, "B") = Null2String(rsSA!NAYM)

            OF_WORK_DAY = 0
            DAY_PRESENT = 0
            DAY_ABSENT = 0
            MIN_LATE = 0

            Set RSATTEND = gconDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE EMPNO = '" & Null2String(rsSA!EMPNO) & "' AND MONTH(DATETODAY) = " & FMON & " AND YEAR(DATETODAY) = " & cboYear & "")
            Do While Not RSATTEND.EOF
                OF_WORK_DAY = OF_WORK_DAY + 1
                If COMPANY_CODE = "HGC" Then
                    If Null2String(RSATTEND!INAM) = "" Then
                        DAY_ABSENT = DAY_ABSENT + 1
                    Else
                        If Null2String(RSATTEND!OUTAM) = "" Then
                            DAY_ABSENT = DAY_ABSENT + 1
                        Else
                            DAY_PRESENT = DAY_PRESENT + 1

                            MIN_ATTEND = 0
                            MIN_ATTEND = DateDiff("N", RSATTEND!INAM, RSATTEND!OUTAM)
                            If MIN_ATTEND < 450 Then
                                MIN_LATE = MIN_LATE + (450 - MIN_ATTEND)
                            End If
                        End If
                    End If
                End If

                RSATTEND.MoveNext
            Loop

            xlSheet.Cells(i, "C") = OF_WORK_DAY
            xlSheet.Cells(i, "D") = DAY_PRESENT
            xlSheet.Cells(i, "E") = DAY_ABSENT
            xlSheet.Cells(i, "F") = MIN_LATE

            OF_WORK_DAY = 0
            DAY_PRESENT = 0
            DAY_ABSENT = 0
            MIN_LATE = 0

            Set RSATTEND = gconDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE EMPNO = '" & Null2String(rsSA!EMPNO) & "' AND MONTH(DATETODAY) = " & SMON & " AND YEAR(DATETODAY) = " & cboYear & "")
            Do While Not RSATTEND.EOF
                OF_WORK_DAY = OF_WORK_DAY + 1
                If COMPANY_CODE = "HGC" Then
                    If Null2String(RSATTEND!INAM) = "" Then
                        DAY_ABSENT = DAY_ABSENT + 1
                    Else
                        If Null2String(RSATTEND!OUTAM) = "" Then
                            DAY_ABSENT = DAY_ABSENT + 1
                        Else
                            DAY_PRESENT = DAY_PRESENT + 1

                            MIN_ATTEND = 0
                            MIN_ATTEND = DateDiff("N", RSATTEND!INAM, RSATTEND!OUTAM)
                            If MIN_ATTEND < 450 Then
                                MIN_LATE = MIN_LATE + (450 - MIN_ATTEND)
                            End If
                        End If
                    End If
                End If
                RSATTEND.MoveNext
            Loop

            xlSheet.Cells(i, "H") = OF_WORK_DAY
            xlSheet.Cells(i, "I") = DAY_PRESENT
            xlSheet.Cells(i, "J") = DAY_ABSENT
            xlSheet.Cells(i, "K") = MIN_LATE

            OF_WORK_DAY = 0
            DAY_PRESENT = 0
            DAY_ABSENT = 0
            MIN_LATE = 0

            Set RSATTEND = gconDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE EMPNO = '" & Null2String(rsSA!EMPNO) & "' AND MONTH(DATETODAY) = " & TMON & " AND YEAR(DATETODAY) = " & cboYear & "")
            Do While Not RSATTEND.EOF
                OF_WORK_DAY = OF_WORK_DAY + 1
                If COMPANY_CODE = "HGC" Then
                    If Null2String(RSATTEND!INAM) = "" Then
                        DAY_ABSENT = DAY_ABSENT + 1
                    Else
                        If Null2String(RSATTEND!OUTAM) = "" Then
                            DAY_ABSENT = DAY_ABSENT + 1
                        Else
                            DAY_PRESENT = DAY_PRESENT + 1

                            MIN_ATTEND = 0
                            MIN_ATTEND = DateDiff("N", RSATTEND!INAM, RSATTEND!OUTAM)
                            If MIN_ATTEND < 450 Then
                                MIN_LATE = MIN_LATE + (450 - MIN_ATTEND)
                            End If
                        End If
                    End If
                End If

                RSATTEND.MoveNext
            Loop

            xlSheet.Cells(i, "M") = OF_WORK_DAY
            xlSheet.Cells(i, "N") = DAY_PRESENT
            xlSheet.Cells(i, "O") = DAY_ABSENT
            xlSheet.Cells(i, "P") = MIN_LATE

            i = i + 1
            rsSA.MoveNext
        Loop
    End If

    Set RSTECH = gconDMIS.Execute("SELECT * FROM CSMS_VW_TECHNICIAN ORDER BY TECH_NAME")
    If Not (RSTECH.BOF And RSTECH.EOF) Then
        Do While Not RSTECH.EOF
            xlSheet.Cells(i, "B") = Null2String(RSTECH!TECH_NAME)

            OF_WORK_DAY = 0
            DAY_PRESENT = 0
            DAY_ABSENT = 0
            MIN_LATE = 0

            Set RSATTEND = gconDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE EMPNO = '" & Null2String(RSTECH!EMPNO) & "' AND MONTH(DATETODAY) = " & FMON & " AND YEAR(DATETODAY) = " & cboYear & "")
            Do While Not RSATTEND.EOF
                OF_WORK_DAY = OF_WORK_DAY + 1
                If COMPANY_CODE = "HGC" Then
                    If Null2String(RSATTEND!INAM) = "" Then
                        DAY_ABSENT = DAY_ABSENT + 1
                    Else
                        If Null2String(RSATTEND!OUTAM) = "" Then
                            DAY_ABSENT = DAY_ABSENT + 1
                        Else
                            DAY_PRESENT = DAY_PRESENT + 1

                            MIN_ATTEND = 0
                            MIN_ATTEND = DateDiff("N", RSATTEND!INAM, RSATTEND!OUTAM)
                            If MIN_ATTEND < 450 Then
                                MIN_LATE = MIN_LATE + (450 - MIN_ATTEND)
                            End If
                        End If
                    End If
                End If

                RSATTEND.MoveNext
            Loop

            xlSheet.Cells(i, "C") = OF_WORK_DAY
            xlSheet.Cells(i, "D") = DAY_PRESENT
            xlSheet.Cells(i, "E") = DAY_ABSENT
            xlSheet.Cells(i, "F") = MIN_LATE

            OF_WORK_DAY = 0
            DAY_PRESENT = 0
            DAY_ABSENT = 0
            MIN_LATE = 0

            Set RSATTEND = gconDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE EMPNO = '" & Null2String(RSTECH!EMPNO) & "' AND MONTH(DATETODAY) = " & SMON & " AND YEAR(DATETODAY) = " & cboYear & "")
            Do While Not RSATTEND.EOF
                OF_WORK_DAY = OF_WORK_DAY + 1
                If COMPANY_CODE = "HGC" Then
                    If Null2String(RSATTEND!INAM) = "" Then
                        DAY_ABSENT = DAY_ABSENT + 1
                    Else
                        If Null2String(RSATTEND!OUTAM) = "" Then
                            DAY_ABSENT = DAY_ABSENT + 1
                        Else
                            DAY_PRESENT = DAY_PRESENT + 1

                            MIN_ATTEND = 0
                            MIN_ATTEND = DateDiff("N", RSATTEND!INAM, RSATTEND!OUTAM)
                            If MIN_ATTEND < 450 Then
                                MIN_LATE = MIN_LATE + (450 - MIN_ATTEND)
                            End If
                        End If
                    End If
                End If

                RSATTEND.MoveNext
            Loop

            xlSheet.Cells(i, "H") = OF_WORK_DAY
            xlSheet.Cells(i, "I") = DAY_PRESENT
            xlSheet.Cells(i, "J") = DAY_ABSENT
            xlSheet.Cells(i, "K") = MIN_LATE

            OF_WORK_DAY = 0
            DAY_PRESENT = 0
            DAY_ABSENT = 0
            MIN_LATE = 0

            Set RSATTEND = gconDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE EMPNO = '" & Null2String(RSTECH!EMPNO) & "' AND MONTH(DATETODAY) = " & TMON & " AND YEAR(DATETODAY) = " & cboYear & "")
            Do While Not RSATTEND.EOF
                OF_WORK_DAY = OF_WORK_DAY + 1
                If COMPANY_CODE = "HGC" Then
                    If Null2String(RSATTEND!INAM) = "" Then
                        DAY_ABSENT = DAY_ABSENT + 1
                    Else
                        If Null2String(RSATTEND!OUTAM) = "" Then
                            DAY_ABSENT = DAY_ABSENT + 1
                        Else
                            DAY_PRESENT = DAY_PRESENT + 1

                            MIN_ATTEND = 0
                            MIN_ATTEND = DateDiff("N", RSATTEND!INAM, RSATTEND!OUTAM)
                            If MIN_ATTEND < 450 Then
                                MIN_LATE = MIN_LATE + (450 - MIN_ATTEND)
                            End If
                        End If
                    End If
                End If

                RSATTEND.MoveNext
            Loop

            xlSheet.Cells(i, "M") = OF_WORK_DAY
            xlSheet.Cells(i, "N") = DAY_PRESENT
            xlSheet.Cells(i, "O") = DAY_ABSENT
            xlSheet.Cells(i, "P") = MIN_LATE

            i = i + 1
            RSTECH.MoveNext
        Loop
    End If

    Set rsSA = Nothing
    Set RSTECH = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    'On Error GoTo Errorcode

    Call PrintAttendance
    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("V", "TECHNICIAN REPORT", "", "", "", "TECHNICIAN ATTENDANCE : " & cboMonth & " " & cboYear, "", "")
    'NEW LOG AUDIT-----------------------------------------------------


    Screen.MousePointer = 0
    Exit Sub

    Dim rsTechnicianAttendance                         As ADODB.Recordset
    Set rsTechnicianAttendance = New ADODB.Recordset
    Set rsTechnicianAttendance = gconDMIS.Execute("SELECT * from HRMS_Attend where Month(DateToday) = '" & What_month(cboMonth.Text) & "' AND Year(DateToday) =" & cboYear.Text)
    If Not rsTechnicianAttendance.BOF And Not rsTechnicianAttendance.EOF Then
        rptTechnicianAttendance.Formulas(3) = "MonthAttendance = '" & cboMonth.Text & "'"
        rptTechnicianAttendance.Formulas(4) = "YearAttendance = '" & cboYear.Text & "'"

        'JUN 02/05/2005
        rptTechnicianAttendance.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptTechnicianAttendance.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptTechnicianAttendance.Formulas(2) = "Printedby = '" & LOGNAME & "'"
        PrintSQLReport rptTechnicianAttendance, CSMS_REPORT_PATH & "TechnicianAttendanceReport.rpt", "Month({HRMS_Attend.DateToday}) = " & What_month(cboMonth.Text) & " AND Year({HRMS_Attend.DateToday}) = " & cboYear.Text, CSMS_REPORT_CONNECTION, 1

        LogAudit "V", "TECHNICIAN ATTENDANCE - REPORTS ", cboMonth & cboYear
    Else
        ShowNoRecord
    End If

    Screen.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
    fillcbomonth cboMonth
    FillCboMoreYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
End Sub

