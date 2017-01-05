VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMSUnitsReceivedWeeklyPerformanceReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UR Weekly"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3105
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMSUnitsReceivedWeeklyPerformanceReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2160
   ScaleWidth      =   3105
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   960
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   600
      Width           =   2175
   End
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
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Select month from the list"
      Top             =   90
      Width           =   1965
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
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Select year from the list"
      Top             =   480
      Width           =   1965
   End
   Begin VB.ComboBox cboWeek 
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
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select week from the list"
      Top             =   870
      Width           =   1965
   End
   Begin Crystal.CrystalReport rpt 
      Left            =   210
      Top             =   1500
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Units Received Weekly Performance Report"
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
      Height          =   750
      Left            =   2085
      MouseIcon       =   "frmCSMSUnitsReceivedWeeklyPerformanceReport.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSUnitsReceivedWeeklyPerformanceReport.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Close Window"
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   750
      Left            =   1365
      MouseIcon       =   "frmCSMSUnitsReceivedWeeklyPerformanceReport.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSUnitsReceivedWeeklyPerformanceReport.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Height          =   225
      Left            =   300
      TabIndex        =   7
      Top             =   180
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Height          =   225
      Left            =   420
      TabIndex        =   6
      Top             =   540
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Week"
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
      Height          =   225
      Left            =   330
      TabIndex        =   5
      Top             =   930
      Width           =   465
   End
End
Attribute VB_Name = "frmCSMSUnitsReceivedWeeklyPerformanceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X_MON                                               As String
Dim X_TUE                                               As String
Dim X_WEN                                               As String
Dim X_THU                                               As String
Attribute X_THU.VB_VarUserMemId = 1073938435
Dim X_FRI                                               As String
Dim X_SAT                                               As String
Dim X_SUN                                               As String
Dim xlApp                                               As Excel.Application
Attribute xlApp.VB_VarUserMemId = 1073938439
Dim xlBook                                              As Excel.Workbook
Attribute xlBook.VB_VarUserMemId = 1073938440
Dim xlSheet                                             As Excel.Worksheet
Attribute xlSheet.VB_VarUserMemId = 1073938441

Function CHECKIFHYUNDAI(PLATE_NO As String) As Boolean
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT MAKE FROM CSMS_CUSVEH WHERE PLATE_NO = '" & PLATE_NO & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        If UCase(rstmp!Make) = "HYUNDAI" Then
            CHECKIFHYUNDAI = True
        ElseIf Null2String(rstmp!Make) = "" Then
            CHECKIFHYUNDAI = True
        Else
            CHECKIFHYUNDAI = False
        End If
    Else
        CHECKIFHYUNDAI = True
    End If

    Set rstmp = Nothing
End Function

Sub GETDATERANGE()
    X_SAT = Day(Combo2.Text) & "-" & Left(MonthName(Month(Combo2)), 3)
    X_FRI = Day(DateAdd("D", -1, Combo2.Text)) & "-" & Left(MonthName(Month(DateAdd("D", -1, Combo2.Text))), 3)
    X_THU = Day(DateAdd("D", -2, Combo2.Text)) & "-" & Left(MonthName(Month(DateAdd("D", -2, Combo2.Text))), 3)
    X_WEN = Day(DateAdd("D", -3, Combo2.Text)) & "-" & Left(MonthName(Month(DateAdd("D", -3, Combo2.Text))), 3)
    X_TUE = Day(DateAdd("D", -4, Combo2.Text)) & "-" & Left(MonthName(Month(DateAdd("D", -4, Combo2.Text))), 3)
    X_MON = Day(DateAdd("D", -5, Combo2.Text)) & "-" & Left(MonthName(Month(DateAdd("D", -5, Combo2.Text))), 3)
    X_SUN = Day(DateAdd("D", -6, Combo2.Text)) & "-" & Left(MonthName(Month(DateAdd("D", -6, Combo2.Text))), 3)
End Sub

Sub AddUnitServiced()

End Sub

Sub FillcboWeek()
    cboWeek.AddItem "1st Week"
    cboWeek.AddItem "2nd Week"
    cboWeek.AddItem "3rd Week"
    cboWeek.AddItem "4th Week"
    cboWeek.AddItem "5th Week"
End Sub

Private Sub cboMonth_click()
'    Dim dbDatex
'    Dim d
'    Dim I
'
'    'COMMENT BY  : MJP 010508 1004AM
'    'DESCRIPTION : 2008 VALUE SHOULD BE CHANGE TO THE YEAR OF THE COMPUTER
'    'dbDatex = firstDay(DateSerial(2008, What_month(cboMonth), 1))
'    'COMMENT BY  : MJP 010508 1004AM
'
'    'UPDATE BY   : MJP 010508 1004AM
'    'DESCRIPTION : 2008 VALUE SHOULD BE CHANGE TO THE YEAR OF THE COMPUTER
'    dbDatex = firstDay(DateSerial(Year(Date), What_month(cboMonth), 1))
'    'UPDATE BY   : MJP 010508 1004AM
'
'    Combo1.Clear
'    Combo2.Clear
'    d = 1
'    cboWeek.Clear
'    Do While What_month(cboMonth) = Month(dbDatex)
'        Combo1.AddItem dbDatex
'        I = I + 1
'        d = 7 - Weekday(dbDatex)
'        cboWeek.AddItem I & " week"
'        Combo2.AddItem DateAdd("d", 7 - Weekday(dbDatex), dbDatex)
'        dbDatex = DateAdd("d", d + 1, dbDatex)
'    Loop
'
'    cboWeek.ListIndex = 0

    Dim dbDatex
    Dim d
    Dim i

    'COMMENT BY  : MJP 010509 1006AM
    'DESCRIPTION : 2008 VALUE SHOULD BE CHANGE TO THE YEAR OF THE COMPUTER
    'dbDatex = firstDay(DateSerial(2008, What_month(cboMonth), 1))
    'COMMENT BY  : MJP 010509 1006AM

    'UPDATE BY   : MJP 010509 1006AM
    'DESCRIPTION : 2008 VALUE SHOULD BE CHANGE TO THE YEAR OF THE COMPUTER
    'dbDatex = firstDay(DateSerial(Year(Date), What_month(cboMonth), 1))
    
    
    'UPDATE BY   : AXP 01042010 2:30 PM
    'cboyear is already Year if we use Date it will get system date in this case we are looking for Year As designated by User
    dbDatex = firstDay(DateSerial(cboYear, What_month(cboMonth), 1))
    
    Combo1.Clear
    Combo2.Clear
    d = 1
    cboWeek.Clear
    Do While What_month(cboMonth) = Month(dbDatex)
        Combo1.AddItem dbDatex
        i = i + 1
        d = 7 - Weekday(dbDatex)
        cboWeek.AddItem i & " week"
        Combo2.AddItem DateAdd("d", 7 - Weekday(dbDatex), dbDatex)
        Debug.Print "-" & DateAdd("d", 7 - Weekday(dbDatex), dbDatex) & vbCrLf
        
        dbDatex = DateAdd("d", d + 1, dbDatex)
    Loop
    If cboWeek.ListCount > 0 Then
    cboWeek.ListIndex = 0
    End If
End Sub

Private Sub cboWeek_Click()
    Combo1.ListIndex = cboWeek.ListIndex
    Combo2.ListIndex = cboWeek.ListIndex
End Sub



Private Sub cboYear_click()
     If cboYear.ListIndex <> -1 And cboMonth.ListIndex <> -1 Then
        cboMonth_click
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "UNITS RECEIVE WEEKLY PERFORMANCE REPORT") = False Then Exit Sub

    If cboWeek.Text = "" Then
        MsgBox "Choose a Week", vbInformation, "CSMS"
        cboWeek.SetFocus
        Exit Sub
    End If

    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT REP_OR FROM CSMS_REPOR WHERE DTE_COMP IS NOT NULL AND DTE_COMP BETWEEN '" & Combo1.Text & "' AND '" & Combo2.Text & "' AND TRANSTYPE = 'R' ORDER BY DTE_COMP, REP_OR ASC")
    If (rstmp.BOF And rstmp.EOF) Then
        ShowNoRecord
        Exit Sub
    End If
    Set rstmp = Nothing
    
    Screen.MousePointer = 11
    'Load frmSplash
    'frmSplash.labCon.Caption = "Calculating Workshop Weekly Performance Details"
    'frmSplash.Show
    DoEvents

    '    Set xlApp = CreateObject("Excel.Application")
    '    Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "WEEKLY UNIT RECEIVED.xls")
    '    Set xlSheet = xlBook.Worksheets(1)

    Call GETDATERANGE

    '    xlSheet.Cells(7, "Q") = Combo1 & " - " & Combo2
    '    xlSheet.Cells(7, "C") = COMPANY_NAME                                    'DEALER NAME


    Dim Month_Value                                    As Integer
    Dim Year_Value                                     As Integer
    Dim Lastday_Of_The_Month                           As Integer
    Dim Day_Value                                      As Integer
    Dim DateRange                                      As String
    Dim Monday_Of_The_Week                             As Date
    Dim Tuesday_Of_The_Week                            As Date
    Dim Wednesday_Of_The_Week                          As Date
    Dim Thursday_Of_The_Week                           As Date
    Dim Friday_Of_The_Week                             As Date
    Dim Saturday_Of_The_Week                           As Date
    Dim Sunday_Of_The_Week                             As Date
    Dim COUNTER                                        As Integer
    Dim SQL                                            As String
    Dim temprs                                         As ADODB.Recordset

    Month_Value = Val(What_month(cboMonth.Text))
    Year_Value = Val(cboYear.Text)
    Lastday_Of_The_Month = Day(lastDay(DateSerial(Year_Value, Month_Value, 1)))
    COUNTER = 0

    For Day_Value = 1 To Lastday_Of_The_Month
        If WeekdayName(Weekday(DateSerial(Year_Value, Month_Value, Day_Value))) = "Monday" Then
            COUNTER = COUNTER + 1
            If Val(Left(cboWeek, 1)) = COUNTER Then
                Monday_Of_The_Week = DateSerial(Year_Value, Month_Value, Day_Value)
                Tuesday_Of_The_Week = DateSerial(Year_Value, Month_Value, Day_Value) + 1
                Wednesday_Of_The_Week = DateSerial(Year_Value, Month_Value, Day_Value) + 2
                Thursday_Of_The_Week = DateSerial(Year_Value, Month_Value, Day_Value) + 3
                Friday_Of_The_Week = DateSerial(Year_Value, Month_Value, Day_Value) + 4
                Saturday_Of_The_Week = DateSerial(Year_Value, Month_Value, Day_Value) + 5
                Sunday_Of_The_Week = DateSerial(Year_Value, Month_Value, Day_Value) + 6
                Exit For
            End If
        End If
        '        If Day_Value = Lastday_Of_The_Month And Val(Left(cboWeek, 1)) <> counter Then
        '            MsgBox "There's no 5th week on the Month and Year selected!", vbCritical, "Error"
        '
        '            Exit Sub
        '        End If
    Next Day_Value
    RPT.Reset
    RPT.Formulas(0) = "WeekNumber = '" & cboWeek.Text & "'"
    RPT.Formulas(1) = "MonthReport = '" & cboMonth.Text & "'"
    RPT.Formulas(2) = "YearReport = '" & cboYear.Text & "'"

    'WeekdayName

    Dim rsREPOR                                        As New ADODB.Recordset
    Dim rsDet                                          As New ADODB.Recordset
    Dim INSLABOR                                       As Double

    Dim MON_FPM As Double: Dim MON_A1 As Double: Dim MON_A2 As Double: Dim MON_B As Double: Dim MON_C As Double: Dim MON_X As Double
    Dim TUE_FPM As Double: Dim TUE_A1 As Double: Dim TUE_A2 As Double: Dim TUE_B As Double: Dim TUE_C As Double: Dim TUE_X As Double
    Dim WEN_FPM As Double: Dim WEN_A1 As Double: Dim WEN_A2 As Double: Dim WEN_B As Double: Dim WEN_C As Double: Dim WEN_X As Double
    Dim THU_FPM As Double: Dim THU_A1 As Double: Dim THU_A2 As Double: Dim THU_B As Double: Dim THU_C As Double: Dim THU_X As Double
    Dim FRI_FPM As Double: Dim FRI_A1 As Double: Dim FRI_A2 As Double: Dim FRI_B As Double: Dim FRI_C As Double: Dim FRI_X As Double
    Dim SAT_FPM As Double: Dim SAT_A1 As Double: Dim SAT_A2 As Double: Dim SAT_B As Double: Dim SAT_C As Double: Dim SAT_X As Double
    Dim SUN_FPM As Double: Dim SUN_A1 As Double: Dim SUN_A2 As Double: Dim SUN_B As Double: Dim SUN_C As Double: Dim SUN_X As Double

    Dim MON_D As Double: Dim MON_E As Double: Dim MON_F As Double: Dim MON_G As Double: Dim MON_Y As Double: Dim MON_DPI As Double
    Dim TUE_D As Double: Dim TUE_E As Double: Dim TUE_F As Double: Dim TUE_G As Double: Dim TUE_Y As Double: Dim TUE_DPI As Double
    Dim WEN_D As Double: Dim WEN_E As Double: Dim WEN_F As Double: Dim WEN_G As Double: Dim WEN_Y As Double: Dim WEN_DPI As Double
    Dim THU_D As Double: Dim THU_E As Double: Dim THU_F As Double: Dim THU_G As Double: Dim THU_Y As Double: Dim THU_DPI As Double
    Dim FRI_D As Double: Dim FRI_E As Double: Dim FRI_F As Double: Dim FRI_G As Double: Dim FRI_Y As Double: Dim FRI_DPI As Double
    Dim SAT_D As Double: Dim SAT_E As Double: Dim SAT_F As Double: Dim SAT_G As Double: Dim SAT_Y As Double: Dim SAT_DPI As Double
    Dim SUN_D As Double: Dim SUN_E As Double: Dim SUN_F As Double: Dim SUN_G As Double: Dim SUN_Y As Double: Dim SUN_DPI As Double

    Dim MON_IN As Double: Dim TUE_IN As Double: Dim WEN_IN As Double: Dim THU_IN As Double: Dim FRI_IN As Double: Dim SAT_IN As Double: Dim SUN_IN As Double
    Dim MON_IN_O As Double: Dim TUE_IN_O As Double: Dim WEN_IN_O As Double: Dim THU_IN_O As Double: Dim FRI_IN_O As Double: Dim SAT_IN_O As Double: Dim SUN_IN_O As Double

    Dim TRIG_GJ_FPM As Integer: Dim TRIG_GJ_PMS As Integer: Dim TRIG_GJ_CST As Integer: Dim TRIG_GJ_WAR As Integer: Dim TRIG_GJ_INT As Integer: Dim TRIG_GJ_OTH As Integer
    Dim TRIG_BP_INS As Integer: Dim TRIG_BP_CST As Integer: Dim TRIG_BP_WRT As Integer: Dim TRIG_BP_INT As Integer: Dim TRIG_BP_OTH As Integer

    Dim TRIGER_INS                                     As String

    Set rsREPOR = New ADODB.Recordset
    'Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_COMP IS NOT NULL AND DTE_COMP BETWEEN '" & Monday_Of_The_Week & "' AND '" & Sunday_Of_The_Week & "' AND TRANSTYPE = 'R' ORDER BY DTE_COMP, REP_OR ASC ")

    Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_COMP IS NOT NULL AND DTE_COMP BETWEEN '" & Combo1.Text & "' AND '" & Combo2.Text & "' AND TRANSTYPE = 'R' ORDER BY DTE_COMP, REP_OR ASC")
    'Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_COMP IS NOT NULL AND DTE_COMP BETWEEN '6/1/2008' AND '6/7/2008' AND TRANSTYPE = 'R' ORDER BY DTE_COMP, REP_OR ASC")
    'Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_COMP IS NOT NULL AND DTE_COMP = '6/5/2008' AND TRANSTYPE = 'R' ORDER BY DTE_COMP, REP_OR ASC")
    'Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_COMP IS NOT NULL AND REP_OR = 'R-00000540' AND TRANSTYPE = 'R' ORDER BY DTE_COMP, REP_OR ASC")
    If Not (rsREPOR.BOF And rsREPOR.EOF) Then
        Do While Not rsREPOR.EOF
            TRIG_GJ_FPM = 0: TRIG_GJ_PMS = 0: TRIG_GJ_CST = 0: TRIG_GJ_WAR = 0: TRIG_GJ_INT = 0: TRIG_GJ_OTH = 0
            TRIG_BP_INS = 0: TRIG_BP_CST = 0: TRIG_BP_WRT = 0: TRIG_BP_INT = 0: TRIG_BP_OTH = 0

            TRIGER_INS = ""
            If Not MonthName(Month(rsREPOR!dte_comp)) = cboMonth.Text Then GoTo CONT_NEXT

            INSLABOR = NumericVal(rsREPOR!PARTLABOR)

            Set rsDet = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & rsREPOR!REP_OR & "' AND LIVIL = '1' ORDER BY WCODE,LINE_NO ASC")
            If Not (rsDet.BOF And rsDet.EOF) Then
                If CHECKIFHYUNDAI(Null2String(rsREPOR!PLATE_NO)) = True Then
                    Do While Not rsDet.EOF
                        If Null2String(rsDet!JOBTYPE) = "GJ" Or Null2String(rsDet!JOBTYPE) = "CND" Or Null2String(rsDet!JOBTYPE) = "" Then
                            If Null2String(rsDet!wCode) = "" Then
                                If INSLABOR > 0 Then
                                    If INSLABOR > (rsDet!DET_AMT - NumericVal(rsDet!Discount_2)) Then
                                        INSLABOR = INSLABOR - (rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                                    Else
                                        INSLABOR = 0
                                    End If
                                End If
                                If TRIG_GJ_CST = 0 Then
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_A2 = MON_A2 + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_A2 = TUE_A2 + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_A2 = WEN_A2 + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_A2 = THU_A2 + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_A2 = FRI_A2 + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_A2 = SAT_A2 + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_A2 = SUN_A2 + 1
                                    TRIG_GJ_CST = 1
                                End If
                            ElseIf Null2String(rsDet!wCode) = "W" Then
                                If TRIG_GJ_WAR = 0 Then
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_B = MON_B + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_B = TUE_B + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_B = WEN_B + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_B = THU_B + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_B = FRI_B + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_B = SAT_B + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_B = SUN_B + 1
                                    TRIG_GJ_WAR = 1
                                End If
                            ElseIf Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                                If TRIG_GJ_INT = 0 Then
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_C = MON_C + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_C = TUE_C + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_C = WEN_C + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_C = THU_C + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_C = FRI_C + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_C = SAT_C + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_C = SUN_C + 1
                                    TRIG_GJ_INT = 1
                                End If
                            End If
                        ElseIf Null2String(rsDet!JOBTYPE) = "PMS" Then
                            If Null2String(rsDet!STATUS1) = "" Then    'OLD VERSION TAGGING OF FPM
                                If Left(Null2String(rsDet!DETDSC), 5) = "1,000" Or Left(Null2String(rsDet!DETDSC), 5) = "5,000" Then
                                    If TRIG_GJ_FPM = 0 Then
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_FPM = MON_FPM + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_FPM = TUE_FPM + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_FPM = WEN_FPM + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_FPM = THU_FPM + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_FPM = FRI_FPM + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_FPM = SAT_FPM + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_FPM = SUN_FPM + 1
                                        TRIG_GJ_FPM = 1
                                    End If
                                Else
                                    If TRIG_GJ_PMS = 0 Then
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_A1 = MON_A1 + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_A1 = TUE_A1 + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_A1 = WEN_A1 + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_A1 = THU_A1 + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_A1 = FRI_A1 + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_A1 = SAT_A1 + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_A1 = SUN_A1 + 1
                                        TRIG_GJ_PMS = 1
                                    End If
                                End If
                            Else
                                If Null2String(rsDet!STATUS1) = "Y" Then
                                    If TRIG_GJ_FPM = 0 Then
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_FPM = MON_FPM + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_FPM = TUE_FPM + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_FPM = WEN_FPM + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_FPM = THU_FPM + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_FPM = FRI_FPM + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_FPM = SAT_FPM + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_FPM = SUN_FPM + 1
                                        TRIG_GJ_FPM = 1
                                    End If
                                Else
                                    If TRIG_GJ_PMS = 0 Then
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_A1 = MON_A1 + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_A1 = TUE_A1 + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_A1 = WEN_A1 + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_A1 = THU_A1 + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_A1 = FRI_A1 + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_A1 = SAT_A1 + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_A1 = SUN_A1 + 1
                                        TRIG_GJ_PMS = 1
                                    End If
                                End If
                            End If
                        ElseIf Null2String(rsDet!JOBTYPE) = "BP" Then
                            If Null2String(rsDet!wCode) = "" Then
                                If INSLABOR > 0 Then
                                    If INSLABOR >= (rsDet!DET_AMT - NumericVal(rsDet!Discount_2)) Then
                                        INSLABOR = INSLABOR - (rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                                        If TRIG_BP_INS = 0 Then
                                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_D = MON_D + 1
                                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_D = TUE_D + 1
                                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_D = WEN_D + 1
                                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_D = THU_D + 1
                                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_D = FRI_D + 1
                                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_D = SAT_D + 1
                                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_D = SUN_D + 1
                                            TRIG_BP_INS = 1
                                        End If
                                    Else
                                        INSLABOR = 0
                                        If TRIG_BP_INS = 0 Then
                                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_D = MON_D + 1
                                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_D = TUE_D + 1
                                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_D = WEN_D + 1
                                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_D = THU_D + 1
                                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_D = FRI_D + 1
                                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_D = SAT_D + 1
                                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_D = SUN_D + 1
                                            TRIG_BP_INS = 1
                                        End If
                                    End If
                                Else
                                    If TRIG_BP_CST = 0 Then
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_E = MON_E + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_E = TUE_E + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_E = WEN_E + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_E = THU_E + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_E = FRI_E + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_E = SAT_E + 1
                                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_E = SUN_E + 1
                                        TRIG_BP_CST = 1
                                    End If
                                End If
                            ElseIf Null2String(rsDet!wCode) = "W" Then
                                If TRIG_BP_WRT = 0 Then
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_F = MON_F + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_F = TUE_F + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_F = WEN_F + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_F = THU_F + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_F = FRI_F + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_F = SAT_F + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_F = SUN_F + 1
                                    TRIG_BP_WRT = 1
                                End If
                            ElseIf Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                                If TRIG_BP_INT = 0 Then
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_G = MON_G + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_G = TUE_G + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_G = WEN_G + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_G = THU_G + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_G = FRI_G + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_G = SAT_G + 1
                                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_G = SUN_G + 1
                                    TRIG_BP_INT = 1
                                End If
                            End If
                        End If

                        rsDet.MoveNext
                    Loop
                Else
                    Do While Not rsDet.EOF
                        If Null2String(rsDet!JOBTYPE) = "BP" Then
                            If TRIG_BP_OTH = 0 Then
                                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_Y = MON_Y + 1
                                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_Y = TUE_Y + 1
                                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_Y = WEN_Y + 1
                                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_Y = THU_Y + 1
                                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_Y = FRI_Y + 1
                                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_Y = SAT_Y + 1
                                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_Y = SUN_Y + 1
                                TRIG_BP_OTH = 1
                            End If
                        Else
                            If TRIG_GJ_OTH = 0 Then
                                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_X = MON_X + 1
                                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_X = TUE_X + 1
                                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_X = WEN_X + 1
                                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_X = THU_X + 1
                                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_X = FRI_X + 1
                                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_X = SAT_X + 1
                                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_X = SUN_X + 1
                                TRIG_GJ_OTH = 1
                            End If
                        End If

                        rsDet.MoveNext
                    Loop
                End If
                Set rsDet = Nothing
            End If

CONT_NEXT:

            rsREPOR.MoveNext
        Loop
        Set rsREPOR = Nothing
    End If

    RPT.Formulas(3) = "GeneralJob_FPI_Mon = " & MON_FPM
    RPT.Formulas(4) = "GeneralJob_FPI_TUE = " & TUE_FPM
    RPT.Formulas(5) = "GeneralJob_FPI_WED = " & WEN_FPM
    RPT.Formulas(6) = "GeneralJob_FPI_THU = " & THU_FPM
    RPT.Formulas(7) = "GeneralJob_FPI_FRI = " & FRI_FPM
    RPT.Formulas(8) = "GeneralJob_FPI_SAT = " & SAT_FPM
    RPT.Formulas(9) = "GeneralJob_FPI_SUN = " & SUN_FPM

    RPT.Formulas(10) = "GeneralJob_A1_Mon = " & MON_A1
    RPT.Formulas(11) = "GeneralJob_A1_TUE = " & TUE_A1
    RPT.Formulas(12) = "GeneralJob_A1_WED = " & WEN_A1
    RPT.Formulas(13) = "GeneralJob_A1_THU = " & THU_A1
    RPT.Formulas(14) = "GeneralJob_A1_FRI = " & FRI_A1
    RPT.Formulas(15) = "GeneralJob_A1_SAT = " & SAT_A1
    RPT.Formulas(16) = "GeneralJob_A1_SUN = " & SUN_A1

    RPT.Formulas(17) = "GeneralJob_A2_Mon = " & MON_A2
    RPT.Formulas(18) = "GeneralJob_A2_TUE = " & TUE_A2
    RPT.Formulas(19) = "GeneralJob_A2_WED = " & WEN_A2
    RPT.Formulas(20) = "GeneralJob_A2_THU = " & THU_A2
    RPT.Formulas(21) = "GeneralJob_A2_FRI = " & FRI_A2
    RPT.Formulas(22) = "GeneralJob_A2_SAT = " & SAT_A2
    RPT.Formulas(23) = "GeneralJob_A2_SUN = " & SUN_A2

    RPT.Formulas(24) = "GeneralJob_B_Mon = " & MON_B
    RPT.Formulas(25) = "GeneralJob_B_TUE = " & TUE_B
    RPT.Formulas(26) = "GeneralJob_B_WED = " & WEN_B
    RPT.Formulas(27) = "GeneralJob_B_THU = " & THU_B
    RPT.Formulas(28) = "GeneralJob_B_FRI = " & FRI_B
    RPT.Formulas(29) = "GeneralJob_B_SAT = " & SAT_B
    RPT.Formulas(30) = "GeneralJob_B_SUN = " & SUN_B

    RPT.Formulas(31) = "GeneralJob_C_Mon = " & MON_C
    RPT.Formulas(32) = "GeneralJob_C_TUE = " & TUE_C
    RPT.Formulas(33) = "GeneralJob_C_WED = " & WEN_C
    RPT.Formulas(34) = "GeneralJob_C_THU = " & THU_C
    RPT.Formulas(35) = "GeneralJob_C_FRI = " & FRI_C
    RPT.Formulas(36) = "GeneralJob_C_SAT = " & SAT_C
    RPT.Formulas(37) = "GeneralJob_C_SUN = " & SUN_C

    RPT.Formulas(38) = "GeneralJob_X_Mon = " & MON_X
    RPT.Formulas(39) = "GeneralJob_X_TUE = " & TUE_X
    RPT.Formulas(40) = "GeneralJob_X_WED = " & WEN_X
    RPT.Formulas(41) = "GeneralJob_X_THU = " & THU_X
    RPT.Formulas(42) = "GeneralJob_X_FRI = " & FRI_X
    RPT.Formulas(43) = "GeneralJob_X_SAT = " & SAT_X
    RPT.Formulas(44) = "GeneralJob_X_SUN = " & SUN_X

    RPT.Formulas(45) = "BodyAndPaint_D_Mon = " & MON_D
    RPT.Formulas(46) = "BodyAndPaint_D_TUE = " & TUE_D
    RPT.Formulas(47) = "BodyAndPaint_D_WED = " & WEN_D
    RPT.Formulas(48) = "BodyAndPaint_D_THU = " & THU_D
    RPT.Formulas(49) = "BodyAndPaint_D_FRI = " & FRI_D
    RPT.Formulas(50) = "BodyAndPaint_D_SAT = " & SAT_D
    RPT.Formulas(51) = "BodyAndPaint_D_SUN = " & SUN_D

    RPT.Formulas(52) = "BodyAndPaint_E_Mon = " & MON_E
    RPT.Formulas(53) = "BodyAndPaint_E_TUE = " & TUE_E
    RPT.Formulas(54) = "BodyAndPaint_E_WED = " & WEN_E
    RPT.Formulas(55) = "BodyAndPaint_E_THU = " & THU_E
    RPT.Formulas(56) = "BodyAndPaint_E_FRI = " & FRI_E
    RPT.Formulas(57) = "BodyAndPaint_E_SAT = " & SAT_E
    RPT.Formulas(58) = "BodyAndPaint_E_SUN = " & SUN_E

    RPT.Formulas(59) = "BodyAndPaint_F_Mon = " & MON_F
    RPT.Formulas(60) = "BodyAndPaint_F_TUE = " & TUE_F
    RPT.Formulas(61) = "BodyAndPaint_F_WED = " & WEN_F
    RPT.Formulas(62) = "BodyAndPaint_F_THU = " & THU_F
    RPT.Formulas(63) = "BodyAndPaint_F_FRI = " & FRI_F
    RPT.Formulas(64) = "BodyAndPaint_F_SAT = " & SAT_F
    RPT.Formulas(65) = "BodyAndPaint_F_SUN = " & SUN_F

    RPT.Formulas(66) = "BodyAndPaint_G_Mon = " & MON_G
    RPT.Formulas(67) = "BodyAndPaint_G_TUE = " & TUE_G
    RPT.Formulas(68) = "BodyAndPaint_G_WED = " & WEN_G
    RPT.Formulas(69) = "BodyAndPaint_G_THU = " & THU_G
    RPT.Formulas(70) = "BodyAndPaint_G_FRI = " & FRI_G
    RPT.Formulas(71) = "BodyAndPaint_G_SAT = " & SAT_G
    RPT.Formulas(72) = "BodyAndPaint_G_SUN = " & SUN_G

    RPT.Formulas(73) = "BodyAndPaint_Y_Mon = " & MON_Y
    RPT.Formulas(74) = "BodyAndPaint_Y_TUE = " & TUE_Y
    RPT.Formulas(75) = "BodyAndPaint_Y_WED = " & WEN_Y
    RPT.Formulas(76) = "BodyAndPaint_Y_THU = " & THU_Y
    RPT.Formulas(77) = "BodyAndPaint_Y_FRI = " & FRI_Y
    RPT.Formulas(78) = "BodyAndPaint_Y_SAT = " & SAT_Y
    RPT.Formulas(79) = "BodyAndPaint_Y_SUN = " & SUN_Y

    'Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_RECD BETWEEN '" & Monday_Of_The_Week & "' AND '" & Sunday_Of_The_Week & "' and transtype = 'R' ORDER BY DTE_RECD ")
    Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_RECD BETWEEN '" & Combo1.Text & "' AND '" & Combo2.Text & "' and transtype = 'R' ORDER BY DTE_RECD ")
    'Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_RECD BETWEEN '" & "6/3/2008" & "' AND '" & "6/3/2008" & "' and transtype = 'R' ORDER BY DTE_RECD ")
    If Not (rsREPOR.BOF And rsREPOR.EOF) Then
        Do While Not rsREPOR.EOF
            If Not MonthName(Month(rsREPOR!DTE_RECD)) = cboMonth.Text Then GoTo CONT_NEXT1

            If CHECKIFHYUNDAI(Null2String(rsREPOR!PLATE_NO)) = True Then
                If UCase(WeekdayName(Weekday(rsREPOR!DTE_RECD))) = "MONDAY" Then MON_IN = MON_IN + 1
                If UCase(WeekdayName(Weekday(rsREPOR!DTE_RECD))) = "TUESDAY" Then TUE_IN = TUE_IN + 1
                If UCase(WeekdayName(Weekday(rsREPOR!DTE_RECD))) = "WEDNESDAY" Then WEN_IN = WEN_IN + 1
                If UCase(WeekdayName(Weekday(rsREPOR!DTE_RECD))) = "THURSDAY" Then THU_IN = THU_IN + 1
                If UCase(WeekdayName(Weekday(rsREPOR!DTE_RECD))) = "FRIDAY" Then FRI_IN = FRI_IN + 1
                If UCase(WeekdayName(Weekday(rsREPOR!DTE_RECD))) = "SATURDAY" Then SAT_IN = SAT_IN + 1
                If UCase(WeekdayName(Weekday(rsREPOR!DTE_RECD))) = "SUNDAY" Then MON_IN = MON_IN + 1
            Else
                If UCase(WeekdayName(Weekday(rsREPOR!DTE_RECD))) = "MONDAY" Then MON_IN_O = MON_IN_O + 1
                If UCase(WeekdayName(Weekday(rsREPOR!DTE_RECD))) = "TUESDAY" Then TUE_IN_O = TUE_IN_O + 1
                If UCase(WeekdayName(Weekday(rsREPOR!DTE_RECD))) = "WEDNESDAY" Then WEN_IN_O = WEN_IN_O + 1
                If UCase(WeekdayName(Weekday(rsREPOR!DTE_RECD))) = "THURSDAY" Then THU_IN_O = THU_IN_O + 1
                If UCase(WeekdayName(Weekday(rsREPOR!DTE_RECD))) = "FRIDAY" Then FRI_IN_O = FRI_IN_O + 1
                If UCase(WeekdayName(Weekday(rsREPOR!DTE_RECD))) = "SATURDAY" Then SAT_IN_O = SAT_IN_O + 1
                If UCase(WeekdayName(Weekday(rsREPOR!DTE_RECD))) = "SUNDAY" Then MON_IN_O = MON_IN_O + 1
            End If

CONT_NEXT1:
            rsREPOR.MoveNext
        Loop
    End If
    Set rsREPOR = Nothing

    RPT.Formulas(80) = "UnitsReceived_IN_MON = " & MON_IN
    RPT.Formulas(81) = "UnitsReceived_IN_TUE = " & TUE_IN
    RPT.Formulas(82) = "UnitsReceived_IN_WED = " & WEN_IN
    RPT.Formulas(83) = "UnitsReceived_IN_THU = " & THU_IN
    RPT.Formulas(84) = "UnitsReceived_IN_FRI = " & FRI_IN
    RPT.Formulas(85) = "UnitsReceived_IN_SAT = " & SAT_IN
    RPT.Formulas(86) = "UnitsReceived_IN_SUN = " & SUN_IN

    RPT.Formulas(87) = "UnitsReceived_OTHERBRANDS_IN_MON = " & MON_IN_O
    RPT.Formulas(175) = "UnitsReceived_OTHERBRANDS_IN_TUE = " & TUE_IN_O
    RPT.Formulas(88) = "UnitsReceived_OTHERBRANDS_IN_WED = " & WEN_IN_O
    RPT.Formulas(89) = "UnitsReceived_OTHERBRANDS_IN_THU = " & THU_IN_O
    RPT.Formulas(90) = "UnitsReceived_OTHERBRANDS_IN_FRI = " & FRI_IN_O
    RPT.Formulas(91) = "UnitsReceived_OTHERBRANDS_IN_SAT = " & SAT_IN_O
    RPT.Formulas(92) = "UnitsReceived_OTHERBRANDS_IN_SUN = " & SUN_IN_O

    Dim MON_OUT As Double: Dim TUE_OUT As Double: Dim WEN_OUT As Double: Dim THU_OUT As Double: Dim FRI_OUT As Double: Dim SAT_OUT As Double: Dim SUN_OUT As Double
    Dim MON_OUT_O As Double: Dim TUE_OUT_O As Double: Dim WEN_OUT_O As Double: Dim THU_OUT_O As Double: Dim FRI_OUT_O As Double: Dim SAT_OUT_O As Double: Dim SUN_OUT_O As Double

    'Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_REL BETWEEN '" & Monday_Of_The_Week & "' AND '" & Sunday_Of_The_Week & "' AND TRANSTYPE = 'R' ORDER BY DTE_REL ")
    Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_REL BETWEEN '" & Combo1.Text & "' AND '" & Combo2.Text & "' AND TRANSTYPE = 'R' ORDER BY DTE_REL ")
    'Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_REL BETWEEN '" & "6/3/2008" & "' AND '" & "6/3/2008" & "' AND TRANSTYPE = 'R' ORDER BY DTE_REL ")
    If Not (rsREPOR.BOF And rsREPOR.EOF) Then
        Do While Not rsREPOR.EOF
            If Not MonthName(Month(rsREPOR!dte_rel)) = cboMonth.Text Then GoTo CONT_NEXT2

            If CHECKIFHYUNDAI(Null2String(rsREPOR!PLATE_NO)) = True Then
                If UCase(WeekdayName(Weekday(rsREPOR!dte_rel))) = "MONDAY" Then MON_OUT = MON_OUT + 1
                If UCase(WeekdayName(Weekday(rsREPOR!dte_rel))) = "TUESDAY" Then TUE_OUT = TUE_OUT + 1
                If UCase(WeekdayName(Weekday(rsREPOR!dte_rel))) = "WEDNESDAY" Then WEN_OUT = WEN_OUT + 1
                If UCase(WeekdayName(Weekday(rsREPOR!dte_rel))) = "THURSDAY" Then THU_OUT = THU_OUT + 1
                If UCase(WeekdayName(Weekday(rsREPOR!dte_rel))) = "FRIDAY" Then FRI_OUT = FRI_OUT + 1
                If UCase(WeekdayName(Weekday(rsREPOR!dte_rel))) = "SATURDAY" Then SAT_OUT = SAT_OUT + 1
                If UCase(WeekdayName(Weekday(rsREPOR!dte_rel))) = "SUNDAY" Then SUN_OUT = SUN_OUT + 1
            Else
                If UCase(WeekdayName(Weekday(rsREPOR!dte_rel))) = "MONDAY" Then MON_OUT_O = MON_OUT_O + 1
                If UCase(WeekdayName(Weekday(rsREPOR!dte_rel))) = "TUESDAY" Then TUE_OUT_O = TUE_OUT_O + 1
                If UCase(WeekdayName(Weekday(rsREPOR!dte_rel))) = "WEDNESDAY" Then WEN_OUT_O = WEN_OUT_O + 1
                If UCase(WeekdayName(Weekday(rsREPOR!dte_rel))) = "THURSDAY" Then THU_OUT_O = THU_OUT_O + 1
                If UCase(WeekdayName(Weekday(rsREPOR!dte_rel))) = "FRIDAY" Then FRI_OUT_O = FRI_OUT_O + 1
                If UCase(WeekdayName(Weekday(rsREPOR!dte_rel))) = "SATURDAY" Then SAT_OUT_O = SAT_OUT_O + 1
                If UCase(WeekdayName(Weekday(rsREPOR!dte_rel))) = "SUNDAY" Then SUN_OUT_O = SUN_OUT_O + 1
            End If

CONT_NEXT2:
            rsREPOR.MoveNext
        Loop
    End If
    Set rsREPOR = Nothing

    RPT.Formulas(93) = "UnitsReceived_OUT_MON = " & MON_OUT
    RPT.Formulas(94) = "UnitsReceived_OUT_TUE = " & TUE_OUT
    RPT.Formulas(95) = "UnitsReceived_OUT_WED = " & WEN_OUT
    RPT.Formulas(96) = "UnitsReceived_OUT_THU = " & THU_OUT
    RPT.Formulas(97) = "UnitsReceived_OUT_FRI = " & FRI_OUT
    RPT.Formulas(98) = "UnitsReceived_OUT_SAT = " & SAT_OUT
    RPT.Formulas(99) = "UnitsReceived_OUT_SUN = " & SUN_OUT

    RPT.Formulas(100) = "UnitsReceived_otherbrandS_OUT_MON = " & MON_OUT_O
    RPT.Formulas(101) = "UnitsReceived_otherbrandS_OUT_TUE = " & TUE_OUT_O
    RPT.Formulas(102) = "UnitsReceived_otherbrandS_OUT_WED = " & WEN_OUT_O
    RPT.Formulas(103) = "UnitsReceived_otherbrandS_OUT_THU = " & THU_OUT_O
    RPT.Formulas(104) = "UnitsReceived_otherbrandS_OUT_FRI = " & FRI_OUT_O
    RPT.Formulas(105) = "UnitsReceived_otherbrandS_OUT_SAT = " & SAT_OUT_O
    RPT.Formulas(106) = "UnitsReceived_otherbrandS_OUT_SUN = " & SUN_OUT_O

    RPT.Formulas(158) = "UnitsServiced_Date_Mon = '" & X_MON & "'"
    RPT.Formulas(159) = "UnitsServiced_Date_Tue = '" & X_TUE & "'"
    RPT.Formulas(160) = "UnitsServiced_Date_Wed = '" & X_WEN & "'"
    RPT.Formulas(161) = "UnitsServiced_Date_Thu = '" & X_THU & "'"
    RPT.Formulas(162) = "UnitsServiced_Date_Fri = '" & X_FRI & "'"
    RPT.Formulas(163) = "UnitsServiced_Date_Sat = '" & X_SAT & "'"
    RPT.Formulas(164) = "UnitsServiced_Date_Sun = '" & X_SUN & "'"

    RPT.Formulas(165) = "UnitsReceived_Date_Mon = '" & X_MON & "'"
    RPT.Formulas(166) = "UnitsReceived_Date_Tue = '" & X_TUE & "'"
    RPT.Formulas(167) = "UnitsReceived_Date_Wed = '" & X_WEN & "'"
    RPT.Formulas(168) = "UnitsReceived_Date_Thu = '" & X_THU & "'"
    RPT.Formulas(169) = "UnitsReceived_Date_Fri = '" & X_FRI & "'"
    RPT.Formulas(170) = "UnitsReceived_Date_Sat = '" & X_SAT & "'"
    RPT.Formulas(171) = "UnitsReceived_Date_Sun = '" & X_SUN & "'"

    RPT.Formulas(172) = "PeriodCovered = '" & Combo1.Text & "-" & Combo2.Text & " '"
    DoEvents
    'JUN 01/05/2008
    RPT.WindowTitle = "Units Received Weekly Performance Report"
    RPT.Formulas(173) = "Company Name = '" & COMPANY_NAME & "'"
    RPT.Formulas(174) = "Company Address = '" & COMPANY_ADDRESS & "'"

    PrintSQLReport RPT, CSMS_REPORT_PATH & "UnitsReceivedWeeklyPerformanceReport.rpt", "", CSMS_REPORT_CONNECTION, 1

    'LogAudit "V", "UNIT RECEIVED WEEKLY PERFORMANCE REPORT", cboWeek & cboMonth & cboYear
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("G", "UNITS RECEIVE WEEKLY PERFORMANCE REPORT", "", "", "", cboMonth & " " & cboWeek & " " & cboYear, "", "")
    'NEW LOG AUDIT-----------------------------------------------------


    '    xlApp.Windows.ITEM(1).Caption = "WEEKLY UNIT RECEIVED REPORT FOR THE MONTH OF " & cboMonth & " " & cboYear & " " & cboWeek
    '    xlApp.Visible = True
    '    Set xlApp = Nothing

    Screen.MousePointer = 0
    'Unload frmSplash
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (UNITS RECEIVE WEEKLY PERFORMANCE REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "UNITS RECEIVE WEEKLY PERFORMANCE REPORT", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
    fillcbomonth cboMonth
    FillCboMoreYear cboYear
    FillcboWeek
    'cboWeek.Text = "1st Week"
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)

End Sub

