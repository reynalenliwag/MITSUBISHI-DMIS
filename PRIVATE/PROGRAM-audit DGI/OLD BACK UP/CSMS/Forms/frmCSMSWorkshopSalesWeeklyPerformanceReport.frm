VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMSWorkshopSalesWeeklyPerformanceReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Workshop Weekly"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3240
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMSWorkshopSalesWeeklyPerformanceReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2175
   ScaleWidth      =   3240
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
      Left            =   915
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select month from the list"
      Top             =   60
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   600
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5490
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   570
      Width           =   2175
   End
   Begin VB.ComboBox cboWeek 
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
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select week from the list"
      Top             =   840
      Width           =   1395
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
      Left            =   915
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Select year from the list"
      Top             =   450
      Width           =   2175
   End
   Begin Crystal.CrystalReport rpt 
      Left            =   180
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Workshop Sales Weekly Performance Report"
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
      Left            =   2310
      MouseIcon       =   "frmCSMSWorkshopSalesWeeklyPerformanceReport.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSWorkshopSalesWeeklyPerformanceReport.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   1290
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   795
      Left            =   1590
      MouseIcon       =   "frmCSMSWorkshopSalesWeeklyPerformanceReport.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSWorkshopSalesWeeklyPerformanceReport.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   1290
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Week"
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
      Height          =   255
      Left            =   105
      TabIndex        =   7
      Top             =   900
      Width           =   735
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   105
      TabIndex        =   6
      Top             =   510
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   105
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmCSMSWorkshopSalesWeeklyPerformanceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim X_MON As String: Dim X_TUE As String: Dim X_WEN    As String
Attribute X_TUE.VB_VarUserMemId = 1073938432
Attribute X_WEN.VB_VarUserMemId = 1073938432
Dim X_THU As String: Dim X_FRI As String: Dim X_SAT As String: Dim X_SUN As String
Attribute X_THU.VB_VarUserMemId = 1073938435
Attribute X_FRI.VB_VarUserMemId = 1073938435
Attribute X_SAT.VB_VarUserMemId = 1073938435
Attribute X_SUN.VB_VarUserMemId = 1073938435


Dim MON_FPM As Double: Dim MON_A1 As Double: Dim MON_A2 As Double: Dim MON_B As Double: Dim MON_C As Double: Dim MON_X As Double: Dim MON_A0 As Double
Attribute MON_FPM.VB_VarUserMemId = 1073938439
Attribute MON_A1.VB_VarUserMemId = 1073938439
Attribute MON_A2.VB_VarUserMemId = 1073938439
Attribute MON_B.VB_VarUserMemId = 1073938439
Attribute MON_C.VB_VarUserMemId = 1073938439
Attribute MON_X.VB_VarUserMemId = 1073938439
Attribute MON_A0.VB_VarUserMemId = 1073938439
Dim TUE_FPM As Double: Dim TUE_A1 As Double: Dim TUE_A2 As Double: Dim TUE_B As Double: Dim TUE_C As Double: Dim TUE_X As Double: Dim TUE_A0 As Double
Attribute TUE_FPM.VB_VarUserMemId = 1073938446
Attribute TUE_A1.VB_VarUserMemId = 1073938446
Attribute TUE_A2.VB_VarUserMemId = 1073938446
Attribute TUE_B.VB_VarUserMemId = 1073938446
Attribute TUE_C.VB_VarUserMemId = 1073938446
Attribute TUE_X.VB_VarUserMemId = 1073938446
Attribute TUE_A0.VB_VarUserMemId = 1073938446
Dim WEN_FPM As Double: Dim WEN_A1 As Double: Dim WEN_A2 As Double: Dim WEN_B As Double: Dim WEN_C As Double: Dim WEN_X As Double: Dim WEN_A0 As Double
Attribute WEN_FPM.VB_VarUserMemId = 1073938453
Attribute WEN_A1.VB_VarUserMemId = 1073938453
Attribute WEN_A2.VB_VarUserMemId = 1073938453
Attribute WEN_B.VB_VarUserMemId = 1073938453
Attribute WEN_C.VB_VarUserMemId = 1073938453
Attribute WEN_X.VB_VarUserMemId = 1073938453
Attribute WEN_A0.VB_VarUserMemId = 1073938453
Dim THU_FPM As Double: Dim THU_A1 As Double: Dim THU_A2 As Double: Dim THU_B As Double: Dim THU_C As Double: Dim THU_X As Double: Dim THU_A0 As Double
Attribute THU_FPM.VB_VarUserMemId = 1073938460
Attribute THU_A1.VB_VarUserMemId = 1073938460
Attribute THU_A2.VB_VarUserMemId = 1073938460
Attribute THU_B.VB_VarUserMemId = 1073938460
Attribute THU_C.VB_VarUserMemId = 1073938460
Attribute THU_X.VB_VarUserMemId = 1073938460
Attribute THU_A0.VB_VarUserMemId = 1073938460
Dim FRI_FPM As Double: Dim FRI_A1 As Double: Dim FRI_A2 As Double: Dim FRI_B As Double: Dim FRI_C As Double: Dim FRI_X As Double: Dim FRI_A0 As Double
Attribute FRI_FPM.VB_VarUserMemId = 1073938467
Attribute FRI_A1.VB_VarUserMemId = 1073938467
Attribute FRI_A2.VB_VarUserMemId = 1073938467
Attribute FRI_B.VB_VarUserMemId = 1073938467
Attribute FRI_C.VB_VarUserMemId = 1073938467
Attribute FRI_X.VB_VarUserMemId = 1073938467
Attribute FRI_A0.VB_VarUserMemId = 1073938467
Dim SAT_FPM As Double: Dim SAT_A1 As Double: Dim SAT_A2 As Double: Dim SAT_B As Double: Dim SAT_C As Double: Dim SAT_X As Double: Dim SAT_A0 As Double
Attribute SAT_FPM.VB_VarUserMemId = 1073938474
Attribute SAT_A1.VB_VarUserMemId = 1073938474
Attribute SAT_A2.VB_VarUserMemId = 1073938474
Attribute SAT_B.VB_VarUserMemId = 1073938474
Attribute SAT_C.VB_VarUserMemId = 1073938474
Attribute SAT_X.VB_VarUserMemId = 1073938474
Attribute SAT_A0.VB_VarUserMemId = 1073938474
Dim SUN_FPM As Double: Dim SUN_A1 As Double: Dim SUN_A2 As Double: Dim SUN_B As Double: Dim SUN_C As Double: Dim SUN_X As Double: Dim SUN_A0 As Double
Attribute SUN_FPM.VB_VarUserMemId = 1073938481
Attribute SUN_A1.VB_VarUserMemId = 1073938481
Attribute SUN_A2.VB_VarUserMemId = 1073938481
Attribute SUN_B.VB_VarUserMemId = 1073938481
Attribute SUN_C.VB_VarUserMemId = 1073938481
Attribute SUN_X.VB_VarUserMemId = 1073938481
Attribute SUN_A0.VB_VarUserMemId = 1073938481

Dim MON_D As Double: Dim MON_E As Double: Dim MON_F As Double: Dim MON_G As Double: Dim MON_Y As Double: Dim MON_DPI As Double
Attribute MON_D.VB_VarUserMemId = 1073938488
Attribute MON_E.VB_VarUserMemId = 1073938488
Attribute MON_F.VB_VarUserMemId = 1073938488
Attribute MON_G.VB_VarUserMemId = 1073938488
Attribute MON_Y.VB_VarUserMemId = 1073938488
Attribute MON_DPI.VB_VarUserMemId = 1073938488
Dim TUE_D As Double: Dim TUE_E As Double: Dim TUE_F As Double: Dim TUE_G As Double: Dim TUE_Y As Double: Dim TUE_DPI As Double
Attribute TUE_D.VB_VarUserMemId = 1073938494
Attribute TUE_E.VB_VarUserMemId = 1073938494
Attribute TUE_F.VB_VarUserMemId = 1073938494
Attribute TUE_G.VB_VarUserMemId = 1073938494
Attribute TUE_Y.VB_VarUserMemId = 1073938494
Attribute TUE_DPI.VB_VarUserMemId = 1073938494
Dim WEN_D As Double: Dim WEN_E As Double: Dim WEN_F As Double: Dim WEN_G As Double: Dim WEN_Y As Double: Dim WEN_DPI As Double
Attribute WEN_D.VB_VarUserMemId = 1073938500
Attribute WEN_E.VB_VarUserMemId = 1073938500
Attribute WEN_F.VB_VarUserMemId = 1073938500
Attribute WEN_G.VB_VarUserMemId = 1073938500
Attribute WEN_Y.VB_VarUserMemId = 1073938500
Attribute WEN_DPI.VB_VarUserMemId = 1073938500
Dim THU_D As Double: Dim THU_E As Double: Dim THU_F As Double: Dim THU_G As Double: Dim THU_Y As Double: Dim THU_DPI As Double
Attribute THU_D.VB_VarUserMemId = 1073938506
Attribute THU_E.VB_VarUserMemId = 1073938506
Attribute THU_F.VB_VarUserMemId = 1073938506
Attribute THU_G.VB_VarUserMemId = 1073938506
Attribute THU_Y.VB_VarUserMemId = 1073938506
Attribute THU_DPI.VB_VarUserMemId = 1073938506
Dim FRI_D As Double: Dim FRI_E As Double: Dim FRI_F As Double: Dim FRI_G As Double: Dim FRI_Y As Double: Dim FRI_DPI As Double
Attribute FRI_D.VB_VarUserMemId = 1073938512
Attribute FRI_E.VB_VarUserMemId = 1073938512
Attribute FRI_F.VB_VarUserMemId = 1073938512
Attribute FRI_G.VB_VarUserMemId = 1073938512
Attribute FRI_Y.VB_VarUserMemId = 1073938512
Attribute FRI_DPI.VB_VarUserMemId = 1073938512
Dim SAT_D As Double: Dim SAT_E As Double: Dim SAT_F As Double: Dim SAT_G As Double: Dim SAT_Y As Double: Dim SAT_DPI As Double
Attribute SAT_D.VB_VarUserMemId = 1073938518
Attribute SAT_E.VB_VarUserMemId = 1073938518
Attribute SAT_F.VB_VarUserMemId = 1073938518
Attribute SAT_G.VB_VarUserMemId = 1073938518
Attribute SAT_Y.VB_VarUserMemId = 1073938518
Attribute SAT_DPI.VB_VarUserMemId = 1073938518
Dim SUN_D As Double: Dim SUN_E As Double: Dim SUN_F As Double: Dim SUN_G As Double: Dim SUN_Y As Double: Dim SUN_DPI As Double
Attribute SUN_D.VB_VarUserMemId = 1073938524
Attribute SUN_E.VB_VarUserMemId = 1073938524
Attribute SUN_F.VB_VarUserMemId = 1073938524
Attribute SUN_G.VB_VarUserMemId = 1073938524
Attribute SUN_Y.VB_VarUserMemId = 1073938524
Attribute SUN_DPI.VB_VarUserMemId = 1073938524

Dim MON_IN As Double: Dim TUE_IN As Double: Dim WEN_IN As Double: Dim THU_IN As Double: Dim FRI_IN As Double: Dim SAT_IN As Double: Dim SUN_IN As Double
Attribute MON_IN.VB_VarUserMemId = 1073938530
Attribute TUE_IN.VB_VarUserMemId = 1073938530
Attribute WEN_IN.VB_VarUserMemId = 1073938530
Attribute THU_IN.VB_VarUserMemId = 1073938530
Attribute FRI_IN.VB_VarUserMemId = 1073938530
Attribute SAT_IN.VB_VarUserMemId = 1073938530
Attribute SUN_IN.VB_VarUserMemId = 1073938530
Dim MON_IN_O As Double: Dim TUE_IN_O As Double: Dim WEN_IN_O As Double: Dim THU_IN_O As Double: Dim FRI_IN_O As Double: Dim SAT_IN_O As Double: Dim SUN_IN_O As Double
Attribute MON_IN_O.VB_VarUserMemId = 1073938537
Attribute TUE_IN_O.VB_VarUserMemId = 1073938537
Attribute WEN_IN_O.VB_VarUserMemId = 1073938537
Attribute THU_IN_O.VB_VarUserMemId = 1073938537
Attribute FRI_IN_O.VB_VarUserMemId = 1073938537
Attribute SAT_IN_O.VB_VarUserMemId = 1073938537
Attribute SUN_IN_O.VB_VarUserMemId = 1073938537
Dim INS_LABOR_TMP As Currency: Dim INS_PART_TMP As Currency: Dim INS_MAT_TMP As Currency
Attribute INS_LABOR_TMP.VB_VarUserMemId = 1073938544
Attribute INS_PART_TMP.VB_VarUserMemId = 1073938544
Attribute INS_MAT_TMP.VB_VarUserMemId = 1073938544
Dim MPR_AMOUNT                                         As Currency
Attribute MPR_AMOUNT.VB_VarUserMemId = 1073938547

Dim Month_Value                                        As Integer
Attribute Month_Value.VB_VarUserMemId = 1073938548
Dim Year_Value                                         As Integer
Attribute Year_Value.VB_VarUserMemId = 1073938549
Dim Lastday_Of_The_Month                               As Integer
Attribute Lastday_Of_The_Month.VB_VarUserMemId = 1073938550
Dim Day_Value                                          As Integer
Attribute Day_Value.VB_VarUserMemId = 1073938551
Dim DateRange                                          As String
Attribute DateRange.VB_VarUserMemId = 1073938552
Dim Monday_Of_The_Week                                 As Date
Attribute Monday_Of_The_Week.VB_VarUserMemId = 1073938553
Dim Tuesday_Of_The_Week                                As Date
Attribute Tuesday_Of_The_Week.VB_VarUserMemId = 1073938554
Dim Wednesday_Of_The_Week                              As Date
Attribute Wednesday_Of_The_Week.VB_VarUserMemId = 1073938555
Dim Thursday_Of_The_Week                               As Date
Attribute Thursday_Of_The_Week.VB_VarUserMemId = 1073938556
Dim Friday_Of_The_Week                                 As Date
Attribute Friday_Of_The_Week.VB_VarUserMemId = 1073938557
Dim Saturday_Of_The_Week                               As Date
Attribute Saturday_Of_The_Week.VB_VarUserMemId = 1073938558
Dim Sunday_Of_The_Week                                 As Date
Attribute Sunday_Of_The_Week.VB_VarUserMemId = 1073938559

Dim COUNTER                                            As Integer
Attribute COUNTER.VB_VarUserMemId = 1073938560
Dim SQL                                                As String
Attribute SQL.VB_VarUserMemId = 1073938561
Dim temprs                                             As ADODB.Recordset
Attribute temprs.VB_VarUserMemId = 1073938562

Dim rsREPOR                                            As New ADODB.Recordset
Attribute rsREPOR.VB_VarUserMemId = 1073938563
Dim rsDet                                              As New ADODB.Recordset
Attribute rsDet.VB_VarUserMemId = 1073938564

Dim xlApp                                              As Excel.Application
Attribute xlApp.VB_VarUserMemId = 1073938565
Dim xlBook                                             As Excel.Workbook
Attribute xlBook.VB_VarUserMemId = 1073938566
Dim xlSheet                                            As Excel.Worksheet
Attribute xlSheet.VB_VarUserMemId = 1073938567

Function CHECKIFHYUNDAI(PLATE_NO As String) As Boolean
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT MAKE FROM CSMS_CUSVEH WHERE PLATE_NO = '" & PLATE_NO & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        If UCase(Null2String(rstmp!Make)) = "HYUNDAI" Then
            CHECKIFHYUNDAI = True
        ElseIf Null2String(rstmp!Make) = "" Then
            CHECKIFHYUNDAI = False
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


    '    X_SAT = Day(Combo2.Text) & "-" & Left(cboMonth, 3)
    '    X_FRI = Day(DateAdd("D", -1, Combo2.Text)) & "-" & Left(MonthName(Month(DateAdd("D", -1, Combo2.Text))), 3)
    '    X_THU = Day(DateAdd("D", -2, Combo2.Text)) & "-" & Left(MonthName(Month(DateAdd("D", -2, Combo2.Text))), 3)
    '    X_WEN = Day(DateAdd("D", -3, Combo2.Text)) & "-" & Left(MonthName(Month(DateAdd("D", -3, Combo2.Text))), 3)
    '    X_TUE = Day(DateAdd("D", -4, Combo2.Text)) & "-" & Left(MonthName(Month(DateAdd("D", -4, Combo2.Text))), 3)
    '    X_MON = Day(DateAdd("D", -5, Combo2.Text)) & "-" & Left(MonthName(Month(DateAdd("D", -5, Combo2.Text))), 3)
    '    X_SUN = Day(DateAdd("D", -6, Combo2.Text)) & "-" & Left(MonthName(Month(DateAdd("D", -6, Combo2.Text))), 3)
End Sub

Sub IfLivilIsThree()
    If rsDet!LIVIL = "3" Then
        If Null2String(rsDet!JOBTYPE) = "GJ" Or Null2String(rsDet!JOBTYPE) = "CND" Or Null2String(rsDet!JOBTYPE) = "" Then
            If Null2String(rsDet!wCode) = "" Then
                If INS_MAT_TMP > 0 Then
                    If INS_MAT_TMP >= MPR_AMOUNT Then
                        INS_MAT_TMP = INS_MAT_TMP - MPR_AMOUNT

                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_A0 = MON_A0 + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_A0 = TUE_A0 + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_A0 = WEN_A0 + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_A0 = THU_A0 + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_A0 = FRI_A0 + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_A0 = SAT_A0 + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_A0 = SUN_A0 + MPR_AMOUNT
                    Else
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_A0 = MON_A0 + INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_A0 = TUE_A0 + INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_A0 = WEN_A0 + INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_A0 = THU_A0 + INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_A0 = FRI_A0 + INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_A0 = SAT_A0 + INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_A0 = SUN_A0 + INS_MAT_TMP

                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_A2 = (MON_A2 + MPR_AMOUNT) - INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_A2 = (TUE_A2 + MPR_AMOUNT) - INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_A2 = (WEN_A2 + MPR_AMOUNT) - INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_A2 = (THU_A2 + MPR_AMOUNT) - INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_A2 = (FRI_A2 + MPR_AMOUNT) - INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_A2 = (SAT_A2 + MPR_AMOUNT) - INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_A2 = (SUN_A2 + MPR_AMOUNT) - INS_MAT_TMP

                        INS_MAT_TMP = 0
                    End If
                End If
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_A2 = MON_A2 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_A2 = TUE_A2 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_A2 = WEN_A2 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_A2 = THU_A2 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_A2 = FRI_A2 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_A2 = SAT_A2 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_A2 = SUN_A2 + MPR_AMOUNT
            ElseIf Null2String(rsDet!wCode) = "W" Then
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_B = MON_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_B = TUE_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_B = WEN_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_B = THU_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_B = FRI_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_B = SAT_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_B = SUN_B + MPR_AMOUNT
            ElseIf Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_C = MON_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_C = TUE_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_C = WEN_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_C = THU_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_C = FRI_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_C = SAT_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_C = SUN_C + MPR_AMOUNT
            End If
        ElseIf Null2String(rsDet!JOBTYPE) = "BP" Then
            If Null2String(rsDet!wCode) = "" Then
                If INS_MAT_TMP > 0 Then
                    If INS_MAT_TMP >= (rsDet!DET_AMT - NumericVal(rsDet!Discount_2)) Then
                        INS_MAT_TMP = INS_MAT_TMP - (rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_D = MON_D + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_D = TUE_D + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_D = WEN_D + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_D = THU_D + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_D = FRI_D + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_D = SAT_D + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_D = SUN_D + MPR_AMOUNT
                    Else
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_D = MON_D + INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_D = TUE_D + INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_D = WEN_D + INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_D = THU_D + INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_D = FRI_D + INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_D = SAT_D + INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_D = SUN_D + INS_MAT_TMP

                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_E = (MON_E + MPR_AMOUNT) - INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_E = (TUE_E + MPR_AMOUNT) - INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_E = (WEN_E + MPR_AMOUNT) - INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_E = (THU_E + MPR_AMOUNT) - INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_E = (FRI_E + MPR_AMOUNT) - INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_E = (SAT_E + MPR_AMOUNT) - INS_MAT_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_E = (SUN_E + MPR_AMOUNT) - INS_MAT_TMP
                        INS_MAT_TMP = 0
                    End If
                Else
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_E = MON_E + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_E = TUE_E + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_E = WEN_E + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_E = THU_E + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_E = FRI_E + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_E = SAT_E + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_E = SUN_E + MPR_AMOUNT
                End If
            ElseIf Null2String(rsDet!wCode) = "W" Then
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_F = MON_F + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_F = TUE_F + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_F = WEN_F + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_F = THU_F + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_F = FRI_F + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_F = SAT_F + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_F = SUN_F + MPR_AMOUNT
            ElseIf Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_G = MON_G + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_G = TUE_G + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_G = WEN_G + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_G = THU_G + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_G = FRI_G + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_G = SAT_G + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_G = SUN_G + MPR_AMOUNT
            End If
        End If
    End If
End Sub

Sub IfLivilIsTwoOrFour()
    If Null2String(rsDet!JOBTYPE) = "BP" Then Debug.Print Null2String(rsREPOR!REP_OR)

    If rsDet!LIVIL = "2" Or rsDet!LIVIL = "4" Then
        If Null2String(rsDet!JOBTYPE) = "GJ" Or Null2String(rsDet!JOBTYPE) = "CND" Or Null2String(rsDet!JOBTYPE) = "" Then
            If Null2String(rsDet!wCode) = "" Then
                If INS_PART_TMP > 0 Then
                    If INS_PART_TMP >= MPR_AMOUNT Then
                        INS_PART_TMP = INS_PART_TMP - MPR_AMOUNT

                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_A0 = MON_A0 + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_A0 = TUE_A0 + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_A0 = WEN_A0 + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_A0 = THU_A0 + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_A0 = FRI_A0 + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_A0 = SAT_A0 + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_A0 = SUN_A0 + MPR_AMOUNT
                    Else
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_A0 = MON_A0 + INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_A0 = TUE_A0 + INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_A0 = WEN_A0 + INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_A0 = THU_A0 + INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_A0 = FRI_A0 + INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_A0 = SAT_A0 + INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_A0 = SUN_A0 + INS_PART_TMP

                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_A2 = (MON_A2 + MPR_AMOUNT) - INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_A2 = (TUE_A2 + MPR_AMOUNT) - INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_A2 = (WEN_A2 + MPR_AMOUNT) - INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_A2 = (THU_A2 + MPR_AMOUNT) - INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_A2 = (FRI_A2 + MPR_AMOUNT) - INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_A2 = (SAT_A2 + MPR_AMOUNT) - INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_A2 = (SUN_A2 + MPR_AMOUNT) - INS_PART_TMP

                        INS_PART_TMP = 0
                    End If
                Else
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_A2 = MON_A2 + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_A2 = TUE_A2 + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_A2 = WEN_A2 + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_A2 = THU_A2 + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_A2 = FRI_A2 + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_A2 = SAT_A2 + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_A2 = SUN_A2 + MPR_AMOUNT
                End If
            ElseIf Null2String(rsDet!wCode) = "W" Then
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_B = MON_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_B = TUE_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_B = WEN_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_B = THU_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_B = FRI_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_B = SAT_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_B = SUN_B + MPR_AMOUNT
            ElseIf Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_C = MON_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_C = TUE_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_C = WEN_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_C = THU_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_C = FRI_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_C = SAT_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_C = SUN_C + MPR_AMOUNT
            End If
        ElseIf Null2String(rsDet!JOBTYPE) = "BP" Then
            If Null2String(rsDet!wCode) = "" Then
                If INS_PART_TMP > 0 Then
                    If INS_PART_TMP >= (rsDet!DET_AMT - NumericVal(rsDet!Discount_2)) Then
                        INS_PART_TMP = INS_PART_TMP - (rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_D = MON_D + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_D = TUE_D + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_D = WEN_D + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_D = THU_D + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_D = FRI_D + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_D = SAT_D + MPR_AMOUNT
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_D = SUN_D + MPR_AMOUNT
                    Else
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_D = MON_D + INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_D = TUE_D + INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_D = WEN_D + INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_D = THU_D + INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_D = FRI_D + INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_D = SAT_D + INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_D = SUN_D + INS_PART_TMP

                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_E = (MON_E + MPR_AMOUNT) - INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_E = (TUE_E + MPR_AMOUNT) - INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_E = (WEN_E + MPR_AMOUNT) - INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_E = (THU_E + MPR_AMOUNT) - INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_E = (FRI_E + MPR_AMOUNT) - INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_E = (SAT_E + MPR_AMOUNT) - INS_PART_TMP
                        If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_E = (SUN_E + MPR_AMOUNT) - INS_PART_TMP
                        INS_PART_TMP = 0
                    End If
                Else
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_E = MON_E + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_E = TUE_E + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_E = WEN_E + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_E = THU_E + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_E = FRI_E + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_E = SAT_E + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_E = SUN_E + MPR_AMOUNT
                End If
            ElseIf Null2String(rsDet!wCode) = "W" Then
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_F = MON_F + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_F = TUE_F + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_F = WEN_F + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_F = THU_F + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_F = FRI_F + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_F = SAT_F + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_F = SUN_F + MPR_AMOUNT
            ElseIf Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_G = MON_G + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_G = TUE_G + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_G = WEN_G + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_G = THU_G + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_G = FRI_G + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_G = SAT_G + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_G = SUN_G + MPR_AMOUNT
            End If
        End If
    End If
End Sub

Sub IfLivilIsONE()
    If Null2String(rsDet!JOBTYPE) = "GJ" Or Null2String(rsDet!JOBTYPE) = "CND" Or Null2String(rsDet!JOBTYPE) = "" Then
        If Null2String(rsDet!wCode) = "" Then
            If INS_LABOR_TMP > 0 Then
                If INS_LABOR_TMP >= MPR_AMOUNT Then
                    INS_LABOR_TMP = INS_LABOR_TMP - (MPR_AMOUNT)

                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_A0 = MON_A0 + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_A0 = TUE_A0 + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_A0 = WEN_A0 + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_A0 = THU_A0 + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_A0 = FRI_A0 + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_A0 = SAT_A0 + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_A0 = SUN_A0 + MPR_AMOUNT
                Else
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_A0 = MON_A0 + INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_A0 = TUE_A0 + INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_A0 = WEN_A0 + INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_A0 = THU_A0 + INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_A0 = FRI_A0 + INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_A0 = SAT_A0 + INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_A0 = SUN_A0 + INS_LABOR_TMP

                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_A2 = (MON_A2 + MPR_AMOUNT) - INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_A2 = (TUE_A2 + MPR_AMOUNT) - INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_A2 = (WEN_A2 + MPR_AMOUNT) - INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_A2 = (THU_A2 + MPR_AMOUNT) - INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_A2 = (FRI_A2 + MPR_AMOUNT) - INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_A2 = (SAT_A2 + MPR_AMOUNT) - INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_A2 = (SUN_A2 + MPR_AMOUNT) - INS_LABOR_TMP

                    INS_LABOR_TMP = 0
                End If
            Else
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_A2 = MON_A2 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_A2 = TUE_A2 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_A2 = WEN_A2 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_A2 = THU_A2 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_A2 = FRI_A2 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_A2 = SAT_A2 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_A2 = SUN_A2 + MPR_AMOUNT
            End If
        ElseIf Null2String(rsDet!wCode) = "W" Then
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_B = MON_B + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_B = TUE_B + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_B = WEN_B + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_B = THU_B + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_B = FRI_B + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_B = SAT_B + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_B = SUN_B + MPR_AMOUNT
        ElseIf Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_C = MON_C + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_C = TUE_C + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_C = WEN_C + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_C = THU_C + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_C = FRI_C + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_C = SAT_C + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_C = SUN_C + MPR_AMOUNT
        End If
    ElseIf Null2String(rsDet!JOBTYPE) = "PMS" Then
        If Null2String(rsDet!STATUS1) = "Y" And Null2String(rsDet!wCode) = "W" Then
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_FPM = MON_FPM + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_FPM = TUE_FPM + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_FPM = WEN_FPM + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_FPM = THU_FPM + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_FPM = FRI_FPM + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_FPM = SAT_FPM + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_FPM = SUN_FPM + MPR_AMOUNT
        Else
            If Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_C = MON_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_C = TUE_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_C = WEN_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_C = THU_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_C = FRI_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_C = SAT_C + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_C = SUN_C + MPR_AMOUNT
            ElseIf Null2String(rsDet!wCode) = "W" Then
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_B = MON_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_B = TUE_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_B = WEN_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_B = THU_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_B = FRI_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_B = SAT_B + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_B = SUN_B + MPR_AMOUNT
            Else
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_A1 = MON_A1 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_A1 = TUE_A1 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_A1 = WEN_A1 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_A1 = THU_A1 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_A1 = FRI_A1 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_A1 = SAT_A1 + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_A1 = SUN_A1 + MPR_AMOUNT
            End If
        End If
    ElseIf Null2String(rsDet!JOBTYPE) = "BP" Then
        If Null2String(rsDet!wCode) = "" Then
            If INS_LABOR_TMP > 0 Then
                If INS_LABOR_TMP >= MPR_AMOUNT Then
                    INS_LABOR_TMP = INS_LABOR_TMP - MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_D = MON_D + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_D = TUE_D + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_D = WEN_D + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_D = THU_D + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_D = FRI_D + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_D = SAT_D + MPR_AMOUNT
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_D = SUN_D + MPR_AMOUNT
                Else
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_D = MON_D + INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_D = TUE_D + INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_D = WEN_D + INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_D = THU_D + INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_D = FRI_D + INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_D = SAT_D + INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_D = SUN_D + INS_LABOR_TMP

                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_E = (MON_E + MPR_AMOUNT) - INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_E = (TUE_E + MPR_AMOUNT) - INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_E = (WEN_E + MPR_AMOUNT) - INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_E = (THU_E + MPR_AMOUNT) - INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_E = (FRI_E + MPR_AMOUNT) - INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_E = (SAT_E + MPR_AMOUNT) - INS_LABOR_TMP
                    If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_E = (SUN_E + MPR_AMOUNT) - INS_LABOR_TMP
                    INS_LABOR_TMP = 0
                End If
            Else
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_E = MON_E + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_E = TUE_E + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_E = WEN_E + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_E = THU_E + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_E = FRI_E + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_E = SAT_E + MPR_AMOUNT
                If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_E = SUN_E + MPR_AMOUNT
            End If
        ElseIf Null2String(rsDet!wCode) = "W" Then
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_F = MON_F + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_F = TUE_F + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_F = WEN_F + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_F = THU_F + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_F = FRI_F + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_F = SAT_F + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_F = SUN_F + MPR_AMOUNT
        ElseIf Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_G = MON_G + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_G = TUE_G + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_G = WEN_G + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_G = THU_G + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_G = FRI_G + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_G = SAT_G + MPR_AMOUNT
            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_G = SUN_G + MPR_AMOUNT
        End If
    End If
End Sub

Sub FillcboWeek()
    cboWeek.AddItem "1st Week"
    cboWeek.AddItem "2nd Week"
    cboWeek.AddItem "3rd Week"
    cboWeek.AddItem "4th Week"
    cboWeek.AddItem "5th Week"
End Sub

Private Sub cboMonth_click()
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
    'AXP 01042010 UPDATE FOR ON CHANGE OF YEAR
    
    If cboYear.ListIndex <> -1 And cboMonth.ListIndex <> -1 Then
        cboMonth_click
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Sub PrintInExcel()
    Screen.MousePointer = 11
'    Load frmSplash
'    frmSplash.labCon.Caption = "Calculating Workshop Weekly Performance Details"
'    frmSplash.Show
    DoEvents

    Call GETDATERANGE
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "WEEKLY WORKSHOP.xlt")
    Set xlSheet = xlBook.Worksheets(1)

    Call GETDATERANGE

    xlSheet.Cells(7, "M") = Combo1 & " - " & Combo2
    xlSheet.Cells(7, "C") = COMPANY_NAME              'DEALER NAME

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
    Next Day_Value

    MON_FPM = 0: MON_A1 = 0: MON_A2 = 0: MON_B = 0: MON_C = 0: MON_X = 0: MON_A0 = 0
    TUE_FPM = 0: TUE_A1 = 0: TUE_A2 = 0: TUE_B = 0: TUE_C = 0: TUE_X = 0: TUE_A0 = 0
    WEN_FPM = 0: WEN_A1 = 0: WEN_A2 = 0: WEN_B = 0: WEN_C = 0: WEN_X = 0: WEN_A0 = 0
    THU_FPM = 0: THU_A1 = 0: THU_A2 = 0: THU_B = 0: THU_C = 0: THU_X = 0: THU_A0 = 0
    FRI_FPM = 0: FRI_A1 = 0: FRI_A2 = 0: FRI_B = 0: FRI_C = 0: FRI_X = 0: FRI_A0 = 0
    SAT_FPM = 0: SAT_A1 = 0: SAT_A2 = 0: SAT_B = 0: SAT_C = 0: SAT_X = 0: SAT_A0 = 0
    SUN_FPM = 0: SUN_A1 = 0: SUN_A2 = 0: SUN_B = 0: SUN_C = 0: SUN_X = 0: SUN_A0 = 0


    MON_D = 0: MON_E = 0: MON_F = 0: MON_G = 0: MON_Y = 0: MON_DPI = 0
    TUE_D = 0: TUE_E = 0: TUE_F = 0: TUE_G = 0: TUE_Y = 0: TUE_DPI = 0
    WEN_D = 0: WEN_E = 0: WEN_F = 0: WEN_G = 0: WEN_Y = 0: WEN_DPI = 0
    THU_D = 0: THU_E = 0: THU_F = 0: THU_G = 0: THU_Y = 0: THU_DPI = 0
    FRI_D = 0: FRI_E = 0: FRI_F = 0: FRI_G = 0: FRI_Y = 0: FRI_DPI = 0
    SAT_D = 0: SAT_E = 0: SAT_F = 0: SAT_G = 0: SAT_Y = 0: SAT_DPI = 0
    SUN_D = 0: SUN_E = 0: SUN_F = 0: SUN_G = 0: SUN_Y = 0: SUN_DPI = 0

    Set rsREPOR = gconDMIS.Execute("SELECT DTE_COMP,PARTLABOR,PARTPARTS,PARTACCESSORIES,PARTMATERIALS,REP_OR,CSMS_REPOR.PLATE_NO,MAKE FROM CSMS_REPOR INNER JOIN CSMS_CUSVEH ON CSMS_REPOR.PLATE_NO = CSMS_CUSVEH.PLATE_NO WHERE DTE_COMP BETWEEN '" & Combo1.Text & "' AND '" & Combo2.Text & "' AND TRANSTYPE = 'R' ORDER BY MAKE,REP_OR")
    'Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_COMP IS NOT NULL AND DTE_COMP BETWEEN '" & Combo1.Text & "' AND '" & Combo2.Text & "' AND TRANSTYPE = 'R' ORDER BY DTE_COMP ")
    'Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_COMP IS NOT NULL AND DTE_COMP BETWEEN '" & "6/22/2008" & "' AND '" & "6/28/2008" & "' AND TRANSTYPE = 'R' ORDER BY DTE_COMP ")
    'Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_COMP IS NOT NULL AND DTE_COMP = '" & "9/5/2008" & "' AND TRANSTYPE = 'R' ORDER BY DTE_COMP ")
    'Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_COMP IS NOT NULL AND REP_OR = 'R-00001915' AND TRANSTYPE = 'R' ORDER BY DTE_COMP ")
    If Not (rsREPOR.BOF And rsREPOR.EOF) Then
        Do While Not rsREPOR.EOF
            If Not MonthName(Month(rsREPOR!dte_comp)) = cboMonth.Text Then GoTo CONT_NEXT
            INS_LABOR_TMP = NumericVal(rsREPOR!PARTLABOR)
            INS_PART_TMP = NumericVal(rsREPOR!PARTPARTS) + NumericVal(rsREPOR!PARTACCESSORIES)
            INS_MAT_TMP = NumericVal(rsREPOR!PARTMATERIALS)

            Set rsDet = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & rsREPOR!REP_OR & "' ORDER BY LIVIL,LINE_NO ASC")
            If Not (rsDet.BOF And rsDet.EOF) Then
                If CHECKIFHYUNDAI(Null2String(rsREPOR!PLATE_NO)) = True Then
                    Do While Not rsDet.EOF
                        MPR_AMOUNT = (NumericVal(rsDet!DET_AMT) - NumericVal(rsDet!Discount_2))
                        If rsDet!LIVIL = "1" Then Call IfLivilIsONE
                        If rsDet!LIVIL = "2" Or rsDet!LIVIL = "4" Then Call IfLivilIsTwoOrFour
                        If rsDet!LIVIL = "3" Then Call IfLivilIsThree

                        rsDet.MoveNext
                    Loop
                Else
                    Do While Not rsDet.EOF
                        MPR_AMOUNT = (NumericVal(rsDet!DET_AMT) - NumericVal(rsDet!Discount_2))
                        If Null2String(rsDet!JOBTYPE) = "BP" Then
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_Y = MON_Y + NumericVal(rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_Y = TUE_Y + NumericVal(rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_Y = WEN_Y + NumericVal(rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_Y = THU_Y + NumericVal(rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_Y = FRI_Y + NumericVal(rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_Y = SAT_Y + NumericVal(rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_Y = SUN_Y + NumericVal(rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                        Else
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_X = MON_X + NumericVal(rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_X = TUE_X + NumericVal(rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_X = WEN_X + NumericVal(rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_X = THU_X + NumericVal(rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_X = FRI_X + NumericVal(rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_X = SAT_X + NumericVal(rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_X = SUN_X + NumericVal(rsDet!DET_AMT - NumericVal(rsDet!Discount_2))
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

    xlSheet.Cells(13, "B") = X_MON: xlSheet.Cells(14, "B") = X_TUE: xlSheet.Cells(15, "B") = X_WEN: xlSheet.Cells(16, "B") = X_THU: xlSheet.Cells(17, "B") = X_FRI: xlSheet.Cells(18, "B") = X_SAT: xlSheet.Cells(12, "B") = X_SUN
    xlSheet.Cells(26, "B") = X_MON: xlSheet.Cells(27, "B") = X_TUE: xlSheet.Cells(28, "B") = X_WEN: xlSheet.Cells(29, "B") = X_THU: xlSheet.Cells(30, "B") = X_FRI: xlSheet.Cells(31, "B") = X_SAT: xlSheet.Cells(25, "B") = X_SUN
    xlSheet.Cells(13, "C") = MON_A0: xlSheet.Cells(14, "C") = TUE_A0: xlSheet.Cells(15, "C") = WEN_A0: xlSheet.Cells(16, "C") = THU_A0: xlSheet.Cells(17, "C") = FRI_A0: xlSheet.Cells(18, "C") = SAT_A0: xlSheet.Cells(12, "C") = SUN_A0
    xlSheet.Cells(13, "D") = MON_FPM: xlSheet.Cells(14, "D") = TUE_FPM: xlSheet.Cells(15, "D") = WEN_FPM: xlSheet.Cells(16, "D") = THU_FPM: xlSheet.Cells(17, "D") = FRI_FPM: xlSheet.Cells(18, "D") = SAT_FPM: xlSheet.Cells(12, "D") = SUN_FPM
    xlSheet.Cells(13, "E") = MON_A1: xlSheet.Cells(14, "E") = TUE_A1: xlSheet.Cells(15, "E") = WEN_A1: xlSheet.Cells(16, "E") = THU_A1: xlSheet.Cells(17, "E") = FRI_A1: xlSheet.Cells(18, "E") = SAT_A1: xlSheet.Cells(12, "D") = SUN_A1
    xlSheet.Cells(13, "F") = MON_A2: xlSheet.Cells(14, "F") = TUE_A2: xlSheet.Cells(15, "F") = WEN_A2: xlSheet.Cells(16, "F") = THU_A2: xlSheet.Cells(17, "F") = FRI_A2: xlSheet.Cells(18, "F") = SAT_A2: xlSheet.Cells(12, "F") = SUN_A2
    xlSheet.Cells(13, "G") = MON_B: xlSheet.Cells(14, "G") = TUE_B: xlSheet.Cells(15, "G") = WEN_B: xlSheet.Cells(16, "G") = THU_B: xlSheet.Cells(17, "G") = FRI_B: xlSheet.Cells(18, "G") = SAT_B: xlSheet.Cells(12, "G") = SUN_B
    xlSheet.Cells(13, "H") = MON_C: xlSheet.Cells(14, "H") = TUE_C: xlSheet.Cells(15, "H") = WEN_C: xlSheet.Cells(16, "H") = THU_C: xlSheet.Cells(17, "H") = FRI_C: xlSheet.Cells(18, "H") = SAT_C: xlSheet.Cells(12, "H") = SUN_C
    xlSheet.Cells(13, "I") = MON_X: xlSheet.Cells(14, "I") = TUE_X: xlSheet.Cells(15, "I") = WEN_X: xlSheet.Cells(16, "I") = THU_X: xlSheet.Cells(17, "I") = FRI_X: xlSheet.Cells(18, "I") = SAT_X: xlSheet.Cells(12, "I") = SUN_X
    xlSheet.Cells(13, "J") = MON_D: xlSheet.Cells(14, "J") = TUE_D: xlSheet.Cells(15, "J") = WEN_D: xlSheet.Cells(16, "J") = THU_D: xlSheet.Cells(17, "J") = FRI_D: xlSheet.Cells(18, "J") = SAT_D: xlSheet.Cells(12, "J") = SUN_D
    xlSheet.Cells(13, "K") = MON_E: xlSheet.Cells(14, "K") = TUE_E: xlSheet.Cells(15, "K") = WEN_E: xlSheet.Cells(16, "K") = THU_E: xlSheet.Cells(17, "K") = FRI_E: xlSheet.Cells(18, "K") = SAT_E: xlSheet.Cells(12, "K") = SUN_E
    xlSheet.Cells(13, "L") = MON_F: xlSheet.Cells(14, "L") = TUE_F: xlSheet.Cells(15, "L") = WEN_F: xlSheet.Cells(16, "L") = THU_F: xlSheet.Cells(17, "L") = FRI_F: xlSheet.Cells(18, "L") = SAT_F: xlSheet.Cells(12, "L") = SUN_F
    xlSheet.Cells(13, "M") = MON_G: xlSheet.Cells(14, "M") = TUE_G: xlSheet.Cells(15, "M") = WEN_G: xlSheet.Cells(16, "M") = THU_G: xlSheet.Cells(17, "M") = FRI_G: xlSheet.Cells(18, "M") = SAT_G: xlSheet.Cells(12, "M") = SUN_G
    xlSheet.Cells(13, "N") = MON_Y: xlSheet.Cells(14, "N") = TUE_Y: xlSheet.Cells(15, "N") = WEN_Y: xlSheet.Cells(16, "N") = THU_Y: xlSheet.Cells(17, "N") = FRI_Y: xlSheet.Cells(18, "N") = SAT_Y: xlSheet.Cells(12, "N") = SUN_Y


    xlApp.Windows.ITEM(1).Caption = "WEEKLY WORKSHOP REPORT FOR THE MONTH OF " & cboMonth & " " & cboYear & " " & cboWeek
    xlApp.Visible = True
    Set xlApp = Nothing

    Screen.MousePointer = 0
    Unload frmSplash
End Sub

Private Sub cmdPrint_Click()
    'On Error Resume Next
    Dim rstmp                                                               As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT REP_OR FROM CSMS_REPOR WHERE DTE_COMP IS NOT NULL AND DTE_COMP BETWEEN '" & Combo1.Text & "' AND '" & Combo2.Text & "' AND TRANSTYPE = 'R' ORDER BY DTE_COMP, REP_OR ASC")
    If (rstmp.BOF And rstmp.EOF) Then
        ShowNoRecord
        Exit Sub
    End If
    Set rstmp = Nothing

    'UPDATE BY   : MJP 09172008
    'DESCRIPTION : WORKSHOP WEEKLY PRINT IN EXCEL
    
        Call PrintInExcel
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("G", "WORKSHOP SALES WEEKLY PERFORMANCE REPORT", "", "", "", cboMonth & " " & cboWeek & " " & cboYear, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        Exit Sub
    'UPDATE BY  : MJP 09172008

    Screen.MousePointer = 11
    Load frmSplash
    frmSplash.labCon.Caption = "Calculating Workshop Weekly Performance Details"
    frmSplash.Show
    DoEvents

    Call GETDATERANGE

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
    Next Day_Value
    RPT.Reset
    RPT.Formulas(0) = "WeekNumber ='" & cboWeek.Text & "'": RPT.Formulas(1) = "MonthReport ='" & cboMonth.Text & "'": RPT.Formulas(2) = "YearReport ='" & cboYear.Text & "'"

    MON_FPM = 0: MON_A1 = 0: MON_A2 = 0: MON_B = 0: MON_C = 0: MON_X = 0
    TUE_FPM = 0: TUE_A1 = 0: TUE_A2 = 0: TUE_B = 0: TUE_C = 0: TUE_X = 0
    WEN_FPM = 0: WEN_A1 = 0: WEN_A2 = 0: WEN_B = 0: WEN_C = 0: WEN_X = 0
    THU_FPM = 0: THU_A1 = 0: THU_A2 = 0: THU_B = 0: THU_C = 0: THU_X = 0
    FRI_FPM = 0: FRI_A1 = 0: FRI_A2 = 0: FRI_B = 0: FRI_C = 0: FRI_X = 0
    SAT_FPM = 0: SAT_A1 = 0: SAT_A2 = 0: SAT_B = 0: SAT_C = 0: SAT_X = 0
    SUN_FPM = 0: SUN_A1 = 0: SUN_A2 = 0: SUN_B = 0: SUN_C = 0: SUN_X = 0


    MON_D = 0: MON_E = 0: MON_F = 0: MON_G = 0: MON_Y = 0: MON_DPI = 0
    TUE_D = 0: TUE_E = 0: TUE_F = 0: TUE_G = 0: TUE_Y = 0: TUE_DPI = 0
    WEN_D = 0: WEN_E = 0: WEN_F = 0: WEN_G = 0: WEN_Y = 0: WEN_DPI = 0
    THU_D = 0: THU_E = 0: THU_F = 0: THU_G = 0: THU_Y = 0: THU_DPI = 0
    FRI_D = 0: FRI_E = 0: FRI_F = 0: FRI_G = 0: FRI_Y = 0: FRI_DPI = 0
    SAT_D = 0: SAT_E = 0: SAT_F = 0: SAT_G = 0: SAT_Y = 0: SAT_DPI = 0
    SUN_D = 0: SUN_E = 0: SUN_F = 0: SUN_G = 0: SUN_Y = 0: SUN_DPI = 0

    Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_COMP IS NOT NULL AND DTE_COMP BETWEEN '" & Combo1.Text & "' AND '" & Combo2.Text & "' AND TRANSTYPE = 'R' ORDER BY DTE_COMP ")
    'Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_COMP IS NOT NULL AND DTE_COMP BETWEEN '" & "6/22/2008" & "' AND '" & "6/28/2008" & "' AND TRANSTYPE = 'R' ORDER BY DTE_COMP ")
    'Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_COMP IS NOT NULL AND DTE_COMP = '" & "6/11/2008" & "' AND TRANSTYPE = 'R' ORDER BY DTE_COMP ")
    'Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_COMP IS NOT NULL AND REP_OR = 'R-00000512' AND TRANSTYPE = 'R' ORDER BY DTE_COMP ")
    If Not (rsREPOR.BOF And rsREPOR.EOF) Then
        Do While Not rsREPOR.EOF
            If Not MonthName(Month(rsREPOR!dte_comp)) = cboMonth.Text Then GoTo CONT_NEXT
            INS_LABOR_TMP = NumericVal(rsREPOR!PARTLABOR)
            INS_PART_TMP = NumericVal(rsREPOR!PARTPARTS) + NumericVal(rsREPOR!PARTACCESSORIES)
            INS_MAT_TMP = NumericVal(rsREPOR!PARTMATERIALS)

            Set rsDet = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & rsREPOR!REP_OR & "' ORDER BY LIVIL,LINE_NO")
            If Not (rsDet.BOF And rsDet.EOF) Then
                If CHECKIFHYUNDAI(Null2String(rsREPOR!PLATE_NO)) = True Then
                    Do While Not rsDet.EOF
                        MPR_AMOUNT = (NumericVal(rsDet!DET_AMT) - NumericVal(rsDet!Discount_2))
                        If rsDet!LIVIL = "1" Then Call IfLivilIsONE
                        If rsDet!LIVIL = "2" Or rsDet!LIVIL = "4" Then Call IfLivilIsTwoOrFour
                        If rsDet!LIVIL = "3" Then Call IfLivilIsThree

                        rsDet.MoveNext
                    Loop
                Else
                    Do While Not rsDet.EOF
                        MPR_AMOUNT = (NumericVal(rsDet!DET_AMT) - NumericVal(rsDet!Discount_2))
                        If Null2String(rsDet!JOBTYPE) = "BP" Then
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_Y = MON_Y + MPR_AMOUNT
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_Y = TUE_Y + MPR_AMOUNT
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_Y = WEN_Y + MPR_AMOUNT
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_Y = THU_Y + MPR_AMOUNT
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_Y = FRI_Y + MPR_AMOUNT
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_Y = SAT_Y + MPR_AMOUNT
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_Y = SUN_Y + MPR_AMOUNT
                        Else
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "MONDAY" Then MON_X = MON_X + MPR_AMOUNT
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "TUESDAY" Then TUE_X = TUE_X + MPR_AMOUNT
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "WEDNESDAY" Then WEN_X = WEN_X + MPR_AMOUNT
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "THURSDAY" Then THU_X = THU_X + MPR_AMOUNT
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "FRIDAY" Then FRI_X = FRI_X + MPR_AMOUNT
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SATURDAY" Then SAT_X = SAT_X + MPR_AMOUNT
                            If UCase(WeekdayName(Weekday(rsREPOR!dte_comp))) = "SUNDAY" Then SUN_X = SUN_X + MPR_AMOUNT
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

    RPT.Formulas(3) = "GeneralJob_FPM_Mon = " & MON_FPM
    RPT.Formulas(4) = "GeneralJob_FPM_TUE = " & TUE_FPM
    RPT.Formulas(5) = "GeneralJob_FPM_WED = " & WEN_FPM
    RPT.Formulas(6) = "GeneralJob_FPM_THU = " & THU_FPM
    RPT.Formulas(7) = "GeneralJob_FPM_FRI = " & FRI_FPM
    RPT.Formulas(8) = "GeneralJob_FPM_SAT = " & SAT_FPM
    RPT.Formulas(9) = "GeneralJob_FPM_SUN = " & SUN_FPM

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

    RPT.Formulas(73) = "BodyAndPaint_Y_Mon = " & MON_Y: RPT.Formulas(74) = "BodyAndPaint_Y_TUE = " & TUE_Y: RPT.Formulas(75) = "BodyAndPaint_Y_WED = " & WEN_Y: RPT.Formulas(76) = "BodyAndPaint_Y_THU = " & THU_Y: RPT.Formulas(77) = "BodyAndPaint_Y_FRI = " & FRI_Y: RPT.Formulas(78) = "BodyAndPaint_Y_SAT = " & SAT_Y: RPT.Formulas(79) = "BodyAndPaint_Y_SUN = " & SUN_Y
    RPT.Formulas(108) = "WorkshopSales_Date_Mon ='" & X_MON & "'": RPT.Formulas(109) = "WorkshopSales_Date_Tue ='" & X_TUE & "'": RPT.Formulas(110) = "WorkshopSales_Date_Wed ='" & X_WEN & "'": RPT.Formulas(111) = "WorkshopSales_Date_Thu ='" & X_THU & "'": RPT.Formulas(112) = "WorkshopSales_Date_Fri ='" & X_FRI & "'": RPT.Formulas(113) = "WorkshopSales_Date_Sat ='" & X_SAT & "'": RPT.Formulas(114) = "WorkshopSales_Date_Sun ='" & X_SUN & "'"
    RPT.Formulas(115) = "WSServiced_Sales_Date_Mon ='" & X_MON & "'": RPT.Formulas(116) = "WSServiced_Sales_Date_Tue ='" & X_TUE & "'": RPT.Formulas(117) = "WSServiced_Sales_Date_Wed ='" & X_WEN & "'": RPT.Formulas(118) = "WSServiced_Sales_Date_Thu ='" & X_THU & "'": RPT.Formulas(119) = "WSServiced_Sales_Date_Fri ='" & X_FRI & "'": RPT.Formulas(120) = "WSServiced_Sales_Date_Sat ='" & X_SAT & "'": RPT.Formulas(121) = "WSServiced_Sales_Date_Sun ='" & X_SUN & "'"

    'JUN 02/05/2005
    RPT.WindowTitle = "Workshop Sales Weekly Performance Report": RPT.Formulas(122) = "COMPANY NAME = '" & COMPANY_NAME & "'": RPT.Formulas(123) = "COMPANY ADDRESS = '" & COMPANY_ADDRESS & "'": RPT.Formulas(124) = "PeriodsCovered ='" & Combo1.Text & "-" & Combo2.Text & "'"
    PrintSQLReport RPT, CSMS_REPORT_PATH & "WorkshopSalesWeeklyPerformanceReport.rpt", "", CSMS_REPORT_CONNECTION, 1

    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("G", "WORKSHOP SALES WEEKLY PERFORMANCE REPORT", "", "", "", cboMonth & " " & cboWeek & " " & cboYear, "", "")
    'NEW LOG AUDIT-----------------------------------------------------


    Unload frmSplash
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (WORKSHOP SALES WEEKLY PERFORMANCE REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "WORKSHOP SALES WEEKLY PERFORMANCE REPORT", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Screen.MousePointer = 0
    fillcbomonth cboMonth
    cboMonth.Refresh
    
    FillCboMoreYear cboYear

    'FillcboWeek
    'cboWeek.Text = "1st Week"
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
End Sub

