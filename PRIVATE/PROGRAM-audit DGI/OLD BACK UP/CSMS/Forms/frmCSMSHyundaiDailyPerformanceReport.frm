VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMSHyundaiMonthlyPerformanceReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MPR Report"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   3570
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMSHyundaiDailyPerformanceReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1815
   ScaleWidth      =   3570
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1470
      TabIndex        =   16
      Top             =   2910
      Width           =   1815
   End
   Begin VB.OptionButton Option3 
      Caption         =   "DATE"
      Height          =   375
      Left            =   180
      TabIndex        =   15
      Top             =   2910
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1470
      TabIndex        =   14
      Top             =   2520
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      Caption         =   "RO NO"
      Height          =   375
      Left            =   180
      TabIndex        =   13
      Top             =   2490
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   90
      ScaleHeight     =   735
      ScaleWidth      =   2835
      TabIndex        =   6
      Top             =   3450
      Width           =   2865
      Begin VB.OptionButton optExcel 
         Caption         =   "By Excel Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   8
         Top             =   90
         Value           =   -1  'True
         Width           =   2265
      End
      Begin VB.OptionButton Option1 
         Caption         =   "By Crystal Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   450
         Visible         =   0   'False
         Width           =   2325
      End
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
      ItemData        =   "frmCSMSHyundaiDailyPerformanceReport.frx":1082
      Left            =   915
      List            =   "frmCSMSHyundaiDailyPerformanceReport.frx":1084
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select month from the list"
      Top             =   105
      Width           =   2595
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
      TabIndex        =   1
      ToolTipText     =   "Select year from the list"
      Top             =   495
      Width           =   2595
   End
   Begin Crystal.CrystalReport rptHyundaiMonthlyPerformanceReport 
      Left            =   405
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Hyundai Dealer Monthly Performance Report"
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
      Left            =   2760
      MouseIcon       =   "frmCSMSHyundaiDailyPerformanceReport.frx":1086
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSHyundaiDailyPerformanceReport.frx":11D8
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   930
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   795
      Left            =   2040
      MouseIcon       =   "frmCSMSHyundaiDailyPerformanceReport.frx":1623
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSHyundaiDailyPerformanceReport.frx":1775
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Report"
      Top             =   930
      Width           =   735
   End
   Begin VB.PictureBox picPROG 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   60
      ScaleHeight     =   1095
      ScaleWidth      =   3405
      TabIndex        =   9
      Top             =   450
      Visible         =   0   'False
      Width           =   3435
      Begin MSComctlLib.ProgressBar prb 
         Height          =   315
         Left            =   30
         TabIndex        =   10
         Top             =   330
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblCAP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "STATUS"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   690
         Width           =   3255
      End
      Begin XtremeShortcutBar.ShortcutCaption SRT 
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   3405
         _Version        =   655364
         _ExtentX        =   6006
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "COMPUTING MPR DETAILS..."
         ForeColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   4210752
         ForeColor       =   -2147483633
      End
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
      Left            =   75
      TabIndex        =   5
      Top             =   135
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
      Left            =   75
      TabIndex        =   4
      Top             =   525
      Width           =   735
   End
End
Attribute VB_Name = "frmCSMSHyundaiMonthlyPerformanceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlApp                                               As Excel.Application
Dim xlBook                                              As Excel.Workbook
Dim xlSheet                                             As Excel.Worksheet
Dim WithEvents xlBook2                                  As Excel.Workbook
Attribute xlBook2.VB_VarHelpID = -1
Dim EMP_COUNT                                           As Double
Dim TOTAL_WORKDAY                                       As Double
Dim TOTAL_PERS                                          As Double
Dim DIRE_PERS                                           As Double
Dim GEN_JOB_TECH                                        As Double
Dim BODY_AND_PAINT                                      As Double
Dim SER_ADV                                             As Double
Attribute SER_ADV.VB_VarUserMemId = 1073938441
Dim WARRANTY                                            As Double
Dim INHOUSE_INST                                        As Double
Dim BILL_STAFF                                          As Double
Dim MANTH                                               As Integer
Dim YEER                                                As Integer
'UPDATED BY: JUN ---------------------------------------------------------------------------------------
'DATE UPDATED: 12-04-2008
'DESCRIPTION: AVAILABLE HOURS FOR GJ, BP AND QS
Dim GJ_Available_hrs                                    As Double
Dim BP_Available_hrs                                    As Double
Dim QS_Available_hrs                                    As Double
Dim vSC_QUICKSERVICE                                    As Integer
Attribute vSC_QUICKSERVICE.VB_VarUserMemId = 1073938440
Dim vSC_GENERALJOB                                      As Integer
Dim vSC_BODYANDPAINT                                    As Integer
Dim TOTAL_CAPACITY                                      As Integer
'UPDATED BY: JUN ---------------------------------------------------------------------------------------

'SERVICE PERSONNEL VARIABLES****************************************************************************
Dim vGJMaster                                           As Integer
Dim vGJExpert                                           As Integer
Dim vGJCertified                                        As Integer
Dim vGJNew                                              As Integer
Dim IHTechPaint                                         As Integer
Dim IHTechTinsmith                                      As Integer
Dim CONTech                                             As Integer
Dim SAMas                                               As Integer
Dim SACert                                              As Integer
Dim SANew                                               As Integer
Dim ForeMan                                             As Integer
Dim Warr                                                As Integer
Dim INIns                                               As Integer
Dim BilStf                                              As Integer
Dim Other                                               As Integer
'SERVICE PERSONNEL VARIABLES****************************************************************************

Dim REC                                                 As Integer
Dim REL                                                 As Integer
Dim FPM                                                 As Integer
Dim GJ_CUSTOMERPAID_10K                                 As Integer
Dim GJOTHERHYUNDAI                                      As Integer
Dim GJWARRANTYUNIT                                      As Integer
Dim GJINTERNALUNIT                                      As Integer
Dim GJOTHERBRAND                                        As Integer
Dim BPINSURANCE                                         As Integer
Dim BPCUSTOMERPAID                                      As Integer
Dim BPWARRANTYUNIT                                      As Integer
Dim BPINTERNALUNIT                                      As Integer
Dim BPOTHERBRAND                                        As Integer
Dim PDI                                                 As Integer

Dim UNIT_RECIEVED_OB_GJ                                 As Integer
Dim BP_SERVICED_INSURANCE                               As Integer
Dim BP_SERVICED_CUSTOMER                                As Integer
Dim BP_SERVICED_WARRANTY                                As Integer
Dim BP_SERVICED_INTERNAL                                As Integer
Dim UNIT_RECIEVED_OB_BP                                 As Integer
Dim TRIGER_INS                                          As String
Dim TRIG_GJ_INS                                         As Integer
Dim TRIG_GJ_FPM                                         As Integer
Dim TRIG_GJ_PMS                                         As Integer
Dim TRIG_GJ_CST                                         As Integer
Dim TRIG_GJ_WAR                                         As Integer
Dim TRIG_GJ_INT                                         As Integer
Dim TRIG_GJ_OTH                                         As Integer
Dim TRIG_BP_INS                                         As Integer
Dim TRIG_BP_CST                                         As Integer
Dim TRIG_BP_WRT                                         As Integer
Dim TRIG_BP_INT                                         As Integer
Dim TRIG_BP_OTH                                         As Integer
Dim VPDI_CNT                                            As Integer

Dim JOB_COUNT                                           As Integer
Dim GJ_CNT                                              As Integer
Dim BP_CNT                                              As Integer
Dim GJ_SERVICED_INSURANCE                               As Integer
Dim vFPM                                                As Integer
Dim vPM10K                                              As Integer
Dim vGJ                                                 As Integer
Dim vGJ_WARRANTY                                        As Integer
Dim vGJ_INTERNAL                                        As Integer
Dim GJ_TRG                                              As Integer
Dim BP_TRG                                              As Integer
Dim vGJINS                                              As Integer
Dim vMAJOR_BP                                           As Integer
Dim vMINOR_BP                                           As Integer

Dim vLASTMONTH                                          As Integer
Dim vLAST3MONTH                                         As Integer

Dim PRODUCTIVITY_BP_AVAILABLE_HRS                       As Double
Dim PRODUCTIVITY_BP_ACTUAL_HRS                          As Double
Dim PRODUCTIVITY_BP_SOLD_HRS                            As Double
Dim PRODUCTIVITY_BP_BACKJOB_HRS                         As Double

Dim PRODUCTIVITY_GJ_AVAILABLE_HRS                       As Double
Dim PRODUCTIVITY_GJ_ACTUAL_HRS                          As Double
Dim PRODUCTIVITY_GJ_SOLD_HRS                            As Double
Dim PRODUCTIVITY_GJ_BACKJOB_HRS                         As Double

Dim PRODUCTIVITY_QS_AVAILABLE_HRS                       As Double
Dim PRODUCTIVITY_QS_ACTUAL_HRS                          As Double
Dim PRODUCTIVITY_QS_SOLD_HRS                            As Double
Dim PRODUCTIVITY_QS_BACKJOB_HRS                         As Double

Dim GJ_N_LABOR                                          As Double
Dim GJ_C_LABOR                                          As Double
Dim GJ_W_LABOR                                          As Double
Dim GJ_I_LABOR                                          As Double
Dim BP_I_LABOR                                          As Double
Dim BP_C_LABOR                                          As Double
Dim BP_W_LABOR                                          As Double
Dim BP_N_LABOR                                          As Double
Dim GJ_N_PARTS                                          As Double
Dim GJ_C_PARTS                                          As Double
Dim GJ_W_PARTS                                          As Double
Dim GJ_I_PARTS                                          As Double
Dim BP_I_PARTS                                          As Double
Dim BP_C_PARTS                                          As Double
Dim BP_W_PARTS                                          As Double
Dim BP_N_PARTS                                          As Double
Dim INSLABOR                                            As Double
Dim GJ_N_CHEM                                           As Double
Dim GJ_C_CHEM                                           As Double
Dim GJ_W_CHEM                                           As Double
Dim GJ_I_CHEM                                           As Double
Dim BP_I_CHEM                                           As Double
Dim BP_C_CHEM                                           As Double
Dim BP_W_CHEM                                           As Double
Dim BP_N_CHEM                                           As Double
Dim GJ_N_OTHER                                          As Double
Dim GJ_C_OTHER                                          As Double
Dim GJ_W_OTHER                                          As Double
Dim GJ_I_OTHER                                          As Double
Dim BP_I_OTHER                                          As Double
Dim BP_C_OTHER                                          As Double
Dim BP_W_OTHER                                          As Double
Dim BP_N_OTHER                                          As Double

Dim GJ_OTHER_BRAND                                      As Double
Dim BP_OTHER_BRAND                                      As Double
Dim GJ_OTHER_LABOR                                      As Double
Dim BP_OTHER_LABOR                                      As Double
Dim GJ_OTHER_PART                                       As Double
Dim BP_OTHER_PART                                       As Double
Dim GJ_OTHER_MAT                                        As Double
Dim BP_OTHER_MAT                                        As Double

Dim INS_LABOR                                           As Double
Dim INS_LABOR_TMP                                       As Double
Dim INS_PART                                            As Double
Dim INS_PART_TMP                                        As Double
Dim INS_MAT                                             As Double
Dim INS_MAT_TMP                                         As Double
Dim INS_ACC                                             As Double
Dim INS_ACC_TMP                                         As Double
Dim MPR_AMOUNT                                          As Double

Dim TRIG_GP_INS                                         As Integer

Function ComputeAvalaibleHRs(VTYPE As String) As Double
    Dim rstmp                                          As New ADODB.Recordset
    Dim RSATTEND                                       As New ADODB.Recordset
    Dim TIME_LOG                                       As Double
    Dim AVAIL_TIME                                     As Double

    If VTYPE = "GJ" Then
        Set rstmp = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE (substring(CSMS_POSITION,1,1) = 1 OR substring(CSMS_POSITION,2,1) = 1 OR substring(CSMS_POSITION,3,1) = 1 OR substring(CSMS_POSITION,4,1) = 1)")
    Else
        Set rstmp = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE (substring(CSMS_POSITION,5,1) = 1 OR substring(CSMS_POSITION,6,1) = 1 OR substring(CSMS_POSITION,7,1) = 1)")
    End If

    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            'If rsTmp!empno = "1890" Then
            Set RSATTEND = gconDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE EMPNO = '" & Null2String(rstmp!EMPNO) & "' AND MONTH(DATETODAY) = " & What_month(cboMonth) & " AND YEAR(DATETODAY) = " & cboYear & "")
            If Not (RSATTEND.BOF And RSATTEND.EOF) Then
                Do While Not RSATTEND.EOF
                    If Null2String(RSATTEND!Shift) = "SHIFT1" Then    'IN AM AND OUT AM ONLY
                        If Not Null2String(RSATTEND!INAM) = "" Then
                            If Not Null2String(RSATTEND!OUTAM) = "" Then
                                TIME_LOG = 0
                                TIME_LOG = DateDiff("N", RSATTEND!INAM, RSATTEND!OUTAM)
                                If Hour(RSATTEND!OUTAM) > 12 Then
                                    TIME_LOG = TIME_LOG - 60
                                End If

                                AVAIL_TIME = AVAIL_TIME + TIME_LOG
                            End If
                        End If
                    Else
                        If Not Null2String(RSATTEND!INAM) = "" Then
                            If Not Null2String(RSATTEND!outpm) = "" Then
                                TIME_LOG = 0
                                TIME_LOG = DateDiff("N", RSATTEND!INAM, RSATTEND!outpm)
                                If Hour(RSATTEND!outpm) > 12 Then
                                    TIME_LOG = TIME_LOG - 60
                                End If

                                AVAIL_TIME = AVAIL_TIME + TIME_LOG
                            End If
                        End If
                    End If

                    RSATTEND.MoveNext
                Loop
            End If
            rstmp.MoveNext
        Loop
    End If


    ComputeAvalaibleHRs = AVAIL_TIME / 60
    Set rstmp = Nothing
End Function

Function CHECK_HYUNDAI_DEALER(VREP_OR As String) As Boolean
    Dim rstmp                                          As New ADODB.Recordset
    Dim RSDEL                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("SELECT PLATE_NO FROM CSMS_REPOR WHERE REP_OR = " & N2Str2Null(VREP_OR) & "")
    If Not (rstmp.BOF And rstmp.EOF) Then
        Set RSDEL = gconDMIS.Execute("SELECT MAKE FROM CSMS_CUSVEH WHERE PLATE_NO = '" & rstmp!PLATE_NO & "'  ")
        If Not (RSDEL.BOF And RSDEL.EOF) Then
            If UCase(Null2String(RSDEL!Make)) = "HYUNDAI" Then
                CHECK_HYUNDAI_DEALER = True
            Else
                CHECK_HYUNDAI_DEALER = False
            End If
        End If
    End If
    Set rstmp = Nothing
End Function

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
        CHECKIFHYUNDAI = False
    End If
    Set rstmp = Nothing
End Function

Function COMPUTATION_OF_WORKSHOP_LABOR(vLIVIL As Integer, VWCODE As String, vGJ_SALES_H As Currency, vBP_SALES_H As Currency, MANTH, YEER)
    Dim rstmp                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("select * from CSMS_RO_DET where livil = " & vLIVIL & _
                               " And month(SAVEDATE) = " & MANTH & " and YEAR(SAVEDATE) = " & YEER & _
                               " And transtype = 'R' ORDER BY REP_OR ASC")
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            If Null2String(rstmp!JOBTYPE) = "GJ" Then
                If Left(Null2String(rstmp!DETCDE), 2) = "PM" Then
                    If Left(Null2String(rstmp!DETDSC), 9) = "1,000 KMS" Or Left(Null2String(rstmp!DETDSC), 9) = "5,000 KMS" Then
                    Else
                        If CHECK_HYUNDAI_DEALER(rstmp!REP_OR) = True Then
                            vGJ_SALES_H = vGJ_SALES_H + rstmp!DET_AMT
                        End If
                    End If
                Else
                    If CHECK_HYUNDAI_DEALER(rstmp!REP_OR) = True Then
                        vGJ_SALES_H = vGJ_SALES_H + rstmp!DET_AMT
                    End If
                End If
            ElseIf Null2String(rstmp!JOBTYPE) = "BP" Then
                If CHECK_HYUNDAI_DEALER(rstmp!REP_OR) = True Then
                    vBP_SALES_H = vGJ_SALES_H + rstmp!DET_AMT
                End If
            ElseIf Null2String(rstmp!JOBTYPE) = "CND" Then
                If CHECK_HYUNDAI_DEALER(rstmp!REP_OR) = True Then
                    vGJ_SALES_H = vGJ_SALES_H + rstmp!DET_AMT
                End If
            ElseIf Null2String(rstmp!JOBTYPE) = "PMS" Then
                If Left(Null2String(rstmp!DETDSC), 9) = "1,000 KMS" Or Left(Null2String(rstmp!DETDSC), 9) = "5,000 KMS" Then
                Else
                    If CHECK_HYUNDAI_DEALER(rstmp!REP_OR) = True Then
                        vGJ_SALES_H = vGJ_SALES_H + rstmp!DET_AMT
                    End If
                End If
            Else
                If Left(UCase(Null2String(rstmp!DETCDE)), 2) = "BP" Then
                    vBP_SALES_H = vGJ_SALES_H + rstmp!DET_AMT
                Else
                    If CHECK_HYUNDAI_DEALER(rstmp!REP_OR) = True Then
                        vGJ_SALES_H = vGJ_SALES_H + rstmp!DET_AMT
                    End If
                End If
            End If
            rstmp.MoveNext
        Loop
    End If
End Function

Function COMPUTATION_OF_WARKSHOP_PARTS(vLIVIL As Integer, VWCODE As String, vGJ_SALES_H As Currency, vBP_SALES_H As Currency, MANTH, YEER)
    Dim rstmp                                          As New ADODB.Recordset
    Dim rsDet                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE MONTH(DTE_COMP) = " & _
                                 MANTH & " AND YEAR(DTE_COMP) = " & YEER & " AND TRANSTYPE = 'R' AND INVOICE IS NOT NULL")
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            Set rsDet = gconDMIS.Execute("sELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & rstmp!REP_OR & _
                                         "' AND LIVIL = " & vLIVIL & "")

            rstmp.MoveNext
        Loop
    End If
    Set rstmp = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : ComputeMPR
' Author    : wizweirdo
' Date      : 9/24/2008
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub ComputeMPR()
    Dim rsREPOR                                         As New ADODB.Recordset
    Dim rsDet                                           As New ADODB.Recordset
    Dim rstmp                                           As New ADODB.Recordset


    MANTH = What_month(cboMonth)
    YEER = Val(cboYear)
    Dim RSTMP1                                         As New ADODB.Recordset
    Set RSTMP1 = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE " & _
                    " MONTH(CSMS_REPOR.DTE_COMP) = " & MANTH & " AND " & _
                    " YEAR(CSMS_REPOR.DTE_COMP) = " & YEER & " AND " & _
                    " CSMS_REPOR.TRANSTYPE = 'R'")
    If (RSTMP1.EOF And RSTMP1.BOF) Then
        Call ShowNoRecord
        Screen.MousePointer = 0
        Exit Sub
    End If

    Screen.MousePointer = 11

    'UPDATED BY: JUN------------------
    'DATE UPDATED: 12-04-2008
        Call Kwenta_Available_Hrs
    'UPDATED BY: JUN------------------

    frmMain.Enabled = False
    picPROG.Visible = True
    prb.Value = 0
    DoEvents
    lblCap.Caption = ""
    DoEvents

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "Monthly Performance Report.xlt")
    Set xlSheet = xlBook.Worksheets(1)

    xlSheet.Cells(4, "B") = "For The Month of " & cboMonth & " " & cboYear    'FOR THE MONTH OF MPR
    xlSheet.Cells(5, "B") = COMPANY_NAME              'DEALER NAME

    'WORK DAYS -------------------------------------------------------------------------------
        DoEvents
        SRT.Caption = "Computing Workday..."
        DoEvents
        Call ComputeWorkDays
        xlSheet.Cells(8, "J") = TOTAL_WORKDAY             'TOTAL WORKDAY
    'WORK DAYS -------------------------------------------------------------------------------

    'EMPLOYEE ----------------------------------------------------------------------------------
        DoEvents
        SRT.Caption = "Computing Service Personnel..."
        DoEvents
        Call ComputeServicePersonnel
        
        xlSheet.Cells(13, "J") = vGJMaster                'GENERAL JOB MASTER
        xlSheet.Cells(14, "J") = vGJExpert                'GENERAL JOB EXPERT
        xlSheet.Cells(15, "J") = vGJCertified             'GENERAL JOB CERTIFIED
        xlSheet.Cells(16, "J") = vGJNew                   'GENERAL JOB NEW
        xlSheet.Cells(19, "J") = IHTechPaint              'IN HOUSE TECHNICIAN PAINT
        xlSheet.Cells(20, "J") = IHTechTinsmith           'IN HOUSE TECHNICIAN TINSMITH
        xlSheet.Cells(21, "J") = CONTech                  'CONTRACTUAL TECHNICIAN
        xlSheet.Cells(24, "J") = SAMas                    'SERVICE ADVISER MASTER
        xlSheet.Cells(25, "J") = SACert                   'SERVICE ADVISER CERTIFIED
        xlSheet.Cells(26, "J") = SANew                    'SERVICE ADVISER NEW
        xlSheet.Cells(27, "J") = ForeMan                  'FOREMAN
        xlSheet.Cells(28, "J") = Warr                     'WARRANTY
        xlSheet.Cells(29, "J") = INIns                    'IN HOUSE INSTRUCTOR
        xlSheet.Cells(30, "J") = BilStf                   'BILLING STAFF
        xlSheet.Cells(31, "J") = Other                    'OTHER
    'EMPLOYEE--------------------------------------------------------------------------------

    'UNIT RECEIVED---------------------------------------------------------------------------
        DoEvents
        SRT.Caption = "Computing Unit Received..."
        DoEvents
        Call ComputeUnitReceived
        xlSheet.Cells(41, "J") = REC
    'UNIT RECEIVED---------------------------------------------------------------------------

    'UNIT RELEASED-----------------------------------------------------------------------------
        DoEvents
        SRT.Caption = "Computing Unit Released..."
        DoEvents
        
        Call ComputeUnitReleased
        xlSheet.Cells(42, "J") = REL
    'UNIT RELEASED-----------------------------------------------------------------------------

    'UNIT SERVICE------------------------------------------------------------------------------
        DoEvents
        SRT.Caption = "Computing Unit Serviced..."
        DoEvents

        Call ComputeUnitService
        Call ComputePDIUR
    
        'UPDATED BY: JUN------------------------------------------------
        'DATE UPDATED: 12-13 - 2008
        'DESCRIPTION: DISPLAY SERVICE CAPACITY
        Call Service_Capacity
    
        xlSheet.Cells(35, "J") = TOTAL_CAPACITY
        xlSheet.Cells(36, "J") = vSC_QUICKSERVICE
        xlSheet.Cells(37, "J") = vSC_GENERALJOB
        xlSheet.Cells(38, "J") = vSC_BODYANDPAINT
        'UPDATED BY: JUN------------------------------------------------
    
        xlSheet.Cells(44, "J") = GJ_CNT                   'GJ UNIT RECEIVED SERVICE
        xlSheet.Cells(54, "J") = BP_CNT                   'BP UNIT RECEIVED SERVICE
    
        xlSheet.Cells(46, "J") = GJ_SERVICED_INSURANCE    'GJ UNIT INSURANCE
        xlSheet.Cells(47, "J") = vFPM                     'GJ UNIT FIRST PREVENTIVE MAINTENANCE
        xlSheet.Cells(49, "J") = vPM10K                   'GJ UNIT PREVENTIVE MAINTENANCE
        xlSheet.Cells(50, "J") = vGJ                      'GJ UNIT GENERAL JOB
        xlSheet.Cells(51, "J") = vGJ_WARRANTY             'GJ UNIT WARRANTY
        xlSheet.Cells(52, "J") = vGJ_INTERNAL             'GJ UNIT INTERNAL
        xlSheet.Cells(53, "J") = UNIT_RECIEVED_OB_GJ      'GJ UNIT RECEIVED OTHER BRAND
    
    
        xlSheet.Cells(58, "J") = BP_SERVICED_INSURANCE    'BP UNIT SERVICED INSURANCE
        xlSheet.Cells(59, "J") = BP_SERVICED_CUSTOMER     'BP UNIT SERVICED CUSTOMER
        xlSheet.Cells(60, "J") = BP_SERVICED_WARRANTY     'BP UNIT SERVICED WARRANTY
        xlSheet.Cells(61, "J") = BP_SERVICED_INTERNAL     'BP UNIT SERVICED INTERNAL
        xlSheet.Cells(62, "J") = UNIT_RECIEVED_OB_BP      'BP UNIT RECEIVED OTHER BRAND
        xlSheet.Cells(66, "J") = VPDI_CNT                 'PDI UNIT
    
        xlSheet.Cells(56, "J") = vMAJOR_BP                'BP MAJOR
        xlSheet.Cells(57, "J") = vMINOR_BP                'BP MINOR
    'UNIT SERVICE---------------------------------------------------------------------------------------------

    'VEHICLE SALES RELEASED LAST MONTH------------------------------------------------------------------------
        DoEvents
        SRT.Caption = "Computing Vehicle Released Last Month..."
        DoEvents
                
        Call ComputeVehicleSalesLastMonth
        xlSheet.Cells(67, "J") = NumericVal(vLASTMONTH)   'VEHICLE RELEASED LAST MONTH
    'VEHICLE SALES RELEASED LAST MONTH------------------------------------------------------------------------

    'VEHICLE SALES RELEASED LAST 3 MONTH------------------------------------------------------------------------
        DoEvents
        SRT.Caption = "Computing Vehicle Released Last 3 Month..."
        DoEvents
        
        Call ComputeVehicleSalesLast3Months
        xlSheet.Cells(68, "J") = NumericVal(vLAST3MONTH)  'VEHICLE RELEASED LAST 3 MONTH
    'VEHICLE SALES RELEASED LAST 3 MONTH------------------------------------------------------------------------

    'WORK SHOP---------------------------------------------------------------------------------
        DoEvents
        SRT.Caption = "Computing Workshop Sales..."
        DoEvents
        
        Call ComputeWorkShop
        BP_OTHER_BRAND = BP_OTHER_LABOR + BP_OTHER_PART + BP_OTHER_MAT
        GJ_OTHER_BRAND = GJ_OTHER_LABOR + GJ_OTHER_PART + GJ_OTHER_MAT
    
        xlSheet.Cells(11, "U") = GJ_N_LABOR               'GJ INSURANCE LABOR
        xlSheet.Cells(12, "U") = GJ_N_PARTS               'GJ INSURANCE PART
        xlSheet.Cells(13, "U") = GJ_N_CHEM                'GJ INSURANCE CHEMICAL
        xlSheet.Cells(14, "U") = GJ_N_OTHER               'GJ INSURANCE OTHER
    
    
        xlSheet.Cells(16, "U") = GJ_C_LABOR               'GJ CUSTOMER LABOR
        xlSheet.Cells(17, "U") = GJ_C_PARTS                'GJ CUSTOMER PARTS
        xlSheet.Cells(18, "U") = GJ_C_CHEM                'GJ CUSTOMER CHEMICAL
        xlSheet.Cells(19, "U") = GJ_C_OTHER               'GJ CUSTOMER OTHER
    
        xlSheet.Cells(21, "U") = GJ_W_LABOR               'GJ WARRANTY LABOR
        xlSheet.Cells(22, "U") = GJ_W_PARTS                'GJ WARRANTY PARTS
        xlSheet.Cells(23, "U") = GJ_W_CHEM                'GJ WARRANTY CHEMICALS
        xlSheet.Cells(24, "U") = GJ_W_OTHER               'GJ WARRANTY OTHERS
    
        xlSheet.Cells(26, "U") = GJ_I_LABOR               'GJ OTHER LABOR
        xlSheet.Cells(27, "U") = GJ_I_PARTS                'GJ OTHER PARTS
        xlSheet.Cells(28, "U") = GJ_I_CHEM                'GJ OTHER CHEMICALS
        xlSheet.Cells(29, "U") = GJ_I_OTHER               'GJ OTHER OTHERS
    
        xlSheet.Cells(30, "U") = GJ_OTHER_BRAND           'GJ OTHER BRAND
    
        'BODY AND PAINT
        xlSheet.Cells(34, "U") = BP_I_LABOR               'BP INSURANCE LABOR
        xlSheet.Cells(35, "U") = BP_I_PARTS                'BP INSURANCE PARTS
        xlSheet.Cells(36, "U") = BP_I_OTHER               'BP INSURANCE CHEMICALS
        'xlSheet.Cells(32, "U") = BP_C_OTHER                                                                                    'BP INSURANCE OTHERS
    
        xlSheet.Cells(39, "U") = BP_C_LABOR               'BP CUSTOMER LABOR
        xlSheet.Cells(40, "U") = BP_C_PARTS                'BP CUSTOMER PARTS
        xlSheet.Cells(41, "U") = BP_C_OTHER               'BP CUSTOMER CHEMICALS
        'xlSheet.Cells(37, "U") = WS_BP_CUST_OTHERS                                                                             'BP CUSTOMER OTHER
    
        xlSheet.Cells(44, "U") = BP_W_LABOR               'BP WARRANTY LABOR
        xlSheet.Cells(45, "U") = BP_W_PARTS                'BP WARRANTY PARTS
        xlSheet.Cells(46, "U") = BP_W_OTHER               'BP WARRANTY CHEMICAL
        'xlSheet.Cells(42, "U") = WS_BP_WARR_OTHERS                                                                             'BP WARRANTY OTHERS
    
        xlSheet.Cells(49, "U") = BP_N_LABOR               'BP OTHER LABOR
        xlSheet.Cells(50, "U") = BP_N_PARTS                'BP OTHER PARTS
        xlSheet.Cells(451, "U") = BP_N_OTHER              'BP OTHER CHEMICALS
        'xlSheet.Cells(47, "U") = WS_BP_OTHER_OTHERS                                                                            'BP OTHER OTHERS
    
        xlSheet.Cells(53, "U") = BP_OTHER_BRAND           'BP OTHER BRAND
    'WORK SHOP---------------------------------------------------------------------------------


    'QUICK SERVICE----------------------------------------------------------------------------
        DoEvents
        SRT.Caption = "Computing Quick Service Productivity..."
        DoEvents
        
        Call ComputeQuickService
        'xlSheet.Cells(65, "U") = NumericVal(ComputeAvalaibleHRs("GJ"))              'GJ AVAILABLE HRS
        'xlSheet.Cells(65, "U") = PRODUCTIVITY_GJ_AVAILABLE_HRS
        
        'UPDATED BY: JUN-----------------------------------------------
        'DATE UPDATED: 12-04-2008
        xlSheet.Cells(60, "U") = NumericVal(QS_Available_hrs)
        'UPDATED BY: JUN-----------------------------------------------
        xlSheet.Cells(61, "U") = NumericVal(PRODUCTIVITY_QS_ACTUAL_HRS)    'GJ ACTUAL HRS
        xlSheet.Cells(62, "U") = NumericVal(PRODUCTIVITY_QS_SOLD_HRS)    'GJ SOLD HRS
        xlSheet.Cells(68, "U") = NumericVal(PRODUCTIVITY_QS_BACKJOB_HRS)    'GJ BACK JOB HRS
    'QUICK SERVICE----------------------------------------------------------------------------

    'GJ PRODUCTIVITY-------------------------------------------------------------------------
        DoEvents
        SRT.Caption = "Computing GJ Productivity..."
        DoEvents
                
        Call ComputeGJProductivity
        'ORIGINAL CODE ---- 12-04-2008------------------------
        ' xlSheet.Cells(65, "U") = NumericVal(ComputeAvalaibleHRs("GJ"))              'GJ AVAILABLE HRS
        'ORIGINAL CODE ---- 12-04-2008------------------------
    
        'xlSheet.Cells(65, "U") = PRODUCTIVITY_GJ_AVAILABLE_HRS
    
        'UPDATED BY: JUN-----------------------------------------------
        'DATE UPDATED: 12-04-2008
        xlSheet.Cells(65, "U") = NumericVal(GJ_Available_hrs)
        'UPDATED BY: JUN-----------------------------------------------
    
        xlSheet.Cells(66, "U") = NumericVal(PRODUCTIVITY_GJ_ACTUAL_HRS)    'GJ ACTUAL HRS
        xlSheet.Cells(67, "U") = NumericVal(PRODUCTIVITY_GJ_SOLD_HRS)    'GJ SOLD HRS
        xlSheet.Cells(68, "U") = NumericVal(PRODUCTIVITY_GJ_BACKJOB_HRS)    'GJ BACK JOB HRS
    'GJ PRODUCTIVITY-------------------------------------------------------------------------

    'BP PRODUCTIVITY-------------------------------------------------------------------------
        DoEvents
        SRT.Caption = "Computing BP Productivity..."
        DoEvents
        
        Call ComputeBPProductivity
        'Original Code-------------12-04-2008
        'xlSheet.Cells(70, "U") = NumericVal(ComputeAvalaibleHRs("BP"))              'GJ AVAILABLE HRS
        'Original Code-------------12-04-2008
    
        'xlSheet.Cells(70, "U") = PRODUCTIVITY_BP_AVAILABLE_HRS
    
        'UPDATED BY: JUN-----------------------------------------------
        'DATE UPDATED: 12-04-2008
        xlSheet.Cells(70, "U") = NumericVal(BP_Available_hrs)
        'UPDATED BY: JUN-----------------------------------------------
        xlSheet.Cells(71, "U") = NumericVal(PRODUCTIVITY_BP_ACTUAL_HRS)    'BP ACTUAL HRS
        xlSheet.Cells(72, "U") = NumericVal(PRODUCTIVITY_BP_SOLD_HRS)    'BP SOLD HRS
        xlSheet.Cells(73, "U") = NumericVal(PRODUCTIVITY_BP_BACKJOB_HRS)    'BP BACK JOB HRS
    'BP PRODUCTIVITY-------------------------------------------------------------------------

    Screen.MousePointer = 0
    DoEvents
    prb.Value = 0
    lblCap.Caption = ""
    picPROG.Visible = False
    frmMain.Enabled = True
    SRT.Caption = ""
    DoEvents

    xlApp.Windows.ITEM(1).Caption = "MONTHLY PERFORMANCE REPORT FOR THE MONTH OF " & cboMonth & " " & cboYear
    xlApp.Visible = True
    Set xlApp = Nothing
    Screen.MousePointer = 0

    Set rsREPOR = Nothing: Set rsDet = Nothing
    Screen.MousePointer = 0

    On Error GoTo 0
    Exit Sub

ComputeMPR_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ComputeMPR of Form frmCSMSHyundaiMonthlyPerformanceReport"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    If optExcel.Value Then
        picPROG.Visible = True
        picPROG.ZOrder 0
        
        Call InitializeVariables
        Call ComputeMPR

        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("G", "MONTHLY PERFORMANCE REPORT", "", "", "", cboMonth & " " & cboYear, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
    End If
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (MONTHLY PERFORMANCE REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "MONTHLY PERFORMANCE REPORT", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    
    Call fillcbomonth(cboMonth)
    Call FillCboMoreYear(cboYear)
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub

Sub Kwenta_Available_Hrs()
    'UPDATED BY: JUN
    'DATE UPDATED: 12-04 -2008
    'DESCRIPTION: COMPUTATION OF AVAILABLE HRS FOR GJ, BP AND QUICK SERVICE

    Dim count_GJ_Technician                            As Integer
    Dim count_BP_Technician                            As Integer
    Dim count_QS_Technician                            As Integer

    count_GJ_Technician = 0
    count_BP_Technician = 0
    count_QS_Technician = 0

    Dim RSTECH                                         As New ADODB.Recordset
    Set RSTECH = gconDMIS.Execute("Select CSMS_POSITION from HRMS_Empinfo where IS_TECHNICIAN = 1")
    If Not RSTECH.EOF And Not RSTECH.BOF Then
        Do While Not RSTECH.EOF
            If Mid(Null2String(RSTECH!CSMS_POSITION), 1, 1) = "1" Or Mid(Null2String(RSTECH!CSMS_POSITION), 2, 1) = "1" Or Mid(Null2String(RSTECH!CSMS_POSITION), 3, 1) = "1" Or Mid(Null2String(RSTECH!CSMS_POSITION), 4, 1) = "1" Then
                count_GJ_Technician = count_GJ_Technician + 1
            ElseIf Mid(Null2String(RSTECH!CSMS_POSITION), 5, 1) = "1" Or Mid(Null2String(RSTECH!CSMS_POSITION), 6, 1) = "1" Then
                count_BP_Technician = count_BP_Technician + 1
            ElseIf Mid(Null2String(RSTECH!CSMS_POSITION), 8, 1) = "1" Then
                count_QS_Technician = count_QS_Technician + 1
            End If
            RSTECH.MoveNext
        Loop
    End If
    Set RSTECH = Nothing

    Dim rsCSMS_EMPINFO                                 As New ADODB.Recordset
    Set rsCSMS_EMPINFO = gconDMIS.Execute("Select EMPNO,CSMS_POSITION from CSMS_EMPINFO WHERE IS_TECHNICIAN = 1")
    If Not rsCSMS_EMPINFO.EOF And Not rsCSMS_EMPINFO.BOF Then
        Do While Not rsCSMS_EMPINFO.EOF
            If IsExistInHRMS(Null2String(LTrim(RTrim(rsCSMS_EMPINFO!EMPNO)))) = True Then
                rsCSMS_EMPINFO.MoveNext
            Else
                If Mid(Null2String(rsCSMS_EMPINFO!CSMS_POSITION), 1, 1) = "1" Or Mid(Null2String(rsCSMS_EMPINFO!CSMS_POSITION), 2, 1) = "1" Or Mid(Null2String(rsCSMS_EMPINFO!CSMS_POSITION), 3, 1) = "1" Or Mid(Null2String(rsCSMS_EMPINFO!CSMS_POSITION), 4, 1) = "1" Then
                    count_GJ_Technician = count_GJ_Technician + 1
                ElseIf Mid(Null2String(rsCSMS_EMPINFO!CSMS_POSITION), 5, 1) = "1" Or Mid(Null2String(rsCSMS_EMPINFO!CSMS_POSITION), 6, 1) = "1" Then
                    count_BP_Technician = count_BP_Technician + 1
                ElseIf Mid(Null2String(rsCSMS_EMPINFO!CSMS_POSITION), 8, 1) = "1" Then
                    count_QS_Technician = count_QS_Technician + 1
                End If
                rsCSMS_EMPINFO.MoveNext
            End If
        Loop
    End If
    Set rsCSMS_EMPINFO = Nothing

    GJ_Available_hrs = ((count_GJ_Technician * 7.5) * 26)
    BP_Available_hrs = ((count_BP_Technician * 7.5) * 26)
    QS_Available_hrs = ((count_QS_Technician * 7.5) * 26)
End Sub

Function IsExistInHRMS(xEMPNO As String) As Boolean
    'UPDATED BY: JUN
    'DATE UPDATED: 12-08-2008
    'DESCRIPTION: CHECK IF EMPNO IS EXISTING IN HRMS
    Dim rsHR                                           As ADODB.Recordset
    Set rsHR = gconDMIS.Execute("Select EMPNO from HRMS_EMPINFO where IS_TECHNICIAN = 1 and EMPNO = '" & xEMPNO & "'")
    If Not rsHR.EOF And Not rsHR.BOF Then
        IsExistInHRMS = True
    Else
        IsExistInHRMS = False
    End If
    Set rsHR = Nothing
End Function

Sub Service_Capacity()
    Dim rsCapacity                                     As ADODB.Recordset
    Set rsCapacity = gconDMIS.Execute("SELECT SC_QUICKSERVICE,SC_GENERALJOB, SC_BODYANDPAINT FROM CSMS_SERVICE_PERSONNEL_MAINTENANCE")
    If Not rsCapacity.EOF And Not rsCapacity.EOF Then
        vSC_QUICKSERVICE = N2Str2IntZero(rsCapacity!SC_QUICKSERVICE)
        vSC_GENERALJOB = N2Str2IntZero(rsCapacity!SC_GENERALJOB)
        vSC_BODYANDPAINT = N2Str2IntZero(rsCapacity!SC_BODYANDPAINT)

        TOTAL_CAPACITY = vSC_QUICKSERVICE + vSC_GENERALJOB + vSC_BODYANDPAINT
    End If
    Set rsCapacity = Nothing
End Sub

Sub ComputeUnitReleased()
    Dim rstmp                                           As New ADODB.Recordset
    'Set RSTMP = gconDMIS.Execute("select COUNT(*) AS TOTAL_RELEASED FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND REP_OR = '" & XRONOX & "'")
    Set rstmp = gconDMIS.Execute("select COUNT(*) AS TOTAL_RELEASED FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND MONTH(DTE_REL)= " & MANTH & " AND  YEAR(DTE_REL)= " & YEER)
    If Not (rstmp.EOF And rstmp.BOF) Then
        REL = N2Str2Zero(rstmp!TOTAL_RELEASED)
    End If
End Sub

Sub ComputeUnitReceived()
    Dim rstmp                                           As New ADODB.Recordset
    'Set RSTMP = gconDMIS.Execute("select COUNT(*) AS TOTAL_RECEIVED FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND REP_OR = '" & XRONOX & "'")
    Set rstmp = gconDMIS.Execute("select COUNT(*) AS TOTAL_RECEIVED FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND MONTH(DTE_RECD)= " & MANTH & " AND  YEAR(DTE_RECD)= " & YEER)
    If Not (rstmp.EOF And rstmp.BOF) Then
        REC = N2Str2Zero(rstmp!TOTAL_RECEIVED)
    End If
End Sub

Sub ComputeWorkDays()
    Dim rstmp                                           As New ADODB.Recordset
    Dim vX                                              As Integer
    TOTAL_WORKDAY = 0
    For vX = 1 To Day(lastDay(DateSerial(YEER, MANTH, 1)))
        Set rstmp = gconDMIS.Execute("Select DTE_RECD, Rep_OR From CSMS_REPOR Where TRANSTYPE = 'R' AND DTE_RECD = '" & DateSerial(YEER, MANTH, vX) & "'")
        If Not (rstmp.BOF And rstmp.EOF) Then
            TOTAL_WORKDAY = TOTAL_WORKDAY + 1
        End If
    Next
End Sub

Sub ComputeServicePersonnel()
    Dim rstmp                                       As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("Select * From CSMS_SERVICE_PERSONNEL_MAINTENANCE")
    If Not (rstmp.BOF And rstmp.EOF) Then
        vGJMaster = rstmp!GJ_TECH_MASTER
        vGJExpert = rstmp!GJ_TECH_EXPERT
        vGJCertified = rstmp!GJ_TECH_CERTIFIED
        vGJNew = rstmp!GJ_TECH_NEW
        IHTechPaint = rstmp!BP_TECH_PAINT
        IHTechTinsmith = rstmp!BP_TECH_TINSMITH
        CONTech = rstmp!BP_TECH_CONTR
        SAMas = rstmp!SA_MASTER
        SACert = rstmp!SA_CERTIFIED
        SANew = rstmp!SA_NEW
        ForeMan = rstmp!ForeMan
        Warr = rstmp!WARRANTY
        INIns = rstmp!IH_INSTRUCTOR
        BilStf = rstmp!BILLING_STAFF
        Other = rstmp!OTHERS
    End If
    Set rstmp = Nothing
    '†-----------------------------------------------------------------------------------------
    'NEW SERVICE
    '    Dim rsSERVICE_PERSON                                              As New ADODB.Recordset
    '    Set rsSERVICE_PERSON = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE DEPTCODE = 'SERVICE'")
    '    If Not (rsSERVICE_PERSON.BOF And rsSERVICE_PERSON.EOF) Then
    '        Do While Not rsSERVICE_PERSON.EOF
    '            If Mid(Null2String(rsSERVICE_PERSON!CSMS_POSITION), 1, 1) = 1 Then vGJMaster = vGJMaster + 1
    '            If Mid(Null2String(rsSERVICE_PERSON!CSMS_POSITION), 2, 1) = 1 Then vGJExpert = vGJExpert + 1
    '            If Mid(Null2String(rsSERVICE_PERSON!CSMS_POSITION), 3, 1) = 1 Then vGJCertified = vGJCertified + 1
    '            If Mid(Null2String(rsSERVICE_PERSON!CSMS_POSITION), 4, 1) = 1 Then vGJNew = vGJNew + 1
    '            If Mid(Null2String(rsSERVICE_PERSON!CSMS_POSITION), 5, 1) = 1 Then IHTechPaint = IHTechPaint + 1
    '            If Mid(Null2String(rsSERVICE_PERSON!CSMS_POSITION), 6, 1) = 1 Then IHTechTinsmith = IHTechTinsmith + 1
    '            If Mid(Null2String(rsSERVICE_PERSON!CSMS_POSITION), 7, 1) = 1 Then CONTech = CONTech + 1
    '            If Mid(Null2String(rsSERVICE_PERSON!CSMS_POSITION), 8, 1) = 1 Then SAMas = SAMas + 1
    '            If Mid(Null2String(rsSERVICE_PERSON!CSMS_POSITION), 9, 1) = 1 Then SACert = SACert + 1
    '            If Mid(Null2String(rsSERVICE_PERSON!CSMS_POSITION), 10, 1) = 1 Then SANew = SANew + 1
    '            If Mid(Null2String(rsSERVICE_PERSON!CSMS_POSITION), 11, 1) = 1 Then ForeMan = ForeMan + 1
    '            If Mid(Null2String(rsSERVICE_PERSON!CSMS_POSITION), 12, 1) = 1 Then Warr = Warr + 1
    '            If Mid(Null2String(rsSERVICE_PERSON!CSMS_POSITION), 13, 1) = 1 Then INIns = INIns + 1
    '            If Mid(Null2String(rsSERVICE_PERSON!CSMS_POSITION), 14, 1) = 1 Then BilStf = BilStf + 1
    '            If Mid(Null2String(rsSERVICE_PERSON!CSMS_POSITION), 15, 1) = 1 Then Other = Other + 1
    '
    '            rsSERVICE_PERSON.MoveNext
    '        Loop
    '    End If
    '    Set rsSERVICE_PERSON = Nothing
    '-----------------------------------------------------------------------------------------†
End Sub

Sub ComputeUnitService()
    Dim rsREPOR                     As New ADODB.Recordset
    Dim rsDet                       As New ADODB.Recordset
    Dim XDATEX                      As String
                
'    If Option2.Value = True Then
'        XDATEX = "REP_OR = " & N2Str2Null(Text1) & ""
'    Else
'        XDATEX = "DTE_COMP =  " & N2Str2Null(Text2) & ""
'    End If
'    Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND " & XDATEX & " AND INVOICE <> 'PDI RO' ORDER BY DTE_COMP, REP_OR ASC")
    
    Set rsREPOR = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND MONTH(DTE_COMP)= " & MANTH & " AND  YEAR(DTE_COMP)= " & YEER & " AND INVOICE <> 'PDI RO' ORDER BY DTE_COMP, REP_OR ASC")
    If Not (rsREPOR.BOF And rsREPOR.EOF) Then
        prb.Max = rsREPOR.RecordCount
        prb.Value = 0
        Do While Not rsREPOR.EOF
            DoEvents
            prb.Value = prb.Value + 1
            lblCap.Caption = "RO NO: " & Null2String(rsREPOR!REP_OR)
            DoEvents
    
            TRIG_GJ_FPM = 0: TRIG_GJ_PMS = 0: TRIG_GJ_CST = 0: TRIG_GJ_WAR = 0: TRIG_GJ_INT = 0: TRIG_GJ_OTH = 0
            TRIG_BP_INS = 0: TRIG_BP_CST = 0: TRIG_BP_WRT = 0: TRIG_BP_INT = 0: TRIG_BP_OTH = 0
            INSLABOR = NumericVal(rsREPOR!PARTLABOR)
            TRIGER_INS = ""
            GJ_TRG = 0: BP_TRG = 0
    
            Set rsDet = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE LIVIL = '1' AND REP_OR = '" & rsREPOR!REP_OR & "' AND TRANSTYPE = 'R' ORDER BY LIVIL,LINE_NO")
            If Not (rsDet.BOF And rsDet.EOF) Then
                If CHECKIFHYUNDAI(Null2String(rsREPOR!PLATE_NO)) = True Then
                    Do While Not rsDet.EOF
                        If Null2String(rsDet!JOBTYPE) <> "BP" And Null2String(rsDet!JOBTYPE) <> "PMS" Then
                            If Null2String(rsDet!wCode) = "" Then
                                If INSLABOR > 0 Then
                                    If INSLABOR >= MPR_AMOUNT Then
                                        INSLABOR = INSLABOR - MPR_AMOUNT
                                        If TRIG_GP_INS = 0 Then
                                            GJ_SERVICED_INSURANCE = GJ_SERVICED_INSURANCE + 1    'N2Str2Zero(RSTMP!TOTAL_INSURANCE)
                                            TRIG_GP_INS = 1
                                        End If
                                    Else
                                        If TRIG_GP_INS = 0 Then
                                            GJ_SERVICED_INSURANCE = GJ_SERVICED_INSURANCE + 1    'N2Str2Zero(RSTMP!TOTAL_INSURANCE)
                                            TRIG_GP_INS = 1
                                        End If
                                        INSLABOR = 0
                                    End If
                                Else
                                    If TRIG_GJ_CST = 0 Then
                                        vGJ = vGJ + 1
                                        TRIG_GJ_CST = 1
                                    End If
                                End If
                            ElseIf Null2String(rsDet!wCode) = "W" Then
                                If TRIG_GJ_WAR = 0 Then
                                    vGJ_WARRANTY = vGJ_WARRANTY + 1
                                    TRIG_GJ_WAR = 1
                                End If
                            ElseIf Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                                If TRIG_GJ_INT = 0 Then
                                    vGJ_INTERNAL = vGJ_INTERNAL + 1
                                    TRIG_GJ_INT = 1
                                End If
                            End If
    
                            If GJ_TRG = 0 Then
                                GJ_TRG = 1
                                GJ_CNT = GJ_CNT + 1
                            End If
                        ElseIf Null2String(rsDet!JOBTYPE) = "PMS" Then
                            If Null2String(rsDet!STATUS1) = "" Then    'OLD VERSION TAGGING OF FPM
                                If Left(Null2String(rsDet!DETDSC), 5) = "1,000" Or Left(Null2String(rsDet!DETDSC), 5) = "5,000" Then
                                    If TRIG_GJ_FPM = 0 Then
                                        vFPM = vFPM + 1
                                        TRIG_GJ_FPM = 1
                                    End If
                                Else
                                    If Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                                        If TRIG_GJ_INT = 0 Then
                                            vGJ_INTERNAL = vGJ_INTERNAL + 1
                                            TRIG_GJ_INT = 1
                                        End If
                                    Else
                                        If TRIG_GJ_PMS = 0 Then
                                            vPM10K = vPM10K + 1
                                            TRIG_GJ_PMS = 1
                                        End If
                                    End If
                                End If
                            Else                      'NEW VERSION OF TAGGING OF FPM
                                If Null2String(rsDet!STATUS1) = "Y" Then
                                    If TRIG_GJ_FPM = 0 Then
                                        vFPM = vFPM + 1
                                        TRIG_GJ_FPM = 1
                                    End If
                                Else
                                    If Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                                        If TRIG_GJ_INT = 0 Then
                                            vGJ_INTERNAL = vGJ_INTERNAL + 1
                                            TRIG_GJ_INT = 1
                                        End If
                                    Else
                                        If TRIG_GJ_PMS = 0 Then
                                            vPM10K = vPM10K + 1
                                            TRIG_GJ_PMS = 1
                                        End If
                                    End If
                                End If
                            End If
    
                            If GJ_TRG = 0 Then
                                GJ_TRG = 1
                                GJ_CNT = GJ_CNT + 1
                            End If
                        ElseIf Null2String(rsDet!JOBTYPE) = "BP" Then
                            If Null2String(rsDet!wCode) = "" Then
                                If INSLABOR > 0 Then
                                    If INSLABOR >= MPR_AMOUNT Then
                                        INSLABOR = INSLABOR - MPR_AMOUNT
                                        If TRIG_BP_INS = 0 Then
                                            BP_SERVICED_INSURANCE = BP_SERVICED_INSURANCE + 1    'N2Str2Zero(RSTMP!TOTAL_INSURANCE)
                                            TRIG_BP_INS = 1
                                        End If
                                    Else
                                        INSLABOR = 0
                                        If TRIG_BP_INS = 0 Then
                                            BP_SERVICED_INSURANCE = BP_SERVICED_INSURANCE + 1    'N2Str2Zero(RSTMP!TOTAL_INSURANCE)
                                            TRIG_BP_INS = 1
                                        End If
                                    End If
                                Else
                                    If TRIG_BP_CST = 0 Then
                                        BP_SERVICED_CUSTOMER = BP_SERVICED_CUSTOMER + 1
                                        TRIG_BP_CST = 1
                                    End If
                                End If
                            ElseIf Null2String(rsDet!wCode) = "W" Then
                                If TRIG_BP_WRT = 0 Then
                                    BP_SERVICED_WARRANTY = BP_SERVICED_WARRANTY + 1
                                    TRIG_BP_WRT = 1
                                End If
                            ElseIf Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                                If TRIG_BP_INT = 0 Then
                                    BP_SERVICED_INTERNAL = BP_SERVICED_INTERNAL + 1
                                    TRIG_BP_INT = 1
                                End If
                            End If
    
                            If BP_TRG = 0 Then
                                BP_TRG = 1
                                BP_CNT = BP_CNT + 1
                            End If
                            If Null2String(rsDet!transtatus) = "" Or Null2String(rsDet!transtatus) = "M" Then
                                vMAJOR_BP = vMAJOR_BP + 1
                            Else
                                vMINOR_BP = vMINOR_BP + 1
                            End If
                        End If
    
                        rsDet.MoveNext
                    Loop
                Else
                    Do While Not rsDet.EOF
                        If Null2String(rsDet!JOBTYPE) = "BP" Then
                            If TRIG_BP_OTH = 0 Then
                                UNIT_RECIEVED_OB_BP = UNIT_RECIEVED_OB_BP + 1
                                TRIG_BP_OTH = 1
                            End If
    
                            If BP_TRG = 0 Then
                                BP_TRG = 1
                            End If
    
                            'If Null2String(RSDET!transtatus) = "" Or Null2String(RSDET!transtatus) = "M" Then
                            '    vMAJOR_BP = vMAJOR_BP + 1
                            'Else
                            '    vMINOR_BP = vMINOR_BP + 1
                            'End If
                        Else
                            If TRIG_GJ_OTH = 0 Then
                                UNIT_RECIEVED_OB_GJ = UNIT_RECIEVED_OB_GJ + 1
                                TRIG_GJ_OTH = 1
                            End If
    
                            If GJ_TRG = 0 Then
                                GJ_TRG = 1
                            End If
                        End If
    
                        rsDet.MoveNext
                    Loop
                End If
            End If
            rsREPOR.MoveNext
        Loop
    End If
End Sub

Sub ComputePDIUR()
    Dim RSPDI                                          As New ADODB.Recordset
    Set RSPDI = gconDMIS.Execute("SELECT COUNT(*) AS COUNT_PDI FROM CSMS_REPOR WHERE TRANSTYPE = 'R' AND MONTH(DTE_COMP)= " & MANTH & " AND  YEAR(DTE_COMP)= " & YEER & " AND INVOICE = 'PDI RO'")
    If Not (RSPDI.BOF And RSPDI.EOF) Then
        VPDI_CNT = NumericVal(RSPDI!COUNT_PDI)
    End If
End Sub

Sub ComputeVehicleSalesLastMonth()
    Dim XMONTH                                          As Integer
    Dim XYEAR                                           As Integer
    Dim rstmp                                           As New ADODB.Recordset
    If What_month(cboMonth) = 1 Then
        XMONTH = 12
        XYEAR = cboYear - 1
    Else
        XMONTH = What_month(cboMonth) - 1
        XYEAR = cboYear
    End If
    
    Set rstmp = gconDMIS.Execute("SELECT COUNT(*) AS TOTAL_RELEASED FROM SMIS_PURCHAGREE WHERE isdaTE(DATERELEASED) = 1 AND MONTH(DATERELEASED) = " & XMONTH & " AND YEAR(DATERELEASED) = " & XYEAR & "")
    If Not (rstmp.BOF And rstmp.EOF) Then
        vLASTMONTH = NumericVal(rstmp!TOTAL_RELEASED)
    End If
End Sub

Sub ComputeVehicleSalesLast3Months()
    Dim rstmp                                   As New ADODB.Recordset
    Dim XMONTH                                  As Integer
    Dim XYEAR                                   As Integer
    
    If What_month(cboMonth) < 4 Then
        If What_month(cboMonth) = 3 Then XMONTH = 12
        If What_month(cboMonth) = 2 Then XMONTH = 11
        If What_month(cboMonth) = 1 Then XMONTH = 10
        XYEAR = cboYear - 1
    Else
        XMONTH = What_month(cboMonth) - 3
        XYEAR = cboYear
    End If

    Set rstmp = gconDMIS.Execute("SELECT COUNT(*) AS TOTAL_RELASED FROM SMIS_PURCHAGREE WHERE ISDATE(DATERELEASED) = 1 AND month(DATERELEASED) = " & XMONTH & " and year(DATERELEASED) = " & XYEAR & "")
    If Not (rstmp.BOF And rstmp.EOF) Then
        vLAST3MONTH = NumericVal(rstmp(0))
    End If
End Sub

Sub ComputeBPProductivity()
    Dim rsREPOR                                 As New ADODB.Recordset
    Dim rsDet                                   As New ADODB.Recordset
    
    PRODUCTIVITY_BP_AVAILABLE_HRS = 7.5 * TOTAL_WORKDAY
    PRODUCTIVITY_BP_ACTUAL_HRS = 0
    PRODUCTIVITY_BP_SOLD_HRS = 0
    PRODUCTIVITY_BP_BACKJOB_HRS = 0
    'Set rsREPOR = gconDMIS.Execute("Select * From CSMS_REPOR where TRANSTYPE = 'R' AND " & XRONOX & "")
'    Set rsREPOR = gconDMIS.Execute("Select * From CSMS_REPOR where TRANSTYPE = 'R' AND MONTH(DTE_COMP) = " & MANTH & " AND YEAR(CSMS_REPOR.DTE_COMP) = " & YEER)
'    If Not (rsREPOR.EOF And rsREPOR.BOF) Then
'        rsREPOR.MoveFirst: PRODUCTIVITY_BP_SOLD_HRS = 0: PRODUCTIVITY_BP_ACTUAL_HRS = 0
'        DoEvents
'        lblCAP.Caption = ""
'        prb.Max = rsREPOR.RecordCount
'        prb.Value = 0
'        DoEvents
'        Do While Not rsREPOR.EOF
'            DoEvents
'            prb.Value = prb.Value + 1
'            lblCAP.Caption = "RO NO: " & Null2String(rsREPOR!rep_OR)
'            DoEvents
'
'            Set RSDET = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE TRANSTYPE = 'R' AND REP_OR = " & N2Str2Null(rsREPOR!rep_OR) & " AND LIVIL = '1' AND JOBTYPE = 'BP'")
'            If Not (RSDET.BOF And RSDET.EOF) Then
'                Do While Not RSDET.EOF
'                    PRODUCTIVITY_BP_SOLD_HRS = PRODUCTIVITY_BP_SOLD_HRS + NumericVal(RSDET!DET_HRS)
'                    PRODUCTIVITY_BP_ACTUAL_HRS = PRODUCTIVITY_BP_ACTUAL_HRS + NumericVal(RSDET!HRSWRK)
'                    RSDET.MoveNext
'                Loop
'            End If
'            rsREPOR.MoveNext
'        Loop
'    End If
'    Set rsREPOR = Nothing
    Dim rstmp                                       As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT " & _
        " CAST(SUM(ISNULL(CSMS_Ro_Det.HRSWRK,0)) AS DECIMAL(18,2)) AS HRSWRK, " & _
        " CAST(SUM(ISNULL(CSMS_Ro_Det.DET_HRS,0)) AS DECIMAL(18,2)) AS DET_HRS " & _
        " FROM CSMS_Repor INNER JOIN " & _
        " CSMS_Ro_Det ON CSMS_Repor.TRANSTYPE = CSMS_Ro_Det.TRANSTYPE AND " & _
        " CSMS_Repor.REP_OR = CSMS_Ro_Det.REP_OR " & _
        " Where " & _
        " (MONTH(CSMS_Repor.DTE_COMP) = " & MANTH & ") " & _
        " AND (CSMS_Repor.TRANSTYPE = 'R') " & _
        " AND (YEAR(CSMS_Repor.DTE_COMP) = " & YEER & ") " & _
        " AND CSMS_Ro_Det.LIVIL = '1' " & _
        " AND CSMS_Ro_Det.JOBTYPE = 'BP' " & _
        " AND CSMS_Ro_Det.TRANSTYPE = 'R'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        PRODUCTIVITY_BP_SOLD_HRS = NumericVal(rstmp!DET_HRS)
        PRODUCTIVITY_BP_ACTUAL_HRS = NumericVal(rstmp!HRSWRK)
    End If
    Set rstmp = Nothing
End Sub

Sub ComputeGJProductivity()
    Dim rsREPOR                             As New ADODB.Recordset
    Dim rsDet                               As New ADODB.Recordset
    
    PRODUCTIVITY_GJ_AVAILABLE_HRS = 7.5 * TOTAL_WORKDAY
    PRODUCTIVITY_GJ_ACTUAL_HRS = 0
    PRODUCTIVITY_GJ_SOLD_HRS = 0
    PRODUCTIVITY_GJ_BACKJOB_HRS = 0

    'Set rsREPOR = gconDMIS.Execute("Select * From CSMS_REPOR where TRANSTYPE = 'R' AND " & XDATEX & "")
'    Set rsREPOR = gconDMIS.Execute("Select * From CSMS_REPOR where TRANSTYPE = 'R' AND MONTH(DTE_COMP) = " & MANTH & " AND YEAR(CSMS_REPOR.DTE_COMP) = " & YEER)
'    If Not (rsREPOR.EOF And rsREPOR.BOF) Then
'        rsREPOR.MoveFirst: PRODUCTIVITY_GJ_ACTUAL_HRS = 0: PRODUCTIVITY_GJ_SOLD_HRS = 0
'        DoEvents
'        lblCAP.Caption = ""
'        prb.Max = rsREPOR.RecordCount
'        prb.Value = 0
'        DoEvents
'        Do While Not rsREPOR.EOF
'            DoEvents
'            prb.Value = prb.Value + 1
'            lblCAP.Caption = "RO NO: " & Null2String(rsREPOR!REP_OR)
'            DoEvents
'            Set RSDET = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE TRANSTYPE = 'R' AND REP_OR = " & N2Str2Null(rsREPOR!REP_OR) & " AND LIVIL = '1' AND (JOBTYPE = 'GJ' or JOBTYPE = 'SR' OR JOBTYPE IS NULL OR JOBTYPE = '' OR (JOBTYPE = 'PMS' AND QUICK_SERVICE = 'N'))")
'            If Not (RSDET.BOF And RSDET.EOF) Then
'                Do While Not RSDET.EOF
'                    PRODUCTIVITY_GJ_SOLD_HRS = PRODUCTIVITY_GJ_SOLD_HRS + NumericVal(RSDET!DET_HRS)
'                    PRODUCTIVITY_GJ_ACTUAL_HRS = PRODUCTIVITY_GJ_ACTUAL_HRS + NumericVal(RSDET!HRSWRK)
'                    RSDET.MoveNext
'                Loop
'            End If
'
'            rsREPOR.MoveNext
'        Loop
'    End If
'    Set rsREPOR = Nothing
    Dim rstmp                                   As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT " & _
        " CAST(SUM(ISNULL(CSMS_Ro_Det.HRSWRK,0)) AS DECIMAL(18,2)) AS HRSWRK, " & _
        " CAST(SUM(ISNULL(CSMS_Ro_Det.DET_HRS,0)) AS DECIMAL(18,2)) AS DET_HRS " & _
        " FROM CSMS_Repor INNER JOIN " & _
        " CSMS_Ro_Det ON CSMS_Repor.TRANSTYPE = CSMS_Ro_Det.TRANSTYPE AND " & _
        " CSMS_Repor.REP_OR = CSMS_Ro_Det.REP_OR " & _
        " Where " & _
        " (MONTH(CSMS_Repor.DTE_COMP) = " & MANTH & ") " & _
        " AND (CSMS_Repor.TRANSTYPE = 'R') " & _
        " AND (YEAR(CSMS_Repor.DTE_COMP) = " & YEER & ") " & _
        " AND (CSMS_Ro_Det.LIVIL = '1') " & _
        " AND (CSMS_Ro_Det.JOBTYPE = 'GJ' " & _
        " OR CSMS_RO_DET.JOBTYPE = 'SR' " & _
        " OR CSMS_Ro_Det.JOBTYPE IS NULL " & _
        " OR CSMS_Ro_Det.JOBTYPE = '' " & _
        " OR (CSMS_Ro_Det.JOBTYPE = 'PMS' AND CSMS_Ro_Det.QUICK_SERVICE = 'N')) " & _
        " AND CSMS_Ro_Det.TRANSTYPE = 'R'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        PRODUCTIVITY_GJ_SOLD_HRS = NumericVal(rstmp!DET_HRS)
        PRODUCTIVITY_GJ_ACTUAL_HRS = NumericVal(rstmp!HRSWRK)
    End If
    Set rstmp = Nothing
End Sub

Sub ComputeQuickService()
    Dim rsREPOR                         As New ADODB.Recordset
    Dim rsDet                           As New ADODB.Recordset
    
    PRODUCTIVITY_QS_AVAILABLE_HRS = 7.5 * TOTAL_WORKDAY
    PRODUCTIVITY_QS_ACTUAL_HRS = 0
    PRODUCTIVITY_QS_SOLD_HRS = 0
    PRODUCTIVITY_QS_BACKJOB_HRS = 0
    
    'Set rsREPOR = gconDMIS.Execute("Select * From CSMS_REPOR where TRANSTYPE = 'R' AND " & XDATEX & "")
'    Set rsREPOR = gconDMIS.Execute("Select * From CSMS_REPOR where TRANSTYPE = 'R' AND MONTH(DTE_COMP) = " & MANTH & " AND YEAR(CSMS_REPOR.DTE_COMP) = " & YEER)
'    If Not (rsREPOR.EOF And rsREPOR.BOF) Then
'        rsREPOR.MoveFirst: PRODUCTIVITY_QS_ACTUAL_HRS = 0: PRODUCTIVITY_QS_SOLD_HRS = 0
'        DoEvents
'        lblCAP.Caption = ""
'        prb.Max = rsREPOR.RecordCount
'        prb.Value = 0
'        DoEvents
'        Do While Not rsREPOR.EOF
'            DoEvents
'            prb.Value = prb.Value + 1
'            lblCAP.Caption = "RO NO: " & Null2String(rsREPOR!REP_OR)
'            DoEvents
'            Set RSDET = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE TRANSTYPE = 'R' AND REP_OR = " & N2Str2Null(rsREPOR!REP_OR) & " AND LIVIL = '1' AND (JOBTYPE = 'PMS' AND QUICK_SERVICE = 'Y')")
'            If Not (RSDET.BOF And RSDET.EOF) Then
'                Do While Not RSDET.EOF
'                    PRODUCTIVITY_QS_SOLD_HRS = PRODUCTIVITY_QS_SOLD_HRS + NumericVal(RSDET!DET_HRS)
'                    PRODUCTIVITY_QS_ACTUAL_HRS = PRODUCTIVITY_QS_ACTUAL_HRS + NumericVal(RSDET!HRSWRK)
'                    RSDET.MoveNext
'                Loop
'            End If
'
'            rsREPOR.MoveNext
'        Loop
'    End If
'    Set rsREPOR = Nothing

    Dim rstmp                                       As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT " & _
        " CAST(SUM(ISNULL(CSMS_Ro_Det.HRSWRK,0)) AS DECIMAL(18,2)) AS HRSWRK, " & _
        " CAST(SUM(ISNULL(CSMS_Ro_Det.DET_HRS,0)) AS DECIMAL(18,2)) AS DET_HRS " & _
        " FROM CSMS_Repor INNER JOIN " & _
        " CSMS_Ro_Det ON CSMS_Repor.TRANSTYPE = CSMS_Ro_Det.TRANSTYPE AND " & _
        " CSMS_Repor.REP_OR = CSMS_Ro_Det.REP_OR " & _
        " Where " & _
        " (MONTH(CSMS_Repor.DTE_COMP) = " & MANTH & ") " & _
        " AND (CSMS_Repor.TRANSTYPE = 'R') " & _
        " AND (YEAR(CSMS_Repor.DTE_COMP) = " & YEER & ") " & _
        " AND CSMS_Ro_Det.LIVIL = '1' " & _
        " AND (JOBTYPE = 'PMS' AND QUICK_SERVICE = 'Y') " & _
        " AND CSMS_Ro_Det.TRANSTYPE = 'R'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        PRODUCTIVITY_QS_SOLD_HRS = NumericVal(rstmp!DET_HRS)
        PRODUCTIVITY_QS_ACTUAL_HRS = NumericVal(rstmp!HRSWRK)
    End If
    Set rstmp = Nothing
End Sub

Sub ComputeWorkShop()
    Dim XDATEX                  As String
    Dim rsREPOR                 As New ADODB.Recordset
    Dim rsDet                   As New ADODB.Recordset
    
    
'    If Option2.Value = True Then
'        XDATEX = "REP_OR = " & N2Str2Null(Text1) & ""
'    Else
'        XDATEX = "DTE_COMP =  " & N2Str2Null(Text2) & ""
'    End If
'    Set rsREPOR = gconDMIS.Execute("SELECT PARTLABOR,PARTPARTS,PARTACCESSORIES,PARTMATERIALS,REP_OR,CSMS_REPOR.PLATE_NO,MAKE FROM CSMS_REPOR INNER JOIN CSMS_CUSVEH ON CSMS_REPOR.PLATE_NO = CSMS_CUSVEH.PLATE_NO WHERE " & XDATEX & " AND TRANSTYPE = 'R' ORDER BY MAKE,REP_OR")
    
    Set rsREPOR = gconDMIS.Execute("SELECT PARTLABOR,PARTPARTS,PARTACCESSORIES,PARTMATERIALS,REP_OR,CSMS_REPOR.PLATE_NO,MAKE FROM CSMS_REPOR INNER JOIN CSMS_CUSVEH ON CSMS_REPOR.PLATE_NO = CSMS_CUSVEH.PLATE_NO WHERE MONTH(DTE_COMP) = " & MANTH & " AND YEAR(DTE_COMP) = " & YEER & " AND TRANSTYPE = 'R' ORDER BY MAKE,REP_OR")
    If Not (rsREPOR.BOF And rsREPOR.EOF) Then
        prb.Max = rsREPOR.RecordCount
        prb.Value = 0
        Do While Not rsREPOR.EOF
            DoEvents
            prb.Value = prb.Value + 1
            lblCap.Caption = "RO NO : " & Null2String(rsREPOR!REP_OR)
            DoEvents

            INS_LABOR_TMP = NumericVal(rsREPOR!PARTLABOR)
            INS_PART_TMP = NumericVal(rsREPOR!PARTPARTS) + NumericVal(rsREPOR!PARTACCESSORIES)
            INS_MAT_TMP = NumericVal(rsREPOR!PARTMATERIALS)

            Set rsDet = gconDMIS.Execute("SELECT JOBTYPE,WCODE,DET_AMT,DISCOUNT_2,STATUS1,DETDSC,DETCDE,LIVIL FROM CSMS_RO_DET WHERE TRANSTYPE = 'R' AND REP_OR = '" & rsREPOR!REP_OR & "' ORDER BY LIVIL,LINE_NO ASC")
            If Not (rsDet.BOF And rsDet.EOF) Then
                If CHECKIFHYUNDAI(Null2String(rsREPOR!PLATE_NO)) = True Then
                    Do While Not rsDet.EOF
                        MPR_AMOUNT = NumericVal(N2Str2Zero(rsDet!DET_AMT) - N2Str2Zero(rsDet!Discount_2))
                        If rsDet!LIVIL = "1" Then     '----LIVIL 1
                            If Null2String(rsDet!JOBTYPE) <> "BP" And Null2String(rsDet!JOBTYPE) <> "PMS" Then
                                If Null2String(rsDet!wCode) = "" Then
                                    If INS_LABOR_TMP > 0 Then
                                        If INS_LABOR_TMP >= MPR_AMOUNT Then
                                            INS_LABOR_TMP = INS_LABOR_TMP - MPR_AMOUNT    '
                                            GJ_N_LABOR = GJ_N_LABOR + MPR_AMOUNT    'GJ INSURANCE LABOR
                                        Else
                                            GJ_N_LABOR = GJ_N_LABOR + INS_LABOR_TMP    'GJ INSURANCE LABOR
                                            GJ_C_LABOR = GJ_C_LABOR + (MPR_AMOUNT - INS_LABOR_TMP)    'GJ CUSTOMER PART
                                            INS_LABOR_TMP = 0
                                        End If
                                    Else
                                        GJ_C_LABOR = GJ_C_LABOR + MPR_AMOUNT    'GJ CUSTOMER LABOR
                                    End If
                                ElseIf Null2String(rsDet!wCode) = "W" Then
                                    GJ_W_LABOR = GJ_W_LABOR + MPR_AMOUNT    'GJ WARRANTY LABOR
                                ElseIf Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                                    GJ_I_LABOR = GJ_I_LABOR + MPR_AMOUNT    'GJ INTERNAL LABOR
                                Else
                                    '
                                End If
                            ElseIf Null2String(rsDet!JOBTYPE) = "PMS" Then
                                If Null2String(rsDet!STATUS1) = "Y" And Null2String(rsDet!wCode) = "W" Then
                                    GJ_I_LABOR = GJ_I_LABOR + MPR_AMOUNT    'GJ INTERNAL LABOR
                                Else
                                    If Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                                        GJ_I_LABOR = GJ_I_LABOR + MPR_AMOUNT    'GJ INTERNAL LABOR
                                    ElseIf Null2String(rsDet!wCode) = "W" Then
                                        GJ_W_LABOR = GJ_W_LABOR + MPR_AMOUNT    'GJ WARRANTY LABOR
                                    Else
                                        GJ_C_LABOR = GJ_C_LABOR + MPR_AMOUNT    'GJ CUSTOMER LABOR
                                    End If
                                End If
                            ElseIf Null2String(rsDet!JOBTYPE) = "BP" Then
                                If Null2String(rsDet!wCode) = "" Then
                                    If INS_LABOR_TMP > 0 Then
                                        If INS_LABOR_TMP >= MPR_AMOUNT Then
                                            INS_LABOR_TMP = INS_LABOR_TMP - MPR_AMOUNT
                                            BP_I_LABOR = BP_I_LABOR + MPR_AMOUNT    'BP INSURANCE LABOR
                                        Else
                                            BP_I_LABOR = BP_I_LABOR + INS_LABOR_TMP    'BP INSURANCE LABOR
                                            BP_C_LABOR = BP_C_LABOR + (MPR_AMOUNT - INS_LABOR_TMP)    'BP CUSTOMER LABOR
                                            INS_LABOR_TMP = 0
                                        End If
                                    Else
                                        BP_C_LABOR = BP_C_LABOR + MPR_AMOUNT    'BP CUSTOMER LABOR
                                    End If
                                ElseIf Null2String(rsDet!wCode) = "W" Then
                                    BP_W_LABOR = BP_W_LABOR + MPR_AMOUNT    'BP WARRANTY LABOR
                                ElseIf Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                                    BP_N_LABOR = BP_N_LABOR + MPR_AMOUNT    'BP INTERNAL LABOR
                                Else
                                    '
                                End If
                            End If
                        ElseIf rsDet!LIVIL = "2" Or rsDet!LIVIL = "4" Then
                            If Null2String(rsDet!JOBTYPE) <> "BP" Then
                                If Null2String(rsDet!wCode) = "" Then
                                    If INS_PART_TMP > 0 Then
                                        If INS_PART_TMP >= MPR_AMOUNT Then
                                            INS_PART_TMP = INS_PART_TMP - MPR_AMOUNT
                                            GJ_N_PARTS = GJ_N_PARTS + MPR_AMOUNT    'GJ INSURANCE PART
                                        Else
                                            GJ_N_PARTS = GJ_N_PARTS + INS_PART_TMP    'GJ INSURANCE PART
                                            GJ_C_PARTS = GJ_C_PARTS + (MPR_AMOUNT - INS_PART_TMP)    'GJ CUSTOMER PART
                                            INS_PART_TMP = 0
                                        End If
                                    Else
                                        GJ_C_PARTS = GJ_C_PARTS + MPR_AMOUNT    'GJ CUSTOMER PART
                                    End If
                                ElseIf Null2String(rsDet!wCode) = "W" Then
                                    GJ_W_PARTS = GJ_W_PARTS + MPR_AMOUNT    'GJ WARRANTY PART
                                ElseIf Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                                    GJ_I_PARTS = GJ_I_PARTS + MPR_AMOUNT    'GJ INTERNAL PART
                                Else
                                    '
                                End If
                            ElseIf Null2String(rsDet!JOBTYPE) = "BP" Then
                                If Null2String(rsDet!wCode) = "" Then
                                    If INS_PART_TMP > 0 Then
                                        If INS_PART_TMP >= MPR_AMOUNT Then
                                            INS_PART_TMP = INS_PART_TMP - MPR_AMOUNT
                                            BP_I_PARTS = BP_I_PARTS + MPR_AMOUNT    'BP INSURANCE PART
                                        Else
                                            BP_I_PARTS = BP_I_PARTS + INS_PART_TMP    'BP INSURANCE PART
                                            BP_C_PARTS = BP_C_PARTS + (MPR_AMOUNT - INS_PART_TMP)    'BP CUSTOMER PART
                                            INS_PART_TMP = 0
                                        End If
                                    Else
                                        BP_C_PARTS = BP_C_PARTS + MPR_AMOUNT    'BP CUSTOMER PART
                                    End If
                                ElseIf Null2String(rsDet!wCode) = "W" Then
                                    BP_W_PARTS = BP_W_PARTS + MPR_AMOUNT    'BP WARRANTY PART
                                ElseIf Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                                    BP_N_PARTS = BP_N_PARTS + MPR_AMOUNT    'BP INTERNAL PART
                                Else
                                    '
                                End If
                            Else
                                '
                            End If
                        ElseIf rsDet!LIVIL = "3" Then
                            If Null2String(rsDet!JOBTYPE) <> "BP" Then
                                If Null2String(rsDet!wCode) = "" Then
                                    If INS_MAT_TMP > 0 Then
                                        If INS_MAT_TMP >= MPR_AMOUNT Then
                                            INS_MAT_TMP = INS_MAT_TMP - MPR_AMOUNT
                                            If UCase(Left(Null2String(rsDet!DETCDE), 2)) = "LB" Or UCase(Left(Null2String(rsDet!DETCDE), 2)) = "CH" Then
                                                GJ_N_CHEM = GJ_N_CHEM + MPR_AMOUNT    'GJ CUSTOMER CHEMICAL
                                            Else
                                                GJ_N_OTHER = GJ_N_OTHER + MPR_AMOUNT    'GJ CUSTOMER OTHER
                                            End If
                                        Else
                                            If UCase(Left(Null2String(rsDet!DETCDE), 2)) = "LB" Or UCase(Left(Null2String(rsDet!DETCDE), 2)) = "CH" Then
                                                GJ_C_CHEM = GJ_C_CHEM + INS_MAT_TMP    'GJ INSURANCE CHEMICAL
                                            Else
                                                GJ_C_OTHER = GJ_C_OTHER + INS_MAT_TMP    'GJ INSURANCE OTHER
                                            End If

                                            If UCase(Left(Null2String(rsDet!DETCDE), 2)) = "LB" Or UCase(Left(Null2String(rsDet!DETCDE), 2)) = "CH" Then
                                                GJ_C_CHEM = GJ_C_CHEM + MPR_AMOUNT    'GJ CUSTOMER CHEMICAL
                                            Else
                                                GJ_C_OTHER = GJ_C_OTHER + MPR_AMOUNT    'GJ CUSTOMER OTHER
                                            End If
                                            INS_MAT_TMP = 0
                                        End If
                                    Else
                                        If UCase(Left(Null2String(rsDet!DETCDE), 2)) = "LB" Or UCase(Left(Null2String(rsDet!DETCDE), 2)) = "CH" Then
                                            GJ_C_CHEM = GJ_C_CHEM + MPR_AMOUNT    'GJ CUSTOMER CHEMICAL
                                        Else
                                            GJ_C_OTHER = GJ_C_OTHER + MPR_AMOUNT    'GJ CUSTOMER OTHER
                                        End If
                                    End If
                                ElseIf Null2String(rsDet!wCode) = "W" Then
                                    If UCase(Left(Null2String(rsDet!DETCDE), 2)) = "LB" Or UCase(Left(Null2String(rsDet!DETCDE), 2)) = "CH" Then
                                        GJ_W_CHEM = GJ_W_CHEM + MPR_AMOUNT    'GJ WARRANTY CHEMICAL
                                    Else
                                        GJ_W_OTHER = GJ_W_OTHER + MPR_AMOUNT    'GJ WARRANTY OTHER
                                    End If
                                ElseIf Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                                    If UCase(Left(Null2String(rsDet!DETCDE), 2)) = "LB" Or UCase(Left(Null2String(rsDet!DETCDE), 2)) = "CH" Then
                                        GJ_I_CHEM = GJ_I_CHEM + MPR_AMOUNT    'GJ INTERNAL CHEMICAL
                                    Else
                                        GJ_I_OTHER = GJ_I_OTHER + MPR_AMOUNT    'GJ INTERNAL OTHER
                                    End If
                                Else
                                    '
                                End If
                            ElseIf Null2String(rsDet!JOBTYPE) = "BP" Then
                                If Null2String(rsDet!wCode) = "" Then
                                    If INS_MAT_TMP > 0 Then
                                        If INS_MAT_TMP >= MPR_AMOUNT Then
                                            INS_MAT_TMP = INS_MAT_TMP - MPR_AMOUNT
                                            BP_I_OTHER = BP_I_OTHER + MPR_AMOUNT    'BP INTERNAL OTHER
                                        Else
                                            BP_I_OTHER = BP_I_OTHER + INS_MAT_TMP    'BP INTERNAL OTHER
                                            BP_C_OTHER = BP_C_OTHER + (MPR_AMOUNT - INS_MAT_TMP)    'BP CUSTOMER PART
                                            INS_MAT_TMP = 0
                                        End If
                                    Else
                                        BP_C_OTHER = BP_C_OTHER + MPR_AMOUNT    'BP CUSTOMER OTHER
                                    End If
                                ElseIf Null2String(rsDet!wCode) = "W" Then
                                    BP_W_OTHER = BP_W_OTHER + MPR_AMOUNT    'BP WARRANTY OTHER
                                ElseIf Null2String(rsDet!wCode) = "C" Or Null2String(rsDet!wCode) = "S" Then
                                    BP_N_OTHER = BP_N_OTHER + MPR_AMOUNT    'BP INTERNAL OTHER
                                Else
                                    '
                                End If
                            Else
                                '
                            End If
                        End If

                        rsDet.MoveNext
                    Loop
                Else
                    Do While Not rsDet.EOF
                        MPR_AMOUNT = NumericVal(N2Str2Zero(rsDet!DET_AMT) - N2Str2Zero(rsDet!Discount_2))
                        If rsDet!LIVIL = "1" Then
                            If Null2String(rsDet!JOBTYPE) = "BP" Then
                                BP_OTHER_LABOR = BP_OTHER_LABOR + MPR_AMOUNT    'BP OTHER BRAND LABOR
                            Else
                                GJ_OTHER_LABOR = GJ_OTHER_LABOR + MPR_AMOUNT    'GJ OTHER BRAND LABOR
                            End If
                        End If

                        If rsDet!LIVIL = "2" Or rsDet!LIVIL = "4" Then
                            If Null2String(rsDet!JOBTYPE) = "BP" Then
                                BP_OTHER_PART = BP_OTHER_PART + MPR_AMOUNT    'BP OTHER BRAND PART
                            Else
                                GJ_OTHER_PART = GJ_OTHER_PART + MPR_AMOUNT    'GJ OTHER BRAND PART
                            End If
                        End If

                        If rsDet!LIVIL = "3" Then
                            If Null2String(rsDet!JOBTYPE) = "BP" Then
                                BP_OTHER_MAT = BP_OTHER_MAT + MPR_AMOUNT    'BP OTHER BRAND OTHER
                            Else
                                GJ_OTHER_MAT = GJ_OTHER_MAT + MPR_AMOUNT    'GJ OTHER BRAND OTHER
                            End If
                        End If

                        rsDet.MoveNext
                    Loop
                End If

            End If
            rsREPOR.MoveNext
        Loop
    End If
End Sub

Sub InitializeVariables()
     EMP_COUNT = 0
     TOTAL_WORKDAY = 0
     TOTAL_PERS = 0
     DIRE_PERS = 0
     GEN_JOB_TECH = 0
     BODY_AND_PAINT = 0
     SER_ADV = 0
     WARRANTY = 0
     INHOUSE_INST = 0
     BILL_STAFF = 0
     MANTH = 0
     YEER = 0
     
     GJ_Available_hrs = 0
     BP_Available_hrs = 0
     QS_Available_hrs = 0
     vSC_QUICKSERVICE = 0
     vSC_GENERALJOB = 0
     vSC_BODYANDPAINT = 0
     TOTAL_CAPACITY = 0
    
     vGJMaster = 0
     vGJExpert = 0
     vGJCertified = 0
     vGJNew = 0
     IHTechPaint = 0
     IHTechTinsmith = 0
     CONTech = 0
     SAMas = 0
     SACert = 0
     SANew = 0
     ForeMan = 0
     Warr = 0
     INIns = 0
     BilStf = 0
     Other = 0
    
     REC = 0
     REL = 0
     FPM = 0
     GJ_CUSTOMERPAID_10K = 0
     GJOTHERHYUNDAI = 0
     GJWARRANTYUNIT = 0
     GJINTERNALUNIT = 0
     GJOTHERBRAND = 0
     BPINSURANCE = 0
     BPCUSTOMERPAID = 0
     BPWARRANTYUNIT = 0
     BPINTERNALUNIT = 0
     BPOTHERBRAND = 0
     PDI = 0
    
     UNIT_RECIEVED_OB_GJ = 0
     UNIT_RECIEVED_OB_BP = 0
     TRIGER_INS = ""
     TRIG_GJ_INS = 0
     TRIG_GJ_FPM = 0
     TRIG_GJ_PMS = 0
     TRIG_GJ_CST = 0
     TRIG_GJ_WAR = 0
     TRIG_GJ_INT = 0
     TRIG_GJ_OTH = 0
     TRIG_BP_INS = 0
     TRIG_BP_CST = 0
     TRIG_BP_WRT = 0
     TRIG_BP_INT = 0
     TRIG_BP_OTH = 0
     VPDI_CNT = 0
    
     JOB_COUNT = 0
     GJ_CNT = 0:  BP_CNT = 0
     GJ_TRG = 0:  BP_TRG = 0
     vGJINS = 0
     vMAJOR_BP = 0
     vMINOR_BP = 0
    
     vLAST3MONTH = 0
     vLAST3MONTH = 0
    
     PRODUCTIVITY_BP_AVAILABLE_HRS = 0
     PRODUCTIVITY_BP_ACTUAL_HRS = 0
     PRODUCTIVITY_BP_SOLD_HRS = 0
     PRODUCTIVITY_BP_BACKJOB_HRS = 0
    
     PRODUCTIVITY_GJ_AVAILABLE_HRS = 0
     PRODUCTIVITY_GJ_ACTUAL_HRS = 0
     PRODUCTIVITY_GJ_SOLD_HRS = 0
     PRODUCTIVITY_GJ_BACKJOB_HRS = 0
    
     PRODUCTIVITY_QS_AVAILABLE_HRS = 0
     PRODUCTIVITY_QS_ACTUAL_HRS = 0
     PRODUCTIVITY_QS_SOLD_HRS = 0
     PRODUCTIVITY_QS_BACKJOB_HRS = 0
    
     GJ_N_LABOR = 0
     GJ_C_LABOR = 0
     GJ_W_LABOR = 0
     GJ_I_LABOR = 0
     BP_I_LABOR = 0
     BP_C_LABOR = 0
     BP_W_LABOR = 0
     BP_N_LABOR = 0
     GJ_N_PARTS = 0
     GJ_C_PARTS = 0
     GJ_W_PARTS = 0
     GJ_I_PARTS = 0
     BP_I_PARTS = 0
     BP_C_PARTS = 0
     BP_W_PARTS = 0
     BP_N_PARTS = 0
    
     GJ_N_CHEM = 0
     GJ_C_CHEM = 0
     GJ_W_CHEM = 0
     GJ_I_CHEM = 0
     BP_I_CHEM = 0
     BP_C_CHEM = 0
     BP_W_CHEM = 0
     BP_N_CHEM = 0
     GJ_N_OTHER = 0
     GJ_C_OTHER = 0
     GJ_W_OTHER = 0
     GJ_I_OTHER = 0
     BP_I_OTHER = 0
     BP_C_OTHER = 0
     BP_W_OTHER = 0
     BP_N_OTHER = 0
    
     GJ_OTHER_BRAND = 0
     BP_OTHER_BRAND = 0
     GJ_OTHER_LABOR = 0
     BP_OTHER_LABOR = 0
     GJ_OTHER_PART = 0
     BP_OTHER_PART = 0
     GJ_OTHER_MAT = 0
     BP_OTHER_MAT = 0
    
     INS_LABOR = 0
     INS_LABOR_TMP = 0
     INS_PART = 0
     INS_PART_TMP = 0
     INS_MAT = 0
     INS_MAT_TMP = 0
     INS_ACC = 0
     INS_ACC_TMP = 0
     MPR_AMOUNT = 0
End Sub
