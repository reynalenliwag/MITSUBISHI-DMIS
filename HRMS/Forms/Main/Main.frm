VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7E0552E8-E2C9-4C9E-BC75-F34D871C75F4}#2.0#0"; "wizEncrypt.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#10.4#0"; "COF080~1.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#10.4#0"; "CODEJO~2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO301B~1.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000F&
   Caption         =   "Human Resources Management System"
   ClientHeight    =   8295
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "Main.frx":030A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   4620
      Top             =   3360
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   4680
      Top             =   3960
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3060
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":DFF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":EC4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":EF66
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":F280
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":F59A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":F8B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":FBCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":FEF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1020A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":15CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1600E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":16460
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1677A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":16BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1701E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":17470
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":178C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":17BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":18830
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin wizEncrypt.wizEnc wizEnc1 
      Left            =   4560
      Top             =   8190
      _ExtentX        =   3969
      _ExtentY        =   3969
   End
   Begin Crystal.CrystalReport rptReports 
      Left            =   3720
      Top             =   3900
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8040
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   4939
            MinWidth        =   4939
            Object.ToolTipText     =   "Login Name"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
            Object.ToolTipText     =   "Login Level"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Object.ToolTipText     =   "Current Date"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   4410
            MinWidth        =   4410
            Object.ToolTipText     =   "Login Time"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10:34 AM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
            Object.ToolTipText     =   "Num Lock Status"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Caps Lock Status"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1147
            MinWidth        =   1147
            TextSave        =   "SCRL"
            Object.ToolTipText     =   "Scroll Lock Status"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   7832
            MinWidth        =   7832
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PopupControl PopCntrl 
      Left            =   2610
      Top             =   4005
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   2
      Width           =   140
      Height          =   270
      Animation       =   1
      AnimateDelay    =   125
      ShowDelay       =   2500
      BackgroundBitmap=   60
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   4200
      Top             =   3960
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      DesignerControls=   "Main.frx":1E592
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    HEADOREMP = "EMP_A"
    CenterMe Screen, Me, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "]"
    '& App.Revision & "]"
    ApplyThemes
    ConfigurePopUps
End Sub

Private Sub MDIForm_Resize()
    CenterMe Screen, Me, 0
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If MsgBox("Exit HRMS, Are You Sure?", vbExclamation + vbOKCancel, "Exit System") = vbOK Then
        Dim FRM                                                       As Form
        For Each FRM In Forms
            If Not (FRM Is Nothing) Then
                Unload FRM
            End If
        Next
        CommandBars1.SaveCommandBars MODULENAME, App.TITLE, "Layout"
        UnloadForm Me
    Else
        Cancel = 1
        frmMainMenu.Show
    End If
End Sub

Private Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    'MsgBox Control.ID
    Select Case Control.ID
            ''''''''''''''''''''''''''''''''''''''''''FILES'''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case FILES_SSSBRACKETING
            '1020
            If Module_Access(LOGID, "FILES SSS BRACKETING", "DATA ENTRY") = False Then Exit Sub
            'frmHRMSSSSBracketing.Show

        Case FILES_PHILHEALTHBRACKETING
            '1021
            If Module_Access(LOGID, "FILES PHILHEATH BRACKETING", "DATA ENTRY") = False Then Exit Sub
            'frmHRMSPHBracketing.Show
            frmHRMSTables_PHIC.Show

        Case FILES_SALARYGRADECODES
            '1022
            If Module_Access(LOGID, "FILES SALARY GRADE CODES", "DATA ENTRY") = False Then Exit Sub
            frmHRMSSalaryGrade.Show

        Case FILES_DEPARTMENT
            '1023
            If Module_Access(LOGID, "FILES DEPARTMENT", "DATA ENTRY") = False Then Exit Sub
            frmHRMSDepartment.Show

        Case FILES_GROUPS
            '1024
            If Module_Access(LOGID, "FILES GROUPS", "DATA ENTRY") = False Then Exit Sub
            'frmHRMSGroups.Show

            ''''''''''''''''''''''''''''''''''''''''''EMPLOYEES'''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case EMP_INFO_EMPLOYEEINFO, TOOL_EMPLOYEEINFORMATION
            '1025  -  1083
            If Module_Access(LOGID, "EMPLOYEE INFO", "DATA ENTRY") = False Then Exit Sub
            EMP_TYPE = "EMPLOYEE"
            HEADOREMP = "EMP_A"
            'frmHRMSEmpInfo.Option1.Value = True
            frmHRMSEmpInfo.Show

        Case EMP_INFO_CONTRACTUALINFO
            '1026
            If Module_Access(LOGID, "CONTRACTUAL INFO", "DATA ENTRY") = False Then Exit Sub

            EMP_TYPE = "CONTRACTUAL"
            HEADOREMP = "EMP_A"
            'frmHRMSEmpInfo.Option1.Value = True
            frmHRMSEmpInfo.Show

        Case EMP_INFO_ALLOWANCEBASEEMPLOYEES
            '1027
            If Module_Access(LOGID, "ALLOWANCE BASE INFO", "DATA ENTRY") = False Then Exit Sub
            EMP_TYPE = "ALLOWANCE BASE"
            HEADOREMP = "EMP_A"
            'frmHRMSEmpInfo.Option1.Value = True
            frmHRMSEmpInfo.Show

        Case EMP_INFO_MANAGERSINFO
            '1028
            If Module_Access(LOGID, "MANAGERS INFO", "DATA ENTRY") = False Then Exit Sub
            EMP_TYPE = "EMPLOYEE"
            HEADOREMP = "HEAD"
            'frmHRMSEmpInfo.Option1.Value = True
            frmHRMSEmpInfo.Show

        Case EMP_MAINTAIN_ATTENDANCE, TOOL_NUMBEROFDAYS
            '1029 - 1084
            If Module_Access(LOGID, "EMPLOYEE MAINTAIN ATTENDANCE", "DATA ENTRY") = False Then Exit Sub
            EMP_TYPE = "EMPLOYEE"
            frmHRMSDailyMonitoring.Show

        Case EMP_MAINTAIN_SALARYADVANCE
            '1030
            If Module_Access(LOGID, "EMPLOYEE MAINTAIN ADVANCE", "DATA ENTRY") = False Then Exit Sub
            frmHRMS_Advance.Show

        Case EMP_MAINTAIN_OVERTIME
            '1031   -   1085
            If Module_Access(LOGID, "EMPLOYEE MAINTAIN OVERTIME", "DATA ENTRY") = False Then Exit Sub
            EMP_TYPE = "EMPLOYEE"
            frmHRMSOvertime.Show

        Case EMP_MAINTAIN_DEDUCTIONS
            '1032   -   1086
            If Module_Access(LOGID, "EMPLOYEE MAINTAIN DEDUCTIONS", "DATA ENTRY") = False Then Exit Sub
            EMP_TYPE = "EMPLOYEE"
            DEDUCTION_OPTION = "ATTENDANCE DEDUCTION"
            frmHRMSDeductions.Show

        Case EMP_MAINTAIN_COMMISSION, TOOL_COMMISSION
            '1033   -   1087
            If Module_Access(LOGID, "EMPLOYEE MAINTAIN COMMISSION", "DATA ENTRY") = False Then Exit Sub
            EMP_TYPE = "EMPLOYEE"
            frmHRMSCommission.Show

        Case EMP_MAINTAIN_ATMENTRY, TOOL_ATMDETAILS
            '1034   -   1089
            If Module_Access(LOGID, "EMPLOYEE MAINTAIN ATM ENTRY", "DATA ENTRY") = False Then Exit Sub
            EMP_TYPE = "EMPLOYEE"
            frmHRMSATM.Show

        Case EMP_MAINTAIN_ADJUSTMENT
            '1035   -   1088
            If Module_Access(LOGID, "EMPLOYEE MAINTAIN ADJUSTMENTS", "DATA ENTRY") = False Then Exit Sub
            EMP_TYPE = "EMPLOYEE"
            frmHRMSAdjustment.Show

        Case EMP_LEDGER, TOOL_LEDGER
            '1036   -   1102
            If Module_Access(LOGID, "EMPLOYEE LEDGER", "DATA ENTRY") = False Then Exit Sub
            EMP_TYPE = "EMPLOYEE"
            HEADOREMP = "EMP_A"
            frmHRMSLedger.Show

        Case PROCESS_GENERATEPAYROLL, TOOL_GENERATEPAYROLL
            '1037   -   1091
            If Module_Access(LOGID, "PROCESS GENERATE PAYROLL", "PROCESSING") = False Then Exit Sub
            frmHRMSGenerate.Show

        Case PROCESS_YEARTODATEPROCESSING, TOOL_YEARTODATEPROCESSING
            '1038   -   1092
            If Module_Access(LOGID, "PROCESS YEAR-TO-DATE PROCESS", "PROCESSING") = False Then Exit Sub
            frmHRMSYTDProcessing.Show

        Case PROCESS_BIRALPHALISTPROCESSING
            '1039
            If Module_Access(LOGID, "PROCESS BIR ALPHA-LIST PROCESSING", "PROCESSING") = False Then Exit Sub
            On Error Resume Next
            frmHRMSBIRProcessing.Show

        Case PROCESS_UPDATECOMMISSION
            '1040
            ''''EMPTY

        Case PROCESS_UPDATEATTENDANCE
            '1041
            If Module_Access(LOGID, "PROCESS UPDATE ATTENDANCE", "PROCESSING") = False Then Exit Sub
            On Error Resume Next
            frmHRMSUpDateAttendance.Show



            '''REPORTS
            '=============================================================
        Case RPT_ALPHA_ALPHALISTOFTERMINATEDEMPLOYEESBEFOREDECEMBER
            If Module_Access(LOGID, "REPORT ALPHALIST TERMINATED", "REPORTS") = False Then Exit Sub
            FormYearlyRequest = "ALTERMINATED"
            frmHRMSYearly.Show

        Case RPT_ALPHA_ALPHALISTOFEMPLOYEESWITHPREVIOUSEMPLOYERWITHINAYEAR
            If Module_Access(LOGID, "REPORT ALPHALIST WITH PREVIOUS EMP", "REPORTS") = False Then Exit Sub
            FormYearlyRequest = "ALWITHEMP"
            frmHRMSYearly.Show

        Case RPT_ALPHA_ID_ALPHALISTOFEMPLOYEESWITHNOPREVIOUSEMPLOYERWITHINAYEAR
            If Module_Access(LOGID, "REPORT ALPHALIST W/OUT PREVIOUS EMPLOYEE", "REPORTS") = False Then Exit Sub

            FormYearlyRequest = "ALWITHNOEMP"
            frmHRMSYearly.Show

        Case RPT_ALPH_OTHER_EMPLOYEELISTING
            If Module_Access(LOGID, "REPORT EMPLOYEE LISTING", "REPORTS") = False Then Exit Sub

            rptReports.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
            rptReports.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
            rptReports.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
            PrintSQLReport rptReports, HRMS_REPORT_PATH & "emplist.rpt", "", DMIS_REPORT_Connection, 1

        Case RPT_ALPH_OTHER_EMPLOYEELISTINGWAGE
            If Module_Access(LOGID, "REPORT OTHER EMPLOYEE LISTING WAGE", "REPORTS") = False Then Exit Sub

            rptReports.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
            rptReports.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
            rptReports.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
            PrintSQLReport rptReports, HRMS_REPORT_PATH & "empAge.rpt", "", DMIS_REPORT_Connection, 1

        Case RPT_ALPHA_OTHER_RESIGNEESLISTING
            If Module_Access(LOGID, "REPORT RESIGNESS LISTING", "REPORTS") = False Then Exit Sub

            rptReports.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
            rptReports.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
            rptReports.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
            PrintSQLReport rptReports, HRMS_REPORT_PATH & "resignees.rpt", "", DMIS_REPORT_Connection, 1

        Case RPT_ALPHA_OTHER_LOANSMASTERFILELISTING
            If Module_Access(LOGID, "REPORT LOANS LISTING", "REPORTS") = False Then Exit Sub

            rptReports.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
            rptReports.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
            rptReports.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
            PrintSQLReport rptReports, HRMS_REPORT_PATH & "Loanmas.rpt", "", DMIS_REPORT_Connection, 1

        Case RPT_MONTH_WITHHOLDINGTAXMONTHLYREMITTANCE, TOOL_WITHOLDINGTAXMONTHLYREMITTANCE
            If Module_Access(LOGID, "REPORT WITHHOLDING TAX MONTHLY REMITTANCE", "REPORTS") = False Then Exit Sub
            frmHRMSPHMonthly.Caption = "REPORT WITHHOLDING TAX MONTHLY REMITTANCE"
            frmHRMSPHMonthly.Show
        Case RPT_MONTH_SSSMONTHLYREMITTANCE, TOOL_SSSMONTHLYREMITTANCE
            If Module_Access(LOGID, "REPORT SSS MONTHLY REMITTANCE", "REPORTS") = False Then Exit Sub
            frmHRMSPHMonthly.Caption = "REPORT SSS MONTHLY REMITTANCE"
            frmHRMSPHMonthly.Show

        Case RPT_MONTH_PHILHEALTHMONTHLYREMITTANCE, TOOL_PHILHEALTHMONTHLYREMITTANCE
            If Module_Access(LOGID, "REPORT PHILHEALTH MONTHLY REMITTANCE", "REPORTS") = False Then Exit Sub
            frmHRMSPHMonthly.Caption = "PHILHEALTH MONTHLY REMITTANCE"
            frmHRMSPHMonthly.Show
        Case RPT_MONTH_PAGIBIGMONTHLYREMITTANCE, TOOL_PAGIBIGMONTHLYREMITTANCE
            If Module_Access(LOGID, "REPORT PAG-IBIG MONTHLY REMITTANCE", "REPORTS") = False Then Exit Sub
            frmHRMSPHMonthly.Caption = "REPORT PAG-IBIG MONTHLY REMITTANCE"
            frmHRMSPHMonthly.Show

        Case RPT_MONTH_LOANSMONTHLYREMITTANCE
            If Module_Access(LOGID, "REPORT LOANS MONTHLY REMITTANCE", "REPORTS") = False Then Exit Sub

            frmHRMSLoansMonthly.Show

        Case RPT_MONTH_PRINTMONTHLYPAYROLL
            If Module_Access(LOGID, "REPORT PRINT MONTHLY PAYROLL", "REPORTS") = False Then Exit Sub

            frmHRMSMonthlyPayroll.Show

        Case RPT_SCHD_SCHEDULEOFSSSPREMIUMCONTRIBUTION
            If Module_Access(LOGID, "REPORT SCHEDULE OF SSS PREMIUM CONTRIBUTION", "REPORTS") = False Then Exit Sub

            FormYearlyRequest = "SCHEDSSS"
            frmHRMSYearly.Show

        Case RPT_SCHD_SCHEDULEOFPHILHEALTHPREMIUMCONTRIBUTION
            If Module_Access(LOGID, "REPORT SCHEDULE OF PHILHEALTH PREMIUM CONTRIBUTION", "REPORTS") = False Then Exit Sub

            FormYearlyRequest = "SCHEDPHIC"
            frmHRMSYearly.Show

        Case RPT_SCHD_SCHEDULEOFPAGIBIGPREMIUMCONTRIBUTION
            If Module_Access(LOGID, "REPORT SCHEDULE OF PAGIBIG PREMIUM CONTRIBUTION", "REPORTS") = False Then Exit Sub

            FormYearlyRequest = "SCHEDPAGIBIG"
            frmHRMSYearly.Show

        Case RPT_SCHD_SCHEDULEOFTAXWITHHELD
            If Module_Access(LOGID, "REPORT SCHEDULE OF TAX WITHHELD", "REPORTS") = False Then Exit Sub

            FormYearlyRequest = "SCHEDTAX"
            frmHRMSYearly.Show

        Case RPT_SCHD_SCHEDULEOFOVERTIMEPAY
            If Module_Access(LOGID, "REPORT SCHEDULE OF OVERTIME PAY", "REPORTS") = False Then Exit Sub

            FormYearlyRequest = "SCHEDOVERTIME"
            frmHRMSYearly.Show

        Case RPT_SCHD_SCHEDULEOFPAYROLL
            If Module_Access(LOGID, "REPORT SCHEDULE OF PAYROLL", "REPORTS") = False Then Exit Sub

            FormYearlyRequest = "SCHEDPAYROLL"
            frmHRMSYearly.Show

        Case RPT_SCHD_SCHEDULEOFCOMMISSION
            If Module_Access(LOGID, "REPORT SCHEDULE OF COMMISSION", "REPORTS") = False Then Exit Sub

            FormYearlyRequest = "SCHEDCOMMISSION"
            frmHRMSYearly.Show

        Case RPT_SCHD_SCHEDULEOFCOMMISSIONTAX
            If Module_Access(LOGID, "REPORT SCHEDULE OF COMMISSION TAX", "REPORTS") = False Then Exit Sub

            FormYearlyRequest = "SCHEDCOMMISSIONTAX"
            frmHRMSYearly.Show

        Case RPT_SCHD_THMONTHPAYSCHEDULE
            If Module_Access(LOGID, "REPORT 13TH MONTH PAY SCHEDULE", "REPORTS") = False Then Exit Sub

            frmHRMSPrint13thMonth.Show

        Case RPT_SCHD_SCHEDULEOFTAXDUEREFUND
            If Module_Access(LOGID, "REPORT SCHEDULE OF TAXDUE/REFUND", "REPORTS") = False Then Exit Sub

            FormYearlyRequest = "SCHEDTAXDUEREFUND"
            frmHRMSYearly.Show

        Case RPT_SUMMARY_DEDUCTIONDETAILS
            If Module_Access(LOGID, "REPORT DEDUCTION DETAILS", "REPORTS") = False Then Exit Sub

            frmHRMSDedDetails.Show

        Case RPT_SUMMARY_YEARTODATEDETAILS
            If Module_Access(LOGID, "REPORT YEAR-TO-DATE DETAILS", "REPORTS") = False Then Exit Sub

            frmHRMSPrintYTDProcessing.Show

        Case RPT_MISC_PRINTBLANKEMPLOYEEINFOSHEET
            If Module_Access(LOGID, "REPORT PRINT BLANK EMP. INFO SHEET", "REPORTS") = False Then Exit Sub

            PrintSQLReport rptReports, HRMS_REPORT_PATH & "blankempinfo.rpt", "", DMIS_REPORT_Connection, 1

        Case RPT_PRINTPAYROLLSHEET, TOOL_PRINTPAYROLLSHEET
            If Module_Access(LOGID, "REPORT PRINT PAYROLL SHEET", "REPORTS") = False Then Exit Sub

            frmHRMSPrintPayroll.Show

        Case RPT_PRINTATMADVICE, TOOL_PRINTATMADVICE
            If Module_Access(LOGID, "REPORT PRINT ATM ADVICE", "REPORTS") = False Then Exit Sub

            frmHRMSPrintATM.Show

        Case RPT_PRINTPAYROLLJOURNAL
            If Module_Access(LOGID, "REPORT PRINT PAYROLL JOURNAL", "REPORTS") = False Then Exit Sub

            frmHRMSPayJournal.Show

        Case MAINTAIN_COMPANYPROFILE
            If Module_Access(LOGID, "HRMS PROFILE", "SYSTEM") = False Then Exit Sub

            frmHRMSProfile.Show

        Case MAINTAIN_PASSWORDMAINTENANCE
            If Module_Access(LOGID, "PASSWORD MAINTENANCE", "SYSTEM") = False Then Exit Sub
            frmAccMaintenance.Show


        Case TOOL_CALENDAR

        Case TOOL_CALCULATOR
            'frmToolsCalculator.Show
        Case TOOL_ZIPBACKUPDATABASE

        Case TOOL_COMPACTREPAIRDATABASE

        Case WINDOW_ABOUTTHEAUTHOR, TOOL_ABOUTTHEAUTHOR
            frmAbout.Show

        Case WINDOW_EXIT, TOOL_EXITHRMS
            Unload Me

        Case TOOL_DASHBOARD
            frmMainMenu.Show
        Case 1020
            frmHRMSTables_SSS.Show
        Case 1112
            frmHRMSTables_PAGIBIG.Show
        Case 1113
            frmHRMSTables_Tax.Show
        Case 1114
            frmHRMSPayrollSetup.Show
        Case 1115
            frmHRMS_HolidaySetup.Show
        Case 1116
            frmSETUP_Deduction.Show
        Case 1120
            frmHRMS_Shift_Management.Show
        Case 1118
            frmHRMS_TimeShift.Show
        Case TOOL_OVERTIME
            frmHRMSOTCodes.Show
        Case TOOL_DEDUCTION
            frmHRMSDeductionCodeMaterFile.Show
        Case TOOL_ADJUSTMENT
            frmHRMSCodes_Adjustment.Show
        Case 1119
            frmHRMS_LoanCodes.Show
        Case 1122
            If Module_Access(LOGID, "EMPLOYEE MAINTAIN DEDUCTIONS", "DATA ENTRY") = False Then Exit Sub
            EMP_TYPE = "EMPLOYEE"
            DEDUCTION_OPTION = "OTHER DEDUCTIONS"
            frmHRMSDeductions.Show
    End Select

End Sub

Private Sub ApplyThemes()
    CommandBarsGlobalSettings.App = App
    With CommandBars1
        .LoadDesignerBars
        .LoadCommandBars MODULENAME, App.TITLE, "Layout"
        .PaintManager.ClearTypeTextQuality = True
        .TabWorkspace.ThemedBackColor = False
        .StatusBar.Visible = True
        .Options.SyncFloatingToolbars = True
    End With
    With SkinFramework1
        '.LoadSkin "C:\DMIS 2.0\Styles\royale.cjstyle", ""
        .ApplyWindow Me.hwnd
        '.ApplyOptions = xtpSkinApplyFrame Or xtpSkinApplyColors Or xtpSkinApplyMetrics
        .ApplyOptions = xtpSkinApplyFrame Or xtpSkinApplyColors Or Not xtpSkinApplyMetrics
    End With
    Dim ToolTipContext                                                As ToolTipContext
    Set ToolTipContext = CommandBars1.ToolTipContext
    With ToolTipContext
        .ShowTitleAndDescription True, xtpToolTipIconInfo
        .SetMargin 2, 2, 2, 2
        .MaxTipWidth = 180
        If .IsBalloonStyleSupported Then
            .Style = xtpToolTipBalloon
        Else
            .Style = xtpToolTipOffice2007
        End If
        .ShowShadow = True
    End With
End Sub

''''''''''''''START REGION POPUPCONTROLS''''''''''''''''''''''''''''''''''''
Private Sub ConfigurePopUps()
    Dim Item                                                          As PopupControlItem
    PopCntrl.RemoveAllItems
    'PopCntrl.Icons.AddIcons ImageManager.Icons
    PopCntrl.Icons.AddIcons CommandBars1.Icons
    'PopCntrl.VisualTheme = xtpPopupThemeOffice2003
    'PopCntrl.SetSize 270, 140

    Set Item = PopCntrl.AddItem(245, 8, 265, 20, vbNullString)
    Item.Button = True
    Item.IconIndex = 899
    Item.ID = 707
    Item.HEIGHT = 20
    Item.Width = 20
    Item.CenterIcon
    Set Item = PopCntrl.AddItem(10, 10, 218, 30, vbNullString)
    Item.TextColor = RGB(15, 48, 145)
    Item.Bold = True
    Item.Font.Size = 10
    Item.Hyperlink = False
    Set Item = PopCntrl.AddItem(10, 32, 60, 50, vbNullString)
    Item.HEIGHT = 50
    Item.Width = 50
    Item.IconIndex = 0
    Item.Hyperlink = False

    Set Item = PopCntrl.AddItem(62, 32, 260, 50, vbNullString)
    Item.TextAlignment = DT_WORDBREAK Or DT_LEFT
    Item.HEIGHT = 50
    Item.ID = 655
    Item.Hyperlink = False

    Set Item = PopCntrl.AddItem(20, 85, 260, 105, vbNullString)
    Item.TextAlignment = DT_WORDBREAK Or DT_LEFT
    Item.TextColor = RGB(190, 1, 1)
    Item.HEIGHT = 50
    Item.Font.Size = 7
    Item.Hyperlink = False
End Sub

Private Sub PopCntrl_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    If Item.ID = 707 Then
        PopCntrl.Close
    End If
End Sub

''''''''''''''END REGION POPUPCONTROLS''''''''''''''''''''''''''''''''''''
Private Sub Timer1_Timer()

    If TIMER_REMIND = "" Then
        ReminderModule ""
    Else
        If DateDiff("n", TIMER_REMIND, Now) >= 0 Then
            frmSMIS_Files_Reminders.Show


        End If
    End If
End Sub

