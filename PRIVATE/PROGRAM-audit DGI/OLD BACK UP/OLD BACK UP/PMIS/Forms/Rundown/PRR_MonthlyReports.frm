VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISReports_PRRMonthly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PRR Monthly Reports"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   ForeColor       =   &H00DEDFDE&
   Icon            =   "PRR_MonthlyReports.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3045
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
      Left            =   870
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select year from the list"
      Top             =   510
      Width           =   1965
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
      Left            =   870
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select month from the list"
      Top             =   120
      Width           =   1965
   End
   Begin Crystal.CrystalReport rptIssuances 
      Left            =   0
      Top             =   1050
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Monthly Issuances"
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1620
      MouseIcon       =   "PRR_MonthlyReports.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "PRR_MonthlyReports.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   900
      MouseIcon       =   "PRR_MonthlyReports.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "PRR_MonthlyReports.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Print Report"
      Top             =   960
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   30
      TabIndex        =   4
      Top             =   540
      Width           =   735
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2130
      TabIndex        =   3
      Top             =   3000
      Width           =   495
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   30
      TabIndex        =   2
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmPMISReports_PRRMonthly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ERRORCODE:
    Screen.MousePointer = 11

    If PRR_REPORT = "RETAIL SALES" Then
        If Function_Access(LOGID, "Acess_Print", "REPORTS TOTAL RETAIL SALES") = False Then Exit Sub
        rptIssuances.WindowTitle = "PARTS RUNDOWN REPORT - TOTAL RETAIL SALES"
        rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "RunDown\PRR_Retail_Sales.rpt", "MONTH({ORD_HD.TRANDATE}) = " & What_month(cboMonth.Text) & " and YEAR({ORD_HD.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
    End If
    If PRR_REPORT = "COST OF SALES" Then
        If Module_Access(LOGID, "REPORTS TOTAL COST OF SALES", "REPORTS") = False Then Exit Sub
        rptIssuances.WindowTitle = "PARTS RUNDOWN REPORT - TOTAL COST OF SALES"
        rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "RunDown\PRR_Cost_Of_Sales.rpt", "MONTH({ORD_HD.TRANDATE}) = " & What_month(cboMonth.Text) & " and YEAR({ORD_HD.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
    End If
    If PRR_REPORT = "BEGINNING INVENTORY" Then
        If Module_Access(LOGID, "REPORTS PARTS RUNDOWN BEGINNING INVENTORY REPORT", "REPORTS") = False Then Exit Sub
        rptIssuances.WindowTitle = "PARTS RUNDOWN REPORT - BEGINNING INVENTORY REPORT"
        rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "RunDown\PRR_Beginning_Inventory.rpt", "MONTH({ORD_HD.TRANDATE}) = " & What_month(cboMonth.Text) & " and YEAR({ORD_HD.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
    End If
    If PRR_REPORT = "TOTAL PURCHASES" Then
        If Module_Access(LOGID, "REPORTS PARTS RUNDOWN TOTAL PURCHASES", "REPORTS") = False Then Exit Sub
        rptIssuances.WindowTitle = "PARTS RUNDOWN REPORT - TOTAL PURCHASES REPORT"
        rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "RunDown\PRR_Total_Purchases.rpt", "MONTH({PO_HD.PODATE}) = " & What_month(cboMonth.Text) & " and YEAR({PO_HD.PODATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
    End If
    If PRR_REPORT = "INVENTORY ADJUSTMENTS" Then
        If Module_Access(LOGID, "REPORTS PARTS RUNDOWN INVENTORY ADJUSTMENTS", "REPORTS") = False Then Exit Sub
        rptIssuances.WindowTitle = "PARTS RUNDOWN REPORT - INVENTORY ADJUSTMENTS REPORT"
        rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "RunDown\PRR_Inventory_Adjustments.rpt", "MONTH({ORD_HD.TRANDATE}) = " & What_month(cboMonth.Text) & " and YEAR({ORD_HD.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
    End If
    If PRR_REPORT = "PARTS MAD" Then
        If Module_Access(LOGID, "REPORTS PARTS RUNDOWN INVENTORY ADJUSTMENTS", "REPORTS") = False Then Exit Sub

        rptIssuances.WindowTitle = "PARTS RUNDOWN REPORT - MOVING AVERAGE DEMAND"
        rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "RunDown\PRR_Moving_Average_Demand.rpt", "MONTH({ORD_HD.TRANDATE}) = " & What_month(cboMonth.Text) & " and YEAR({ORD_HD.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
    End If
    If PRR_REPORT = "INV_GROSS_RETURN" Then
        If Module_Access(LOGID, "REPORTS PARTS RUNDOWN INVENTORY GROSS RETURN", "REPORTS") = False Then Exit Sub
        rptIssuances.WindowTitle = "PARTS RUNDOWN REPORT - INVENTORY GROSS RETURN"
        rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "RunDown\PRR_Inventory_Gross_Return.rpt", "MONTH({ORD_HD.TRANDATE}) = " & What_month(cboMonth.Text) & " and YEAR({ORD_HD.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
    End If
    If PRR_REPORT = "FILL RATE" Then
        If Module_Access(LOGID, "REPORTS PARTS RUNDOWN FILL RATE", "REPORTS") = False Then Exit Sub
        rptIssuances.WindowTitle = "PARTS RUNDOWN REPORT - FILL RATE REPORT"
        rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "RunDown\PRR_Fill_Rate.rpt", "MONTH({ORD_HD.TRANDATE}) = " & What_month(cboMonth.Text) & " and YEAR({ORD_HD.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
    End If
    If PRR_REPORT = "ORDERED PARTS" Then
        If Module_Access(LOGID, "REPORTS PARTS RUNDOWN ORDERED PARTS REPORT BY CATEGORY", "REPORTS") = False Then Exit Sub
        rptIssuances.WindowTitle = "PARTS RUNDOWN REPORT - ORDERED PARTS BY CATEGORY"
        rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "RunDown\PRR_Ordered_Parts_By_Category.rpt", "MONTH({ORD_HD.TRANDATE}) = " & What_month(cboMonth.Text) & " and YEAR({ORD_HD.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
    End If
    If PRR_REPORT = "PARTS BACK ORDER" Then
        If Module_Access(LOGID, "REPORTS PARTS RUNDOWN PARTS BACK ORDER", "REPORTS") = False Then Exit Sub
        rptIssuances.WindowTitle = "PARTS RUNDOWN REPORT - PARTS BACK ORDER REPORT"
        rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "RunDown\PRR_Back_Ordered.rpt", "MONTH({ORD_HD.TRANDATE}) = " & What_month(cboMonth.Text) & " and YEAR({ORD_HD.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
    End If

    LogAudit "V", "Print " & PRR_REPORT, cboMonth & "/" & cboYear
    Screen.MousePointer = 0
    Exit Sub
ERRORCODE:
    ShowVBError

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    FillcboYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub

