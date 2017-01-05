VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMS_Reports_SASales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Advisor Sales"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4095
   Icon            =   "ReportServiceAdvisorSales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2175
   ScaleWidth      =   4095
   Begin VB.CheckBox Check1 
      Caption         =   "Summary"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1530
      TabIndex        =   3
      Top             =   1230
      Width           =   975
   End
   Begin VB.ComboBox cboServiceAdvisor 
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
      Left            =   1530
      TabIndex        =   0
      Top             =   90
      Width           =   2475
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
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select month from the list"
      Top             =   480
      Width           =   2475
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
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Select year from the list"
      Top             =   870
      Width           =   2475
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
      Height          =   810
      Left            =   3255
      MouseIcon       =   "ReportServiceAdvisorSales.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "ReportServiceAdvisorSales.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   1275
      Width           =   735
   End
   Begin Crystal.CrystalReport rpt 
      Left            =   870
      Top             =   2520
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
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   2535
      MouseIcon       =   "ReportServiceAdvisorSales.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "ReportServiceAdvisorSales.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   1275
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Service Advisor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   -1005
      TabIndex        =   8
      Top             =   150
      Width           =   2490
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   750
      TabIndex        =   7
      Top             =   540
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   930
      Width           =   735
   End
End
Attribute VB_Name = "frmCSMS_Reports_SASales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Function FindSACODE() As String
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT CODE,NAYM FROM CSMS_VW_EMPNO WHERE NAYM = '" & cboServiceAdvisor.Text & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        FindSACODE = LTrim(RTrim(Null2String(rstmp!Code)))
    End If

    Set rstmp = Nothing
End Function

Private Sub cmdPrint_Click()
    'If Function_Access(LOGID, "Acess_PRINT", "SERVICE ADVISOR SALES") = False Then Exit Sub
    Screen.MousePointer = 11
    'On Error GoTo ErrorCode
    Dim Filter                                         As String

    RPT.Reset
    RPT.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    RPT.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    If UCase(cboServiceAdvisor) <> "ALL" Then
        Filter = " AND UCASE({CSMS_RepairOrder.WRITER}) = '" & UCase(cboServiceAdvisor) & "'"
    End If


    If Check1.Value = 1 Then
        'JUN 01/05/2008
        RPT.WindowTitle = "Service Advisor Sales Summary Report"
        RPT.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        RPT.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        RPT.Formulas(2) = "Printedby = '" & LOGNAME & "'"

        PrintSQLReport RPT, CSMS_REPORT_PATH & "ServiceAdvisorSalesSUM.rpt", "ISNULL({RO.DTE_COMP})=false and  Month({RO.DTE_COMP}) = " & What_month(cboMonth.Text) & " AND Year({RO.DTE_COMP}) = " & cboYear.Text & Filter, DMIS_REPORT_Connection, 1

        'LogAudit "V", "SERVICE ADVISOR SALES SUMMARY - REPORTS ", cboServiceAdvisor & cboMonth & cboYear
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "SERVICE ADVISOR SALES", "", "", "", "SUMMARY - " & cboServiceAdvisor & " " & cboMonth & " " & cboYear, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
    Else
        'JUN 01/05/2008
        RPT.WindowTitle = "Service Advisor Sales Report"
        RPT.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        RPT.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        RPT.Formulas(2) = "Printedby = '" & LOGNAME & "'"

        PrintSQLReport RPT, CSMS_REPORT_PATH & "ServiceAdvisorSales.rpt", "ISNULL({RO.DTE_COMP})=false and  Month({RO.DTE_COMP}) = " & What_month(cboMonth.Text) & " AND Year({RO.DTE_COMP}) = " & cboYear.Text & Filter, DMIS_REPORT_Connection, 1

        'LogAudit "V", "SERVICE ADVISOR SALES - REPORTS ", cboServiceAdvisor & cboMonth & cboYear
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "SERVICE ADVISOR SALES", "", "", "", "DETAIL - " & cboServiceAdvisor & " " & cboMonth & " " & cboYear, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SERVICE ADVISOR SALES)"
            Call frmALL_AuditInquiry.DisplayHistory("", "SERVICE ADVISOR SALES", "PRINTING")

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
    FillCombo
    cboServiceAdvisor.ListIndex = 0
End Sub

Sub FillCombo()
    Dim tmp_value                                      As String
    Dim rsServiceAdvisor                               As ADODB.Recordset
    tmp_value = ""
    Set rsServiceAdvisor = New ADODB.Recordset
    rsServiceAdvisor.Open "Select NAYM from CSMS_VW_EMPNO ORDER BY NAYM ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsServiceAdvisor.EOF And Not rsServiceAdvisor.BOF Then
        cboServiceAdvisor.Clear
        cboServiceAdvisor.AddItem "All"
        Do While Not rsServiceAdvisor.EOF
            cboServiceAdvisor.AddItem Null2String(rsServiceAdvisor!NAYM)
            rsServiceAdvisor.MoveNext
        Loop
    End If
    Set rsServiceAdvisor = Nothing
End Sub

