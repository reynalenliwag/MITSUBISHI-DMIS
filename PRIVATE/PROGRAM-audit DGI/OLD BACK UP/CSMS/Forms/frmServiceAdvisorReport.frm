VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMSServiceAdvisorReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SA Performance Report"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServiceAdvisorReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2160
   ScaleWidth      =   4185
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
      Left            =   1590
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Select year from the list"
      Top             =   870
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
      Left            =   1590
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select month from the list"
      Top             =   480
      Width           =   1965
   End
   Begin VB.ComboBox cboServiceAdvisor 
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
      Left            =   1590
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   2505
   End
   Begin Crystal.CrystalReport rptService_Advisor 
      Left            =   180
      Top             =   1530
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Service Advisor Performance Report"
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   3300
      MouseIcon       =   "frmServiceAdvisorReport.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmServiceAdvisorReport.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   2520
      MouseIcon       =   "frmServiceAdvisorReport.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "frmServiceAdvisorReport.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   1095
      TabIndex        =   7
      Top             =   930
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   960
      TabIndex        =   6
      Top             =   540
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Service Advisor"
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   150
      Width           =   1320
   End
End
Attribute VB_Name = "frmCSMSServiceAdvisorReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboServiceAdvisor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdPrint_Click
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SERVICE ADVISOR REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "SERVICE ADVISOR REPORT", "PRINTING")

    End Select
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SERVICE ADVISOR REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "SERVICE ADVISOR REPORT", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 0
    fillcbomonth cboMonth
    FillCboMoreYear cboYear
    cboMonth.Text = MonthName(Month(Date))

    FillCombo
    cboServiceAdvisor.AddItem "All"
    cboServiceAdvisor.Text = "All"
End Sub

Sub cmdPrint_Click()
    'If Function_Access(LOGID, "Acess_PRINT", "SERVICE ADVISOR REPORT") = False Then Exit Sub
    On Error GoTo Errorcode

    Dim rsService_Advisor                              As ADODB.Recordset
    Set rsService_Advisor = New ADODB.Recordset
    Set rsService_Advisor = gconDMIS.Execute("Select * from CSMS_vw_EMPNO")

    'JUN 02/05/2008
    rptService_Advisor.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptService_Advisor.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptService_Advisor.Formulas(2) = "Printedby = '" & LOGNAME & "'"

    If cboServiceAdvisor.Text = "All" Then
        If Not rsService_Advisor.EOF And Not rsService_Advisor.BOF Then
            Screen.MousePointer = 11
            'PrintSQLReport rptService_Advisor, CSMS_REPORT_PATH & "Service_Advisor_Performance.rpt", "MONTH({CSMS_REPOR.DTE_COMP}) = " & What_month(cboMonth) & " AND YEAR({CSMS_REPOR.DTE_COMP}) = " & cboYEAR & "", CSMS_REPORT_CONNECTION, 1
            PrintSQLReport rptService_Advisor, CSMS_REPORT_PATH & "Service_Advisor_Performance.rpt", "MONTH({REPOR.DTE_COMP}) = " & What_month(cboMonth) & " AND YEAR({REPOR.DTE_COMP}) = " & cboYear & "", CSMS_REPORT_CONNECTION, 1

            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "SERVICE ADVISOR REPORT", "", "", "", "SA NAME: " & cboServiceAdvisor, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
            'LogAudit "V", "SERVICE ADVISOR PERFORMANCE - REPORTS ", cboServiceAdvisor
            Screen.MousePointer = 0
        Else
            ShowNoRecord
            On Error Resume Next
            cboServiceAdvisor.SetFocus
            Exit Sub
        End If
    Else
        If Not rsService_Advisor.EOF And Not rsService_Advisor.BOF Then
            Screen.MousePointer = 11
            PrintSQLReport rptService_Advisor, CSMS_REPORT_PATH & "Service_Advisor_Performance.rpt", "MONTH({REPOR.DTE_COMP}) = " & What_month(cboMonth) & " AND YEAR({REPOR.DTE_COMP}) = " & cboYear & " And {REPOR.RECD_BY} = '" & ReturnSAcode(cboServiceAdvisor) & "'", CSMS_REPORT_CONNECTION, 1

            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "SERVICE ADVISOR REPORT", "", "", "", "SA NAME: " & cboServiceAdvisor, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
            'LogAudit "V", "SERVICE ADVISOR PERFORMANCE - REPORTS ", cboServiceAdvisor
            Screen.MousePointer = 0
        Else
            ShowNoRecord
            On Error Resume Next
            cboServiceAdvisor.SetFocus
            Exit Sub
        End If
    End If
    Exit Sub

Errorcode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Function ReturnSAcode(XXX As String) As String
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT NAYM, CODE FROM CSMS_VW_EMPNO WHERE NAYM = '" & XXX & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        ReturnSAcode = Null2String(rstmp!Code)
    End If
    Set rstmp = Nothing
End Function

Sub FillCombo()
    Dim tmp_value                                      As String
    Dim rsServiceAdvisor                               As ADODB.Recordset
    tmp_value = ""
    Set rsServiceAdvisor = New ADODB.Recordset
    rsServiceAdvisor.Open "Select NAYM from CSMS_vw_EMPNO order by NAYM asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsServiceAdvisor.EOF And Not rsServiceAdvisor.BOF Then
        rsServiceAdvisor.MoveFirst
        cboServiceAdvisor.Clear
        cboServiceAdvisor.AddItem "All"
        Do While Not rsServiceAdvisor.EOF
            If tmp_value = Null2String(rsServiceAdvisor!NAYM) Then
                rsServiceAdvisor.MoveNext
            Else
                cboServiceAdvisor.AddItem Null2String(rsServiceAdvisor!NAYM)
                tmp_value = Null2String(rsServiceAdvisor!NAYM)
                rsServiceAdvisor.MoveNext
            End If
        Loop
    End If
    Set rsServiceAdvisor = Nothing
End Sub



