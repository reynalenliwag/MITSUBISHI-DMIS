VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISReports_Issuances_MAT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Issuance"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MAT_Issuances.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1830
   ScaleWidth      =   3135
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
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select year from the list"
      Top             =   510
      Width           =   2115
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
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select month from the list"
      Top             =   120
      Width           =   2115
   End
   Begin Crystal.CrystalReport rptIssuances 
      Left            =   30
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
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
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
      Height          =   795
      Left            =   2340
      MouseIcon       =   "MAT_Issuances.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "MAT_Issuances.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   960
      Width           =   675
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
      Height          =   795
      Left            =   1680
      MouseIcon       =   "MAT_Issuances.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "MAT_Issuances.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Print Report"
      Top             =   960
      Width           =   675
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
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
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
      Left            =   60
      TabIndex        =   2
      Top             =   210
      Width           =   735
   End
End
Attribute VB_Name = "frmPMISReports_Issuances_MAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSORD_HIST                                                        As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:

    Dim RIV_Filter1                                                   As String
    Dim RO_Filter1                                                    As String
    Dim RIV_Filter2                                                   As String
    Dim RO_Filter2                                                    As String
    Dim LastDate                                                      As String
    Set RSORD_HIST = New ADODB.Recordset

    RSORD_HIST.Open "select trandate from PMIS_Ord_Hist where TYPE = 'M' AND month(TranDate) = " & What_month(cboMonth) & " AND year(TranDate) = " & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSORD_HIST.EOF And Not RSORD_HIST.EOF Then
        Screen.MousePointer = 11
        If ISSREPTYPE = "RIV_INPROCESS" Then
            LastDate = lastDay(CDate(cboMonth.Text & "/1/" & cboYear.Text))
            RIV_Filter1 = "{ORD_HD.TYPE} = 'M' AND (({ORD_HD.TRANDATE} <= DATESERIAL(" & cboYear.Text & "," & What_month(cboMonth.Text) & "," & Day(LastDate) & "))"
            RO_Filter1 = "({repor.DTE_REL} > DATESERIAL(" & cboYear.Text & "," & What_month(cboMonth.Text) & "," & Day(LastDate) & ")))"

            RIV_Filter2 = "{ORD_HD.TYPE} = 'M' AND (({ORD_HD.TRANDATE} <= DATESERIAL(" & cboYear.Text & "," & What_month(cboMonth.Text) & "," & Day(LastDate) & "))"
            RO_Filter2 = "ISNULL({repor.DTE_REL}) = TRUE)"
            rptIssuances.WindowTitle = "RIV IN-PROCESS"
            rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "RIV_InProcess.rpt", RIV_Filter1 & " And " & RO_Filter1, DMIS_REPORT_Connection, 1
            PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "RIV_InProcessUnReleased.rpt", RIV_Filter2 & " And " & RO_Filter2, DMIS_REPORT_Connection, 1

            Call NEW_LogAudit("V", "MATERIALS MONTHLY REPORT", "", "", "", cboMonth & " " & cboYear, "", "")
        Else
            If Function_Access(LOGID, "Acess_Print", "MATERIALS MONTHLY REPORT") = False Then Exit Sub
            rptIssuances.WindowTitle = "MONTHLY ISSUANCE"
            rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "Issuances.rpt", "{ORD_HD.TYPE} = 'M' AND month({Ord_hd.TranDate}) = " & What_month(cboMonth.Text) & " AND year({Ord_hd.TranDate}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
            PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "Issuancesum.rpt", "{ORD_HD.TYPE} = 'M' AND month({Ord_hd.TranDate}) = " & What_month(cboMonth.Text) & " AND year({Ord_hd.TranDate}) = " & cboYear.Text, DMIS_REPORT_Connection, 1

            Call NEW_LogAudit("V", "MATERIALS MONTHLY REPORT", "", "", "", cboMonth & " " & cboYear, "", "")
        End If
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record for the Month of " & cboMonth.Text & " Year " & cboYear.Text
    End If

    Exit Sub
Errorcode:
    ShowVBError

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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (MATERIALS MONTHLY REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "MATERIALS MONTHLY REPORT", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    FillcboYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    If ISSREPTYPE = "RIV_INPROCESS" Then
        Me.Caption = "RIV In-Process"
    Else
        Me.Caption = "Issuance Report"
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReports_Issuances = Nothing
    UnloadForm Me
End Sub

