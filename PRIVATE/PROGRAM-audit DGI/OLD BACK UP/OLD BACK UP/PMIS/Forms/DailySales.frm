VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISReports_DailySales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Sales Report"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   ForeColor       =   &H00DEDFDE&
   Icon            =   "DailySales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2145
   ScaleWidth      =   3135
   Begin VB.CheckBox chkLookInHistory 
      Caption         =   "Look in History File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   930
      Width           =   1755
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
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select year from the list"
      Top             =   510
      Width           =   2055
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
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select month from the list"
      Top             =   120
      Width           =   2055
   End
   Begin Crystal.CrystalReport rptIssuances 
      Left            =   150
      Top             =   1050
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Daily Sales Report (As per Issuance)"
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
      MouseIcon       =   "DailySales.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "DailySales.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Close Window"
      Top             =   1260
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
      MouseIcon       =   "DailySales.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "DailySales.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Print Report"
      Top             =   1260
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
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Width           =   735
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2280
      TabIndex        =   4
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
      Left            =   120
      TabIndex        =   3
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmPMISReports_DailySales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsORD_HIST                                                        As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "DAILY SALES REPORT") = False Then Exit Sub

    On Error GoTo ERRORCODE:
    rptIssuances.Formulas(0) = "CompanyName='" & COMPANY_NAME & "'"
    rptIssuances.Formulas(1) = "CompanyAddress='" & COMPANY_ADDRESS & "'"
    If chkLookInHistory.Value = 1 Then
        Set rsORD_HIST = New ADODB.Recordset
        rsORD_HIST.Open "select trandate from PMIS_Ord_Hist where type = 'P' and month(TranDate) = " & What_month(cboMonth) & " AND year(TranDate) =" & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsORD_HIST.EOF And Not rsORD_HIST.EOF Then
            Screen.MousePointer = 11
            PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "dailysaleshist.rpt", "{Ord_hist.Type} = 'P' and month({Ord_hist.TranDate}) = " & What_month(cboMonth.Text) & " AND year({Ord_hist.TranDate}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
            NEW_LogAudit "V", "DAILY SALES REPORT", "", "", "", cboMonth & " " & cboYear & " HISTORY", "", ""
            Screen.MousePointer = 0
        Else
            MsgSpeechBox "No Record for the Month of " & cboMonth.Text & " Year " & cboYear.Text
        End If
    Else
        Set rsORD_HIST = New ADODB.Recordset
        rsORD_HIST.Open "select trandate from PMIS_Ord_Hd where type = 'P' and month(TranDate) = " & What_month(cboMonth) & " AND year(TranDate) =" & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsORD_HIST.EOF And Not rsORD_HIST.EOF Then
            Screen.MousePointer = 11
            PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "dailysales.rpt", "{Ord_hd.Type} = 'P' and month({Ord_hd.TranDate}) = " & What_month(cboMonth.Text) & " AND year({Ord_hd.TranDate}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
            NEW_LogAudit "V", "DAILY SALES REPORT", "", "", "", cboMonth & " " & cboYear, "", ""
            Screen.MousePointer = 0
        Else
            MsgSpeechBox "No Record for the Month of " & cboMonth.Text & " Year " & cboYear.Text
        End If
    End If

    Exit Sub

ERRORCODE:
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (DAILY SALES REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "DAILY SALES REPORT", "PRINTING")
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
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReports_Issuances = Nothing
    UnloadForm Me
End Sub

