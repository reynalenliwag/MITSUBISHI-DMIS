VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISReports_History_DNP_SRP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "History of Price of Parts"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3090
   ForeColor       =   &H00DEDFDE&
   Icon            =   "HistoryOfPrice_DNPandSRP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3090
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
      Left            =   870
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select year from the list"
      Top             =   240
      Width           =   1965
   End
   Begin Crystal.CrystalReport rptIssuances 
      Left            =   2550
      Top             =   990
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
      Left            =   1530
      MouseIcon       =   "HistoryOfPrice_DNPandSRP.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "HistoryOfPrice_DNPandSRP.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   690
      Width           =   735
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
      Left            =   810
      MouseIcon       =   "HistoryOfPrice_DNPandSRP.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "HistoryOfPrice_DNPandSRP.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   690
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
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
      Top             =   240
      Width           =   735
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   1
      Top             =   3000
      Width           =   495
   End
End
Attribute VB_Name = "frmPMISReports_History_DNP_SRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo ErrorCode:
    Screen.MousePointer = 11
    Dim rsLost_Sales                                                  As ADODB.Recordset
    Set rsLost_Sales = New ADODB.Recordset
    rsLost_Sales.Open "select date_gen from PMIS_StkStat where TYPE = 'P' AND year(date_gen) = " & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
    'rsLost_Sales.Open "SELECT TRANDATE FROM PMIS_ALLDAYTRAN WHERE TYPE = 'P' AND TRANTYPE IN('CSH','CHG','DR','RIV') AND (STATUS = 'P' OR STATUS = 'B') AND YEAR(TRANDATE)= " & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
        
    If Not rsLost_Sales.EOF And Not rsLost_Sales.EOF Then
        Screen.MousePointer = 11
        rptIssuances.WindowTitle = "History of Price of Parts"
        rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptIssuances.Formulas(11) = "monthprint = 'For the Year  " & cboYear.Text & "'"
        'PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "StockReport_HistoryPrice_DNP_SRP.rpt", "Month({PMIS_StkStat.Date_Gen}) = " & What_month(cboMonth) & " AND Year({PMIS_StkStat.Date_Gen}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "StockReport_HistoryPrice_DNP_SRP.rpt", "Year({StkStat.Date_Gen}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record for the Year of " & cboYear.Text
        Screen.MousePointer = 0
    End If
    LogAudit "V", "HISTORY OF PRICE (COST/SRP)", cboYear
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    'fillcbomonth cboMonth
    FillcboYear cboYear
    'cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReports_History_DNP_SRP = Nothing
    UnloadForm Me
End Sub

