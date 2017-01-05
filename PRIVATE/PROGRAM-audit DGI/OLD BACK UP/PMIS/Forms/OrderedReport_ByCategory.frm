VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISReports_OrderedReport_ByCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordered Report by Category"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3090
   ForeColor       =   &H00DEDFDE&
   Icon            =   "OrderedReport_ByCategory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1890
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
      Left            =   900
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
      Width           =   1965
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
      Left            =   1680
      MouseIcon       =   "OrderedReport_ByCategory.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "OrderedReport_ByCategory.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   960
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
      Left            =   960
      MouseIcon       =   "OrderedReport_ByCategory.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "OrderedReport_ByCategory.frx":1531
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
      Top             =   540
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
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmPMISReports_OrderedReport_ByCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo Errorcode:

    Dim rsOR_Cat                                       As ADODB.Recordset
    Set rsOR_Cat = New ADODB.Recordset
    rsOR_Cat.Open "select * from PMIS_AllDayTran where TYPE = 'P' AND TRANTYPE = 'PO' AND month(TRANDATE) = " & What_month(cboMonth) & " AND year(TRANDATE) = " & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly

    If Not rsOR_Cat.EOF And Not rsOR_Cat.EOF Then
        Screen.MousePointer = 11
        rptIssuances.WindowTitle = "Parts Ordered Report by Category"
        rptIssuances.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptIssuances.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptIssuances.Formulas(11) = "forthemonth = '" & cboMonth.Text + " " + cboYear.Text & "'"
        PrintSQLReport rptIssuances, PMIS_REPORT_PATH & "Ordered_Parts_Group_By_Category.rpt", "Month({PMIS_AllDayTran.TRANDATE}) = " & What_month(cboMonth) & " AND Year({PMIS_AllDayTran.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record for the Month of " & cboMonth.Text & " Year " & cboYear.Text
    End If
    LogAudit "V", "ORDERED REPORT BY CATEGORY", cboMonth & "-" & cboYear
    Exit Sub
Errorcode:
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
    FillCboMoreYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReports_OrderedReport_ByCategory = Nothing
    UnloadForm Me
End Sub

