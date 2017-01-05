VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmAMISMonthlyYearly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select a Year"
   ClientHeight    =   1800
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   2865
   ForeColor       =   &H00F5F5F5&
   Icon            =   "MonthlyYearly.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   2865
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
      Height          =   825
      Left            =   1500
      MouseIcon       =   "MonthlyYearly.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "MonthlyYearly.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   900
      Width           =   885
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
      Height          =   825
      Left            =   630
      MouseIcon       =   "MonthlyYearly.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "MonthlyYearly.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   900
      Width           =   885
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
      ForeColor       =   &H00973640&
      Height          =   330
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   1965
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
      ForeColor       =   &H00973640&
      Height          =   330
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   450
      Width           =   1965
   End
   Begin Crystal.CrystalReport rptMonthlyYearly 
      Left            =   30
      Top             =   930
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
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
   Begin VB.Label Label1 
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
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   -90
      TabIndex        =   0
      Top             =   90
      Width           =   735
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   -90
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   6
      Top             =   2940
      Width           =   495
   End
End
Attribute VB_Name = "frmAMISMonthlyYearly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsProfile                                          As ADODB.Recordset

Sub FillcboYear2()
    FillcboYear cboYear
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()


    On Error GoTo ErrorCode
    Dim Deyt_To                                        As String
    rptMonthlyYearly.Reset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptMonthlyYearly.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptMonthlyYearly.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
    End If
    Deyt_To = lastDay(CDate(What_month(cboMonth.Text) & "/1/" & cboYear.Text))
    rptMonthlyYearly.WindowTitle = "Schedule of Depreciation for the month of : " & cboMonth.Text & " year: " & cboYear.Text
    rptMonthlyYearly.ReportTitle = "Schedule of Depreciation for the month of : " & cboMonth.Text & " year: " & cboYear.Text
    rptMonthlyYearly.Formulas(0) = "ToJdate = Date(" & Year(Deyt_To) & "," & Month(Deyt_To) & "," & Day(Deyt_To) & ")"
    PrintSQLReport rptMonthlyYearly, AMIS_REPORT_PATH & "Files\Assets.Rpt", "", DMIS_REPORT_Connection, 1
    Exit Sub

ErrorCode:
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
        frmALL_AuditInquiry.Caption = "ASSETS REPORT"
        Call frmALL_AuditInquiry.DisplayHistory("", "ASSETS REPORT")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    FillcboYear2
    fillcbomonth cboMonth
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub

