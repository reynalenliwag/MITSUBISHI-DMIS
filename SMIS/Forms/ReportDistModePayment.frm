VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_DistModePayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales by Mode of Payment"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   ForeColor       =   &H00FCFCFC&
   Icon            =   "ReportDistModePayment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1800
   ScaleWidth      =   3870
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
      Left            =   2790
      MouseIcon       =   "ReportDistModePayment.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportDistModePayment.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Close Window"
      Top             =   900
      Width           =   765
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
      Left            =   2040
      MouseIcon       =   "ReportDistModePayment.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportDistModePayment.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print Report"
      Top             =   900
      Width           =   765
   End
   Begin VB.ComboBox cboYear2 
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
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2505
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
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   2505
   End
   Begin Crystal.CrystalReport rptGenREP 
      Left            =   30
      Top             =   900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Distribution of Sales by Mode of Financing"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   0   'False
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   240
      TabIndex        =   4
      Top             =   540
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   2
      Top             =   2940
      Width           =   495
   End
End
Attribute VB_Name = "frmSMIS_Report_DistModePayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPurchAgree                                                      As ADODB.Recordset

Sub FillcboYear2()
    FillCboMoreYear cboYear
    FillCboMoreYear cboYear2
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo ErrorCode
    Set rsPurchAgree = New ADODB.Recordset
    rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE year(datereleased) >= " & cboYear.Text & " AND year(datereleased) <=" & cboYear2.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPurchAgree.EOF And Not rsPurchAgree.EOF Then
        Screen.MousePointer = 11
        rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "ModOfFinancing.rpt", "year({PURCHAGREE.datereleased}) >= " & cboYear.Text & " AND year({PURCHAGREE.datereleased}) <= " & cboYear2.Text, DMIS_REPORT_Connection, 1
        
        'UPDATED BY: JUN
        'DATE UPDATED: 09032008 5:00
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "SALES DISTRIBUTION OF SALES AS TO MODE OF PAYMENT", "", "", "", "FROM " & cboYear & " " & "To " & cboYear2, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

        'LogAudit "V", "SALES BY MODE OF FINANCING ", "FOR " & cboYear & " / " & cboYear2
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record for the year " & cboYear.Text
    End If
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SALES DISTRIBUTION OF SALES AS TO MODE OF PAYMENT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "SALES DISTRIBUTION OF SALES AS TO MODE OF PAYMENT", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    FillCboMoreYear cboYear2
    FillCboMoreYear cboYear
'    cboYear.Text = Year(LOGDATE)
'    cboYear2.Text = Year(LOGDATE)
    cboYear.Text = Format(LOGDATE, "yyyy")
    cboYear.Text = Format(LOGDATE, "yyyy")
    Screen.MousePointer = 0
End Sub

