VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_ListReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Registration"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   ForeColor       =   &H00FCFCFC&
   Icon            =   "ReportListReg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1770
   ScaleWidth      =   4770
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
      Left            =   2550
      MouseIcon       =   "ReportListReg.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportListReg.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Close Window"
      Top             =   870
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
      Left            =   1680
      MouseIcon       =   "ReportListReg.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportListReg.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Print Report"
      Top             =   870
      Width           =   885
   End
   Begin VB.TextBox txtYear 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   555
      Left            =   3690
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "9999"
      Top             =   150
      Width           =   945
   End
   Begin VB.ComboBox cboMonth2 
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
      Left            =   660
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   450
      Width           =   1965
   End
   Begin VB.ComboBox cboMonth 
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
      Left            =   660
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   1965
   End
   Begin Crystal.CrystalReport rptGenREP 
      Left            =   30
      Top             =   990
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "List of Registrations"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   2820
      TabIndex        =   6
      Top             =   180
      Width           =   825
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
      Left            =   -180
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   4
      Top             =   2940
      Width           =   495
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
      Left            =   -180
      TabIndex        =   3
      Top             =   90
      Width           =   735
   End
End
Attribute VB_Name = "frmSMIS_Report_ListReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPurchAgree                                                      As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode

    If What_month(cboMonth) > What_month(cboMonth2) Then
        MsgSpeechBox "Error In From - To Months"
        Exit Sub
    End If
    Set rsPurchAgree = New ADODB.Recordset
    rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE month(datereleased) >= " & What_month(cboMonth) & " AND month(datereleased) <=" & What_month(cboMonth2) & " and year(datereleased) = " & txtYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPurchAgree.EOF And Not rsPurchAgree.EOF Then
        Screen.MousePointer = 11

        rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

        PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "unitsreg.rpt", "month({purchagree.datereleased}) >= " & What_month(cboMonth) & " AND month({purchagree.datereleased}) <= " & What_month(cboMonth2) & " and year({purchagree.datereleased}) = " & txtYear.Text, DMIS_REPORT_Connection, 1
        
        'UPDATED BY: JUN
        'DATE UPDATED: 09032008 5:00
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Call NEW_LogAudit("V", "LIST OF UNITS REGISTERED", "", "", "", "FROM " & cboMonth & " " & "TO " & cboMonth2 & "" & txtYear, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Screen.MousePointer = 0
        'LogAudit "V", "LIST OF UNIT REGISTERED", cboMonth & " " & cboMonth2 & " " & txtYear
    Else
        MsgSpeechBox "No Record for the Month of " & cboMonth.Text
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (LIST OF UNITS REGISTERED)"
            Call frmALL_AuditInquiry.DisplayHistory("", "LIST OF UNITS REGISTERED", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    fillcbomonth cboMonth2
    cboMonth.Text = The_month(Month(LOGDATE))
    cboMonth2.Text = The_month(Month(LOGDATE))
    txtYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub

