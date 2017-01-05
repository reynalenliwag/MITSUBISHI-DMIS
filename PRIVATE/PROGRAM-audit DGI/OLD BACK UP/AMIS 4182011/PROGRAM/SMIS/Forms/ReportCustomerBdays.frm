VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_CustomerBDays 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Birthday Celebrant of the Month"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   ForeColor       =   &H00FCFCFC&
   Icon            =   "ReportCustomerBdays.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1650
   ScaleWidth      =   4545
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
      Left            =   3615
      MouseIcon       =   "ReportCustomerBdays.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportCustomerBdays.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Close Window"
      Top             =   660
      Width           =   795
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
      Left            =   2835
      MouseIcon       =   "ReportCustomerBdays.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportCustomerBdays.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   660
      Width           =   795
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   465
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   3105
   End
   Begin Crystal.CrystalReport rptCelebrants 
      Left            =   1950
      Top             =   1020
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Customers Birthday Celebrants of the Month"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   1185
   End
End
Attribute VB_Name = "frmSMIS_Report_CustomerBDays"
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

    On Error GoTo ErrorCode:

    Set rsPurchAgree = New ADODB.Recordset
    rsPurchAgree.Open "select * from SMIS_PurchAgree where Month(Birthdate) = " & What_month(cboMonth), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPurchAgree.EOF And Not rsPurchAgree.BOF Then
        rptCelebrants.Reset
        rptCelebrants.Formulas(0) = "YEER = " & Year(LOGDATE)
        rptCelebrants.Formulas(1) = "CURRENT_DAY = DATE(" & Year(LOGDATE) & "," & Month(LOGDATE) & "," & Day(LOGDATE) & ")"
        '   rptCelebrants.Formulas(2) = "CompanyName " & SetCompName("1")
        '  rptCelebrants.Formulas(3) = "CompanyAddress " & SetCompAdress("1")
        rptCelebrants.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptCelebrants.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptCelebrants, SMIS_REPORT_PATH & "CustomerBday.rpt", "Month({Purchagree.BirthDate}) = " & What_month(cboMonth), DMIS_REPORT_Connection, 1
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "BIRTHDAY CELEBRANTS", "", "", "", cboMonth, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

        'LogAudit "V", "BIRTHDAY CELEBRANT REPORT", "FOR THE " & cboMonth
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Celebrants for the month of " & cboMonth.Text
        Exit Sub
    End If





    Exit Sub
ErrorCode:
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (BIRTHDAY CELEBRANTS)"
            Call frmALL_AuditInquiry.DisplayHistory("", "BIRTHDAY CELEBRANTS", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    cboMonth.Text = The_month(Month(LOGDATE))
    Screen.MousePointer = 0
End Sub

