VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_VSRep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Vehicle Sales"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   ForeColor       =   &H00FCFCFC&
   Icon            =   "ReportVSRep.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2385
   ScaleWidth      =   4470
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
      Height          =   795
      Left            =   2355
      MouseIcon       =   "ReportVSRep.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportVSRep.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Close Window"
      Top             =   1455
      Width           =   825
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
      Height          =   795
      Left            =   1545
      MouseIcon       =   "ReportVSRep.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportVSRep.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Print Report"
      Top             =   1455
      Width           =   825
   End
   Begin VB.ComboBox cboYEAR 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   450
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   870
      Width           =   2025
   End
   Begin VB.ComboBox cboMonth2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   345
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   510
      Width           =   2025
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   345
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   2025
   End
   Begin Crystal.CrystalReport rptGenREP 
      Left            =   30
      Top             =   1590
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Monthly Vehicle Sales Report"
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   915
      TabIndex        =   6
      Top             =   930
      Width           =   510
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1125
      TabIndex        =   5
      Top             =   540
      Width           =   300
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
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   825
      TabIndex        =   3
      Top             =   180
      Width           =   600
   End
End
Attribute VB_Name = "frmSMIS_Report_VSRep"
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

    If What_month(cboMonth) > What_month(cboMonth2) Then
        MsgSpeechBox "Error In From - To Months"
        Exit Sub
    End If
    Set rsPurchAgree = New ADODB.Recordset
    rsPurchAgree.Open "select * from SMIS_PurchAgree where month(datereleased) >= " & What_month(cboMonth) & " AND month(datereleased) <=" & What_month(cboMonth2), gconDMIS, adOpenForwardOnly, adLockReadOnly
    rptGenREP.WindowShowGroupTree = False
    If Not rsPurchAgree.EOF And Not rsPurchAgree.EOF Then
        Screen.MousePointer = 11
        If cboMonth.Text = cboMonth2.Text Then
            rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "vsrep.rpt", "year({purchagree.datereleased}) = " & cboYear.Text & " and month({purchagree.datereleased}) = " & What_month(cboMonth), DMIS_REPORT_Connection, 1
            
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 5:00
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call NEW_LogAudit("V", "VEHICLE SALES", "", "", "", "FROM " & cboMonth & " " & "TO " & cboMonth & " " & cboYear, "", "")
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Else
            rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "vsrep.rpt", "year({purchagree.datereleased}) = " & cboYear.Text & " and month({purchagree.datereleased}) >= " & What_month(cboMonth) & " AND month({purchagree.datereleased}) <= " & What_month(cboMonth2), DMIS_REPORT_Connection, 1
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 5:00
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Call NEW_LogAudit("V", "VEHICLE SALES", "", "", "", "FROM " & cboMonth & " " & "TO " & cboMonth2 & " " & cboYear, "", "")
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        End If
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record for the Month of " & cboMonth.Text
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (VEHICLE SALES)"
            Call frmALL_AuditInquiry.DisplayHistory("", "VEHICLE SALES", "PRINTING")
            
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
    Dim Last_YEAR                                                     As String
    Set rsPurchAgree = New ADODB.Recordset
    rsPurchAgree.Open "select distinct year(datereleased) as YEAR_RELEASED from SMIS_PurchAgree ", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPurchAgree.EOF And Not rsPurchAgree.BOF Then
        cboYear.Clear
        Do While Not rsPurchAgree.EOF
            cboYear.AddItem Null2String(rsPurchAgree!YEAR_RELEASED)
            Last_YEAR = Null2String(rsPurchAgree!YEAR_RELEASED)
            rsPurchAgree.MoveNext
        Loop
        cboYear.Text = Last_YEAR
    End If
    Screen.MousePointer = 0
End Sub

