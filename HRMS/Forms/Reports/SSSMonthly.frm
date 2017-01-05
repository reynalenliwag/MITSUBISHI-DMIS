VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmHRMSSSSMonthly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SSS Monthly Remittance"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4710
   ForeColor       =   &H00D8E9EC&
   Icon            =   "SSSMonthly.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4710
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
      Left            =   2340
      MouseIcon       =   "SSSMonthly.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "SSSMonthly.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   825
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
      Left            =   1470
      MouseIcon       =   "SSSMonthly.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "SSSMonthly.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   825
      Width           =   885
   End
   Begin VB.ComboBox cboYear 
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
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   3420
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   225
      Width           =   1245
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
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   225
      Width           =   2355
   End
   Begin Crystal.CrystalReport rptSSSMonthly 
      Left            =   3330
      Top             =   795
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   2550
      TabIndex        =   2
      Top             =   255
      Width           =   825
   End
End
Attribute VB_Name = "frmHRMSSSSMonthly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:
    'If Function_Access(LOGID, "Acess_Print", "REPORT SSS MONTHLY REMITTANCE") = False Then Exit Sub
    Screen.MousePointer = 11
    rptSSSMonthly.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptSSSMonthly.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptSSSMonthly.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    rptSSSMonthly.Formulas(3) = "PREPARED_BY = '" & PREPARED_BY & "'"
    rptSSSMonthly.Formulas(4) = "CHECKED_BY = '" & CHECKED_BY & "'"
    rptSSSMonthly.Formulas(5) = "APPROVED_BY = '" & APPROVED_BY & "'"
    rptSSSMonthly.Formulas(6) = "PrintedBy = '" & LOGNAME & "'"
    rptSSSMonthly.Formulas(7) = "YEER = '" & cboYear & "'"
    rptSSSMonthly.Formulas(8) = "month = '" & cboMonth & "'"
    PrintSQLReport rptSSSMonthly, HRMS_REPORT_PATH & "SSSremit.rpt", "{sssdet.PAY_YEAR} = " & cboYear.Text & " and {sssdet.Pay_month} = " & What_month(cboMonth) & " and ({sssdet.Cut_off} = '1' OR {sssdet.Cut_off} = '2' )", DMIS_REPORT_Connection, 1
    Call LogAudit("V", "PRINT SSS MONTHLY REPORT", cboMonth & "-" & cboYear)
    Screen.MousePointer = 0
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
    cboMonth.Text = The_month(Month(LOGDATE))
    FillcboYear cboYear
    cboYear.Text = Year(LOGDATE)
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub
