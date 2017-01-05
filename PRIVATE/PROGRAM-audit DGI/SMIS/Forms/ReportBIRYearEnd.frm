VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_BIRYearEnd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BIR Year-End Report"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   ForeColor       =   &H00F5F5F5&
   Icon            =   "ReportBIRYearEnd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3645
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
      Height          =   765
      Left            =   1875
      MouseIcon       =   "ReportBIRYearEnd.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportBIRYearEnd.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   750
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
      Height          =   765
      Left            =   1095
      MouseIcon       =   "ReportBIRYearEnd.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportBIRYearEnd.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   750
      Width           =   795
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00F1F6F5&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   465
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   1965
   End
   Begin Crystal.CrystalReport rptGenREP 
      Left            =   240
      Top             =   1140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "BIR Year-End Report"
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
      BackStyle       =   0  'Transparent
      Caption         =   "YEAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   300
      TabIndex        =   2
      Top             =   240
      Width           =   1005
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   1
      Top             =   2940
      Width           =   495
   End
End
Attribute VB_Name = "frmSMIS_Report_BIRYearEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPurchAgree                                                      As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200712:41
Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "BIR YEAR REPORT") = False Then Exit Sub
    On Error GoTo ErrorCode
    Dim filter                                                        As String
    Set rsPurchAgree = New ADODB.Recordset
    rsPurchAgree.Open "select * from SMIS_MrrInv WHERE (year(datereleased) > " & cboYear.Text & " OR Released = 0)", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPurchAgree.EOF And Not rsPurchAgree.EOF Then
        Screen.MousePointer = 11
        rptGenREP.Reset
        rptGenREP.Formulas(0) = "YEAR_END = " & cboYear.Text
        filter = "((year({VEHICLE.datereceived}) <= " & cboYear.Text & " AND {VEHICLE.Released} = false) OR (year({VEHICLE.datereleased}) > " & cboYear.Text & " AND year({VEHICLE.datereceived}) <= " & cboYear.Text & " AND {VEHICLE.Released} = true))"
        rptGenREP.Formulas(0) = "CompanyName = '" & Company_name & "'"
        rptGenREP.Formulas(1) = "CompanyAddress = '" & Company_Address & "'"
        rptGenREP.WindowTitle = "BIR YEAR END REPORT"
        PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "BIR_YEAREND.rpt", filter, DMIS_REPORT_Connection, 1
        LogAudit "G", "BIR YEAR END", cboYear
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record for the Year of " & cboYear.Text
    End If
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
    FillcboYear cboYear
    Screen.MousePointer = 0
End Sub
