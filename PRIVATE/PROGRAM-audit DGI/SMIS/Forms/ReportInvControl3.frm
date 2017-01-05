VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_InvControl3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Inventory Control"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   ForeColor       =   &H00FCFCFC&
   Icon            =   "ReportInvControl3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4470
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
      Left            =   2190
      MouseIcon       =   "ReportInvControl3.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportInvControl3.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   675
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
      Left            =   1320
      MouseIcon       =   "ReportInvControl3.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportInvControl3.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   675
      Width           =   885
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
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2355
   End
   Begin Crystal.CrystalReport rptReleased 
      Left            =   3210
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "DMC Monthly Inventory Control"
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
      Left            =   3420
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "9999"
      Top             =   90
      Width           =   1005
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
      Left            =   2610
      TabIndex        =   2
      Top             =   150
      Width           =   825
   End
End
Attribute VB_Name = "frmSMIS_Report_InvControl3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMRRINV                                                          As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo ErrorCode:

    If Len(txtYear.Text) = 4 Or txtYear.Text <> "" Then
        Set rsMRRINV = New ADODB.Recordset
        rsMRRINV.Open "select * from SMIS_MrrInv WHERE year(datereceived) = '" & txtYear.Text & "' and month(datereceived) <= " & What_month(cboMonth), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
            Screen.MousePointer = 11
            rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptReleased, SMIS_REPORT_PATH & "invcontrol3.rpt", "(year({VEHICLE.datereceived}) = " & txtYear.Text & " AND month({VEHICLE.datereceived}) <= " & What_month(cboMonth) & " AND {VEHICLE.Released} = false) OR (year({VEHICLE.datereceived}) = " & txtYear.Text & " AND month({VEHICLE.datereceived}) <= " & What_month(cboMonth) & " AND {VEHICLE.Released} = true AND month({VEHICLE.datereleased}) > " & What_month(cboMonth) & ")", DMIS_REPORT_Connection, 1
            LogAudit "V", "MONTHLY INVENTORY CONTROL", cboMonth & "/" & txtYear
            Screen.MousePointer = 0
        Else
            MsgSpeechBox "No Record for " & cboMonth.Text & " " & txtYear.Text
            Exit Sub
        End If
    End If





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    cboMonth.Text = The_month(Month(LOGDATE))
    txtYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub

