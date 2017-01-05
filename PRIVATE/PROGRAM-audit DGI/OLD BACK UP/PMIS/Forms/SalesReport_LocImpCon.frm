VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISReports_SalesReport_Loc_Imp_Con 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Report (Loc,Imp,Con)"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2940
   ForeColor       =   &H00DEDFDE&
   Icon            =   "SalesReport_LocImpCon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   2940
   Begin VB.CheckBox chkHistIssuance 
      Caption         =   "Look in History File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   510
      TabIndex        =   7
      Top             =   900
      Width           =   2415
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   810
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
      Left            =   810
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select month from the list"
      Top             =   120
      Width           =   1965
   End
   Begin Crystal.CrystalReport rptOrderReport 
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
      Left            =   1530
      MouseIcon       =   "SalesReport_LocImpCon.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "SalesReport_LocImpCon.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   1260
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
      MouseIcon       =   "SalesReport_LocImpCon.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "SalesReport_LocImpCon.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Print Report"
      Top             =   1260
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
      Left            =   -30
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
      Left            =   -30
      TabIndex        =   2
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmPMISReports_SalesReport_Loc_Imp_Con"
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

    If chkHistIssuance.Value = False Then
        If What_month(cboMonth) >= Month(Now) Then
            Dim RSPO_HD                                As ADODB.Recordset
            Set RSPO_HD = New ADODB.Recordset
            RSPO_HD.Open "select * from PMIS_Tdaytran where TYPE = 'P' AND month(trandate) = " & What_month(cboMonth) & " AND year(trandate) =" & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not RSPO_HD.EOF And Not RSPO_HD.BOF Then
                Screen.MousePointer = 11
                rptOrderReport.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptOrderReport.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptOrderReport.Formulas(11) = "ForTheMonth = '" & cboMonth.Text + " " + cboYear.Text & "'"
                PrintSQLReport rptOrderReport, PMIS_REPORT_PATH & "SalesReport_LocImpCon.rpt", "{PMIS_Tdaytran.TYPE} = 'P' and month({PMIS_Tdaytran.TRANDATE}) = " & What_month(cboMonth.Text) & " AND year({PMIS_Tdaytran.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
                Screen.MousePointer = 0
            Else
                MsgSpeechBox "No Record for the Month of " & cboMonth.Text & " Year " & cboYear.Text
            End If
        Else
            MsgSpeechBox "No Record for the Month of " & cboMonth.Text & " Year " & cboYear.Text
        End If
    Else
        Dim RSPO_HIST                                  As ADODB.Recordset
        Set RSPO_HIST = New ADODB.Recordset
        RSPO_HIST.Open "select * from PMIS_Daytran where TYPE = 'P' AND month(trandate) = " & What_month(cboMonth) & " AND year(trandate) =" & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPO_HIST.EOF And Not RSPO_HIST.BOF Then
            Screen.MousePointer = 11
            rptOrderReport.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptOrderReport.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptOrderReport.Formulas(11) = "ForTheMonth = '" & cboMonth.Text + " " + cboYear.Text & "'"
            PrintSQLReport rptOrderReport, PMIS_REPORT_PATH & "SalesReport_LocImpCon_Hist.rpt", "{PMIS_Daytran.TYPE} = 'P' AND month({PMIS_Daytran.TRANDATE}) = " & What_month(cboMonth.Text) & " AND year({PMIS_Daytran.TRANDATE}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        Else
            MsgSpeechBox "No Record for the Month of " & cboMonth.Text & " Year " & cboYear.Text
        End If
    End If
    LogAudit "V", "SALES REPORT(LOCAL,IMPORTED,CONSIGNED)", cboMonth & "-" & cboYear
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
    Set frmPMISReports_SalesReport_Loc_Imp_Con = Nothing
    UnloadForm Me
End Sub

