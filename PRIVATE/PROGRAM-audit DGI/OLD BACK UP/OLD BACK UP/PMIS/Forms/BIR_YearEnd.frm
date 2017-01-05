VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISReports_BIR_YearEnd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BIR Year-End Report"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3210
   ForeColor       =   &H00DEDFDE&
   Icon            =   "BIR_YearEnd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   3210
   Begin VB.OptionButton Option3 
      Caption         =   "Materials"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   780
      TabIndex        =   6
      Top             =   810
      Width           =   1995
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Accessories"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   780
      TabIndex        =   5
      Top             =   450
      Width           =   1995
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Parts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   780
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   1995
   End
   Begin VB.ComboBox cboYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1410
      Width           =   1185
   End
   Begin Crystal.CrystalReport rptPrintStkStat 
      Left            =   2700
      Top             =   2310
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "BIR Year End Report"
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
      Left            =   1440
      MouseIcon       =   "BIR_YearEnd.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "BIR_YearEnd.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Close Window"
      Top             =   1920
      Width           =   735
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
      Left            =   720
      MouseIcon       =   "BIR_YearEnd.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "BIR_YearEnd.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Selected Report"
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "YEAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   510
      TabIndex        =   1
      Top             =   1440
      Width           =   885
   End
End
Attribute VB_Name = "frmPMISReports_BIR_YearEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If BIR_YearEnd = "PARTS" Then
        If Function_Access(LOGID, "Acess_Print", "REPORT GOV BIR YEAR REPORT") = False Then Exit Sub
    Else
        If Function_Access(LOGID, "Acess_Print", "MATERIAL INVENTORY MONTHLY BIR YEAR REPORT") = False Then Exit Sub
    End If
    On Error GoTo ERRORCODE:

    Dim rsSTKSTAT                                                     As ADODB.Recordset
    If BIR_YearEnd = "PARTS" Then
        Set rsSTKSTAT = New ADODB.Recordset
        rsSTKSTAT.Open "select * from PMIS_StkStat where month(date_gen) = 12 and year(date_gen) = " & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        Set rsSTKSTAT = New ADODB.Recordset
        rsSTKSTAT.Open "select * from PMIS_StkStat where month(date_gen) = 12 and year(date_gen) = " & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
        Screen.MousePointer = 11
        If Option1.Value = True Then BIR_YearEnd = "PARTS"
        If Option2.Value = True Then BIR_YearEnd = "ACCESSORIES"
        If Option3.Value = True Then BIR_YearEnd = "MATERIALS"
        If BIR_YearEnd = "PARTS" Then
            rptPrintStkStat.ReportTitle = cboYear.Text & " Parts BIR Year-End Report"
            rptPrintStkStat.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptPrintStkStat.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptPrintStkStat, PMIS_REPORT_PATH & "BIR_YearEnd.rpt", "{STKSTAT.TYPE} = 'P' AND YEAR({STKSTAT.DATE_GEN}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        End If
        If BIR_YearEnd = "ACCESSORIES" Then
            rptPrintStkStat.ReportTitle = cboYear.Text & " Accessories BIR Year-End Report"
            rptPrintStkStat.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptPrintStkStat.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptPrintStkStat, PMIS_REPORT_PATH & "BIR_YearEnd.rpt", "{STKSTAT.TYPE} = 'A' AND YEAR({STKSTAT.DATE_GEN}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        End If
        If BIR_YearEnd = "MATERIALS" Then
            rptPrintStkStat.ReportTitle = cboYear.Text & " Materials BIR Year-End Report"
            rptPrintStkStat.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptPrintStkStat.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptPrintStkStat, PMIS_REPORT_PATH & "BIR_YearEnd.rpt", "{STKSTAT.TYPE} = 'M' AND YEAR({STKSTAT.DATE_GEN}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        End If
        LogAudit "V", "REPORT BIR YEAR END"
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "Not Yet Generated!"
    End If

    Exit Sub
ERRORCODE:
    ShowVBError

End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    FillcboYear cboYear
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

