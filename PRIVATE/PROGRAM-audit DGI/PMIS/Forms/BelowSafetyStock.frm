VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISReports_PrintBelowSafetyStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Stocks Below SSL"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   ForeColor       =   &H00DEDFDE&
   Icon            =   "BelowSafetyStock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   3645
   Begin VB.ComboBox cboDate_Gen 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select date from list"
      Top             =   120
      Width           =   2205
   End
   Begin VB.CheckBox chkInclude 
      Caption         =   "Include Negative On Hand"
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
      Height          =   225
      Left            =   450
      TabIndex        =   1
      Top             =   1500
      Width           =   3075
   End
   Begin Crystal.CrystalReport rptPrintStkStat 
      Left            =   3120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Stock Below Safety Stock Level Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
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
      Left            =   1680
      MouseIcon       =   "BelowSafetyStock.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "BelowSafetyStock.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   600
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
      Left            =   960
      MouseIcon       =   "BelowSafetyStock.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "BelowSafetyStock.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "AS OF:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   330
      TabIndex        =   2
      Top             =   150
      Width           =   1035
   End
End
Attribute VB_Name = "frmPMISReports_PrintBelowSafetyStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSTKSTAT                                          As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    If Function_Access(LOGID, "Acess_Print", "REPORTS INTERNAL STOCKS BELOW SAFETY STOCK LEVEL") = False Then Exit Sub

    On Error GoTo Errorcode:

    Set rsSTKSTAT = New ADODB.Recordset
    rsSTKSTAT.Open "select * from PMIS_StkStat", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
        Screen.MousePointer = 11
        rptPrintStkStat.ReportTitle = "STOCKS BELOW SAFETY STOCK LEVEL"
        rptPrintStkStat.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptPrintStkStat.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptPrintStkStat, PMIS_REPORT_PATH & "belowssl.rpt", "{stkstat.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
        LogAudit "V", "Below Saftey Level"
    Else
        MsgSpeechBox "Not Yet Generated!"
    End If

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Set rsSTKSTAT = New ADODB.Recordset
    rsSTKSTAT.Open "select date_gen from PMIS_StkStat group by date_gen order by date_gen desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
        cboDate_Gen.Clear
        Do While Not rsSTKSTAT.EOF
            cboDate_Gen.AddItem Null2Date(rsSTKSTAT!DATE_GEN)
            rsSTKSTAT.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReports_PrintBelowSafetyStock = Nothing
    UnloadForm Me
End Sub

