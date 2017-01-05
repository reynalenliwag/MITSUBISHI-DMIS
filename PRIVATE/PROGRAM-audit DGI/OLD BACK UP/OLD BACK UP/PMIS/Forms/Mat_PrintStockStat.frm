VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmCSMSMat_PrintStockStat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Stock Status"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Mat_PrintStockStat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1545
   ScaleWidth      =   3645
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
      Height          =   750
      Left            =   1425
      MouseIcon       =   "Mat_PrintStockStat.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "Mat_PrintStockStat.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   735
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
      Height          =   750
      Left            =   2190
      MouseIcon       =   "Mat_PrintStockStat.frx":1433
      MousePointer    =   99  'Custom
      Picture         =   "Mat_PrintStockStat.frx":1585
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
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
      ToolTipText     =   "Select date from the list"
      Top             =   150
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
      Top             =   1530
      Width           =   3075
   End
   Begin Crystal.CrystalReport rptPrintStkStat 
      Left            =   3120
      Top             =   570
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Stock Status Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Top             =   180
      Width           =   1035
   End
End
Attribute VB_Name = "frmCSMSMat_PrintStockStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSTKSTAT            As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If IsDate(cboDate_Gen.Text) = True Then
        Set rsSTKSTAT = New ADODB.Recordset
        rsSTKSTAT.Open "select * from PMIS_StkStat where date_gen = '" & cboDate_Gen.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
            Screen.MousePointer = 11
            rptPrintStkStat.ReportTitle = "MATERIALS STOCK STATUS REPORT"
            rptPrintStkStat.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptPrintStkStat.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptPrintStkStat, CSMS_REPORT_PATH & "Mat_stockstatus.rpt", "{stkstat.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        Else
            MsgSpeechBox "Not Yet Generated!"
        End If
    Else
        MsgSpeechBox "Invalid Date Generated!"
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Set rsSTKSTAT = New ADODB.Recordset
    rsSTKSTAT.Open "select date_gen from PMIS_StkStat group by date_gen order by date_gen desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSTKSTAT.EOF And Not rsSTKSTAT.BOF Then
        cboDate_Gen.Clear
        Do While Not rsSTKSTAT.EOF
            cboDate_Gen.AddItem Null2Date(rsSTKSTAT!date_gen)
            rsSTKSTAT.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
End Sub
