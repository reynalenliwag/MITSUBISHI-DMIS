VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_MonthlySalesAnalysis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Sales Analysis "
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Report_MonthlySalesAnalysis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4575
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
      Left            =   2295
      MouseIcon       =   "Report_MonthlySalesAnalysis.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "Report_MonthlySalesAnalysis.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   750
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
      Left            =   1425
      MouseIcon       =   "Report_MonthlySalesAnalysis.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "Report_MonthlySalesAnalysis.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   750
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
      Left            =   3540
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "MMPC Monthly Inventory Control"
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
      Left            =   3450
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "9999"
      Top             =   90
      Width           =   1005
   End
   Begin VB.Label Label1 
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
      TabIndex        =   1
      Top             =   150
      Width           =   825
   End
End
Attribute VB_Name = "frmSMIS_Report_MonthlySalesAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMRRINV                                                          As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'==========================================================================================
'FUNCTION / FEATURE :cmdPrint_Click:
'DATE STARTED       :5/28/200715:40
'LAST UPDATED       :5/28/200715:40
'DATABASE UPDATES   :
'WHO UPDATED        :AXP  5/28/2007
'UDPATING COCODE    :AXP-5282007
'==========================================================================================

'Upating Code       : AXP-0707200712:43
Private Sub cmdPrint_Click()
'UDPATING COCODE    :AXP-5282007
    Dim filter                                                        As String
    On Error GoTo ErrorCode:

    If Len(txtYear.Text) = 4 Or txtYear.Text <> "" Then
        Set rsMRRINV = New ADODB.Recordset
        rsMRRINV.Open "select * from SMIS_PurchAgree WHERE year(DateReleased) = '" & txtYear.Text & "' and month(DateReleased) = " & What_month(cboMonth), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
            Screen.MousePointer = 11
            rptReleased.Formulas(0) = "CompanyName = '" & Company_name & "'"
            rptReleased.Formulas(1) = "CompanyAddress = '" & Company_Address & "'"
            filter = "(year({PurchAgree.DateReleased}) = " & txtYear.Text & " AND month({PurchAgree.DateReleased}) = " & What_month(cboMonth) & " AND {MRRINV.Released} = true)"
            rptReleased.WindowTitle = "Monthly Sales Analysis"
            PrintSQLReport rptReleased, SMIS_REPORT_PATH & "SRA-Monthly.rpt", filter, DMIS_REPORT_Connection, 1
            LogAudit "G", "MONTHLY SALES ANALYSIS", cboMonth & " " & txtYear
            
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    fillcbomonth cboMonth
    cboMonth.Text = The_month(Month(LOGDATE))
    txtYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub
