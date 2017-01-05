VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_SAEPersonal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAE Performance"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   ForeColor       =   &H00FCFCFC&
   Icon            =   "frmSMIS_Report_SAEPersonalPerformance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4515
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   0
      Width           =   4515
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
         MouseIcon       =   "frmSMIS_Report_SAEPersonalPerformance.frx":0E42
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_Report_SAEPersonalPerformance.frx":0F94
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Close Window"
         Top             =   810
         Width           =   885
      End
      Begin VB.ComboBox cboYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   330
         Width           =   1365
      End
      Begin VB.ComboBox cboMonth2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   1455
      End
      Begin VB.ComboBox cboMonth 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   330
         Width           =   1515
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
         MouseIcon       =   "frmSMIS_Report_SAEPersonalPerformance.frx":13DF
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_Report_SAEPersonalPerformance.frx":1531
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print Report"
         Top             =   810
         Width           =   885
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3150
         TabIndex        =   6
         Top             =   0
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1470
         TabIndex        =   5
         Top             =   60
         Width           =   300
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   60
         TabIndex        =   4
         Top             =   30
         Width           =   600
      End
   End
   Begin Crystal.CrystalReport rptGenREP 
      Left            =   7020
      Top             =   1140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Sales Executive Performance"
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
End
Attribute VB_Name = "frmSMIS_Report_SAEPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPurchAgree                                                      As ADODB.Recordset
Dim rsSrep                                                            As ADODB.Recordset
Dim REPORTNAME                                                        As String

Sub ShowSAEVsPROSPECT()
    REPORTNAME = "SAECLOSING"
End Sub

Sub ShowSAETeamVsPROSPECT()
    REPORTNAME = "SAECLOSINGTEAM"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo errcode:
    If What_month(cboMonth) > What_month(cboMonth2) Then
        MsgSpeechBox "Error In From - To Months"
        Exit Sub
    End If
    rptGenREP.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptGenREP.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

    If REPORTNAME = "SAECLOSING" Then
        PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "SAE\saeclosing.rpt", "month({CP.LogInitialInquiry}) >= " & What_month(cboMonth) & " AND month({CP.LogInitialInquiry}) <= " & What_month(cboMonth2) & " AND YEAR({CP.LogInitialInquiry}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        rptGenREP.PageZoom 90

    ElseIf REPORTNAME = "SAECLOSINGTEAM" Then
        PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "SAE\saeclosingbyteam.rpt", "month({CP.LogInitialInquiry}) >= " & What_month(cboMonth) & " AND month({CP.LogInitialInquiry}) <= " & What_month(cboMonth2) & " AND YEAR({CP.LogInitialInquiry}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        rptGenREP.PageZoom 90
    Else
        Set rsPurchAgree = New ADODB.Recordset
        rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE year(datereleased) = " & cboYear.Text & " AND month(datereleased) >= " & What_month(cboMonth) & " AND month(datereleased) <=" & What_month(cboMonth2), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPurchAgree.EOF And Not rsPurchAgree.EOF Then
            Screen.MousePointer = 11
            PrintSQLReport rptGenREP, SMIS_REPORT_PATH & "SAE\saeper.rpt", "month({purchagree.datereleased}) >= " & What_month(cboMonth) & " AND month({purchagree.datereleased}) <= " & What_month(cboMonth2) & " AND YEAR({purchagree.datereleased}) = " & cboYear.Text & " AND {purchagree.salesae} = '" & SAENAME & "' ", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        Else
            MsgSpeechBox "No Record for the Month of " & cboMonth.Text
        End If
    End If
    Exit Sub
errcode:
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
    fillcbomonth cboMonth2
    fillcbomoreyear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboMonth2.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    REPORTNAME = ""
End Sub

