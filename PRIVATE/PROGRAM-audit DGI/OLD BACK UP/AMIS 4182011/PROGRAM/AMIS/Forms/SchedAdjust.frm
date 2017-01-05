VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAMISSchedAdjust 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schedule of Adjustments"
   ClientHeight    =   2445
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   4530
   ForeColor       =   &H00F5F5F5&
   Icon            =   "SchedAdjust.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4530
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
      Left            =   2250
      MouseIcon       =   "SchedAdjust.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "SchedAdjust.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Close Window"
      Top             =   1515
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
      Left            =   1380
      MouseIcon       =   "SchedAdjust.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "SchedAdjust.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Print Report"
      Top             =   1515
      Width           =   885
   End
   Begin VB.Frame picPeriod 
      Height          =   585
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4305
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   810
         TabIndex        =   5
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   48693249
         CurrentDate     =   38216
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   2760
         TabIndex        =   7
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   48693249
         CurrentDate     =   38216
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   30
         TabIndex        =   4
         Top             =   210
         Width           =   675
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   2250
         TabIndex        =   6
         Top             =   210
         Width           =   435
      End
   End
   Begin Crystal.CrystalReport rptAMISTrialBalance 
      Left            =   300
      Top             =   -420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Work Sheet - Trial Balance of Journals"
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
   Begin MSComCtl2.DTPicker dtpAsOF 
      Height          =   405
      Left            =   1440
      TabIndex        =   1
      Top             =   90
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   48693249
      CurrentDate     =   38216
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustments made for the period:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   210
      TabIndex        =   2
      Top             =   630
      Width           =   4155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "As Of:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "frmAMISSchedAdjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function Setacctname(XXX As String) As String
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount where AcctCode = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        Setacctname = Null2String(rsChartAccount!DESCRIPTION)
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200714:10
Private Sub cmdPrint_Click()


    On Error GoTo ErrorCode:

    If IsDate(dtpAsOF) = False Then
        MsgSpeechBox "Error In As of date"
        Exit Sub
    End If
    Screen.MousePointer = 11
    ShowWorkSheetReport "AdjustmentSheet", "FinancialStatement", "(({AMIS_Journal_Det.Jtype} = 'ADJ' or {AMIS_Journal_Det.Jtype} = 'PDJ') AND {AMIS_Journal_Det.Jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") and {AMIS_Journal_Det.Jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")) OR ({AMIS_Journal_Det.Jtype} <> 'ADJ' AND {AMIS_Journal_Det.Jtype} <> 'PDJ' AND {AMIS_Journal_Det.JDate} <= date(" & Year(dtpAsOF) & "," & Month(dtpAsOF) & "," & Day(dtpAsOF) & "))", "Schedule of Adjustments", dtpAsOF, True, 0, 0
    Screen.MousePointer = 0
    LogAudit "V", "SCHEDULE OF ADJUSTMENT"
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
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    dtpAsOF = LOGDATE
    Screen.MousePointer = 0
End Sub

Public Sub ShowWorkSheetReport(ReportName As Variant, ReportFolder As Variant, filter As Variant, ReportHeading As String, REPORT_DATE As String, WithDate As Boolean, SUM_DEBIT_WS As Variant, SUM_CREDIT_WS As Variant)
    Screen.MousePointer = 11
    Dim rsProfile                                      As ADODB.Recordset
    Dim CrystalRpt                                     As Crystal.CrystalReport
    Set CrystalRpt = Me.rptAMISTrialBalance
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        'CrystalRpt.ReportFileName = AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rtp"
        CrystalRpt.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        CrystalRpt.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        If WithDate = True Then CrystalRpt.Formulas(2) = "ReportDate = '" & REPORT_DATE & "'"
        CrystalRpt.ReportTitle = ReportHeading: CrystalRpt.WindowTitle = ReportHeading
        'CrystalRpt.Formulas(3) = "SUM_DEBIT =" & SUM_DEBIT_WS
        'CrystalRpt.Formulas(4) = "SUM_CREDIT =" & SUM_CREDIT_WS
        PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", filter, DMIS_REPORT_Connection, 1
        CrystalRpt.PageZoom 89
    End If
    Screen.MousePointer = 0
End Sub

