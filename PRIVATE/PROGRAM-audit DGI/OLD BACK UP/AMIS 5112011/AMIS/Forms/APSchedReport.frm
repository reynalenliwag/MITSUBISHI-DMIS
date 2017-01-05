VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMISAPSchedReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Schedule of Accounts Payable"
   ClientHeight    =   2790
   ClientLeft      =   180
   ClientTop       =   330
   ClientWidth     =   4635
   ForeColor       =   &H00FFFFFF&
   Icon            =   "APSchedReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2790
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
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
      Left            =   2310
      MouseIcon       =   "APSchedReport.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "APSchedReport.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Close Window"
      Top             =   1815
      Width           =   885
   End
   Begin VB.OptionButton optForthePeriod 
      Caption         =   "Accounts Payable for the Period"
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
      Left            =   660
      TabIndex        =   3
      Top             =   870
      Width           =   3405
   End
   Begin VB.OptionButton optAsOf 
      Caption         =   "Accounts Payable as Of"
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
      Left            =   660
      TabIndex        =   2
      Top             =   570
      Value           =   -1  'True
      Width           =   3405
   End
   Begin Crystal.CrystalReport rptAMISDueReport 
      Left            =   90
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Accounts Receivable Aging Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowNavigationCtls=   -1  'True
      WindowShowCancelBtn=   -1  'True
      WindowShowPrintBtn=   -1  'True
      WindowShowExportBtn=   -1  'True
      WindowShowZoomCtl=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowProgressCtls=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpAsOF 
      Height          =   405
      Left            =   1770
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
      CalendarBackColor=   15857397
      Format          =   52101121
      CurrentDate     =   38216
   End
   Begin VB.Frame picPeriod 
      Height          =   585
      Left            =   210
      TabIndex        =   4
      Top             =   1110
      Width           =   4215
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   780
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
         CalendarBackColor=   15857397
         Format          =   52101121
         CurrentDate     =   38216
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   2730
         TabIndex        =   5
         Top             =   150
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
         CalendarBackColor=   15857397
         Format          =   52101121
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
         Left            =   0
         TabIndex        =   6
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
         Left            =   2190
         TabIndex        =   8
         Top             =   210
         Width           =   435
      End
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
      Left            =   1440
      MouseIcon       =   "APSchedReport.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "APSchedReport.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print Report"
      Top             =   1815
      Width           =   885
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
      Left            =   900
      TabIndex        =   0
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "frmAMISAPSchedReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
    Select Case KeyCode
    Case vbKeyF1 And Shift = 1:
        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        Unload frmALL_AuditInquiry
        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        If optAsOf.Value = True Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SCHEDULE OF ACCOUNTS PAYABLE)"
            Call frmALL_AuditInquiry.DisplayHistory("", "SCHEDULE OF ACCOUNTS PAYABLE", "PRINTING")
        Else
            frmALL_AuditInquiry.Caption = "Audit Inquiry (ACCOUNTS PAYABLE DUE REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "ACCOUNTS PAYABLE DUE REPORT", "PRINTING")
        End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    dtpAsOF = LOGDATE
    picPeriod.Enabled = False
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
    dtpAsOF.Enabled = True
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    Dim rsProfile                                 As ADODB.Recordset
    On Error GoTo Errorcode:

    'rptAMISDueReport.Reset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptAMISDueReport.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptAMISDueReport.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
    End If
    If optAsOf.Value = True Then
        rptAMISDueReport.WindowTitle = "SCHEDULE OF ACCOUNTS PAYABLE AS OF: " & dtpAsOF
        rptAMISDueReport.ReportTitle = "SCHEDULE OF ACCOUNTS PAYABLE AS OF: " & dtpAsOF
        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\APScheduleReport.Rpt", "{Journal_Hd.InvoiceDate} <= Date(" & Year(dtpAsOF) & "," & Month(dtpAsOF) & "," & Day(dtpAsOF) & ")", DMIS_REPORT_Connection, 1
        LogAudit "V", "SCHEDULE OF ACCOUNTS PAYABLE", "As of: " & dtpAsOF
        Call NEW_LogAudit("V", "SCHEDULE OF ACCOUNTS PAYABLE", "", "", "", dtpAsOF, "", "")
    Else
        rptAMISDueReport.WindowTitle = "ACCOUNTS PAYABLE DUE REPORT FOR THE PERIOD: " & dtpFrom & " TO " & dtpTo
        rptAMISDueReport.ReportTitle = "ACCOUNTS PAYABLE DUE REPORT FOR THE PERIOD: " & dtpFrom & " TO " & dtpTo
        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\APScheduleReport.Rpt", "{Journal_Hd.DueDate} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.DueDate} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
        Call NEW_LogAudit("V", "ACCOUNTS PAYABLE DUE REPORT", "", "", "", dtpFrom & " " & dtpTo, "", "")
    End If
    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub optAsOf_Click()
    picPeriod.Enabled = False
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
    dtpAsOF.Enabled = True
End Sub

Private Sub optForthePeriod_Click()
    picPeriod.Enabled = True
    dtpAsOF.Enabled = False
    dtpFrom.Enabled = True
    dtpTo.Enabled = True
End Sub

