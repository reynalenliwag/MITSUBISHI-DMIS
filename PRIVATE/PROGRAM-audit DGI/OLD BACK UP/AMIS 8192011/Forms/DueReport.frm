VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAMISDueReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A/P Due Report"
   ClientHeight    =   2055
   ClientLeft      =   180
   ClientTop       =   330
   ClientWidth     =   4560
   ForeColor       =   &H00FFFFFF&
   Icon            =   "DueReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2055
   ScaleWidth      =   4560
   Begin VB.TextBox txtdescription 
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   450
      Width           =   3915
   End
   Begin VB.ComboBox cboacctcode 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   120
      Width           =   3405
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
      Height          =   825
      Left            =   2190
      MouseIcon       =   "DueReport.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "DueReport.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Close Window"
      Top             =   1125
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
      MouseIcon       =   "DueReport.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "DueReport.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print Report"
      Top             =   1125
      Width           =   885
   End
   Begin VB.OptionButton optForthePeriod 
      Caption         =   "Due Report for the Period"
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
      Left            =   810
      TabIndex        =   3
      Top             =   3750
      Width           =   2745
   End
   Begin VB.OptionButton optAsOf 
      Caption         =   "Due Report as Of"
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
      Left            =   810
      TabIndex        =   2
      Top             =   3450
      Value           =   -1  'True
      Width           =   2745
   End
   Begin Crystal.CrystalReport rptAMISDueReport 
      Left            =   240
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Work Sheet - Trial Balance of Journals"
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
   Begin MSComCtl2.DTPicker dtpAsOF 
      Height          =   405
      Left            =   1440
      TabIndex        =   1
      Top             =   630
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
   Begin VB.Frame picPeriod 
      Height          =   585
      Left            =   90
      TabIndex        =   4
      Top             =   4080
      Width           =   4185
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   780
         TabIndex        =   6
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
         Left            =   2730
         TabIndex        =   8
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
         Left            =   0
         TabIndex        =   5
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
         TabIndex        =   7
         Top             =   210
         Width           =   435
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Description"
      Height          =   375
      Left            =   180
      TabIndex        =   12
      Top             =   150
      Width           =   1125
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
      Top             =   690
      Width           =   765
   End
End
Attribute VB_Name = "frmAMISDueReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboacctcode_Change()
    ReturnAccountCode cboacctcode.Text
End Sub

Private Sub cboacctcode_Click()
    ReturnAccountCode cboacctcode.Text

End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim rsProfile                                      As ADODB.Recordset
    On Error GoTo ErrorCode:

    'rptAMISDueReport.Reset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptAMISDueReport.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptAMISDueReport.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
    End If
    If REPORT_AP = "SCHED" Then
        If optAsOf.Value = True Then
            rptAMISDueReport.WindowTitle = "ACCOUNTS PAYABLE DUE AS OF: " & dtpAsOF
            PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\APDueReport.Rpt", "{AMIS_vw_APAGING.duedate} <= Date(" & Year(dtpAsOF) & "," & Month(dtpAsOF) & "," & Day(dtpAsOF) & ")AND {AMIS_vw_APAGING.Acct_Code} = '" & txtDescription & "'", DMIS_REPORT_Connection, 1
            LogAudit "V", "A/P DUE REPORT", "As of: " & dtpAsOF
        Else
            rptAMISDueReport.WindowTitle = "SCHEDULE OF ACCOUNTS PAYABLE FOR THE PERIOD: " & dtpFrom & " TO " & dtpTo
            PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\APDueReport.Rpt", "{AMIS_vw_APAGING.duedate} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {AMIS_vw_APAGING.duedate} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")AND {AMIS_vw_APAGING.Acct_Code} = '" & txtDescription & "'", DMIS_REPORT_Connection, 1
            LogAudit "V", "A/P DUE REPORT FOR THE PERIOD", dtpFrom & " - " & dtpTo
        End If
    Else
        If optAsOf.Value = True Then
            rptAMISDueReport.WindowTitle = "ACCOUNTS PAYABLE AGING REPORT AS OF: " & dtpAsOF
            rptAMISDueReport.ReportTitle = "ACCOUNTS PAYABLE AGING REPORT AS OF: " & dtpAsOF
            PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\APAgingReport.rpt", "{AMIS_vw_APAGING.jdate} <= Date(" & Year(dtpAsOF) & "," & Month(dtpAsOF) & "," & Day(dtpAsOF) & ")AND{AMIS_vw_APAGING.Acct_Code} = '" & txtDescription & "'", DMIS_REPORT_Connection, 1
            LogAudit "V", "SCHEDULE OF ACCOUNTS PAYABLE", "As of: " & dtpAsOF
        Else
            rptAMISDueReport.WindowTitle = "ACCOUNTS PAYABLE AGING REPORT FOR THE PERIOD: " & dtpFrom & " TO " & dtpTo
            rptAMISDueReport.ReportTitle = "ACCOUNTS PAYABLE AGING REPORT FOR THE PERIOD: " & dtpFrom & " TO " & dtpTo
            PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\APAgingReport.rpt", "{AMIS_vw_APAGING.jdate} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {AMIS_vw_APAGING.jdate} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")AND {AMIS_vw_APAGING.Acct_Code} = '" & txtDescription & "'", DMIS_REPORT_Connection, 1
            LogAudit "V", "ACCOUNTS PAYABLE DUE REPORT", "Period: " & dtpFrom & " - " & dtpTo
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
    If REPORT_AP = "SCHED" Then
        Me.Caption = "Sched of A/P"
    Else
        Me.Caption = "A/P Aging Report"
    End If
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    dtpAsOF = LOGDATE
    picPeriod.Enabled = False
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
    dtpAsOF.Enabled = True
    GetAcctcode
    Screen.MousePointer = 0

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
Sub GetAcctcode()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT * from AMIS_chartaccount where left(acctcode,5)='21-01' or left(acctcode,5)='21-02'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    cboacctcode.Clear

    Do While Not RS.EOF
        cboacctcode.AddItem (RS!DESCRIPTION)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub
Function ReturnAccountCode(Xacct_desc As String)
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "select description,acctcode from AMIS_chartaccount where description = '" & Xacct_desc & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        txtDescription = Null2String(RS!ACCTCODE)
    End If
    Set RS = Nothing
End Function


