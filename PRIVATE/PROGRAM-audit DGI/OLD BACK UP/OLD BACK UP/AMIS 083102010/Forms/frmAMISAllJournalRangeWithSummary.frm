VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMISRangeWithSummary1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voucher Range With Summary"
   ClientHeight    =   6330
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   4740
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmAMISAllJournalRangeWithSummary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   4740
   Begin VB.OptionButton Option7 
      Caption         =   "Adjustment Journal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2250
      Width           =   3855
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Cash Receipt Journal-Deposited"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1890
      Width           =   3855
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Cash Receipt Journal-Undeposited"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1530
      Width           =   3855
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Sales Journal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1170
      Width           =   3855
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Cash Disbursement Journal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   810
      Width           =   3855
   End
   Begin VB.CheckBox chkDetailed 
      Caption         =   "Print Detailed"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1590
      TabIndex        =   6
      Top             =   4290
      Width           =   1875
   End
   Begin VB.CheckBox chkSummary 
      Caption         =   "Print Summary"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1590
      TabIndex        =   5
      Top             =   4020
      Width           =   1875
   End
   Begin VB.CheckBox chkJournal 
      Caption         =   "Print Journal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1590
      TabIndex        =   4
      Top             =   3750
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   405
      Left            =   750
      TabIndex        =   1
      Top             =   3240
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
      Format          =   47185921
      CurrentDate     =   38216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   405
      Left            =   3000
      TabIndex        =   3
      Top             =   3240
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
      Format          =   47185921
      CurrentDate     =   38216
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
      Left            =   2310
      MouseIcon       =   "frmAMISAllJournalRangeWithSummary.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "frmAMISAllJournalRangeWithSummary.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Close Window"
      Top             =   4620
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
      Left            =   1440
      MouseIcon       =   "frmAMISAllJournalRangeWithSummary.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "frmAMISAllJournalRangeWithSummary.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Print Report"
      Top             =   4620
      Width           =   885
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Accounts Payable Journal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   456
      Width           =   3855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "General Journal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   90
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
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
      Left            =   2520
      TabIndex        =   2
      Top             =   3300
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "From :"
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
      Left            =   30
      TabIndex        =   0
      Top             =   3300
      Width           =   675
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   9
      Top             =   3990
      Width           =   495
   End
End
Attribute VB_Name = "frmAMISRangeWithSummary1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                       As ADODB.Recordset
Dim LocalAcess As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200714:10
Private Sub cmdPrint_Click()

    On Error GoTo ErrorCode:

    If dtpFrom > dtpTo Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If
    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select * from AMIS_Journal_HD where JTYPE= '" & REPORT_RANGETYPE & "' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
    
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then

        If REPORT_RANGETYPE = "ADJ" Then
            If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "ADJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Audit Adjustment Journal", False
            If chkSummary.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "ADJSummary", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Audit Adjustment Summary", False
        End If
        If REPORT_RANGETYPE = "APJ" Then
            If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "APJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Accounts Payable Journals", False
            If chkSummary.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "APSummary", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Accounts Payables Summary", False
        End If

        If REPORT_RANGETYPE = "CDJ" Then
            If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "CDJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Cash Disbursement Journals", False
            If chkSummary.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "CDJSummary", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Cash Disbursements Summary", False
        End If

        If REPORT_RANGETYPE = "SJ" Then
            If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "SJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Sales Journals", False
            If chkSummary.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "SJSummary", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Sales Journals Summary", False
        End If
        If REPORT_RANGETYPE = "CRJ" Then
            If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "CRJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Cash Receipts Journals", False
            If chkSummary.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "CRJSummary", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Cash Receipts Summary", False
            If chkDetailed.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "CRJDetailed", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Cash Receipts Detailed", False
        End If

        If REPORT_RANGETYPE = "DRJ" Then
            If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "DRJJournals", "Journals", "{Journal_Hd.jtype}='DRJ' AND {Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Deposited Cash Receipts Journals", False
            If chkSummary.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "DRJSummary", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Deposited Cash Receipts Summary", False
            If chkDetailed.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "DRJDetailed", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Deposited Cash Receipts Detailed", False
        End If

        
    Else
        ShowNoRecord
    End If
    LogAudit "V", "JOURNAL VOUCHER RANGE WITH SUMMARY", REPORT_RANGETYPE & ": " & dtpFrom & "-" & dtpTo
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
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    Screen.MousePointer = 0
    Option1.Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAMISRange = Nothing
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Option1_Click()
    REPORT_RANGETYPE = "GJ"
    
End Sub

Private Sub Option2_Click()
    REPORT_RANGETYPE = "APJ"
End Sub

Private Sub Option3_Click()
REPORT_RANGETYPE = "SJ"
End Sub

Private Sub Option4_Click()
    REPORT_RANGETYPE = "CRJ"
End Sub

Private Sub Option5_Click()
    REPORT_RANGETYPE = "DRJ"
End Sub

Private Sub Option6_Click()
REPORT_RANGETYPE = "CDJ"
End Sub

Private Sub Option7_Click()
REPORT_RANGETYPE = "ADJ"
End Sub
