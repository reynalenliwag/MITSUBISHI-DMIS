VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMISRangeWithSummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Journal Voucher Range With Summary"
   ClientHeight    =   2940
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   4740
   ForeColor       =   &H00FFFFFF&
   Icon            =   "ReportRangeWithSummary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4740
   Begin VB.CheckBox chkCancel 
      Caption         =   "Print Cancelled"
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
      Left            =   1560
      TabIndex        =   13
      Top             =   1830
      Width           =   1875
   End
   Begin VB.CheckBox chkUposted 
      Caption         =   "Print Un-Posted"
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
      Left            =   1560
      TabIndex        =   12
      Top             =   1620
      Width           =   1875
   End
   Begin VB.OptionButton OptJornal 
      Caption         =   "Exportable Journals"
      Height          =   375
      Index           =   1
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   570
      Width           =   2055
   End
   Begin VB.OptionButton OptJornal 
      Caption         =   "Printable Journals"
      Height          =   375
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   570
      Width           =   2115
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
      Left            =   1560
      TabIndex        =   6
      Top             =   1410
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
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
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
      Left            =   1560
      TabIndex        =   4
      Top             =   990
      Width           =   1875
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   405
      Left            =   780
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
      Format          =   56950785
      CurrentDate     =   38216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   405
      Left            =   3000
      TabIndex        =   3
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
      Format          =   56950785
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
      Left            =   2220
      MouseIcon       =   "ReportRangeWithSummary.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportRangeWithSummary.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Close Window"
      Top             =   2070
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
      Left            =   1350
      MouseIcon       =   "ReportRangeWithSummary.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportRangeWithSummary.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Print Report"
      Top             =   2070
      Width           =   885
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
      Top             =   150
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
      Top             =   150
      Width           =   675
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   9
      Top             =   2940
      Width           =   495
   End
End
Attribute VB_Name = "frmAMISRangeWithSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                  As ADODB.Recordset
Dim LocalAcess                                    As String
Dim printable                                     As Boolean
Dim exportable                                    As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Update by BTT 05292008
Private Sub cmdPrint_Click()

    On Error GoTo Errorcode:

    If dtpFrom > dtpTo Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If

    If chkCancel.Value = 0 And chkUposted.Value = 0 And chkDetailed.Value = 0 And chkJournal.Value = 0 And chkSummary.Value = 0 Then
        MsgBox "Please select in the option box below.", vbInformation, "Information"
        Exit Sub
    End If

    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select * from AMIS_Journal_HD where JTYPE= '" & REPORT_RANGETYPE & "' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then

        If REPORT_RANGETYPE = "ADJ" Then
            LocalAcess = "ADJUSTMENT JOURNAL"
            If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "ADJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Audit Adjustment Journal", False
            If chkSummary.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "ADJSummary", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Audit Adjustment Summary", False
        End If
        If REPORT_RANGETYPE = "APJ" Then
            LocalAcess = "ACCOUNTS PAYABLE JOURNAL"
            If printable = True Then                       ' Printable Journal
                If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "APJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_HD.Status} ='P'", "Accounts Payable Journals", False
                If chkSummary.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "APSummary", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Accounts Payables Summary", False
                If chkUposted.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "APJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_HD.Status} ='N'", "Un-Posted Accounts Payable Journals", False
                If chkCancel.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "APJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_HD.Status} ='C'", "Cancelled Accounts Payable Journals", False
            End If
            If exportable = True Then
                If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "APEJournals", "Journals\ExportableJournal", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Accounts Payable Journals", False
            End If
            Call NEW_LogAudit("V", "ACCOUNTS PAYABLE JOURNAL", "", "", "", dtpFrom & " " & dtpTo, "", "")
        End If

        If REPORT_RANGETYPE = "CDJ" Then
            LocalAcess = "CASH DISBURSEMENT JOURNALS"
            If printable = True Then                       ' Printable Journal
                If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "CDJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_HD.Status} ='P'", "Cash Disbursement Journals", False
                If chkSummary.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "CDJSummary", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Cash Disbursements Summary", False
                If chkUposted.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "CDJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_HD.Status} ='N'", "Un-Posted Cash Disbursement Journals", False
                If chkCancel.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "CDJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")  AND {Journal_HD.Status} ='C'", "Cancelled Cash Disbursement Journals", False
            End If
            If exportable = True Then
                If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "CDJEJournals", "Journals\ExportableJournal", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Cash Disbursement Journals", False
            End If
            Call NEW_LogAudit("V", "CASH DISBURSEMENT JOURNALS", "", "", "", dtpFrom & " " & dtpTo, "", "")
        End If

        If REPORT_RANGETYPE = "SJ" Then
            LocalAcess = "SALES JOURNAL"
            If printable = True Then                       ' Printable Journal
                If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "SJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") and {Journal_HD.JType} = 'SJ' AND {Journal_HD.Status} = 'P'", "Sales Journals", False
                If chkSummary.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "SJSummary", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Sales Journals Summary", False
                If chkUposted.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "SJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") and {Journal_HD.JType} = 'SJ' AND {Journal_HD.Status} = 'N'", "Un-Posted Sales Journals", False
                If chkCancel.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "SJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")and {Journal_HD.JType} = 'SJ' AND {Journal_HD.Status} = 'C'", "Cancelled Sales Journals", False
            End If
            If exportable = True Then
                If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "SJEJournals", "Journals\ExportableJournal", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Sales Journals", False
            End If
            Call NEW_LogAudit("V", "SALES JOURNAL", "", "", "", dtpFrom & " " & dtpTo, "", "")
        End If
        If REPORT_RANGETYPE = "CRJ" Then
            LocalAcess = "CASH RECEIPTS JOURNAL"
            If printable = True Then                       ' Printable Journal
                If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "CRJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") and {Journal_HD.JType} = 'CRJ' and {Journal_HD.Status} = 'P'", "Cash Receipts Journals", False
                If chkSummary.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "CRJSummary", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Cash Receipts Summary", False
                If chkDetailed.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "CRJDetailed", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Cash Receipts Detailed", False
                If chkUposted.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "CRJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") and {Journal_HD.JType} = 'CRJ' and {Journal_HD.Status} = 'N' ", " Un-Posted Cash Receipts Journals", False
                If chkCancel.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "CRJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") and {Journal_HD.JType} = 'CRJ' and {Journal_HD.Status} = 'C'", "Cancelled Cash Receipts Journals", False
            End If
            If exportable = True Then
                If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "CRJEJournals", "Journals\ExportableJournal", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Cash Receipts Journals", False
            End If
            Call NEW_LogAudit("V", "CASH RECEIPTS JOURNAL", "", "", "", dtpFrom & " " & dtpTo, "", "")
        End If

        If REPORT_RANGETYPE = "DRJ" Then
            LocalAcess = "DEPOSITED CASH RECEIPTS JOURNAL"
            If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "DRJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Deposited Cash Receipts Journals", False
            If chkSummary.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "DRJSummary", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Deposited Cash Receipts Summary", False
            If chkDetailed.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "DRJDetailed", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Deposited Cash Receipts Detailed", False
            Call NEW_LogAudit("V", "DEPOSITED CASH RECEIPTS JOURNAL", "", "", "", dtpFrom & " " & dtpTo, "", "")
        End If

        If REPORT_RANGETYPE = "GJ" Then
            LocalAcess = "GENERAL JOURNAL"
            If printable = True Then                       ' Printable Journal
                If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "GJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_HD.Status} ='P'", "General Journals", False
                If chkSummary.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "GJSummary", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_HD.Status} ='P'", "General Summary", False
                If chkUposted.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "GJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_HD.Status} ='N'", "Un-Posted General Journals", False
                If chkCancel.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "GJJournals", "Journals", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_HD.Status} ='C'", "Cancelled General Journals", False
            End If
            If exportable = True Then
                If chkJournal.Value = 1 Then ShowRangeReport dtpFrom, dtpTo, "GJEJournals", "Journals\ExportableJournal", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "General Journals", False
            End If
            Call NEW_LogAudit("V", "ACCOUNTS PAYABLE JOURNAL", "", "", "", dtpFrom & " " & dtpTo, "", "")
        End If
        'Unload Me
    Else
        ShowNoRecord
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
    Select Case KeyCode
    Case vbKeyF1 And Shift = 1:
        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        Unload frmALL_AuditInquiry
        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        If REPORT_RANGETYPE = "APJ" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (ACCOUNTS PAYABLE JOURNAL)"
            Call frmALL_AuditInquiry.DisplayHistory("", "ACCOUNTS PAYABLE JOURNAL", "PRINTING")
        ElseIf REPORT_RANGETYPE = "CDJ" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (CASH DISBURSEMENT JOURNALS)"
            Call frmALL_AuditInquiry.DisplayHistory("", "CASH DISBURSEMENT JOURNALS", "PRINTING")
        ElseIf REPORT_RANGETYPE = "SJ" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SALES JOURNAL)"
            Call frmALL_AuditInquiry.DisplayHistory("", "SALES JOURNAL", "PRINTING")
        ElseIf REPORT_RANGETYPE = "CRJ" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (CASH RECEIPTS JOURNAL)"
            Call frmALL_AuditInquiry.DisplayHistory("", "CASH RECEIPTS JOURNAL", "PRINTING")
        ElseIf REPORT_RANGETYPE = "DRJ" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (DEPOSITED CASH RECEIPTS JOURNAL)"
            Call frmALL_AuditInquiry.DisplayHistory("", "DEPOSITED CASH RECEIPTS JOURNAL", "PRINTING")
        ElseIf REPORT_RANGETYPE = "JVS" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (DEPOSITED CASH RECEIPTS JOURNAL)"
            Call frmALL_AuditInquiry.DisplayHistory("", "DEPOSITED CASH RECEIPTS JOURNAL", "PRINTING")
        End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAMISRange = Nothing
End Sub

Private Sub OptJornal_Click(Index As Integer)
    If OptJornal(0).Value = True Then
        OptJornal(0).BackColor = &HFFFF00
        printable = True
    Else
        printable = False
        OptJornal(0).BackColor = &HE0E0E0
    End If
    If OptJornal(1).Value = True Then
        exportable = True
        OptJornal(1).BackColor = &HFFFF00
        chkJournal.Value = 1
        chkSummary.Enabled = False
        chkCancel.Enabled = False
        chkDetailed.Enabled = False
        chkUposted.Enabled = False
    Else
        exportable = False
        OptJornal(1).BackColor = &HE0E0E0
        chkJournal.Value = 0
        chkSummary.Enabled = True
        chkCancel.Enabled = True
        chkDetailed.Enabled = True
        chkUposted.Enabled = True
    End If
End Sub

