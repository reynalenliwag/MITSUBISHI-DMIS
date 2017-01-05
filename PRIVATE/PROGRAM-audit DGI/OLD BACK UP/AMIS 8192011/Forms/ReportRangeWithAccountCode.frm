VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAMISRangeWithAccountCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Journal Voucher Summary"
   ClientHeight    =   2265
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   4830
   ForeColor       =   &H00FFFFFF&
   Icon            =   "ReportRangeWithAccountCode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4830
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
      Left            =   2445
      MouseIcon       =   "ReportRangeWithAccountCode.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportRangeWithAccountCode.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Close Window"
      Top             =   1380
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
      Left            =   1575
      MouseIcon       =   "ReportRangeWithAccountCode.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportRangeWithAccountCode.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Print Report"
      Top             =   1380
      Width           =   885
   End
   Begin VB.ComboBox cboAcct_Code 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   330
      Left            =   60
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   450
      Width           =   4695
   End
   Begin VB.TextBox txtDescription 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00701E2A&
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Top             =   90
      Width           =   3495
   End
   Begin Crystal.CrystalReport rptAMISrange 
      Left            =   120
      Top             =   1710
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   405
      Left            =   810
      TabIndex        =   4
      Top             =   840
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
      Format          =   20578305
      CurrentDate     =   38216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   405
      Left            =   3060
      TabIndex        =   6
      Top             =   840
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
      Format          =   20578305
      CurrentDate     =   38216
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
      Left            =   2580
      TabIndex        =   5
      Top             =   900
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
      Left            =   90
      TabIndex        =   3
      Top             =   900
      Width           =   675
   End
   Begin VB.Label Label34 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Account No."
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
      Height          =   225
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   9
      Top             =   2970
      Width           =   495
   End
End
Attribute VB_Name = "frmAMISRangeWithAccountCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                       As ADODB.Recordset
Dim rsChartAccount                                     As ADODB.Recordset
Dim xJOURNALTYPE                                       As String

Function SetAccountName(VVV As Variant) As String
    Set rsChartAccount = New ADODB.Recordset
    rsChartAccount.Open "Select AcctCode,Description from AMIS_ChartAccount where Description = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        SetAccountName = Null2String(rsChartAccount!ACCTCODE)
    End If
End Function

Sub LoadJournal(XXX As String)
    xJOURNALTYPE = XXX
End Sub

Sub InitCbo(xJType As String)
    Set rsChartAccount = New ADODB.Recordset
    rsChartAccount.Open "select Description from AMIS_ChartAccount AC INNER JOIN (SELECT DISTINCT ACCT_CODE FROM AMIS_JOURNAL_DET WHERE JTYPE='" & xJType & "') DET ON AC.ACCTCODE=DET.ACCT_CODE order by DESCRIPTION asc ", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        Combo_Loadval cboAcct_Code, rsChartAccount
    End If
End Sub

Private Sub cboAcct_Code_Change()
    txtDescription.Text = SetAccountName(cboAcct_Code.Text)
End Sub

Private Sub cboAcct_Code_Click()
    txtDescription.Text = SetAccountName(cboAcct_Code.Text)
End Sub

Private Sub cboAcct_Code_LostFocus()
    txtDescription.Text = SetAccountName(cboAcct_Code.Text)
End Sub

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
    rsJournal_HD.Open "select * from AMIS_Journal_HD where (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        If REPORT_RANGETYPE = "APJ" Then
            If Function_Access(LOGID, "Acess_Print", "ACCOUNTS PAYABLE LEDGER CODE RUNNING BALANCE") = False Then Exit Sub
            ShowRangeReport dtpFrom, dtpTo, "APJLedgerCodeRunningBalance", "RunningBalance", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_Det.Acct_Code} = '" & txtDescription.Text & "'and {Journal_Det.jtype} = 'APJ'", "Ledger Code Running Balance", False
            Call NEW_LogAudit("V", "ACCOUNTS PAYABLE LEDGER CODE RUNNING BALANCE", "", "", "", dtpFrom & " " & dtpTo, "", "")
        ElseIf REPORT_RANGETYPE = "CDJ" Then
            If Function_Access(LOGID, "Acess_Print", "CASH DISBURSEMENT LEDGER CODE RUNNING BALANCE") = False Then Exit Sub
            ShowRangeReport dtpFrom, dtpTo, "CDJLedgerCodeRunningBalance", "RunningBalance", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_Det.Acct_Code} = '" & txtDescription.Text & "' and {Journal_Det.jtype} = 'CDJ'", "Ledger Code Running Balance", False
            Call NEW_LogAudit("V", "CASH DISBURSEMENT LEDGER CODE RUNNING BALANCE", "", "", "", dtpFrom & " " & dtpTo, "", "")
        ElseIf REPORT_RANGETYPE = "SJ" Then
            If Function_Access(LOGID, "Acess_Print", "SALES LEDGER CODE RUNNING BALANCE") = False Then Exit Sub
            ShowRangeReport dtpFrom, dtpTo, "SJLedgerCodeRunningBalance", "RunningBalance", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_Det.Acct_Code} = '" & txtDescription.Text & "' and {Journal_Det.jtype} = 'SJ'", "Ledger Code Running Balance", False
            Call NEW_LogAudit("V", "SALES JOURNALS LEDGER CODE RUNNING BALANCE", "", "", "", dtpFrom & " " & dtpTo, "", "")
        ElseIf REPORT_RANGETYPE = "CRJ" Then
            If Function_Access(LOGID, "Acess_Print", "CASH RECEIPTS LEDGER CODE RUNNING BALANCE") = False Then Exit Sub
            ShowRangeReport dtpFrom, dtpTo, "CRJLedgerCodeRunningBalance", "RunningBalance", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_Det.Acct_Code} = '" & txtDescription.Text & "' and {Journal_Det.jtype} = 'CRJ' ", "Ledger Code Running Balance", False
            Call NEW_LogAudit("V", "CASH RECEIPTS LEDGER CODE RUNNING BALANCE", "", "", "", dtpFrom & " " & dtpTo, "", "")
        ElseIf REPORT_RANGETYPE = "GJ" Then
            If Function_Access(LOGID, "Acess_Print", "GENERAL JOURNAL LEDGER CODE RUNNING BALANCE") = False Then Exit Sub
            ShowRangeReport dtpFrom, dtpTo, "GJLedgerCodeRunningBalance", "RunningBalance", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") AND {Journal_Det.Acct_Code} = '" & txtDescription.Text & "' and {Journal_Det.jtype} = 'GJ'", "Ledger Code Running Balance", False
            Call NEW_LogAudit("V", "GENERAL JOURNAL LEDGER CODE RUNNING BALANCE", "", "", "", dtpFrom & " " & dtpTo, "", "")
        End If


    Else
        ShowNoRecord
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If Me.ActiveControl.Name = "cboAcct_Code" Then
        If cboAcct_Code.Text = "" Then
            Call VBComBoBoxDroppedDown(cboAcct_Code)
        Else
            MoveKeyPress KeyCode
        End If
    Else
        MoveKeyPress KeyCode
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF1 And Shift = 1:
        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        Unload frmALL_AuditInquiry
        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        If REPORT_RANGETYPE = "APJ" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (ACCOUNTS PAYABLE LEDGER CODE RUNNING BALANCE)"
            Call frmALL_AuditInquiry.DisplayHistory("", "ACCOUNTS PAYABLE LEDGER CODE RUNNING BALANCE", "PRINTING")
        ElseIf REPORT_RANGETYPE = "CDJ" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (CASH DISBURSEMENT LEDGER CODE RUNNING BALANCE)"
            Call frmALL_AuditInquiry.DisplayHistory("", "CASH DISBURSEMENT LEDGER CODE RUNNING BALANCE", "PRINTING")
        ElseIf REPORT_RANGETYPE = "SJ" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SALES JOURNALS LEDGER CODE RUNNING BALANCE)"
            Call frmALL_AuditInquiry.DisplayHistory("", "SALES JOURNALS LEDGER CODE RUNNING BALANCE", "PRINTING")
        ElseIf REPORT_RANGETYPE = "CRJ" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (CASH RECEIPTS LEDGER CODE RUNNING BALANCE)"
            Call frmALL_AuditInquiry.DisplayHistory("", "CASH RECEIPTS LEDGER CODE RUNNING BALANCE", "PRINTING")
        ElseIf REPORT_RANGETYPE = "GJ" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (GENERAL JOURNAL LEDGER CODE RUNNING BALANCE)"
            Call frmALL_AuditInquiry.DisplayHistory("", "GENERAL JOURNAL LEDGER CODE RUNNING BALANCE", "PRINTING")
        End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    Call InitCbo(xJOURNALTYPE)
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAMISRange = Nothing
End Sub
