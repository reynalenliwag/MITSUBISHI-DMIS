VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMISRange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Journal Voucher Summary"
   ClientHeight    =   1470
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   4770
   ForeColor       =   &H00FFFFFF&
   Icon            =   "ReportRange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4770
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
      MouseIcon       =   "ReportRange.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportRange.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   585
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
      MouseIcon       =   "ReportRange.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportRange.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   585
      Width           =   885
   End
   Begin Crystal.CrystalReport rptAMISrange 
      Left            =   870
      Top             =   990
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
      Format          =   131137537
      CurrentDate     =   38216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   405
      Left            =   3030
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
      Format          =   131137537
      CurrentDate     =   38216
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
      Left            =   60
      TabIndex        =   0
      Top             =   150
      Width           =   675
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
      Left            =   2550
      TabIndex        =   2
      Top             =   150
      Width           =   435
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   6
      Top             =   2970
      Width           =   495
   End
End
Attribute VB_Name = "frmAMISRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                            As ADODB.Recordset
Public LocalAcess                                           As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200714:10
Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "REGISTER REPORT") = False Then Exit Sub

    On Error GoTo ErrorCode:

    If dtpFrom > dtpTo Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If
    If REPORT_RANGETYPE = "ADJ" Then
        'If Function_Access(LOGID, "Acess_Print", "") = False Then Exit Sub
        Set rsJournal_HD = New ADODB.Recordset
        rsJournal_HD.Open "select * from AMIS_Journal_HD where jtype = 'ADJ' and (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
            ShowRangeReport dtpFrom, dtpTo, "ADJSummary", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Audit Adjustment Summary", False
            Unload Me
        Else
            ShowNoRecord
        End If
    End If
    If REPORT_RANGETYPE = "JVS" Then
        'If Function_Access(LOGID, "Acess_Print", "") = False Then Exit Sub
        Dim xREMARKS                                        As String
        Dim rsJournalDet                                    As ADODB.Recordset
        Set rsJournal_HD = New ADODB.Recordset
        rsJournal_HD.Open "select * from AMIS_Journal_HD where jtype = 'GJ' and (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
            ShowRangeReport dtpFrom, dtpTo, "GJSummary", "Summary", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "Journal Vouchers Summary", False
            Unload Me
        Else
            ShowNoRecord
        End If
        Call NEW_LogAudit("V", "GJSummary Report", "", "", "", dtpFrom & " " & dtpTo, "", "")
    End If
    If REPORT_RANGETYPE = "REC_REGISTER" Then
        'If Function_Access(LOGID, "Acess_Print", "") = False Then Exit Sub
        Set rsJournal_HD = New ADODB.Recordset
        rsJournal_HD.Open "select * from AMIS_Journal_HD where jtype = 'APJ' and (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
            ShowRangeReport dtpFrom, dtpTo, "ReceiptsRegisters", "Registers", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "RECEIVING REPORT REGISTERS", False
            Unload Me
        Else
            ShowNoRecord
        End If
        Call NEW_LogAudit("V", "REC REGISTER REPORT", "", "", "", dtpFrom & " " & dtpTo, "", "")
    End If
    If REPORT_RANGETYPE = "CHECK_REGISTER" Then
        ' If Function_Access(LOGID, "Acess_Print", "") = False Then Exit Sub
        Set rsJournal_HD = New ADODB.Recordset
        rsJournal_HD.Open "select * from AMIS_Journal_HD where jtype = 'CDJ' and (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
            ShowRangeReport dtpFrom, dtpTo, "CheckRegisters", "Registers", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "CHECK REGISTERS", False
            Unload Me
        Else
            ShowNoRecord
        End If
        Call NEW_LogAudit("V", "CHECK REGISTER REPORT", "", "", "", dtpFrom & " " & dtpTo, "", "")
    End If
    If REPORT_RANGETYPE = "INV_REGISTER" Then

        Set rsJournal_HD = New ADODB.Recordset
        rsJournal_HD.Open "select * from AMIS_Journal_HD where jtype = 'SJ' and (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
            ShowRangeReport dtpFrom, dtpTo, "SalesInvoiceRegisters", "Registers", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "INVOICES REGISTERS", False
            Unload Me
        Else
            ShowNoRecord
        End If
        Call NEW_LogAudit("V", "INV REGISTER REPORT", "", "", "", dtpFrom & " " & dtpTo, "", "")
    End If
    If REPORT_RANGETYPE = "OR_REGISTER" Then
        'If Function_Access(LOGID, "Acess_Print", "") = False Then Exit Sub
        Set rsJournal_HD = New ADODB.Recordset
        rsJournal_HD.Open "select * from AMIS_Journal_HD where jtype = 'CRJ' and (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
            ShowRangeReport dtpFrom, dtpTo, "ORRegisters", "Registers", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "O.R. REGISTERS", False
            Unload Me
        Else
            ShowNoRecord
        End If
        Call NEW_LogAudit("V", "OR REGISTER REPORT", "", "", "", dtpFrom & " " & dtpTo, "", "")
    End If
    If REPORT_RANGETYPE = "EX_TAX" Then
        'If Function_Access(LOGID, "Acess_Print", "") = False Then Exit Sub
        Set rsJournal_HD = New ADODB.Recordset
        rsJournal_HD.Open "select * from AMIS_Journal_Det where jtype = 'APJ' and (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
            ShowRangeReport dtpFrom, dtpTo, "ScheduleOfExpandedWTax", "Schedules", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", "ALPHALIST OF PAYEES SUBJECT TO EXPANDED WITHHOLDING TAX", False
            Unload Me
        Else
            ShowNoRecord
        End If
    End If
    ' Update By BTT 07242008
    If REPORT_RANGETYPE = "CANCEL APJ" Then
        'Update By BTT : 06/05/2008
        Dim SQL1                                            As String
        Dim RS                                              As New ADODB.Recordset

        SQL1 = "SELECT * from ALL_CANCEL_TRANSACTION where date_cancelled>='" & CDate(dtpFrom) & "' and date_cancelled<='" & CDate(dtpTo) & "' and application_type='APJ'"

        Set RS = New ADODB.Recordset
        Set RS = gconDMIS.Execute(SQL1)

        If Not RS.EOF And Not RS.BOF Then
            Screen.MousePointer = 11
            rptAMISrange.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptAMISrange.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptAMISrange.Formulas(2) = "fromjdate ='" & dtpFrom & "'"
            rptAMISrange.Formulas(3) = "tojdate ='" & dtpTo & "'"
            rptAMISrange.WindowTitle = "Cancelled Report"
            PrintSQLReport rptAMISrange, AMIS_REPORT_PATH & "CancelReport.rpt", "{ALL_CANCEL_TRANSACTION.DATE_CANCELLED} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {ALL_CANCEL_TRANSACTION.DATE_CANCELLED} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & " )AND {ALL_CANCEL_TRANSACTION.application_type}='APJ' ", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
            LogAudit "V", "Cancelled Report", dtpFrom & "-" & dtpTo
        Else
            ShowNoRecord
        End If
        Call NEW_LogAudit("V", "Cancelled APJ Report", "", "", "", dtpFrom & " " & dtpTo, "", "")
    End If
    If REPORT_RANGETYPE = "CANCEL CDJ" Then
        'Update By BTT : 06/05/2008
        Dim SQL2                                            As String
        Dim rs1                                             As New ADODB.Recordset

        SQL2 = "SELECT * from ALL_CANCEL_TRANSACTION where date_cancelled>='" & CDate(dtpFrom) & "' and date_cancelled<='" & CDate(dtpTo) & "' and application_type='CDJ'"

        Set rs1 = New ADODB.Recordset
        Set rs1 = gconDMIS.Execute(SQL2)

        If Not rs1.EOF And Not rs1.BOF Then
            Screen.MousePointer = 11
            rptAMISrange.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptAMISrange.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptAMISrange.Formulas(2) = "fromjdate ='" & dtpFrom & "'"
            rptAMISrange.Formulas(3) = "tojdate ='" & dtpTo & "'"
            rptAMISrange.WindowTitle = "Cancelled Report"
            PrintSQLReport rptAMISrange, AMIS_REPORT_PATH & "CancelReport.rpt", "{ALL_CANCEL_TRANSACTION.DATE_CANCELLED} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {ALL_CANCEL_TRANSACTION.DATE_CANCELLED} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & " )AND {ALL_CANCEL_TRANSACTION.application_type}='CDJ' ", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
            LogAudit "V", "Cancelled Report", dtpFrom & "-" & dtpTo
        Else
            ShowNoRecord
        End If
    End If
    If REPORT_RANGETYPE = "CANCEL SJ" Then
        'Update By BTT : 06/05/2008
        Dim SQL3                                            As String
        Dim rs2                                             As New ADODB.Recordset

        SQL3 = "SELECT * from ALL_CANCEL_TRANSACTION where date_cancelled>='" & CDate(dtpFrom) & "' and date_cancelled<='" & CDate(dtpTo) & "' and application_type='SJ'"

        Set rs2 = New ADODB.Recordset
        Set rs2 = gconDMIS.Execute(SQL3)

        If Not rs2.EOF And Not rs2.BOF Then
            Screen.MousePointer = 11
            rptAMISrange.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptAMISrange.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptAMISrange.Formulas(2) = "fromjdate ='" & dtpFrom & "'"
            rptAMISrange.Formulas(3) = "tojdate ='" & dtpTo & "'"
            rptAMISrange.WindowTitle = "Cancelled Report"
            PrintSQLReport rptAMISrange, AMIS_REPORT_PATH & "CancelReport.rpt", "{ALL_CANCEL_TRANSACTION.DATE_CANCELLED} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {ALL_CANCEL_TRANSACTION.DATE_CANCELLED} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & " )AND {ALL_CANCEL_TRANSACTION.application_type}='SJ' ", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
            LogAudit "V", "Cancelled Report", dtpFrom & "-" & dtpTo
        Else
            ShowNoRecord
        End If
    End If
    If REPORT_RANGETYPE = "CANCEL CRJ" Then
        'Update By BTT : 06/05/2008
        Dim SQL4                                            As String
        Dim rs3                                             As New ADODB.Recordset

        SQL4 = "SELECT * from ALL_CANCEL_TRANSACTION where date_cancelled>='" & CDate(dtpFrom) & "' and date_cancelled<='" & CDate(dtpTo) & "' and application_type='CRJ'"

        Set rs3 = New ADODB.Recordset
        Set rs3 = gconDMIS.Execute(SQL4)

        If Not rs3.EOF And Not rs3.BOF Then
            Screen.MousePointer = 11
            rptAMISrange.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptAMISrange.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptAMISrange.Formulas(2) = "fromjdate ='" & dtpFrom & "'"
            rptAMISrange.Formulas(3) = "tojdate ='" & dtpTo & "'"
            rptAMISrange.WindowTitle = "Cancelled Report"
            PrintSQLReport rptAMISrange, AMIS_REPORT_PATH & "CancelReport.rpt", "{ALL_CANCEL_TRANSACTION.DATE_CANCELLED} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {ALL_CANCEL_TRANSACTION.DATE_CANCELLED} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & " )AND {ALL_CANCEL_TRANSACTION.application_type}='CRJ' ", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
            LogAudit "V", "Cancelled Report", dtpFrom & "-" & dtpTo
        Else
            ShowNoRecord
        End If
    End If
    If REPORT_RANGETYPE = "CANCEL GJ" Then
        'Update By BTT : 06/05/2008
        Dim SQL5                                            As String
        Dim rs4                                             As New ADODB.Recordset

        SQL5 = "SELECT * from ALL_CANCEL_TRANSACTION where date_cancelled>='" & CDate(dtpFrom) & "' and date_cancelled<='" & CDate(dtpTo) & "' and application_type='GJ'"

        Set rs4 = New ADODB.Recordset
        Set rs4 = gconDMIS.Execute(SQL5)

        If Not rs4.EOF And Not rs4.BOF Then
            Screen.MousePointer = 11
            rptAMISrange.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptAMISrange.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptAMISrange.Formulas(2) = "fromjdate ='" & dtpFrom & "'"
            rptAMISrange.Formulas(3) = "tojdate ='" & dtpTo & "'"
            rptAMISrange.WindowTitle = "Cancelled Report"
            PrintSQLReport rptAMISrange, AMIS_REPORT_PATH & "CancelReport.rpt", "{ALL_CANCEL_TRANSACTION.DATE_CANCELLED} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {ALL_CANCEL_TRANSACTION.DATE_CANCELLED} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & " )AND {ALL_CANCEL_TRANSACTION.application_type}='GJ' ", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
            LogAudit "V", "Cancelled Report", dtpFrom & "-" & dtpTo
        Else
            ShowNoRecord
        End If
    End If

    LogAudit "V", "JOURNAL VOUCHER SUMMARY", dtpFrom & "-" & dtpTo
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Update by BTT
    MoveKeyPress KeyCode
    Select Case KeyCode
    Case vbKeyF1 And Shift = 1:
        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        Unload frmALL_AuditInquiry
        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        If REPORT_RANGETYPE = "CHECK_REGISTER" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (CHECK REGISTER REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "CHECK REGISTER REPORT", "PRINTING")
        ElseIf REPORT_RANGETYPE = "INV_REGISTER" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (INV REGISTER REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "INV REGISTER REPORT", "PRINTING")
        ElseIf REPORT_RANGETYPE = "OR_REGISTER" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (OR REGISTER REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "OR REGISTER REPORT", "PRINTING")
        ElseIf REPORT_RANGETYPE = "CANCEL APJ" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (Cancelled APJ Report)"
            Call frmALL_AuditInquiry.DisplayHistory("", "Cancelled APJ Report", "PRINTING")
        ElseIf REPORT_RANGETYPE = "REC_REGISTER" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (REC REGISTER REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "REC REGISTER REPORT", "PRINTING")
        ElseIf REPORT_RANGETYPE = "JVS" Then
            frmALL_AuditInquiry.Caption = "Audit Inquiry (GJSummary Report)"
            Call frmALL_AuditInquiry.DisplayHistory("", "GJSummary Report", "PRINTING")
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
    LocalAcess = ""
    Set frmAMISRange = Nothing
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

