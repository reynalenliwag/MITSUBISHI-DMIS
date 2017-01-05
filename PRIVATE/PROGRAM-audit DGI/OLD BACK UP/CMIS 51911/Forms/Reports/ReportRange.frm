VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCMISReportRange 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   ForeColor       =   &H00DEDFDE&
   Icon            =   "ReportRange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1530
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
      Left            =   2490
      MouseIcon       =   "ReportRange.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportRange.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Close Window"
      Top             =   630
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
      Left            =   1620
      MouseIcon       =   "ReportRange.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportRange.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Print Report"
      Top             =   630
      Width           =   885
   End
   Begin Crystal.CrystalReport rptCMISReportRange 
      Left            =   120
      Top             =   960
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
      TabIndex        =   0
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
      Format          =   20250625
      CurrentDate     =   38216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   405
      Left            =   3030
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
      Format          =   20250625
      CurrentDate     =   38216
   End
   Begin VB.TextBox txtTo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3030
      TabIndex        =   3
      Top             =   90
      Width           =   1695
   End
   Begin VB.TextBox txtFrom 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   780
      TabIndex        =   2
      Top             =   90
      Width           =   1695
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   180
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2550
      TabIndex        =   5
      Top             =   180
      Width           =   435
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   4
      Top             =   2970
      Width           =   495
   End
End
Attribute VB_Name = "frmCMISReportRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If dtpFrom > dtpTo Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If

    With rptCMISReportRange
        .Formulas(0) = "DEALER_NAME = '" & COMPANY_NAME & "'"
        .Formulas(1) = "DEALER_ADDRESS = '" & COMPANY_ADDRESS & "'"
        .Formulas(2) = "PREPAREDBY= '" & PreparedBy & "'"
        .Formulas(3) = "NOTEDBY= '" & NotedBy & "'"
        .Formulas(4) = "CHECKEDBY= '" & CheckedBy & "'"
        .Formulas(5) = "PRINTEDBY= " & N2Str2Null(LOGNAME)
    End With

    If CMIS_Report_Range = "Cash In Flow Report" Then
        frmCMISTypeOfReport.Show vbModal
        If CMIS_Type_Of_Report = "SUMMARY" Then
            PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH & "Cash_In_Flow_Summary.rpt", "{OFF_HD.OR_DATE} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") and {OFF_HD.OR_DATE} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
            'NEW LOG AUDIT---------------------------------------------------------
                Call NEW_LogAudit("V", "REPORT CASH INFLOW PER DATE OF RANGE", "", "", "", "DATE RANGE: " & dtpFrom.Value & "-" & dtpTo.Value & " - SUMMARY", "", "")
            'NEW LOG AUDIT---------------------------------------------------------
        End If
        If CMIS_Type_Of_Report = "DETAILED" Then
            PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH & "Cash_In_Flow_Detailed.rpt", "{OFF_HD.OR_DATE} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") and {OFF_HD.OR_DATE} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
            'NEW LOG AUDIT---------------------------------------------------------
                Call NEW_LogAudit("V", "REPORT CASH INFLOW PER DATE OF RANGE", "", "", "", "DATE RANGE: " & dtpFrom.Value & "-" & dtpTo.Value & " - DETAILED", "", "")
            'NEW LOG AUDIT---------------------------------------------------------
        End If
    End If

    If CMIS_Report_Range = "Customer Over-Payment Report" Then
        PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH & "Customer_Over_Payment_Report.rpt", "{OFF_DT.ORDATE} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") and {OFF_DT.ORDATE} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
        'NEW LOG AUDIT---------------------------------------------------------
            Call NEW_LogAudit("V", "REPORT CUSTOMER OVER PAYMENT REPORT", "", "", "", "DATE RANGE: " & dtpFrom.Value & "-" & dtpTo.Value & " - DETAILED", "", "")
        'NEW LOG AUDIT---------------------------------------------------------
    End If

    If CMIS_Report_Range = "Card Listing Report On Hand" Then
        PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH & "Summary_Of_Credit_Card.rpt", "{OFF_HD.DEPOSIT} = false and {OFF_HD.OR_DATE} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") and {OFF_HD.OR_DATE} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
        'NEW LOG AUDIT---------------------------------------------------------
            Call NEW_LogAudit("V", "REPORT CARD LISTING REPORT ON-HAND", "", "", "", "DATE RANGE: " & dtpFrom.Value & "-" & dtpTo.Value & " - DETAILED", "", "")
        'NEW LOG AUDIT---------------------------------------------------------
    End If

    If CMIS_Report_Range = "Credit Card Bank Deposit Report" Then
        PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH & "Credit_Card_Bank_Deposit_Summary.rpt", "{CMIS_BankDepo.DATDEPOSIT} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") and {CMIS_BankDepo.DATDEPOSIT} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
        'NEW LOG AUDIT---------------------------------------------------------
            Call NEW_LogAudit("V", "REPORT CARD BANK DEPOSIT REPORT", "", "", "", "DATE RANGE: " & dtpFrom.Value & "-" & dtpTo.Value & " - DETAILED", "", "")
        'NEW LOG AUDIT---------------------------------------------------------
    End If

    If CMIS_Report_Range = "Daily Transmittal Report" Then
        frmCMISTypeOfReport.Show vbModal
        If CMIS_Type_Of_Report = "SUMMARY" Then
            'PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH, "{OFF"
        End If
        If CMIS_Type_Of_Report = "DETAILED" Then
            'PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH, "{OFF"
        End If
    End If

    If CMIS_Report_Range = "Cash Receipts Journal Report - Summary" Then
        PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH & "Cash_Receipts_System_Summary.rpt", "{OFF_HD.OR_DATE} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") and {OFF_HD.OR_DATE} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
    End If

    If CMIS_Report_Range = "Cash Receipts Journal Report - Detail" Then
        PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH & "Cash_Receipts_System_Detailed.rpt", "{OFF_HD.OR_DATE} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") and {OFF_HD.OR_DATE} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
    End If

    If CMIS_Report_Range = "Cash On Hand Report" Then
    End If

    If CMIS_Report_Range = "OutPut Tax Report" Then
    End If

    If CMIS_Report_Range = "Corporate Tax Report" Then
    End If

    If CMIS_Report_Range = "Sales Discount Report" Then
    End If

    If CMIS_Report_Range = "Journal with Reference Code" Then
    End If

    If CMIS_Report_Range = "Petty Cash Summary Report" Then
        'PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH & "Petty_Cash_Replenishment_Detailed.rpt", "{PETTY.PETTY_DATE} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") and {PETTY.PETTY_DATE} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
        'PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH & "Petty_Cash_Replenishment_Summary.rpt", "{PETTY.PETTY_DATE} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") and {PETTY.PETTY_DATE} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
        rptCMISReportRange.WindowTitle = "Petty Cash Replenishment Report"
        PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH & "Petty_Cash_Replenishment_Detailed.rpt", "{PETTY.PCF_NUMBER} >= '" & Format(txtFrom.Text, "000000") & "' and {PETTY.PCF_NUMBER} <= '" & Format(txtTo.Text, "000000") & "'", DMIS_REPORT_Connection, 1
        PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH & "Petty_Cash_Replenishment_Summary.rpt", "{PETTY.PCF_NUMBER} >= '" & Format(txtFrom.Text, "000000") & "' and {PETTY.PCF_NUMBER} <= '" & Format(txtTo.Text, "000000") & "'", DMIS_REPORT_Connection, 1
        
        'NEW LOG AUDIT---------------------------------------------------------
            Call NEW_LogAudit("V", "REPORT REPLENISH PETTY CASH SUMMARY REPORT", "", "", "", "DATE RANGE: " & dtpFrom.Value & "-" & dtpTo.Value & " - DETAILED", "", "")
        'NEW LOG AUDIT---------------------------------------------------------
    End If

    If CMIS_Report_Range = "L.T.O. Summary Report" Then
        rptCMISReportRange.WindowTitle = "L.T.O. Replenishment Report"
        PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH & "LTO_Replenishment_Detailed.rpt", "{PETTY.PCF_NUMBER} >= '" & Format(txtFrom.Text, "000000") & "' and {PETTY.PCF_NUMBER} <= '" & Format(txtTo.Text, "000000") & "'", DMIS_REPORT_Connection, 1
        PrintSQLReport rptCMISReportRange, CMIS_REPORT_PATH & "LTO_Replenishment_Summary.rpt", "{PETTY.PCF_NUMBER} >= '" & Format(txtFrom.Text, "000000") & "' and {PETTY.PCF_NUMBER} <= '" & Format(txtTo.Text, "000000") & "'", DMIS_REPORT_Connection, 1
        
        'NEW LOG AUDIT---------------------------------------------------------
            Call NEW_LogAudit("V", "REPORT REPLENISH LTO SUMMARY REPORT", "", "", "", "DATE RANGE: " & dtpFrom.Value & "-" & dtpTo.Value & " - DETAILED", "", "")
        'NEW LOG AUDIT---------------------------------------------------------
    End If

    'LogAudit "V", CMIS_Report_Range, "DATE RANGE: " & dtpFrom.Value & " - " & dtpTo.Value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
             
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            If CMIS_Report_Range = "Cash In Flow Report" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (REPORT CASH INFLOW PER DATE OF RANGE)"
                Call frmALL_AuditInquiry.DisplayHistory("", "REPORT CASH INFLOW PER DATE OF RANGE", "PRINTING")
            ElseIf CMIS_Report_Range = "Customer Over-Payment Report" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (REPORT CUSTOMER OVER PAYMENT REPORT)"
                Call frmALL_AuditInquiry.DisplayHistory("", "REPORT CUSTOMER OVER PAYMENT REPORT", "PRINTING")
            ElseIf CMIS_Report_Range = "Cash On Hand Report" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (REPORT CASHRECIEPT CASH ON HAND)"
                Call frmALL_AuditInquiry.DisplayHistory("", "REPORT CASHRECIEPT CASH ON HAND", "PRINTING")
            ElseIf CMIS_Report_Range = "Card Listing Report On Hand" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (REPORT CARD LISTING REPORT ON-HAND)"
                Call frmALL_AuditInquiry.DisplayHistory("", "REPORT CARD LISTING REPORT ON-HAND", "PRINTING")
            ElseIf CMIS_Report_Range = "Credit Card Bank Deposit Report" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (REPORT CARD BANK DEPOSIT REPORT)"
                Call frmALL_AuditInquiry.DisplayHistory("", "REPORT CARD BANK DEPOSIT REPORT", "PRINTING")
            ElseIf CMIS_Report_Range = "OutPut Tax Report" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (REPORT CASHRECIEPT OUT PUT TAX)"
                Call frmALL_AuditInquiry.DisplayHistory("", "REPORT CASHRECIEPT OUT PUT TAX", "PRINTING")
            ElseIf CMIS_Report_Range = "Corporate Tax Report" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (REPORT CASHRECIEPT CORPORATE TAX)"
                Call frmALL_AuditInquiry.DisplayHistory("", "REPORT CASHRECIEPT CORPORATE TAX", "PRINTING")
            ElseIf CMIS_Report_Range = "Sales Discount Report" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (REPORT CASHRECIEPT SALES DISCOUNT)"
                Call frmALL_AuditInquiry.DisplayHistory("", "REPORT CASHRECIEPT SALES DISCOUNT", "PRINTING")
            ElseIf CMIS_Report_Range = "Petty Cash Summary Report" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (REPORT REPLENISH PETTY CASH SUMMARY REPORT)"
                Call frmALL_AuditInquiry.DisplayHistory("", "REPORT REPLENISH PETTY CASH SUMMARY REPORT", "PRINTING")
            ElseIf CMIS_Report_Range = "L.T.O. Summary Report" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (REPORT REPLENISH LTO SUMMARY REPORT)"
                Call frmALL_AuditInquiry.DisplayHistory("", "REPORT REPLENISH LTO SUMMARY REPORT", "PRINTING")
            Else
            
            End If
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpFrom.Value = firstDay(Now)
    dtpTo.Value = Now

    Me.Caption = CMIS_Report_Range
    If CMIS_Report_Range = "Petty Cash Summary Report" Then
        txtFrom.Visible = True: txtTo.Visible = True
        dtpFrom.Visible = False: dtpTo.Visible = False
    ElseIf CMIS_Report_Range = "L.T.O. Summary Report" Then
        txtFrom.Visible = True: txtTo.Visible = True
        dtpFrom.Visible = False: dtpTo.Visible = False
    Else
        txtFrom.Visible = False: txtTo.Visible = False
        dtpFrom.Visible = True: dtpTo.Visible = True
    End If
    Dim rsPCF_NUMBER                                                  As ADODB.Recordset
    If CMIS_Report_Range = "Petty Cash Summary Report" Then
        Set rsPCF_NUMBER = New ADODB.Recordset
        Set rsPCF_NUMBER = gconDMIS.Execute("Select * from CMIS_Petty Where PCF_NUMBER IS NOT NULL Order by PCF_NUMBER ASC")
        If Not rsPCF_NUMBER.EOF And Not rsPCF_NUMBER.BOF Then
            txtFrom.Text = Format(Null2String(rsPCF_NUMBER!PCF_NUMBER), "000000")
            rsPCF_NUMBER.MoveLast
            txtTo.Text = Format(Null2String(rsPCF_NUMBER!PCF_NUMBER), "000000")
        End If
    End If
    If CMIS_Report_Range = "L.T.O. Summary Report" Then
        Set rsPCF_NUMBER = New ADODB.Recordset
        Set rsPCF_NUMBER = gconDMIS.Execute("Select * from CMIS_LTOPondo Where PCF_NUMBER IS NOT NULL Order by PCF_NUMBER ASC")
        If Not rsPCF_NUMBER.EOF And Not rsPCF_NUMBER.BOF Then
            txtFrom.Text = Format(Null2String(rsPCF_NUMBER!PCF_NUMBER), "000000")
            rsPCF_NUMBER.MoveLast
            txtTo.Text = Format(Null2String(rsPCF_NUMBER!PCF_NUMBER), "000000")
        End If
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCMISReportRange = Nothing
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

