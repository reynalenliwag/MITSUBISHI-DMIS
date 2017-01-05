VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPMISReports_PartsIssuedToCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Issued Report"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   ForeColor       =   &H00DEDFDE&
   Icon            =   "PartsIssuedToCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3480
   ScaleWidth      =   3705
   Begin VB.OptionButton optByPartNumber 
      Caption         =   "By Part Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   180
      TabIndex        =   11
      Top             =   360
      Width           =   3105
   End
   Begin VB.OptionButton optCustomer 
      Caption         =   "By Customer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   10
      Top             =   60
      Value           =   -1  'True
      Width           =   3105
   End
   Begin VB.ComboBox cboTranPartNo 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   180
      TabIndex        =   0
      Text            =   "cboTranPartNo"
      Top             =   1050
      Width           =   3405
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   345
      Left            =   810
      TabIndex        =   2
      Top             =   1890
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52625409
      CurrentDate     =   39232
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   345
      Left            =   810
      TabIndex        =   1
      Top             =   1500
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   52625409
      CurrentDate     =   39203
   End
   Begin VB.CheckBox chkHistReceipts 
      Caption         =   "Look in History File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   810
      TabIndex        =   3
      Top             =   2280
      Width           =   2655
   End
   Begin Crystal.CrystalReport rptReceipts 
      Left            =   90
      Top             =   2700
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Transaction Listing - Receipts"
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
      Height          =   795
      Left            =   1860
      MouseIcon       =   "PartsIssuedToCustomer.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "PartsIssuedToCustomer.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   2610
      Width           =   675
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
      Height          =   795
      Left            =   1200
      MouseIcon       =   "PartsIssuedToCustomer.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "PartsIssuedToCustomer.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   2610
      Width           =   675
   End
   Begin VB.Label Label34 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   180
      TabIndex        =   9
      Top             =   780
      Width           =   1515
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
      Left            =   0
      TabIndex        =   8
      Top             =   1560
      Width           =   765
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
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   765
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1800
      TabIndex        =   6
      Top             =   2970
      Width           =   495
   End
End
Attribute VB_Name = "frmPMISReports_PartsIssuedToCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRR_HD                                                           As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "PARTS ISSUED TO CUSTOMER") = False Then Exit Sub
    On Error GoTo ERRORCODE:

    Dim FDate                                                         As Date
    Dim TDate                                                         As Date

    FDate = CDate(dtpFromDate.Value)
    TDate = CDate(dtpToDate.Value)
    
    If optByPartNumber.Value = True Then
            If chkHistReceipts.Value = 1 Then
                Set rsRR_HD = New ADODB.Recordset
                rsRR_HD.Open "select * from PMIS_ORD_Hist where TYPE = '" & PARTS_ISSUED_TO_CUSTOMER_TYPE & "' AND (TRANDATE >= '" & FDate & "' AND TRANDATE <= '" & TDate & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not rsRR_HD.EOF And Not rsRR_HD.EOF Then
                    Screen.MousePointer = 11
                    rptReceipts.WindowTitle = "STOCK ISSUED TO CUSTOMER - HISTORY" '
                    rptReceipts.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                    rptReceipts.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                    rptReceipts.Formulas(12) = "mindate = '" & FDate & "'"
                    rptReceipts.Formulas(11) = "maxdate = '" & TDate & "'"
                    If LTrim(RTrim(cboTranPartNo.Text)) = "ALL" Then
                        PrintSQLReport rptReceipts, PMIS_REPORT_PATH & "PartsIssuedtoCustomer_Hist.rpt", "{ORD_HD.TYPE} = '" & PARTS_ISSUED_TO_CUSTOMER_TYPE & "' AND {ORD_HD.TRANDATE} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {ORD_HD.TRANDATE} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", DMIS_REPORT_Connection, 1
                        Screen.MousePointer = 0
                    Else
                        PrintSQLReport rptReceipts, PMIS_REPORT_PATH & "PartsIssuedtoCustomer_Hist.rpt", "{ORD_HD.TYPE} = '" & PARTS_ISSUED_TO_CUSTOMER_TYPE & "' AND {ORD_HD.TRANDATE} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {ORD_HD.TRANDATE} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ") AND {TDAYTRAN.STOCK_ORD} = " & N2Str2Null(Trim(cboTranPartNo.Text)), DMIS_REPORT_Connection, 1
                        Screen.MousePointer = 0
                    End If
                    NEW_LogAudit "V", "PARTS ISSUED TO CUSTOMER", "", "", "", dtpFromDate & " - " & dtpToDate, "HISTORY", ""
                Else
                    ShowNoRecord
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            Else
                Set rsRR_HD = New ADODB.Recordset
                rsRR_HD.Open "select * from PMIS_ORD_Hd where TYPE = '" & PARTS_ISSUED_TO_CUSTOMER_TYPE & "' AND (TRANDATE >= '" & FDate & "' AND TRANDATE <= '" & TDate & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not rsRR_HD.EOF And Not rsRR_HD.EOF Then
                    Screen.MousePointer = 11
                    rptReceipts.WindowTitle = "STOCK ISSUED TO CUSTOMER"
                    rptReceipts.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                    rptReceipts.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                    rptReceipts.Formulas(12) = "mindate = '" & FDate & "'"
                    rptReceipts.Formulas(11) = "maxdate = '" & TDate & "'"
                    If LTrim(RTrim(cboTranPartNo.Text)) = "ALL" Then
                        PrintSQLReport rptReceipts, PMIS_REPORT_PATH & "PartsIssuedtoCustomer.rpt", "{ORD_HD.TYPE} = '" & PARTS_ISSUED_TO_CUSTOMER_TYPE & "' AND {ORD_HD.TRANDATE} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {ORD_HD.TRANDATE} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", DMIS_REPORT_Connection, 1
                        Screen.MousePointer = 0
                    Else
                        PrintSQLReport rptReceipts, PMIS_REPORT_PATH & "PartsIssuedtoCustomer.rpt", "{ORD_HD.TYPE} = '" & PARTS_ISSUED_TO_CUSTOMER_TYPE & "' AND {ORD_HD.TRANDATE} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {ORD_HD.TRANDATE} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ") AND {TDAYTRAN.STOCK_ORD} = " & N2Str2Null(Trim(cboTranPartNo.Text)), DMIS_REPORT_Connection, 1
                        Screen.MousePointer = 0
                    End If
                    NEW_LogAudit "V", "PARTS ISSUED TO CUSTOMER", "", "", "", dtpFromDate & " - " & dtpToDate, "", ""
                    Screen.MousePointer = 0
                Else
                    ShowNoRecord
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        Else
            If chkHistReceipts.Value = 1 Then
                Set rsRR_HD = New ADODB.Recordset
                If LTrim(RTrim(cboTranPartNo.Text)) = "ALL" Then
                    rsRR_HD.Open "select * from PMIS_ORD_Hist where TYPE = '" & PARTS_ISSUED_TO_CUSTOMER_TYPE & "' AND (TRANDATE >= '" & FDate & "' AND TRANDATE <= '" & TDate & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
                    If Not rsRR_HD.EOF And Not rsRR_HD.EOF Then
                        Screen.MousePointer = 11
                        rptReceipts.WindowTitle = "CUSTOMER ISSUED STOCK - HISTORY" '
                        rptReceipts.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                        rptReceipts.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                        rptReceipts.Formulas(12) = "mindate = '" & FDate & "'"
                        rptReceipts.Formulas(11) = "maxdate = '" & TDate & "'"
                        PrintSQLReport rptReceipts, PMIS_REPORT_PATH & "CustomerIssuedParts_Hist.rpt", "{ORD_HD.TYPE} = '" & PARTS_ISSUED_TO_CUSTOMER_TYPE & "' AND {ORD_HD.TRANDATE} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {ORD_HD.TRANDATE} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", DMIS_REPORT_Connection, 1
                        Screen.MousePointer = 0
                        NEW_LogAudit "V", "CUSTOMER ISSUED STOCK", "", "", "", dtpFromDate & " - " & dtpToDate, "HISTORY", ""
                    Else
                        ShowNoRecord
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                Else
                    rsRR_HD.Open "select * from PMIS_ORD_Hist where TYPE = '" & PARTS_ISSUED_TO_CUSTOMER_TYPE & "' AND (TRANDATE >= '" & FDate & "' AND TRANDATE <= '" & TDate & "') AND CUSTNAME = " & N2Str2Null(Trim(cboTranPartNo.Text)), gconDMIS, adOpenForwardOnly, adLockReadOnly
                    If Not rsRR_HD.EOF And Not rsRR_HD.EOF Then
                        Screen.MousePointer = 11
                        rptReceipts.WindowTitle = "CUSTOMER ISSUED STOCK - HISTORY" '
                        rptReceipts.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                        rptReceipts.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                        rptReceipts.Formulas(12) = "mindate = '" & FDate & "'"
                        rptReceipts.Formulas(11) = "maxdate = '" & TDate & "'"
                        PrintSQLReport rptReceipts, PMIS_REPORT_PATH & "CustomerIssuedParts_Hist.rpt", "{ORD_HD.TYPE} = '" & PARTS_ISSUED_TO_CUSTOMER_TYPE & "' AND {ORD_HD.TRANDATE} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {ORD_HD.TRANDATE} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ") AND {ORD_HD.CUSTNAME} = " & N2Str2Null(Trim(cboTranPartNo.Text)), DMIS_REPORT_Connection, 1
                        Screen.MousePointer = 0
                        NEW_LogAudit "V", "CUSTOMER ISSUED STOCK", "", "", "", dtpFromDate & " - " & dtpToDate, "HISTORY", ""
                    Else
                        ShowNoRecord
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                End If
            Else
                Set rsRR_HD = New ADODB.Recordset
                If LTrim(RTrim(cboTranPartNo.Text)) = "ALL" Then
                    rsRR_HD.Open "select * from PMIS_ORD_Hd where TYPE = '" & PARTS_ISSUED_TO_CUSTOMER_TYPE & "' AND (TRANDATE >= '" & FDate & "' AND TRANDATE <= '" & TDate & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
                    If Not rsRR_HD.EOF And Not rsRR_HD.EOF Then
                        Screen.MousePointer = 11
                        rptReceipts.WindowTitle = "CUSTOMER ISSUED STOCK"
                        rptReceipts.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                        rptReceipts.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                        rptReceipts.Formulas(12) = "mindate = '" & FDate & "'"
                        rptReceipts.Formulas(11) = "maxdate = '" & TDate & "'"
                        PrintSQLReport rptReceipts, PMIS_REPORT_PATH & "CustomerIssuedParts.rpt", "{ORD_HD.TYPE} = '" & PARTS_ISSUED_TO_CUSTOMER_TYPE & "' AND {ORD_HD.TRANDATE} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {ORD_HD.TRANDATE} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", DMIS_REPORT_Connection, 1
                        Screen.MousePointer = 0
                        NEW_LogAudit "V", "CUSTOMER ISSUED STOCK", "", "", "", dtpFromDate & " - " & dtpToDate, "", ""
                    Else
                        ShowNoRecord
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                Else
                    rsRR_HD.Open "select * from PMIS_ORD_Hd where TYPE = '" & PARTS_ISSUED_TO_CUSTOMER_TYPE & "' AND (TRANDATE >= '" & FDate & "' AND TRANDATE <= '" & TDate & "') AND CUSTNAME = " & N2Str2Null(Trim(cboTranPartNo.Text)), gconDMIS, adOpenForwardOnly, adLockReadOnly
                    If Not rsRR_HD.EOF And Not rsRR_HD.EOF Then
                        Screen.MousePointer = 11
                        rptReceipts.WindowTitle = "CUSTOMER ISSUED STOCK"
                        rptReceipts.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                        rptReceipts.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                        rptReceipts.Formulas(12) = "mindate = '" & FDate & "'"
                        rptReceipts.Formulas(11) = "maxdate = '" & TDate & "'"
                        PrintSQLReport rptReceipts, PMIS_REPORT_PATH & "CustomerIssuedParts.rpt", "{ORD_HD.TYPE} = '" & PARTS_ISSUED_TO_CUSTOMER_TYPE & "' AND {ORD_HD.TRANDATE} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {ORD_HD.TRANDATE} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ") AND {ORD_HD.CUSTNAME} = " & N2Str2Null(Trim(cboTranPartNo.Text)), DMIS_REPORT_Connection, 1
                        Screen.MousePointer = 0
                        NEW_LogAudit "V", "CUSTOMER ISSUED STOCK", "", "", "", dtpFromDate & " - " & dtpToDate, "", ""
                    Else
                        ShowNoRecord
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                End If
            End If
    End If
    Exit Sub
ERRORCODE:
    ShowVBError
    Screen.MousePointer = 0
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (PARTS ISSUED TO CUSTOMER)"
            Call frmALL_AuditInquiry.DisplayHistory("", "PARTS ISSUED TO CUSTOMER", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    LOADCUSTOMER
    dtpFromDate.Value = firstDay(LOGDATE)
    dtpToDate.Value = LOGDATE
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReports_PartsIssuedToCustomer = Nothing
    UnloadForm Me
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub
Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Sub LOADPARTS()
    Dim rsParts As ADODB.Recordset
    Set rsParts = New ADODB.Recordset
    rsParts.Open "SELECT STOCKNO FROM PMIS_STOCKMAS WHERE TYPE = '" & PARTS_ISSUED_TO_CUSTOMER_TYPE & "' ORDER BY STOCKNO ASC", gconDMIS, adOpenKeyset
    cboTranPartNo.Clear
    If Not rsParts.EOF And Not rsParts.BOF Then
        rsParts.MoveFirst
        Do While Not rsParts.EOF
            cboTranPartNo.AddItem Null2String(rsParts!STOCKNO)
            rsParts.MoveNext
        Loop
    End If
    cboTranPartNo.AddItem "ALL", 0
    cboTranPartNo.ListIndex = 0
    Set rsParts = Nothing
End Sub

Private Sub optByPartNumber_Click()
    LOADPARTS
End Sub

Private Sub optCustomer_Click()
    LOADCUSTOMER
End Sub

Sub LOADCUSTOMER()
    Dim rsOrdCustomer  As ADODB.Recordset
    Set rsOrdCustomer = New ADODB.Recordset
    rsOrdCustomer.Open "SELECT DISTINCT CUSTNAME FROM PMIS_vw_ISS_HISTORY WHERE (STATUS = 'P' OR STATUS = 'B') AND TYPE = '" & PARTS_ISSUED_TO_CUSTOMER_TYPE & "' ORDER BY CUSTNAME ASC", gconDMIS, adOpenKeyset
    cboTranPartNo.Clear
    If Not rsOrdCustomer.BOF And Not rsOrdCustomer.EOF Then
        rsOrdCustomer.MoveFirst
        Do While Not rsOrdCustomer.EOF
            cboTranPartNo.AddItem Null2String(rsOrdCustomer!custname)
            rsOrdCustomer.MoveNext
        Loop
    End If
    cboTranPartNo.AddItem "ALL", 0
    cboTranPartNo.ListIndex = 0
    Set rsOrdCustomer = Nothing
End Sub
