VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMISProfitbyVoucher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gross Profit by Voucher"
   ClientHeight    =   1905
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboInvoiceType 
      Height          =   315
      Left            =   2310
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   840
      Width           =   3075
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Generate Report"
      Height          =   375
      Left            =   3690
      TabIndex        =   0
      Top             =   1380
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   315
      Left            =   735
      TabIndex        =   1
      Top             =   300
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM dd, yyyy"
      Format          =   113508355
      CurrentDate     =   38148
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   315
      Left            =   3300
      TabIndex        =   2
      Top             =   300
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM dd, yyyy"
      Format          =   113508355
      CurrentDate     =   38148
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Type"
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   930
      Width           =   930
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2850
      TabIndex        =   4
      Top             =   330
      Width           =   405
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   300
      Width           =   645
   End
End
Attribute VB_Name = "frmAMISProfitbyVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xInvoiceType                                            As String

Private Sub cboInvoiceType_Click()
    If cboInvoiceType = "Parts Invoice" Then
        xInvoiceType = "PI"
    ElseIf cboInvoiceType = "Service Invoice" Then
        xInvoiceType = "SI"
    ElseIf cboInvoiceType = "Vehicle Invoice" Then
        xInvoiceType = "VI"
    Else
    End If
End Sub

Private Sub cmdCheck_Click()
    GenerateReport
End Sub

Private Sub Form_Load()
    DateRange
    InitCbo
End Sub

Sub DateRange()
    Dim rsJournalHD                                         As ADODB.Recordset
    Set rsJournalHD = New ADODB.Recordset
    rsJournalHD.Open "SELECT * FROM (SELECT MIN(JDATE) AS MINDATE,MAX(JDATE) AS MAXDATE FROM AMIS_JOURNAL_HD)T WHERE MINDATE IS NOT NULL", gconDMIS, adOpenForwardOnly
    If Not rsJournalHD.EOF And Not rsJournalHD.BOF Then
        dtFrom.Value = rsJournalHD!MinDate
        dtTo.Value = rsJournalHD!MaxDate
    End If
    Set rsJournalHD = Nothing
End Sub

Sub InitCbo()
    With cboInvoiceType
        .AddItem "Parts Invoice"
        .AddItem "Service Invoice"
        .AddItem "Vehicle Invoice"
    End With
End Sub

Sub GenerateReport()

'    Dim xlApplication As Excel.Application
'    Dim xlWorkbook As Excel.Workbook
'    Dim xlWorksheet As Excel.Worksheet
'    Dim xlRange As Excel.Range
'    Dim xCounter As Integer
'    Dim xPrevious As String
'    Dim xNew As String
'    Dim xDebit As Double
'    Dim xCredit As Double
'    Dim rsGenerateReport As ADODB.Recordset
'
'    Set xlApplication = CreateObject("Excel.Application")
'    Set xlWorkbook = xlApplication.Workbooks.Open(AMIS_REPORT_PATH & "\Journals\GrossProfit.xlt")
'    Set xlWorksheet = xlWorkbook.Worksheets(1)
'    xlWorksheet.Cells(1, "A") = COMPANY_NAME
'    xlWorksheet.Cells(1, "A").Font.Bold = True
'    xlWorksheet.Cells(2, "A") = COMPANY_ADDRESS
'    xlWorksheet.Cells(2, "A").Font.Bold = True
'    xlWorksheet.Cells(3, "A") = "From: " & Format(dtFrom.Value, "mm/dd/yyyy") & " To: " & Format(dtTo.Value, "mm/dd/yyyy")
'    xlWorksheet.Cells(3, "A").Font.Bold = True
'    Screen.MousePointer = 11
'    xCounter = 5
'    Set rsGenerateReport = New ADODB.Recordset
'    rsGenerateReport.Open "SELECT HD.JDATE,HD.VOUCHERNO,HD.JTYPE,HD.INVOICETYPE,HD.INVOICENO,DET.DEBIT,Det.CREDIT " & _
     '                            "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
     '                            "ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE " & _
     '                            "WHERE HD.JTYPE='SJ' AND HD.JDATE BETWEEN '" & dtFrom.Value & "' AND '" & dtTo.Value & "' AND LEFT(ACCT_CODE,5) IN ('41-02','61-02')" & _
     '                            "ORDER BY HD.JDATE,HD.VOUCHERNO", gconDMIS, adOpenKeyset
'    If Not rsGenerateReport.EOF And Not rsGenerateReport.BOF Then
'        Do While Not rsGenerateReport.EOF
'            xNew = rsGenerateReport!VoucherNo
'
'            xlWorksheet.Cells(xCounter, "A") = rsGenerateReport!JDate
'            xlWorksheet.Cells(xCounter, "B") = rsGenerateReport!VoucherNo
'            xlWorksheet.Cells(xCounter, "C") = rsGenerateReport!InvoiceType
'            xlWorksheet.Cells(xCounter, "D") = rsGenerateReport!INVOICENO
'            xlWorksheet.Cells(xCounter, "E") = (rsGenerateReport!DEBIT)
'            xlWorksheet.Cells(xCounter, "F") = (rsGenerateReport!CREDIT)
'            xDebit = xDebit + (rsGenerateReport!DEBIT)
'            xCredit = xCredit + (rsGenerateReport!CREDIT)
'
'
'
'            xCounter = xCounter + 1
'
'
'            If xNew <> xPrevious Then
'                xlWorksheet.Cells(xCounter, "E") = (xDebit)
'                xlWorksheet.Cells(xCounter, "F") = (xCredit)
'                xlWorksheet.Cells(xCounter, "G") = (xCredit - xDebit)
'            End If
'
'            xPrevious = rsGenerateReport!VoucherNo
'            rsGenerateReport.MoveNext
'
'
'
'
'            DoEvents
'        Loop
'
'
'    End If
''    xlWorksheet.Cells(xCounter, "E") = NumericVal(xDebit)
''    xlWorksheet.Cells(xCounter, "F") = NumericVal(xCredit)
''    xlWorksheet.Cells(xCounter, "G") = NumericVal(xCredit - xDebit)
'    xlApplication.Visible = True
'    Set xlApplication = Nothing
'    Set xlWorkbook = Nothing
'    Set xlWorksheet = Nothing
'    Set xlRange = Nothing
'    Set rsGenerateReport = Nothing
'    Screen.MousePointer = 0

    Dim rsJournal_HD                                        As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select * from AMIS_Journal_HD where JTYPE= 'SJ' AND (jdate >= '" & dtFrom & "' AND jdate <= '" & dtTo & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
    '    rsJournal_HD.Open "SELECT HD.JDATE,HD.VOUCHERNO,HD.JTYPE,HD.INVOICETYPE,HD.INVOICENO,DET.DEBIT,Det.CREDIT " & _
         "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
         "ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE " & _
         "WHERE HD.JTYPE='SJ' AND HD.JDATE BETWEEN '" & dtFrom.Value & "' AND '" & dtTo.Value & "' AND LEFT(ACCT_CODE,5) IN ('41-02','61-02')" & _
         "ORDER BY HD.JDATE,HD.VOUCHERNO", gconDMIS, adOpenKeyset
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        ShowRangeReport dtFrom, dtTo, "GrossProfit", "Journals", "{AMIS_Journal_Hd.jdate} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") AND {AMIS_Journal_Hd.jdate} <= date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")", "Gross Profit", False
    End If
End Sub
