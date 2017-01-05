VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_PMISReports_BIRinvoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice List Report"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   Icon            =   "frmPMIS_Report_BIRINVOICE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2310
   ScaleWidth      =   3930
   Begin VB.ComboBox cbotype 
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
      ItemData        =   "frmPMIS_Report_BIRINVOICE.frx":06EA
      Left            =   1560
      List            =   "frmPMIS_Report_BIRINVOICE.frx":06EC
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Select month from the list"
      Top             =   240
      Width           =   2265
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
      Left            =   2520
      MouseIcon       =   "frmPMIS_Report_BIRINVOICE.frx":06EE
      MousePointer    =   99  'Custom
      Picture         =   "frmPMIS_Report_BIRINVOICE.frx":0840
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print this Record"
      Top             =   1440
      Width           =   675
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
      Left            =   3195
      MouseIcon       =   "frmPMIS_Report_BIRINVOICE.frx":0CDF
      MousePointer    =   99  'Custom
      Picture         =   "frmPMIS_Report_BIRINVOICE.frx":0E31
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Close Window"
      Top             =   1440
      Width           =   675
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   345
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   2250
      _ExtentX        =   3969
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
      Format          =   100794369
      CurrentDate     =   39203
   End
   Begin Crystal.CrystalReport rptInvoice 
      Left            =   240
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Transaction Listing - Issuances"
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
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   345
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   2250
      _ExtentX        =   3969
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
      Format          =   100794369
      CurrentDate     =   39232
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   1455
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Report Type :"
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
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frm_PMISReports_BIRinvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsInvoice                                      As New ADODB.Recordset
Dim Xstatus                                        As String
Dim FDate                                          As Date
Dim TDate                                          As Date

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If App.EXEName = "PMIS" Then
        Call PrintMe_PMIS
    ElseIf App.EXEName = "SMIS" Then
        Call PrintMe_SMIS
    Else
        Call PrintMe_CSMS
    End If
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Call InitCbo
    cbotype.ListIndex = 0
    dtpFromDate.Value = LOGDATE
    dtpToDate.Value = LOGDATE
End Sub
Sub InitCbo()
cbotype.Clear
    If App.EXEName = "PMIS" Then
         cbotype.AddItem "Posted"
         cbotype.AddItem "Unposted"
         cbotype.AddItem "Cancelled"
         cbotype.AddItem "All"
    ElseIf App.EXEName = "SMIS" Then
         cbotype.AddItem "Invoiced"
         cbotype.AddItem "Released"
         cbotype.AddItem "Cancelled"
         cbotype.AddItem "All"
    Else
        cbotype.AddItem "Invoiced"
        cbotype.AddItem "Cancelled"
        cbotype.AddItem "All"
    End If
End Sub

Sub PrintMe_CSMS()
FDate = CDate(dtpFromDate.Value)
TDate = CDate(dtpToDate.Value)

Xstatus = ""
If cbotype.Text = "Invoiced" Then
    Xstatus = "P"
ElseIf cbotype.Text = "Cancelled" Then
    Xstatus = "C"
End If

If Xstatus <> "" Then
    Set rsInvoice = gconDMIS.Execute("select * from CSMS_VW_INVOICELIST where INVOICENO is not null and isnumeric(RIGHT(INVOICENO,6)) = 1 " & _
                                     " and ((INVOICEDATE between '" & FDate & "' and '" & TDate & "' ) or (CANCELLEDDATE between   '" & FDate & "' and '" & TDate & "' )) AND STATUS = '" & Xstatus & "'")
Else
    Set rsInvoice = gconDMIS.Execute("select * from CSMS_VW_INVOICELIST where INVOICENO is not null and isnumeric(RIGHT(INVOICENO,6)) = 1 " & _
                                     " and ((INVOICEDATE between '" & FDate & "' and '" & TDate & "' ) or (CANCELLEDDATE   between '" & FDate & "' and '" & TDate & "' ))")
End If

If Not (rsInvoice.EOF And rsInvoice.BOF) Then
    Screen.MousePointer = 11
    rptInvoice.ReportTitle = "Invoice List - " & cbotype.Text
    rptInvoice.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptInvoice.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptInvoice.Formulas(2) = "mindate = '" & FDate & "'"
    rptInvoice.Formulas(3) = "maxdate = '" & TDate & "'"
    rptInvoice.Formulas(4) = "PrintedBy = '" & LOGNAME & "'"
    If Xstatus <> "" Then
        PrintSQLReport rptInvoice, CSMS_REPORT_PATH & "Invoicelist.rpt", "{CSMS_VW_INVOICELIST.STATUS}  = '" & Xstatus & "' AND  (({CSMS_VW_INVOICELIST.INVOICEDATE} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {CSMS_VW_INVOICELIST.INVOICEDATE} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")) or ({CSMS_VW_INVOICELIST.CANCELLEDDATE} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {CSMS_VW_INVOICELIST.CANCELLEDDATE}<= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")))", DMIS_REPORT_Connection, 1
    Else
        PrintSQLReport rptInvoice, CSMS_REPORT_PATH & "Invoicelist.rpt", "({CSMS_VW_INVOICELIST.INVOICEDATE} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {CSMS_VW_INVOICELIST.INVOICEDATE} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")) or ({CSMS_VW_INVOICELIST.CANCELLEDDATE} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {CSMS_VW_INVOICELIST.CANCELLEDDATE}<= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & "))", DMIS_REPORT_Connection, 1
    End If
    
       ' PrintSQLReport rptInvoice, CSMS_REPORT_PATH & "Invoicelist.rpt", "{CSMS_VW_INVOICELIST.INVOICEDATE} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {CSMS_VW_INVOICELIST.INVOICEDATE} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", DMIS_REPORT_Connection, 1

    Screen.MousePointer = 0

Else
    ShowNoRecord
End If

End Sub

Sub PrintMe_SMIS()
FDate = CDate(dtpFromDate.Value)
TDate = CDate(dtpToDate.Value)

Xstatus = ""
If cbotype.Text = "Invoiced" Then
    Xstatus = "U"
ElseIf cbotype.Text = "Released" Then
    Xstatus = "P"
ElseIf cbotype.Text = "Cancelled" Then
    Xstatus = "C"
End If
If Xstatus <> "" Then
    Set rsInvoice = gconDMIS.Execute("Select * from SMIS_SALESORDER where STATUS = '" & Xstatus & "' and Invoiceddate between '" & FDate & "' and '" & TDate & "'")
Else
    Set rsInvoice = gconDMIS.Execute("Select * from SMIS_SALESORDER where Invoiceddate between '" & FDate & "' and '" & TDate & "'")
End If


If Not (rsInvoice.EOF And rsInvoice.BOF) Then
    Screen.MousePointer = 11
    rptInvoice.ReportTitle = "Invoice List - " & cbotype.Text
    rptInvoice.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptInvoice.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptInvoice.Formulas(2) = "mindate = '" & FDate & "'"
    rptInvoice.Formulas(3) = "maxdate = '" & TDate & "'"
    rptInvoice.Formulas(4) = "PrintedBy = '" & LOGNAME & "'"
    If Xstatus <> "" Then
        PrintSQLReport rptInvoice, SMIS_REPORT_PATH & "Invoicelist.rpt", "{SMIS_SalesOrder.STATUS} = '" & Xstatus & "' AND  {SMIS_SalesOrder.InvoicedDate} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {SMIS_SalesOrder.InvoicedDate}<= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", DMIS_REPORT_Connection, 1
    Else
        PrintSQLReport rptInvoice, SMIS_REPORT_PATH & "Invoicelist.rpt", "{SMIS_SalesOrder.InvoicedDate} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {SMIS_SalesOrder.InvoicedDate} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", DMIS_REPORT_Connection, 1
    End If
    Screen.MousePointer = 0
Else
    ShowNoRecord
End If

End Sub
Sub PrintMe_PMIS()


FDate = CDate(dtpFromDate.Value)
TDate = CDate(dtpToDate.Value)


Xstatus = ""
If cbotype.Text = "Posted" Then
    Xstatus = "P"
ElseIf cbotype.Text = "Unposted" Then
    Xstatus = "N"
ElseIf cbotype.Text = "Cancelled" Then
    Xstatus = "C"
End If
If Xstatus <> "" Then
    Set rsInvoice = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY where type = 'P' and TRANDATE between '" & FDate & "' and  '" & TDate & "' and status = '" & Xstatus & "'")
Else
    Set rsInvoice = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY where type = 'P' and TRANDATE between '" & FDate & "' and  '" & TDate & "'")
End If

If Not (rsInvoice.EOF And rsInvoice.BOF) Then
    Screen.MousePointer = 11
    rptInvoice.ReportTitle = "Invoice List - " & cbotype.Text
    rptInvoice.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptInvoice.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptInvoice.Formulas(2) = "mindate = '" & FDate & "'"
    rptInvoice.Formulas(3) = "maxdate = '" & TDate & "'"
    rptInvoice.Formulas(4) = "PrintedBy = '" & LOGNAME & "'"
    If Xstatus <> "" Then
        PrintSQLReport rptInvoice, PMIS_REPORT_PATH & "Invoicelist.rpt", " {PMIS_vw_ISS_HISTORY.STATUS} = '" & Xstatus & "' AND  {PMIS_vw_ISS_HISTORY.TYPE} = 'P' AND {PMIS_vw_ISS_HISTORY.TRANDATE} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {PMIS_vw_ISS_HISTORY.TRANDATE}<= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", DMIS_REPORT_Connection, 1
    Else
        PrintSQLReport rptInvoice, PMIS_REPORT_PATH & "Invoicelist.rpt", "{PMIS_vw_ISS_HISTORY.TYPE} = 'P' AND {PMIS_vw_ISS_HISTORY.TRANDATE} >= date(" & Year(FDate) & "," & Month(FDate) & "," & Day(FDate) & ") AND {PMIS_vw_ISS_HISTORY.TRANDATE} <= date(" & Year(TDate) & "," & Month(TDate) & "," & Day(TDate) & ")", DMIS_REPORT_Connection, 1
    End If
'    NEW_LogAudit "V", "PARTS TRANSACTION LISTING", "", "", "", dtpFromDate & " - " & dtpToDate, "HISTORY", ""
    Screen.MousePointer = 0
Else
    ShowNoRecord
End If
End Sub

