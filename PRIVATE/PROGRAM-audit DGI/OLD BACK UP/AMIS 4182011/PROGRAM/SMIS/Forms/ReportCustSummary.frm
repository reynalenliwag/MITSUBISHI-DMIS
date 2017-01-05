VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_CustSummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers Summary"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   ForeColor       =   &H00FCFCFC&
   Icon            =   "ReportCustSummary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1575
   ScaleWidth      =   4815
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
      Left            =   2340
      MouseIcon       =   "ReportCustSummary.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportCustSummary.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   645
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
      Left            =   1470
      MouseIcon       =   "ReportCustSummary.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "ReportCustSummary.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   645
      Width           =   885
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   465
      Left            =   3390
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   1365
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   465
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   2355
   End
   Begin Crystal.CrystalReport rptReleased 
      Left            =   3540
      Top             =   990
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Units Released"
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
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   2520
      TabIndex        =   2
      Top             =   90
      Width           =   825
   End
End
Attribute VB_Name = "frmSMIS_Report_CustSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPurchAgree                                                      As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()


    On Error GoTo ErrorCode:

    If Len(cboYear.Text) = 4 Or cboYear.Text <> "" Then
        Set rsPurchAgree = New ADODB.Recordset

        rsPurchAgree.Open "select * from SMIS_PurchAgree WHERE year(datereleased) = '" & cboYear.Text & "' and month(datereleased) =" & What_month(cboMonth), gconDMIS, adOpenForwardOnly, adLockReadOnly


        If Not rsPurchAgree.EOF And Not rsPurchAgree.BOF Then
            Screen.MousePointer = 11
            rptReleased.Reset
            rptReleased.ReportTitle = Me.Caption

            If CUST_REPT_TYPE = "1" Then
                rptReleased.WindowTitle = " Customer Vehicles Summary"
                rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptReleased, SMIS_REPORT_PATH & "VSCustomer.rpt", "year({PurchAgree.datereleased}) = " & cboYear.Text & " and month({PurchAgree.datereleased}) = " & What_month(cboMonth), DMIS_REPORT_Connection, 1
                'LogAudit "V", "CUSTOMER SUMMARY REPORT", "FOR " & cboYear & " AND " & cboMonth
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                 Call NEW_LogAudit("V", "VEHICLE SALES CUSTOMERS SUMMARY", "", "", "", "CUSTOMER SUMMARY REPORT -" & " " & cboMonth & " " & cboYear, "", "")
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

                rptReleased.PageZoom 72
            Else

                rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptReleased.WindowTitle = " Customer With Insurance Policies"
                PrintSQLReport rptReleased, SMIS_REPORT_PATH & "VSCustomerWithInsuranceOnly.rpt", "year({PurchAgree.datereleased}) = " & cboYear.Text & " and month({PurchAgree.datereleased}) = " & What_month(cboMonth), DMIS_REPORT_Connection, 1
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
                 Call NEW_LogAudit("V", "CUSTOMERS WITH INSURANCE POLICIES", "", "", "", cboMonth & " " & cboYear, "", "")
                'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

                'LogAudit "V", "CUSTOMER WITH INSURANCE POLICIES ", "FOR " & cboYear & " AND " & cboMonth
            End If
            Screen.MousePointer = 0
        Else
            MsgSpeechBox "No Record for " & cboMonth.Text & " " & cboYear.Text
        End If
    End If

    Exit Sub
ErrorCode:
    ShowVBError
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
            
            If CUST_REPT_TYPE = 1 Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (VEHICLE SALES CUSTOMERS SUMMARY)"
                Call frmALL_AuditInquiry.DisplayHistory("", "VEHICLE SALES CUSTOMERS SUMMARY", "PRINTING")
            Else
                frmALL_AuditInquiry.Caption = "Audit Inquiry (CUSTOMERS WITH INSURANCE POLICIES)"
                Call frmALL_AuditInquiry.DisplayHistory("", "CUSTOMERS WITH INSURANCE POLICIES", "PRINTING")
            End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    FillCboMoreYear cboYear
    'fillcombo_up cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub

