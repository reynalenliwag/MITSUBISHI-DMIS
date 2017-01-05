VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{7E0552E8-E2C9-4C9E-BC75-F34D871C75F4}#2.0#0"; "wizEncrypt.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#10.4#0"; "CODEJO~2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO301B~1.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#10.4#0"; "COF080~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   Caption         =   "Accounting Management Information System"
   ClientHeight    =   5910
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14610
   Icon            =   "frmMain.frx":0000
   Moveable        =   0   'False
   NegotiateToolbars=   0   'False
   Picture         =   "frmMain.frx":15162
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   1830
      Top             =   90
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   540
      Top             =   120
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin Crystal.CrystalReport rptMain 
      Left            =   1350
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
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
   Begin wizEncrypt.wizEnc wizEnc1 
      Left            =   -1860
      Top             =   6240
      _ExtentX        =   3969
      _ExtentY        =   3969
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5655
      Width           =   14610
      _ExtentX        =   25770
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   11
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Object.ToolTipText     =   "Login Name"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Object.ToolTipText     =   "Login Level"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Object.ToolTipText     =   "Current Date"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   4410
            MinWidth        =   4410
            Object.ToolTipText     =   "Login Time"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11:07 AM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
            Object.ToolTipText     =   "Num Lock Status"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Caps Lock Status"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1147
            MinWidth        =   1147
            TextSave        =   "SCRL"
            Object.ToolTipText     =   "Scroll Lock Status"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   7064
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3352
            MinWidth        =   3352
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1940
            MinWidth        =   1940
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PopupControl PopCntrl 
      Left            =   960
      Top             =   120
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   2
      Width           =   140
      Height          =   270
      Animation       =   1
      AnimateDelay    =   125
      ShowDelay       =   2500
      BackgroundBitmap=   60
   End
   Begin XtremeCommandBars.CommandBars CommandBars1 
      Left            =   120
      Top             =   150
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   6
      DesignerControls=   "frmMain.frx":2CC79
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "]"
    ApplyThemes
    ConfigurePopUps
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If MsgBox("Exit AMIS, Are You Sure?", vbExclamation + vbOKCancel, "Exit System") = vbOK Then
        Dim frm                                             As Form
        For Each frm In Forms
            If Not (frm Is Nothing) Then
                Unload frm
            End If
        Next
        CommandBars1.SaveCommandBars MODULENAME, App.TITLE, "Layout"
    Else
        Cancel = 1
        frmMainMenu.Show
    End If
End Sub

Private Sub CommandBars1_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
        'Accounts
    Case FILES_AC_CHARTOFACCOUNTS, TOOL_CHARTOFACCOUNTS
        'AXP-07082007-000001
        If Module_Access(LOGID, "CHART OF ACCOUNTS", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmAMISFILESChartOfAccount
        
    Case FILES_AC_ACCOUNTTYPES
        'AXP-07082007-000001
        If Module_Access(LOGID, "ACCOUNT TYPES", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmAMISFILESAccType
        
    Case FILES_AC_ACCOUNTCLASSIFICATION
        'AXP-07082007-000001
        If Module_Access(LOGID, "ACCOUNT CLASSIFICATION", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmAMISFILESHeader
        
    Case FILES_AC_EXTENDEDCLASSIFICATION
        'AXP-07082007-000001
        If Module_Access(LOGID, "EXTENDED CLASSIFICATION", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmAMISFILESSubHeader
        
    Case FILES_AC_ACCOUNTSUBTOTALS
        'AXP-07082007-000001
        If Module_Access(LOGID, "ACCOUNT SUB TOTALS", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmAMISFILESTitleCode
        
    Case FILES_AC_DEPARTMENTCODES
        'AXP-07082007-000001
        If Module_Access(LOGID, "DEPARTMENT CODES", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmAMISFILESDepartment
        
    Case FILES_AC_ACCOUNTENTRIESTEMPLATES
        'AXP-07082007-000001
        If Module_Access(LOGID, "ACCOUNT ENTRIES TEMPLATES", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmAMISMASTERFILESTemplates
                
    Case FILES_MAS_CUSTOMERS
        'AXP-07082007-000001
        If Module_Access(LOGID, "CUSTOMER", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmAllCustomer
        
    Case FILES_MAS_VENDORS
        'AXP-07082007-000001
        If Module_Access(LOGID, "VENDORS", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmAMISMASTERFILEVendor
        
    Case FILES_MAS_BANKS
        'AXP-07082007-000001
        If Module_Access(LOGID, "BANKS", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmAMISMASTERFILEBanks
        
    Case FILES_MAS_INVOICETYPES
        'AXP-07082007-000001
        If Module_Access(LOGID, "INVOICE TYPES", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmAMISMASTERFILEInvoiceType
        
    Case FILES_MAS_TERMSOFPAYMENT
        'AXP-07082007-000001
        If Module_Access(LOGID, "TERMS OF PAYMENT", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmAMISMASTERFILEPayTerm
        
    Case FILES_MAS_ATCCODE
        'AXP-07082007-000001
        If Module_Access(LOGID, "ATC CODES", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmAMISMASTERFILEATC
        
    Case FILES_OB_ACCOUNTSOPENINGBALANCE
        'AXP-07082007-000001
        If Module_Access(LOGID, "ACCOUNT OPENING BALANCE", "TRANSACTION") = False Then Exit Sub
        On Error Resume Next
        Call frmAMISJournalEntry_OPB.LOADJOURNAL("OPB")
        FormExistsShow frmAMISJournalEntry_OPB
        
    Case FILES_OB_CUSTOMEROPENINGBALANCE
        'AXP-07082007-000001
        If Module_Access(LOGID, "CUSTOMER OPENING BALANCE", "DATA ENTRY") = False Then Exit Sub
        On Error Resume Next
        JOURNALTYPE = "COB"
        FormExistsShow frmAMISCustomerAROpening
        
    Case FILES_OB_VENDOROPENINGBALANCE
        'AXP-07082007-000001
        If Module_Access(LOGID, "VENDOR OPENING BALANCE", "DATA ENTRY") = False Then Exit Sub
        On Error Resume Next
        JOURNALTYPE = "VPJ"
        FormExistsShow frmAMISVendorAPOpening
        
    Case FILES_OB_BANKOPENINGBALANCE
        'AXP-07082007-000001
        If Module_Access(LOGID, "BANK OPENING BALANCE", "DATA ENTRY") = False Then Exit Sub
        On Error Resume Next
        JOURNALTYPE = "BOB"
        FormExistsShow frmAMISbanksOpening
        
    Case FILES_ADJ_CLIENTADJUSTINGJOURNALENTRIES
        'AXP-07082007-000001
        If Module_Access(LOGID, "CLIENT ADJUSTING JOURNAL ENTRIES", "TRANSACTION") = False Then Exit Sub
        JOURNALTYPE = "ADJ"
        On Error Resume Next
        FormExistsShow frmAMISJournalEntry
        
    Case FILES_ADJ_PROPOSEDADJUSTINGJOURNALENTRIES
        'AXP-07082007-000001
        If Module_Access(LOGID, "PROPOSED ADJUSTING JOURNAL ENTRIES", "TRANSACTION") = False Then Exit Sub
        JOURNALTYPE = "PDJ"
        On Error Resume Next
        FormExistsShow frmAMISJournalEntry
        
    Case FILES_ADJ_CUSTOMERADJUSTMENTS
        'AXP-07082007-000001
        If Module_Access(LOGID, "CUSTOMER ADJUSTMENT", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmAMISCustomerAdjustment
        
    Case FILES_ADJ_VENDORADJUSTMENTS
        'AXP-07082007-000001
        If Module_Access(LOGID, "VENDOR ADJUSTMENTS", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmAMISVendorAdjustment
        
    Case FILES_ADJ_CLOSINGENTRIES
        If Module_Access(LOGID, "CLOSING ENTRIES", "TRANSACTION") = False Then Exit Sub
        JOURNALTYPE = "CLO"
        On Error Resume Next
        FormExistsShow frmAMISJournalEntry
        
    Case FILES_ASSETSREGISTRY
        If Module_Access(LOGID, "ASSETS REGISTRY", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmAMISDATAAssets
        
    Case FILES_BANKRECONCILIATION
        If Module_Access(LOGID, "BANK RECONCILIATION", "DATA ENTRY") = False Then Exit Sub
        FormExistsShow frmReconcileAccount
        
    Case JR_ACCOUNTSPAYABLEJOURNAL, TOOL_ACCOUNTSPAYABLEJOURNAL
        If Module_Access(LOGID, "ACCOUNTS PAYABLE JOURNAL", "TRANSACTION") = False Then Exit Sub
        On Error Resume Next
        Call frmAMISJournalEntry_APJ.LOADJOURNAL("APJ")
        FormExistsShow frmAMISJournalEntry_APJ
        
    Case JR_CASHDISBURSEMENTJOURNAL, TOOL_CASHDISBURSEMENTJOURNAL
        If Module_Access(LOGID, "CASH DISBURSEMENT JOURNAL", "TRANSACTION") = False Then Exit Sub
        On Error Resume Next
        Call frmAMISJournalEntry_CDJ.LOADJOURNAL("CDJ")
        FormExistsShow frmAMISJournalEntry_CDJ
        
    Case JR_SALESJOURNAL
        If Module_Access(LOGID, "SALES JOURNAL", "TRANSACTION") = False Then Exit Sub
        On Error Resume Next
        Call frmAMISJournalEntry_SJ.LOADJOURNAL("SJ")
        FormExistsShow frmAMISJournalEntry_SJ
        
    Case JR_CASHRECEIPTSJOURNAL, TOOL_CASHRECEIPTSJOURNAL
        If Module_Access(LOGID, "CASH RECEIPTS JOURNAL", "TRANSACTION") = False Then Exit Sub
        On Error Resume Next
        Call frmAMISJournalEntry_CRJ.LOADJOURNAL("CRJ")
        FormExistsShow frmAMISJournalEntry_CRJ
        
    Case JR_GENERALJOURNAL, TOOL_GENERALJOURNAL
        If Module_Access(LOGID, "GENERAL JOURNAL", "TRANSACTION") = False Then Exit Sub
        On Error Resume Next
        Call frmAMISJournalEntry_GJ.LOADJOURNAL("GJ")
        FormExistsShow frmAMISJournalEntry_GJ
               
    Case LEDG_ACCCOUNTSGENERALLEDGER, TOOL_ACCOUNTSGENERALLEDGER
        If Module_Access(LOGID, "ACCOUNT GENERAL LEDGER", "INQUIRY") = False Then Exit Sub
        FormExistsShow frmAMISLEDGERAccounts
        
    Case LEDG_CUSTOMERSARLEDGER, TOOL_CUSTOMERSLEDGER
        If Module_Access(LOGID, "CUSTOMER A/R LEDGER", "INQUIRY") = False Then Exit Sub
        CUST_LEDGER_TYPE = "ARLEDGER"
        FormExistsShow frmAMIS_ARLEDGER
        
    Case LEDG_CUSTOMERSDEPOSIT
        If Module_Access(LOGID, "CUSTOMER DEPOSIT LEDGER", "INQUIRY") = False Then Exit Sub
        CUST_LEDGER_TYPE = "CUSTDEPOSIT"
        
    Case LEDG_VENDORSSUBSIDIARYLEDGER, TOOL_SUPPLIERSLEDGER
        If Module_Access(LOGID, "VENDOR SUBSIDIARY LEDGER", "INQUIRY") = False Then Exit Sub
        FormExistsShow frmAMIS_APLEDGER
        
    Case PROC_IntegrateCustomerMasterFile
        
    Case PROC_ExtractPurchaseEntriesToBIRRelief
        If Module_Access(LOGID, "EXTRACT PURCHASE ENTRIES TO BIR RELIEF", "PROCESSING") = False Then Exit Sub
        EXTRACT_TYPE = "PURCHASES"
        FormExistsShow frmBIRExtract
        
    Case PROC_ExtractSalesToBIRRelief
        If Module_Access(LOGID, "EXTRACT SALES ENTRIES TO BIR RELIEF", "PROCESSING") = False Then Exit Sub
        EXTRACT_TYPE = "SALES"
        FormExistsShow frmBIRExtract
        
    Case PROC_ImportPurchases
        If Module_Access(LOGID, "IMPORT PURCHASES", "PROCESSING") = False Then Exit Sub
        FormExistsShow frmAPJImport
        
    Case PROC_ImportPurchasesFromPartsVehiclesAccessories
        If Module_Access(LOGID, "IMPORT PURCHASES", "PROCESSING") = False Then Exit Sub
        FormExistsShow frmAPJImport
        
    Case PROC_ImportSalesEntriesFromPartsserviceAndSales
        If Module_Access(LOGID, "IMPORT SALES ENTRIES", "PROCESSING") = False Then Exit Sub
        FormExistsShow frmSALESImport
        
    Case PROC_ImportCashReceiptsEntriesFromORSystem
        If Module_Access(LOGID, "IMPORT CASH RECEIPTS", "PROCESSING") = False Then Exit Sub
        FormExistsShow frmCRJImport
        
    Case PROC_ImportInventoryAdjustment
        If Module_Access(LOGID, "IMPORT INVENTORY ADJUSTMENTS", "PROCESSING") = False Then Exit Sub
        FormExistsShow FrmGJImport
        
    Case RPT_AC_CHARTOFACCOUNTS
        If Module_Access(LOGID, "CHART OF ACCOUNTS LIST", "REPORTS") = False Then Exit Sub
        ShowReport "ChartofAccounts", "AccountFiles", "", "Chart of Accounts", "AS OF: " & LOGDATE, True

    Case RPT_AC_CUSTOMERMASTERLIST
        If Module_Access(LOGID, "CUSTOMER MASTER LIST", "REPORTS") = False Then Exit Sub
        ShowReport "Customers", "Files", "", "Customers Master List", "AS OF: " & LOGDATE, True

    Case RPT_AC_SUPPLIERMASTERLIST
        If Module_Access(LOGID, "SUPPLIER MASTER LIST", "REPORTS") = False Then Exit Sub
        ShowReport "Suppliers", "Files", "", "Suppliers Master List", "AS OF: " & LOGDATE, True

    Case RPT_JR_ACCOUNTSPAYABLEJOURNAL, TOOL_APJJOURNALSSUMMARY
        If Module_Access(LOGID, "ACCOUNTS PAYABLE JOURNAL", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "APJ"
        FormExistsShow frmAMISRangeWithSummary
        frmAMISRangeWithSummary.Caption = "Accounts Payable Journal"

    Case RPT_JR_LEDGERCODERUNNINGBALANCE
        If Module_Access(LOGID, "ACCOUNTS PAYABLE LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "APJ"
        FormExistsShow frmAMISRangeWithAccountCode
        frmAMISRangeWithAccountCode.Caption = "Accounts Payable Ledger Code Running Balance"

    Case RPT_JR_ACCOUNTDETAILBYSUPPLIER
        If Module_Access(LOGID, "ACCOUNTS DETAIL BY SUPPLIERS", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "APJ"
        FormExistsShow frmAMISDetailBySupplierWithAccountCode
        frmAMISDetailBySupplierWithAccountCode.Caption = "Accounts Payable Detail Report By Supplier"

    Case RPT_JR_ACCOUNTSPAYABLEDUEREPORT
        If Module_Access(LOGID, "ACCOUNTS PAYABLE DUE REPORT", "REPORTS") = False Then Exit Sub
        REPORT_AP = "SCHED"
        FormExistsShow frmAMISDueReport
    Case RPT_JR_ACCOUNTSPAYABLEAGINGREPORT
        If Module_Access(LOGID, "ACCOUNTS PAYABLE AGING REPORT", "REPORTS") = False Then Exit Sub
        REPORT_AP = "AGING"
        FormExistsShow frmAMISDueReport
    Case RPT_JR_SCHEDULEOFACCOUNTSPAYABLE
        If Module_Access(LOGID, "SCHEDULE OF ACCOUNTS PAYABLE", "REPORTS") = False Then Exit Sub
        
    Case RPT_JR_RECEIVINGREPORTREGISTER
        If Module_Access(LOGID, "RECEIVING REPORT REGISTER", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "REC_REGISTER"
        FormExistsShow frmAMISDetailBySupplierWithAccountCode
        frmAMISDetailBySupplierWithAccountCode.Caption = "Receiving Report Registers"
        
    Case RPT_CSHDIS_CASHDISBURSEMENTJOURNAL, TOOL_CDJJOURNALSSUMMARY
        If Module_Access(LOGID, "CASH DISBURSEMENT JOURNAL", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "CDJ"
        FormExistsShow frmAMISRangeWithSummary
        frmAMISRangeWithSummary.Caption = "Cash Disbursement Journal"

    Case RPT_CSHDIS_LEDGERCODERUNNINGBALANCE
        If Module_Access(LOGID, "CASH DISBURSEMENT LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "CDJ"
        FormExistsShow frmAMISRangeWithAccountCode
        frmAMISRangeWithAccountCode.Caption = "Cash Disbursement Ledger Code Running Balance"
        
    Case RPT_CSHDIS_CHECKREGISTER
        If Module_Access(LOGID, "REGISTER REPORT", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "CHECK_REGISTER"
        FormExistsShow frmAMISRange
        frmAMISRange.Caption = "Check Registers"
        DoEvents
        
    Case RPT_SJR_SALESJOURNAL, TOOL_SJJOURNALSSUMMARY
        If Module_Access(LOGID, "SALES JOURNAL", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "SJ"
        FormExistsShow frmAMISRangeWithSummary
        frmAMISRangeWithSummary.Caption = "Sales Journal"
        
    Case RPT_SJR_ACCOUNTDETAILBYCUSTOMER
        If Module_Access(LOGID, "ACCOUNTS DETAIL BY CUSTOMER", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "SJ"
        FormExistsShow frmAMISDetailBySupplierWithAccountCode
        frmAMISDetailBySupplierWithAccountCode.Caption = "Accounts Detail Report By Customer"

    Case RPT_SJR_LEDGERCODERUNNINGBALANCE
        If Module_Access(LOGID, "SALES JOURNALS LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "SJ"
        FormExistsShow frmAMISRangeWithAccountCode
        frmAMISRangeWithAccountCode.Caption = "Sales Journal Ledger Code Running Balance"

    Case RPT_SJR_SCHEDULEOFACCOUNTSRECEIVABLE
        If Module_Access(LOGID, "SCHEDULE OF ACCOUNTS RECEIVABLE", "REPORTS") = False Then Exit Sub
        Report_AR = "SCHED"
        FormExistsShow frmNEW_ARSchedReport

    Case RPT_SJR_ACCOUNTSRECEIVABLEAGINGREPORT
        If Module_Access(LOGID, "ACCOUNTS RECEIVABLE AGING REPORT", "REPORTS") = False Then Exit Sub
        Report_AR = "AGING"
        FormExistsShow frmNEW_ARSchedReport
        
    Case RPT_SJR_INVOICEREGISTER
        If Module_Access(LOGID, "REGISTER REPORT", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "INV_REGISTER"
        FormExistsShow frmAMISRange
        frmAMISRange.Caption = "Invoices Registers"
        DoEvents
        
    Case RPT_JR_SALEINVTYPE_SALES_VEHICLESALESINVOICES
        If Module_Access(LOGID, "VEHICLE SALES INVOICE", "REPORTS") = False Then Exit Sub
        INVOICE_Type = "VEHICLE"
        FormExistsShow frmAMISSalesByInvoiceType
    Case RPT_JR_SALEINVTYPE_SALES_HYUNDAIVEHICLESALESINVOICES
        If Module_Access(LOGID, "HYUNDAI VEHICLE SALES INVOICE", "REPORTS") = False Then Exit Sub
        
    Case RPT_JR_PARTINVTYPE_SALES_PARTSCASHINVOICES
        If Module_Access(LOGID, "PARTS CASH INVOICE", "REPORTS") = False Then Exit Sub
        INVOICE_Type = "PARTS-CASH"
        FormExistsShow frmAMISSalesByInvoiceType
        
    Case RPT_JR_PARTINVTYPE_SALES_PARTSCHARGEINVOICES
        If Module_Access(LOGID, "PARTS CASH INVOICE", "REPORTS") = False Then Exit Sub
        INVOICE_Type = "PARTS-CHARGE"
        FormExistsShow frmAMISSalesByInvoiceType
    Case RPT_JR_PARTINVTYPE_SALES_HYUNDAIPARTSCASHINVOICES
        If Module_Access(LOGID, "HYUNDAI PARTS CASH INVOICE", "REPORTS") = False Then Exit Sub
        
    Case RPT_JR_PARTINVTYPE_SALES_HYUNDAIPARTSCHARGEINVOICES
        If Module_Access(LOGID, "HYUNDAI PARTS CHARGE INVOICE", "REPORTS") = False Then Exit Sub
        
    Case RPT_JR_SERIVICEINVTYPE_SALES_SERVICECASHINVOICES
        If Module_Access(LOGID, "SERVICE CASH INVOICE", "REPORTS") = False Then Exit Sub
        INVOICE_Type = "SERVICE-CASH"
        FormExistsShow frmAMISSalesByInvoiceType
    Case RPT_JR_SERIVICEINVTYPE_SALES_SERVICECHARGEINVOICES
        If Module_Access(LOGID, "SERVICE CHARGE INVOICE", "REPORTS") = False Then Exit Sub
        INVOICE_Type = "SERVICE-CHARGE"
        FormExistsShow frmAMISSalesByInvoiceType
        
    Case RPT_JR_SERIVICEINVTYPE_SALES_HYUNDAISERVICECASHINVOICES
        If Module_Access(LOGID, "HYUNDAI SERVICE CASH INVOICE", "REPORTS") = False Then Exit Sub
        
    Case RPT_JR_SERIVICEINVTYPE_SALES_HYUNDAISERVICECHARGEINVOICES
        If Module_Access(LOGID, "HYUNDAI SERVICE CHARGE INVOICE", "REPORTS") = False Then Exit Sub
        
    Case RPT_SJR_UNUSEDINVOICES
        If Module_Access(LOGID, "UNUSED INVOICES", "REPORTS") = False Then Exit Sub
        FormExistsShow frmAMISProcessUnusedInvoices
        
    Case RPT_CASHREC_CASHRECEIPTSJOURNAL, TOOL_CRJJOURNALSSUMMARY
        If Module_Access(LOGID, "CASH RECEIPTS JOURNAL", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "CRJ"
        FormExistsShow frmAMISRangeWithSummary
        frmAMISRangeWithSummary.Caption = "Cash Receipts Journal"
        
    Case RPT_CASHREC_LEDGERCODERUNNINGBALANCE
        If Module_Access(LOGID, "CASH RECEIPTS LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "CRJ"
        FormExistsShow frmAMISRangeWithAccountCode
        frmAMISRangeWithAccountCode.Caption = "Cash Receipts Ledger Code Running Balance"
        
    Case RPT_CASHREC_ORREGISTER, 1195
        If Module_Access(LOGID, "REGISTER REPORT", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "OR_REGISTER"
        FormExistsShow frmAMISRange
        frmAMISRange.Caption = "O.R. Registers"
        DoEvents
        
    Case RPT_CASHREC_UNUSEDOR, 1196
        If Module_Access(LOGID, "UNUSED OR", "REPORTS") = False Then Exit Sub
        FormExistsShow frmAMISProcessUnusedOR
        
    Case RPT_GJR_JOURNALVOUCHERSUMMARY, TOOL_GJJOURNALSSUMMARY
        If Module_Access(LOGID, "GENERAL JOURNAL SUMMARY", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "JVS"
        FormExistsShow frmAMISRange
        
    Case RPT_GJR_LEDGERCODERUNNINGBALANCE
        If Module_Access(LOGID, "GENERAL JOURNAL LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "GJ"
        FormExistsShow frmAMISRangeWithAccountCode
        frmAMISRangeWithAccountCode.Caption = "Journal Voucher Ledger Code Running Balance"
        
    Case RRT_LEDG_TRIALBALANCE
        If Module_Access(LOGID, "TRIAL BALANCE", "REPORTS") = False Then Exit Sub

    Case RPT_LEDG_ACCOUNTSPAYABLELEDGER
        If Module_Access(LOGID, "GENERAL LEDGER", "REPORTS") = False Then Exit Sub

    Case RPT_LEDG_SCHEDULEOFACCOUNTSPAYABLE
        If Module_Access(LOGID, "SCHEDULE OF ACCOUNTS PAYABLE", "REPORTS") = False Then Exit Sub
        
    Case RPT_FS_WORKSHEET, TOOL_WORKSHEET
        If Module_Access(LOGID, "WORKSHEET", "REPORTS") = False Then Exit Sub
        FormExistsShow frmAMISWorkSheet
        
    Case RPT_FS_TRIALBALANCE, TOOL_TRIALBALANCE
        If Module_Access(LOGID, "TRIAL BALANCE", "REPORTS") = False Then Exit Sub
        FormExistsShow frmAMISTrialBalance
        
    Case RPT_FS_INCOMESTATEMENT
        If Module_Access(LOGID, "INCOME STATEMENT", "REPORTS") = False Then Exit Sub
        FormExistsShow frmAMISIncomeStatement
        
    Case RPT_FS_STATEMENTOFOWNERSEQUITY
        If Module_Access(LOGID, "STATEMENT OF OWNERS EQUITY", "REPORTS") = False Then Exit Sub

    Case RPT_FS_STATEMENTOFCASHFLOW
        If Module_Access(LOGID, "STATEMENT OF CASH FLOW", "REPORTS") = False Then Exit Sub

    Case RPT_FS_SCHEDULEOFADJUSTMENTS
        If Module_Access(LOGID, "SCHEDULE OF ADJUSTMENTS", "REPORTS") = False Then Exit Sub
        FormExistsShow frmAMISSchedAdjust
        
    Case RPT_FS_FINANCIALSTATEMENTS, TOOL_FINANCIALSTATEMENTS
        If Module_Access(LOGID, "FINANCIAL STATEMENTS", "REPORTS") = False Then Exit Sub
        FormExistsShow frmAMISFinancialStatements
        
    Case RPT_FS_HYUNDAIFINANCIALSTATEMENTS
        If Module_Access(LOGID, "FINANCIAL STATMENT HUYNDAI", "REPORTS") = False Then Exit Sub
        
    Case RPT_DISB_DAILYCHECKDISBURSEMENTREPORT
        If Module_Access(LOGID, "DAILY DISBURSEMENT REPORT", "REPORTS") = False Then Exit Sub
        
    Case RPT_DISB_CHECKDISBURSEMENTREPORT
        If Module_Access(LOGID, "CHECK DISBURSEMENT REPORT", "REPORTS") = False Then Exit Sub
        
    Case RPT_EXRPT_SCHEDULEOFADMINISTRATIVEEXPENSE
        If Module_Access(LOGID, "SCHEDULE OF ADMINISTRATIVE EXPENSE", "REPORTS") = False Then Exit Sub
        REPORT_EXPENSETYPE = "ADMIN"
        FormExistsShow frmAMISExpenseReport
        
    Case RPT_EXRPT_SCHEDULEOFSELLINGEXPENSE
        If Module_Access(LOGID, "SCHEDULE OF SELLING EXPENSE", "REPORTS") = False Then Exit Sub
        REPORT_EXPENSETYPE = "SELLING"
        FormExistsShow frmAMISExpenseReport
        
    Case RPT_BKAC_JR_ALL
        If Module_Access(LOGID, "BOOK OF ACCOUNTS ALL", "REPORTS") = False Then Exit Sub
        
    Case RPT_BKAC_JR_CHECKDISBURSEMENTJOURNAL
        If Module_Access(LOGID, "CHECK DISBURSEMENT JOURNAL", "REPORTS") = False Then Exit Sub
        
    Case RPT_BKAC_JR_CASHRECEIPTJOURNAL
        If Module_Access(LOGID, "CASH RECEIPTS JOURNAL", "REPORTS") = False Then Exit Sub
        
    Case RPT_BKAC_LEDGER
        If Module_Access(LOGID, "LEDGER", "REPORTS") = False Then Exit Sub
        
    Case RPT_BKAC_COLUMNARCASHRECEIPT
        If Module_Access(LOGID, "COLUMNAR (CASH RECEIPT)", "REPORTS") = False Then Exit Sub
        
    Case RPT_BKAC_COLUMNARCASHDISBURSEMENT
        If Module_Access(LOGID, "COLUMNAR (CASH DISBURSEMENT)", "REPORTS") = False Then Exit Sub
        
    Case RPT_BKAC_SUBSIDIARYSALESJOURNAL
        If Module_Access(LOGID, "SUBSIDIARY SALES JOURNAL", "REPORTS") = False Then Exit Sub
        
    Case RPT_BKAC_SUBSIDIARYPURCHASEJOURNAL
        If Module_Access(LOGID, "SUBSIDIARY PURCHASE JOURNAL", "REPORTS") = False Then Exit Sub
                
    Case RPT_DEPASSET_ASSETLIST
        If Module_Access(LOGID, "ASSETLIST", "REPORTS") = False Then Exit Sub
        Dim rsProfile                                       As ADODB.Recordset
        rptMain.Reset
        Set rsProfile = New ADODB.Recordset
        Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
        If Not (rsProfile.EOF And rsProfile.BOF) Then
            rptMain.Formulas(3) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
            rptMain.Formulas(4) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
            rptMain.ReportTitle = "LIST OF ASSETS"
        End If
        PrintReport rptMain, AMIS_REPORT_PATH & "\Files\ListOfAssets.rpt", "", 1
        
    Case RPT_DEPASSET_DEPRECIATEDASSET
        If Module_Access(LOGID, "DEPRECIATED ASSET", "REPORTS") = False Then Exit Sub
        
    Case RPT_DEPASSET_SCHEDULEOFDEPRECIATION
        If Module_Access(LOGID, "SCHEDULE OF DEPRECIATION", "REPORTS") = False Then Exit Sub
        FormExistsShow frmAMISMonthlyYearly
        
    Case RPT_SCH_SCHEDULEOFINCOMETAXESWHELDFROMSUPPLIERS
        If Module_Access(LOGID, "SCHEDULES OF INCOME TAX W/HELD FROM SUPPLIERS", "REPORTS") = False Then Exit Sub
        FormExistsShow frmAMISYearly
        
    Case RPT_SCH_SCHEDULEOFPAYEESSUBJECTTOEXPANDEDWITHHOLDINGTAX
        If Module_Access(LOGID, "REGISTER REPORT", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "EX_TAX"
        FormExistsShow frmAMISRange
        frmAMISRange.Caption = "Schedule of Payees Subject to Expanded Withholding Tax"
        
    Case RPT_AUDI_AUDITADJUSTMENTSUMMARY
        If Module_Access(LOGID, "AUDIT ADJUSTMENT SUMMARY", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "ADJ"
        FormExistsShow frmAMISRange
        
    Case RPT_AUDIT_AUDITADJUSTMENTJOURNAL
        If Module_Access(LOGID, "ADJUSTMENT JOURNAL", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "ADJ"
        FormExistsShow frmAMISRangeWithSummary
        frmAMISRangeWithSummary.Caption = "Audit Adjustment Journal"

    Case MAINTAIN_SYSTEMSETUP
        If Module_Access(LOGID, "SYSTEM SETUP", "SYSTEM") = False Then Exit Sub
        FormExistsShow frmAMISProfile
        
    Case MAINTAIN_PASSWORDMAINTENANCE
        FormExistsShow frmAccMaintenance
        
    Case WINDOW_ABOUT, TOOL_ABOUTTHEAUTHOR
        FormExistsShow frmAbout
        
    Case WINDOW_EXIT, TOOL_EXITSYSTEM
        Unload Me
        
    Case TOOL_ACCOUNTSRECEIVABLEJOURNAL
        If Module_Access(LOGID, "ACCOUNTS RECEIVABLE JOURNAL", "TRANSACTION") = False Then Exit Sub
        JOURNALTYPE = "SJ"
        On Error Resume Next
        FormExistsShow frmAMISJournalEntry
        
    Case 1199
        FormExistsShow frmMainMenu
        frmMainMenu.ZOrder 0
    Case 1203
        If Module_Access(LOGID, "DEPOSITED CASH RECEIPTS JOURNAL", "REPORTS") = False Then Exit Sub
        REPORT_RANGETYPE = "DRJ"
        FormExistsShow frmAMISRangeWithSummary
        frmAMISRangeWithSummary.Caption = "Deposited Cash Receipts Journal"
        
    Case Else
        Debug.Print Control.ID
    End Select
End Sub

Private Sub ApplyThemes()
    CommandBarsGlobalSettings.App = App
    With CommandBars1
        .LoadDesignerBars
        .LoadCommandBars MODULENAME, App.TITLE, "Layout"
        .PaintManager.ClearTypeTextQuality = True
        .TabWorkspace.ThemedBackColor = False
        .StatusBar.Visible = True
        .Options.SyncFloatingToolbars = True
    End With
    
    With SkinFramework1
        .LoadSkin "C:\DMIS 2.0\Styles\royale.cjstyle", ""
        .ApplyWindow Me.hwnd
        .ApplyOptions = xtpSkinApplyFrame Or xtpSkinApplyColors Or Not xtpSkinApplyMetrics
    End With
    
    Dim ToolTipContext                                      As ToolTipContext
    Set ToolTipContext = CommandBars1.ToolTipContext
    With ToolTipContext
        .ShowTitleAndDescription True, xtpToolTipIconInfo
        .SetMargin 2, 2, 2, 2
        .MaxTipWidth = 180
        If .IsBalloonStyleSupported Then
            .Style = xtpToolTipBalloon
        Else
            .Style = xtpToolTipOffice2007
        End If
        .ShowShadow = True
    End With
End Sub

''''''''''''''START REGION POPUPCONTROLS''''''''''''''''''''''''''''''''''''
Private Sub ConfigurePopUps()
    Dim Item                                                As PopupControlItem
    PopCntrl.RemoveAllItems
    PopCntrl.Icons.AddIcons CommandBars1.Icons

    Set Item = PopCntrl.AddItem(245, 8, 265, 20, vbNullString)
    Item.Button = True
    Item.IconIndex = 899
    Item.ID = 707
    Item.Height = 20
    Item.Width = 20
    Item.CenterIcon
    Set Item = PopCntrl.AddItem(10, 10, 218, 30, vbNullString)
    Item.TextColor = RGB(15, 48, 145)
    Item.Bold = True
    Item.Font.Size = 10
    Item.Hyperlink = False
    Set Item = PopCntrl.AddItem(10, 32, 60, 50, vbNullString)
    Item.Height = 50
    Item.Width = 50
    Item.IconIndex = 0
    Item.Hyperlink = False

    Set Item = PopCntrl.AddItem(62, 32, 260, 50, vbNullString)
    Item.TextAlignment = DT_WORDBREAK Or DT_LEFT
    Item.Height = 50
    Item.ID = 655
    Item.Hyperlink = False

    Set Item = PopCntrl.AddItem(20, 85, 260, 105, vbNullString)
    Item.TextAlignment = DT_WORDBREAK Or DT_LEFT
    Item.TextColor = RGB(190, 1, 1)
    Item.Height = 50
    Item.Font.Size = 7
    Item.Hyperlink = False
End Sub

Private Sub PopCntrl_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    If Item.ID = 707 Then
        PopCntrl.Close
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    If TIMER_REMIND = "" Then
        ReminderModule ""
    Else
        If DateDiff("n", TIMER_REMIND, Now) >= 0 Then
            frmSMIS_Files_Reminders.Show
        End If
    End If
End Sub

