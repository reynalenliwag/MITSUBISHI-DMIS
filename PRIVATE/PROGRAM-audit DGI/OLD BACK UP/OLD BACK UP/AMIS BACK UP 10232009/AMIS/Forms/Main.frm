VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{7E0552E8-E2C9-4C9E-BC75-F34D871C75F4}#2.0#0"; "WIZENCRYPT.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#10.4#0"; "CODEJO~2.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Accounting Management Information System"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   3555
   ClientWidth     =   12555
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "Main.frx":1CFA
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   12555
      TabIndex        =   2
      Top             =   0
      Width           =   12555
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   2610
      Top             =   6210
      _Version        =   655364
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin Crystal.CrystalReport rptMain 
      Left            =   5160
      Top             =   2160
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
      Left            =   150
      Top             =   2490
      _ExtentX        =   3969
      _ExtentY        =   3969
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   12525
      TabIndex        =   0
      Top             =   225
      Width           =   12555
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   570
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   12555
         _ExtentX        =   22146
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   24
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Chart of Accounts"
               Object.ToolTipText     =   "Chart of Accounts"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Accounts Payable Journal"
               Object.ToolTipText     =   "Accounts Payable Journal"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Cash Disbursement Journal"
               Object.ToolTipText     =   "Cash Disbursement Journal"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Accounts Receivable Journal"
               Object.ToolTipText     =   "Accounts Receivable Journal"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Cash Receipts Journal"
               Object.ToolTipText     =   "Cash Receipts Journal"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "General Journal"
               Object.ToolTipText     =   "General Journal"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Accounts General Ledger"
               Object.ToolTipText     =   "Accounts General Ledger"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Customers Ledger"
               Object.ToolTipText     =   "Customers Ledger"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Suppliers Ledger"
               Object.ToolTipText     =   "Suppliers Ledger"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Chart of Account Listing"
               Object.ToolTipText     =   "APJ - Journals/Summary"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Customer Listing"
               Object.ToolTipText     =   "CDJ - Journals/Summary"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Supplier Listing"
               Object.ToolTipText     =   "SJ - Journals/Summary"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Work Sheet"
               Object.ToolTipText     =   "CRJ - Journals/Summary"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "GJ - Journals/Summary"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Trial Balance"
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Income Statement"
               Object.ToolTipText     =   "Trial Balance"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Balance Sheet"
               Object.ToolTipText     =   "Work Sheet"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "Statement of Cash Flow"
               Object.ToolTipText     =   "Financial Statements"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "About the Author"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Exit System"
               ImageIndex      =   19
            EndProperty
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "Main.frx":241D3C
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5430
      Top             =   1380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   65535
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":241E9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":242778
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":242A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":242DAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2430C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2433E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2436FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":243A14
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":243D2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":244048
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":244922
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2451FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":245AD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2463B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2466CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2469E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":246E36
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":247150
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":24746A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7665
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
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
            TextSave        =   "2:05 PM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
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
            Object.Width           =   1147
            MinWidth        =   1147
            TextSave        =   "SCRL"
            Object.ToolTipText     =   "Scroll Lock Status"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "&Files"
      Begin VB.Menu mnuAccountMenus 
         Caption         =   "&Accounts"
         Begin VB.Menu mnuChartOfAccount 
            Caption         =   "&Chart of Accounts"
         End
         Begin VB.Menu mnuAccountTypes 
            Caption         =   "&Account Types"
         End
         Begin VB.Menu mnuHeader 
            Caption         =   "Account &Classification"
         End
         Begin VB.Menu mnuAccountSubHeaders 
            Caption         =   "&Extended Classification"
         End
         Begin VB.Menu mnuAccountTitles 
            Caption         =   "Account &Sub-Totals"
         End
         Begin VB.Menu mnuDepartment 
            Caption         =   "&Department Codes"
         End
         Begin VB.Menu mnuTemplates 
            Caption         =   "Account Entries &Templates"
         End
      End
      Begin VB.Menu mnuMasterFiles 
         Caption         =   "&Master Files"
         Begin VB.Menu mnuCustomers 
            Caption         =   "&Customers"
         End
         Begin VB.Menu mnuVendorMaster 
            Caption         =   "&Vendors"
         End
         Begin VB.Menu mnuBanksMaster 
            Caption         =   "&Banks"
         End
         Begin VB.Menu mnuInvoiceTypes 
            Caption         =   "&Invoice Types"
         End
         Begin VB.Menu mnuTermsofPayment 
            Caption         =   "&Terms of Payment"
         End
         Begin VB.Menu mnuTaxRateCode 
            Caption         =   "&ATC Code"
         End
      End
      Begin VB.Menu mnuOpeningBalances 
         Caption         =   "&Opening Balances"
         Begin VB.Menu mnuOpeningBalance 
            Caption         =   "&Accounts Opening Balance"
         End
         Begin VB.Menu mnuCustOpenBalance 
            Caption         =   "&Customer Opening Balance"
         End
         Begin VB.Menu mnuVendorOpeningBalance 
            Caption         =   "&Vendor Opening Balance"
         End
      End
      Begin VB.Menu mnuAdjustments 
         Caption         =   "&Adjustments"
         Begin VB.Menu mnuAuditAdjustments 
            Caption         =   "&Client Adjusting Journal Entries"
         End
         Begin VB.Menu mnuPAJE 
            Caption         =   "&Proposed Adjusting Journal Entries"
         End
         Begin VB.Menu mnuCustomerCreditMemo 
            Caption         =   "&Customer Adjustments"
         End
         Begin VB.Menu mnuVendorAdjustments 
            Caption         =   "&Vendor Adjustments"
         End
         Begin VB.Menu mnuClosingEntries 
            Caption         =   "&Closing Entries"
         End
      End
      Begin VB.Menu mnuMYOBMenu 
         Caption         =   "&MYOB Menu"
         Visible         =   0   'False
         Begin VB.Menu mnuMYOBGeneralLedger 
            Caption         =   "MYOB &General Ledger"
         End
      End
      Begin VB.Menu mnuAssetRegistry 
         Caption         =   "A&ssets Registry"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReconcile 
         Caption         =   "Bank Reconciliation"
      End
   End
   Begin VB.Menu mnuJournals 
      Caption         =   "&Journals"
      Begin VB.Menu mnuAccountsPayableJournal 
         Caption         =   "Accounts &Payable Journal"
      End
      Begin VB.Menu mnuCashDisbursement 
         Caption         =   "Cash &Disbursement Journal"
      End
      Begin VB.Menu mnuAccountsReceivableJournal 
         Caption         =   "&Sales Journal"
      End
      Begin VB.Menu mnuCashSalesJournal 
         Caption         =   "Cash &Receipts Journal"
      End
      Begin VB.Menu mnuGeneralJournal 
         Caption         =   "&General Journal"
      End
   End
   Begin VB.Menu mnuLedger 
      Caption         =   "&Ledgers"
      Begin VB.Menu mnuGeneralLedger 
         Caption         =   "&Acccounts General Ledger"
      End
      Begin VB.Menu mnuCustomersLedger 
         Caption         =   "&Customers A/R Ledger"
      End
      Begin VB.Menu mnuCustomersDeposit 
         Caption         =   "Customers &Deposit"
      End
      Begin VB.Menu mnuVendorLedger 
         Caption         =   "&Vendors Subsidiary Ledger"
      End
   End
   Begin VB.Menu mnuProcessing 
      Caption         =   "&Processing"
      Begin VB.Menu mnuIntegrateCustomerMasterFile 
         Caption         =   "&Integrate Customer Master File"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBirExtraction 
         Caption         =   "&Extract Purchase Entries to BIR Relief"
      End
      Begin VB.Menu mnuExtractSalesBIRRelief 
         Caption         =   "&Extract Sales to BIR Relief"
      End
      Begin VB.Menu ImportPurchases 
         Caption         =   "Import &Purchases from Parts, Vehicles & Accessories"
      End
      Begin VB.Menu mnuImportPartsServiceSales 
         Caption         =   "Import &Sales Entries from Parts, Service and Sales"
      End
      Begin VB.Menu mnuImportCashReceipts 
         Caption         =   "Import &Cash Receipts Entries from OR System"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuAccounts 
         Caption         =   "&Accounts"
         Begin VB.Menu mnuR_ChartOfAccounts 
            Caption         =   "Chart of &Accounts"
         End
         Begin VB.Menu mnuCustomerMasterList 
            Caption         =   "&Customer Master List"
         End
         Begin VB.Menu mnuSupplierList 
            Caption         =   "&Supplier Master List"
         End
      End
      Begin VB.Menu mnuR_Journals 
         Caption         =   "&Journals"
         Begin VB.Menu mnuR_AccountsPayable 
            Caption         =   "&Accounts Payable"
            Begin VB.Menu mnuR_AccountsPayableJournal 
               Caption         =   "&Accounts Payable Journal"
            End
            Begin VB.Menu mnuLedgerCodeRunningBalance 
               Caption         =   "&Ledger Code Running Balance"
            End
            Begin VB.Menu mnuR_AccountDetailBySupplier 
               Caption         =   "Account Detail by &Supplier"
            End
            Begin VB.Menu mnuR_AccountsPayableDueReport 
               Caption         =   "Accounts Payable &Due Report"
            End
            Begin VB.Menu mnuAccountsPayableAgingReport 
               Caption         =   "Accounts &Payable Aging Report"
            End
            Begin VB.Menu mnuScheduleOfAP 
               Caption         =   "&Schedule of Accounts Payable"
            End
            Begin VB.Menu mnuReceivingReportRegister 
               Caption         =   "Receiving Report &Register"
            End
         End
         Begin VB.Menu mnuR_CashDisbursement 
            Caption         =   "&Cash Disbursement"
            Begin VB.Menu mnuR_CashDisbursementJournal 
               Caption         =   "&Cash Disbursement Journal"
            End
            Begin VB.Menu mnuRCD_LedgerCodeRunningBalance 
               Caption         =   "&Ledger Code Running Balance"
            End
            Begin VB.Menu mnuCheckRegister 
               Caption         =   "C&heck Register"
            End
         End
         Begin VB.Menu mnuSales 
            Caption         =   "&Sales"
            Begin VB.Menu mnuSalesJournal 
               Caption         =   "&Sales Journal"
            End
            Begin VB.Menu mnuAccountDetailByCustomer 
               Caption         =   "Account Detail by &Customer"
            End
            Begin VB.Menu mnuSJ_LedgerCodeRunningBalance 
               Caption         =   "&Ledger Code Running Balance"
            End
            Begin VB.Menu mnuScheduleOfAR 
               Caption         =   "Schedule of Accounts &Receivable"
            End
            Begin VB.Menu mnuAccountsReceivableAgingReport 
               Caption         =   "Accounts Receivable &Aging Report"
            End
            Begin VB.Menu mnuInvoiceRegister 
               Caption         =   "&Invoice Register"
            End
            Begin VB.Menu mnuSalesByInvoiceType 
               Caption         =   "Sales by &Invoice Type"
               Begin VB.Menu mnuSalesInvoices 
                  Caption         =   "&Sales Invoices"
                  Begin VB.Menu mnuVehicleSalesInvoices 
                     Caption         =   "&Vehicle Sales Invoices"
                  End
                  Begin VB.Menu mnuVehicleSalesInvoicesH 
                     Caption         =   "&Hyundai Vehicle Sales Invoices"
                     Visible         =   0   'False
                  End
               End
               Begin VB.Menu mnuPartsInvoice 
                  Caption         =   "&Parts Invoices"
                  Begin VB.Menu mnuPartsCashInvoices 
                     Caption         =   "Parts &Cash Invoices"
                  End
                  Begin VB.Menu mnuPartsChargeInvoices 
                     Caption         =   "Parts C&harge Invoices"
                  End
                  Begin VB.Menu mnuHyundaiPartsCI 
                     Caption         =   "&Hyundai Parts Cash Invoices"
                     Visible         =   0   'False
                  End
                  Begin VB.Menu mnuHyundaiPCG 
                     Caption         =   "H&yundai Parts Charge Invoices"
                     Visible         =   0   'False
                  End
               End
               Begin VB.Menu mnuServiceInvoices 
                  Caption         =   "S&ervice Invoices"
                  Begin VB.Menu mnuServiceInvoiceCash 
                     Caption         =   "Service &Cash Invoices"
                  End
                  Begin VB.Menu mnuServiceInvoiceCharge 
                     Caption         =   "Service C&harge Invoices"
                  End
                  Begin VB.Menu mnuHyundaiSCI 
                     Caption         =   "&Hyundai Service Cash Invoices"
                     Visible         =   0   'False
                  End
                  Begin VB.Menu mnuHyundaiSCGI 
                     Caption         =   "H&yundai Service Charge Invoices"
                     Visible         =   0   'False
                  End
               End
            End
            Begin VB.Menu mnuUnusedInvoices 
               Caption         =   "&Unused Invoices"
            End
         End
         Begin VB.Menu mnuCashReceipts 
            Caption         =   "Cash &Receipts"
            Begin VB.Menu mnuCashReceiptsJournal 
               Caption         =   "&Cash Receipts Journal"
            End
            Begin VB.Menu mnuCRJ_LedgeCodeRunningBalance 
               Caption         =   "Ledger Code Running Balance"
            End
            Begin VB.Menu mnuORRegister 
               Caption         =   "&OR Register"
            End
            Begin VB.Menu mnuUnusedOR 
               Caption         =   "&Unused OR"
            End
         End
         Begin VB.Menu mnuR_GeneralJournal 
            Caption         =   "&General Journal"
            Begin VB.Menu mnuR_JournalVoucherSummary 
               Caption         =   "&Journal Voucher Summary"
            End
            Begin VB.Menu mnuRGJ_LedgerCodeRunningBalance 
               Caption         =   "&Ledger Code Running Balance"
            End
         End
      End
      Begin VB.Menu mnuLedgers 
         Caption         =   "&Ledgers"
         Visible         =   0   'False
         Begin VB.Menu mnuR_TrialBalance 
            Caption         =   "&Trial Balance"
         End
         Begin VB.Menu mnuR_GeneralLedger 
            Caption         =   "&General Ledger"
         End
         Begin VB.Menu mnuR_ScheduleOfAccountsPayable 
            Caption         =   "Schedule of Accounts &Payable"
         End
         Begin VB.Menu mnuR_AccountsPayableLedger 
            Caption         =   "&Accounts Payable Ledger"
         End
      End
      Begin VB.Menu mnuFinancialStatement 
         Caption         =   "&Financial Statement"
         Begin VB.Menu mnuFS_WorkSheet 
            Caption         =   "&Work Sheet"
         End
         Begin VB.Menu mnuTrialBalance 
            Caption         =   "&Trial Balance"
         End
         Begin VB.Menu mnuIncomeStatement 
            Caption         =   "&Income Statement"
         End
         Begin VB.Menu mnuBalanceSheet 
            Caption         =   "&Balance Sheet"
         End
         Begin VB.Menu mnuStatementOfOwnersEquity 
            Caption         =   "&Statement of Owners Equity"
         End
         Begin VB.Menu mnuStatementOfCashFlow 
            Caption         =   "Statement of &Cash Flow"
         End
         Begin VB.Menu mnuScheduleOfAdjustments 
            Caption         =   "&Schedule of Adjustments"
         End
         Begin VB.Menu mnuFinancialStatements 
            Caption         =   "&Financial Statements"
         End
         Begin VB.Menu mnuHyundaiFS 
            Caption         =   "&Hyundai Financial Statements"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuDisbursementReport 
         Caption         =   "&Disbursement Report"
         Begin VB.Menu mnuDailyCheckDisbursementReport 
            Caption         =   "&Daily Check Disbursement Report"
         End
         Begin VB.Menu mnuCheckDisbursementReport 
            Caption         =   "&Check Disbursement Report"
         End
      End
      Begin VB.Menu mnuRLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExpenseReport 
         Caption         =   "&Expense Report"
         Begin VB.Menu mnuScheduleOfAdministrativeExpense 
            Caption         =   "Schedule of &Administrative Expense"
         End
         Begin VB.Menu mnuScheduleOfAdminExpense 
            Caption         =   "Schedule of &Selling Expense"
         End
      End
      Begin VB.Menu mnuRLine2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBooksOfAccount 
         Caption         =   "&Books of Account"
         Visible         =   0   'False
         Begin VB.Menu mnuBOA_Journal 
            Caption         =   "&Journal"
            Begin VB.Menu mnuBOA_All 
               Caption         =   "&All"
            End
            Begin VB.Menu mnuBOA_CheckDisbursement 
               Caption         =   "&Check Disbursement Journal"
            End
            Begin VB.Menu mnuBOA_CashReceiptJournal 
               Caption         =   "Cash &Receipt Journal"
            End
         End
         Begin VB.Menu mnuBOA_Ledger 
            Caption         =   "&Ledger"
         End
         Begin VB.Menu mnuBOA_ColumnarCashReceipt 
            Caption         =   "Columnar (Cash Receipt)"
         End
         Begin VB.Menu mnuBOA_ColumnarCashDisbursement 
            Caption         =   "Columnar (Cash Disbursement)"
         End
         Begin VB.Menu mnuBOA_SubsidiarySalesJournal 
            Caption         =   "Subsidiary &Sales Journal"
         End
         Begin VB.Menu mnuBOA_SubsidiaryPurchaseJournal 
            Caption         =   "Subsidiary &Purchase Journal"
         End
      End
      Begin VB.Menu mnuDepreciationAssetReport 
         Caption         =   "Depreciation Asset &Report"
         Begin VB.Menu mnuAssetList 
            Caption         =   "Asset &List"
         End
         Begin VB.Menu mnuDepreciatedAsset 
            Caption         =   "&Depreciated Asset"
         End
         Begin VB.Menu mnuDAR_Line 
            Caption         =   "-"
         End
         Begin VB.Menu mnuScheduleOfDepreciation 
            Caption         =   "&Schedule of Depreciation"
         End
      End
      Begin VB.Menu mnuSchedules 
         Caption         =   "&Schedules"
         Begin VB.Menu mnuSchedIncomeTaxesWheldSupplier 
            Caption         =   "&Schedule of Income Taxes W/Held from Suppliers"
         End
         Begin VB.Menu mnuSchedExpandedWTax 
            Caption         =   "Schedule of Payees Subject to &Expanded Withholding Tax "
         End
      End
      Begin VB.Menu mnuAuditAdjustmentReports 
         Caption         =   "Audit Ad&justment Reports"
         Begin VB.Menu mnuAuditSummary 
            Caption         =   "Audit Adjustment &Summary"
         End
         Begin VB.Menu mnuAuditAdjustmentJournal 
            Caption         =   "Audit Adjustment &Journal"
         End
      End
   End
   Begin VB.Menu mnuMaintenance 
      Caption         =   "&Maintenance"
      Begin VB.Menu mnuCompany 
         Caption         =   "&System Setup"
      End
      Begin VB.Menu mnuPassword 
         Caption         =   "&User Modules"
      End
      Begin VB.Menu mnuPasswordMaintenance 
         Caption         =   "&Password Maintenance"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuUpdateCustomerCodeControl 
         Caption         =   "&Update Customer Code Control"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ImportPurchases_Click()
Screen.MousePointer = 11
frmAPJImport.Show
Screen.MousePointer = 0
End Sub

Private Sub MDIForm_Load()
    'ApplySkin
    
    CenterMe Screen, Me, 0
    
End Sub

Private Sub MDIForm_Resize()
CenterMe Screen, Me, 0
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If MsgBox("Exit AMIS? Are you Sure?", vbQuestion + vbYesNo, "Exit Application") = vbYes Then
    SkinFramework1.RemoveAllWindows
   End
Else
   Cancel = 1
End If
End Sub

Private Sub mnuAbout_Click()
Screen.MousePointer = 11
'frmAbout.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuByEmployee_Click()
Screen.MousePointer = 11
frmPrintByEmployee.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuAccountDetailByCustomer_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "SJ"
frmAMISDetailBySupplierWithAccountCode.Show
frmAMISDetailBySupplierWithAccountCode.Caption = "Accounts Detail Report By Customer"
Screen.MousePointer = 0
End Sub

Private Sub mnuAccountsPayableAgingReport_Click()

End Sub

Private Sub mnuAccountsPayableJournal_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "ACCOUNTS PAYABLE JOURNAL") = False Then Exit Sub
End If
Screen.MousePointer = 11
JOURNALTYPE = "APJ"
On Error Resume Next
Unload frmAMISJournalEntry
frmAMISJournalEntry.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuAccountsReceivableAgingReport_Click()
Screen.MousePointer = 11
Report_Ar = "AGING"
frmAMISARSchedReport.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuAccountsReceivableJournal_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "SALES JOURNAL") = False Then Exit Sub
End If
Screen.MousePointer = 11
JOURNALTYPE = "SJ"
On Error Resume Next
Unload frmAMISJournalEntry
frmAMISJournalEntry.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuAccountSubHeaders_Click()
Screen.MousePointer = 11
frmAMISFILESSubHeader.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuAccountTitles_Click()
Screen.MousePointer = 11
frmAMISFILESTitleCode.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuAccountTypes_Click()
Screen.MousePointer = 11
frmAMISFILESAccType.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuAdjustments_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "ADJUSTMENTS") = False Then Exit Sub
End If
End Sub

Private Sub mnuAssetList_Click()
Screen.MousePointer = 11
Dim rsProfile As ADODB.Recordset
rptMain.Reset
Set rsProfile = New ADODB.Recordset
Set rsProfile = gconDMIS.Execute("Select * from AMIS_Profile")
If Not (rsProfile.EOF And rsProfile.BOF) Then
   rptMain.Formulas(3) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
   rptMain.Formulas(4) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
   rptMain.ReportTitle = "LIST OF ASSETS"
End If
PrintReport rptMain, AMIS_REPORT_PATH & "\Files\ListOfAssets.rpt", "", 1
Screen.MousePointer = 0
End Sub

Private Sub mnuAssetRegistry_Click()
Screen.MousePointer = 11
frmAMISDATAAssets.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuAuditAdjustmentJournal_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "ADJ"
frmAMISRangeWithSummary.Show
frmAMISRangeWithSummary.Caption = "Audit Adjustment Journal"
Screen.MousePointer = 0
End Sub

Private Sub mnuAuditAdjustments_Click()
Screen.MousePointer = 11
JOURNALTYPE = "ADJ"
On Error Resume Next
Unload frmAMISJournalEntry
frmAMISJournalEntry.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuAuditSummary_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "ADJ"
frmAMISRange.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuBalanceSheet_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "BALANCE SHEET") = False Then Exit Sub
End If
Screen.MousePointer = 11
frmAMISBalanceSheet.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuBanksMaster_Click()
Screen.MousePointer = 11
frmAMISMASTERFILEBanks.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuBirExtraction_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "BIR DATA EXTRACTION") = False Then Exit Sub
End If
Screen.MousePointer = 11
EXTRACT_TYPE = "PURCHASE"
frmBIRExtract.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuCashDisbursement_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "CASH DISBURSEMENT JOURNAL") = False Then Exit Sub
End If
Screen.MousePointer = 11
JOURNALTYPE = "CDJ"
On Error Resume Next
Unload frmAMISJournalEntry
frmAMISJournalEntry.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuCashReceipts_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "CASH RECEIPTS") = False Then Exit Sub
End If
End Sub

Private Sub mnuCashReceiptsJournal_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "CRJ"
frmAMISRangeWithSummary.Show
frmAMISRangeWithSummary.Caption = "Cash Receipts Journal"
Screen.MousePointer = 0
End Sub

Private Sub mnuCashSalesJournal_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "CASH RECEIPTS JOURNAL") = False Then Exit Sub
End If
Screen.MousePointer = 11
JOURNALTYPE = "CRJ"
On Error Resume Next
Unload frmAMISJournalEntry
frmAMISJournalEntry.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuChartOfAccount_Click()
Screen.MousePointer = 11
frmAMISFILESChartOfAccount.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuCheckRegister_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "CHECK_REGISTER"
frmAMISRange.Show
frmAMISRange.Caption = "Check Registers"
DoEvents
End Sub

Private Sub mnuClosingEntries_Click()
Screen.MousePointer = 11
JOURNALTYPE = "CLO"
On Error Resume Next
Unload frmAMISJournalEntry
frmAMISJournalEntry.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuCompany_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "SYSTEM SETUP") = False Then Exit Sub
End If
Screen.MousePointer = 11
frmAMISProfile.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuCRJ_LedgeCodeRunningBalance_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "CRJ"
frmAMISRangeWithAccountCode.Show
frmAMISRangeWithAccountCode.Caption = "Cash Receipts Ledger Code Running Balance"
Screen.MousePointer = 0
End Sub

Private Sub mnuCustomerCreditMemo_Click()
Screen.MousePointer = 11
frmAMISCustomerAdjustment.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuCustomerMasterList_Click()
ShowReport "Customers", "Files", "", "Customers Master List", "AS OF: " & LOGDATE, True
End Sub

Private Sub mnuCustomers_Click()
Screen.MousePointer = 11
frmALLCustomer.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuCustomersDeposit_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "CUSTOMERS DEPOSIT LEDGER") = False Then Exit Sub
End If
Screen.MousePointer = 11
CUST_LEDGER_TYPE = "CUSTDEPOSIT"
frmAMISLEDGERCustomers.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuCustomersLedger_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "A/R LEDGER") = False Then Exit Sub
End If
Screen.MousePointer = 11
CUST_LEDGER_TYPE = "ARLEDGER"
frmAMISLEDGERCustomers.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuCustOpenBalance_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "CUSTOMER OPENING BALANCE") = False Then Exit Sub
End If
Screen.MousePointer = 11
On Error Resume Next
JOURNALTYPE = "COB"
frmAMISCustomerAROpening.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuDepartment_Click()
Screen.MousePointer = 11
frmAMISFILESDepartment.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuDisbursementReport_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "DISBURSEMENT REPORT") = False Then Exit Sub
End If
End Sub

Private Sub mnuExit_Click()
If MsgBox("Exit AMIS? Are you Sure?", vbQuestion + vbYesNo, "Exit Application") = vbYes Then
   End
End If
End Sub

Private Sub mnuExpenseReport_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "EXPENSE REPORT") = False Then Exit Sub
End If
End Sub

Private Sub mnuExtractSalesBIRRelief_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "BIR DATA EXTRACTION") = False Then Exit Sub
End If
Screen.MousePointer = 11
EXTRACT_TYPE = "SALES"
frmBIRExtract.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuFiles_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "FILES") = False Then Exit Sub
End If
End Sub

Private Sub mnuFinancialStatement_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "FINANCIAL STATEMENT") = False Then Exit Sub
End If
End Sub

Private Sub mnuFinancialStatements_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "FINANCIAL STATEMENTS") = False Then Exit Sub
End If
Screen.MousePointer = 11
frmAMISFinancialStatements.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuFS_WorkSheet_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "WORK SHEET") = False Then Exit Sub
End If
Screen.MousePointer = 11
frmAMISWorkSheet.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuGeneralJournal_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "GENERAL JOURNAL") = False Then Exit Sub
End If
Screen.MousePointer = 11
JOURNALTYPE = "GJ"
On Error Resume Next
Unload frmAMISJournalEntry
frmAMISJournalEntry.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuGeneralLedger_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "GENERAL LEDGER") = False Then Exit Sub
End If
Screen.MousePointer = 11
frmAMISLEDGERAccounts.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuHeader_Click()
Screen.MousePointer = 11
frmAMISFILESHeader.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuHyundaiFS_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "FINANCIAL STATEMENTS") = False Then Exit Sub
End If
Screen.MousePointer = 11
frmAMISHYUNDAIFinancialStatements.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuHyundaiPartsCI_Click()
Screen.MousePointer = 11
INVOICE_Type = "H_PARTS-CASH"
frmAMISSalesByInvoiceType.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuHyundaiPCG_Click()
Screen.MousePointer = 11
INVOICE_Type = "H_PARTS-CHARGE"
frmAMISSalesByInvoiceType.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuHyundaiSCGI_Click()
Screen.MousePointer = 11
INVOICE_Type = "H_SERVICE-CHARGE"
frmAMISSalesByInvoiceType.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuHyundaiSCI_Click()
Screen.MousePointer = 11
INVOICE_Type = "H_SERVICE-CASH"
frmAMISSalesByInvoiceType.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuImportCashReceipts_Click()
Screen.MousePointer = 11
frmCRJImport.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuImportPartsServiceSales_Click()
Screen.MousePointer = 11
frmSALESImport.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuIncomeStatement_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "INCOME STATEMENT") = False Then Exit Sub
End If
Screen.MousePointer = 11
frmAMISIncomeStatement.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuIncomeStatementbyProduct_Click()
Screen.MousePointer = 11
frmAMISIncomeStatementByProduct.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuIncomeStatements_Click()
Screen.MousePointer = 11
frmAMISIncomeStatements.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuIntegrateCustomerMasterFile_Click()
Screen.MousePointer = 11
frmIntegrateALL_CUSTMASTER_AMIS.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuInvoiceRegister_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "INV_REGISTER"
frmAMISRange.Show
frmAMISRange.Caption = "Invoices Registers"
DoEvents
End Sub

Private Sub mnuInvoiceTypes_Click()
Screen.MousePointer = 11
frmAMISMASTERFILEInvoiceType.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuJournals_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "JOURNALS") = False Then Exit Sub
End If
End Sub

Private Sub mnuLedgerCodeRunningBalance_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "APJ"
frmAMISRangeWithAccountCode.Show
frmAMISRangeWithAccountCode.Caption = "Accounts Payable Ledger Code Running Balance"
Screen.MousePointer = 0
End Sub

Private Sub mnuMaintenance_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "MAINTENANCE") = False Then Exit Sub
End If
End Sub

Private Sub mnuMYOBGeneralLedger_Click()
Screen.MousePointer = 11
frmAMISMYOBLEDGERAccounts.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuOpeningBalance_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "OPENING BALANCE") = False Then Exit Sub
End If
Screen.MousePointer = 11
JOURNALTYPE = "OPB"
On Error Resume Next
Unload frmAMISJournalEntry
frmAMISJournalEntry.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuORRegister_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "OR_REGISTER"
frmAMISRange.Show
frmAMISRange.Caption = "O.R. Registers"
DoEvents
End Sub

Private Sub mnuPAJE_Click()
Screen.MousePointer = 11
JOURNALTYPE = "PDJ"
On Error Resume Next
Unload frmAMISJournalEntry
frmAMISJournalEntry.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuPartsCashInvoices_Click()
Screen.MousePointer = 11
INVOICE_Type = "PARTS-CASH"
frmAMISSalesByInvoiceType.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuPartsChargeInvoices_Click()
Screen.MousePointer = 11
INVOICE_Type = "PARTS-CHARGE"
frmAMISSalesByInvoiceType.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuPassword_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "PASSWORD MAINTENANCE") = False Then Exit Sub
End If
Screen.MousePointer = 11
'frmAccMaintenance.Show
frmusers.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuPayeeMaster_Click()
Screen.MousePointer = 11
frmPayee.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuPasswordMaintenance_Click()
Screen.MousePointer = 11
frmAccMaintenance.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuR_AccountDetailBySupplier_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "APJ"
frmAMISDetailBySupplierWithAccountCode.Show
frmAMISDetailBySupplierWithAccountCode.Caption = "Accounts Payable Detail Report By Supplier"
Screen.MousePointer = 0
End Sub

Private Sub mnuR_AccountsPayable_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "ACCOUNTS PAYABLE") = False Then Exit Sub
End If
End Sub

Private Sub mnuR_AccountsPayableDueReport_Click()
Screen.MousePointer = 11
frmAMISDueReport.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuR_AccountsPayableJournal_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "APJ"
frmAMISRangeWithSummary.Show
frmAMISRangeWithSummary.Caption = "Accounts Payable Journal"
Screen.MousePointer = 0
End Sub

Private Sub mnuR_CashDisbursement_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "CASH DISBURSEMENT") = False Then Exit Sub
End If
End Sub

Private Sub mnuR_CashDisbursementJournal_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "CDJ"
frmAMISRangeWithSummary.Show
frmAMISRangeWithSummary.Caption = "Cash Disbursement Journal"
Screen.MousePointer = 0
End Sub

Private Sub mnuR_ChartOfAccounts_Click()
ShowReport "ChartofAccounts", "AccountFiles", "", "Chart of Accounts", "AS OF: " & LOGDATE, True
End Sub

Private Sub mnuR_GeneralJournal_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "JOURNAL VOUCHER") = False Then Exit Sub
End If
End Sub

Private Sub mnuR_JournalVoucherSummary_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "JVS"
frmAMISRange.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuRCD_LedgerCodeRunningBalance_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "CDJ"
frmAMISRangeWithAccountCode.Show
frmAMISRangeWithAccountCode.Caption = "Cash Disbursement Ledger Code Running Balance"
Screen.MousePointer = 0
End Sub

Private Sub mnuReceivingReportRegister_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "REC_REGISTER"
frmAMISDetailBySupplierWithAccountCode.Show
frmAMISDetailBySupplierWithAccountCode.Caption = "Receiving Report Registers"
End Sub

Private Sub mnuReconcile_Click()
    FrmBankRecon.Show
End Sub

Private Sub mnuReports_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "REPORTS") = False Then Exit Sub
End If
End Sub

Private Sub mnuRGJ_LedgerCodeRunningBalance_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "GJ"
frmAMISRangeWithAccountCode.Show
frmAMISRangeWithAccountCode.Caption = "Journal Voucher Ledger Code Running Balance"
Screen.MousePointer = 0
End Sub

Private Sub mnuSales_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "SALES") = False Then Exit Sub
End If
End Sub

Private Sub mnuSalesJournal_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "SJ"
frmAMISRangeWithSummary.Show
frmAMISRangeWithSummary.Caption = "Sales Journal"
Screen.MousePointer = 0
End Sub

Private Sub mnuSchedExpandedWTax_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "EX_TAX"
frmAMISRange.Show
frmAMISRange.Caption = "Schedule of Payees Subject to Expanded Withholding Tax"
Screen.MousePointer = 0
End Sub

Private Sub mnuSchedIncomeTaxesWheldSupplier_Click()
Screen.MousePointer = 11
frmAMISYearly.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuScheduleOfAdjustments_Click()
Screen.MousePointer = 11
frmAMISSchedAdjust.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuScheduleOfAdminExpense_Click()
Screen.MousePointer = 11
REPORT_EXPENSETYPE = "SELLING"
frmAMISExpenseReport.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuScheduleOfAdministrativeExpense_Click()
Screen.MousePointer = 11
REPORT_EXPENSETYPE = "ADMIN"
frmAMISExpenseReport.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuScheduleOfAP_Click()
Screen.MousePointer = 11
frmAMISAPSchedReport.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuScheduleOfAR_Click()
Screen.MousePointer = 11
Report_Ar = "SCHED"
frmAMISARSchedReport.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuScheduleOfDepreciation_Click()
Screen.MousePointer = 11
frmAMISMonthlyYearly.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuServiceInvoiceCash_Click()
Screen.MousePointer = 11
INVOICE_Type = "SERVICE-CASH"
frmAMISSalesByInvoiceType.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuServiceInvoiceCharge_Click()
Screen.MousePointer = 11
INVOICE_Type = "SERVICE-CHARGE"
frmAMISSalesByInvoiceType.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuSJ_LedgerCodeRunningBalance_Click()
Screen.MousePointer = 11
REPORT_RANGETYPE = "SJ"
frmAMISRangeWithAccountCode.Show
frmAMISRangeWithAccountCode.Caption = "Sales Journal Ledger Code Running Balance"
Screen.MousePointer = 0
End Sub

Private Sub mnuStatementOfCashFlow_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "STATEMENT OF CASH FLOW") = False Then Exit Sub
End If
End Sub

Private Sub mnuStatementOfOwnersEquity_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "STATEMENT OF OWNERS EQUITY") = False Then Exit Sub
End If
End Sub

Private Sub mnuSupplierList_Click()
    ShowReport "Suppliers", "Files", "", "Suppliers Master List", "AS OF: " & LOGDATE, True
End Sub

Private Sub mnuTaxRateCode_Click()
Screen.MousePointer = 11
frmAMISMASTERFILEATC.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuTemplates_Click()
Screen.MousePointer = 11
frmAMISMASTERFILESTemplates.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuTermsofPayment_Click()
Screen.MousePointer = 11
frmAMISMASTERFILEPayTerm.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuTrialBalance_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "TRIAL BALANCE") = False Then Exit Sub
End If
Screen.MousePointer = 11
frmAMISTrialBalance.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuUnusedInvoices_Click()
Screen.MousePointer = 11
frmAMISProcessUnusedInvoices.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuUnusedOR_Click()
Screen.MousePointer = 11
frmAMISProcessUnusedOR.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuUpdateCustomerCodeControl_Click()
Screen.MousePointer = 11
frmAMISDATACusctl.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuVehicleSalesInvoices_Click()
Screen.MousePointer = 11
INVOICE_Type = "VEHICLE"
frmAMISSalesByInvoiceType.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuVehicleSalesInvoicesH_Click()
Screen.MousePointer = 11
INVOICE_Type = "H_VEHICLE"
frmAMISSalesByInvoiceType.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuVendorAdjustments_Click()
Screen.MousePointer = 11
frmAMISVendorAdjustment.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuVendorLedger_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "VENDORS LEDGER") = False Then Exit Sub
End If
Screen.MousePointer = 11
frmAMISLEDGERVendors.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuVendorMaster_Click()
Screen.MousePointer = 11
frmAMISMASTERFILEVendor.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuVendorOpeningBalance_Click()
If ApplySecurityValidation = True Then
   If Module_Access(LOGID, "VENDOR OPENING BALANCE") = False Then Exit Sub
End If
Screen.MousePointer = 11
On Error Resume Next
frmAMISVendorAPOpening.Show
Screen.MousePointer = 0
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
       Case 1
            Screen.MousePointer = 11
            frmAMISFILESChartOfAccount.Show
            Screen.MousePointer = 0
       Case 3
            Screen.MousePointer = 11
            JOURNALTYPE = "APJ"
            On Error Resume Next
            Unload frmAMISJournalEntry
            frmAMISJournalEntry.Show
            Screen.MousePointer = 0
       Case 4
            Screen.MousePointer = 11
            JOURNALTYPE = "CDJ"
            On Error Resume Next
            Unload frmAMISJournalEntry
            frmAMISJournalEntry.Show
            Screen.MousePointer = 0
       Case 5
            Screen.MousePointer = 11
            JOURNALTYPE = "SJ"
            On Error Resume Next
            Unload frmAMISJournalEntry
            frmAMISJournalEntry.Show
            Screen.MousePointer = 0
       Case 6
            Screen.MousePointer = 11
            JOURNALTYPE = "CRJ"
            On Error Resume Next
            Unload frmAMISJournalEntry
            frmAMISJournalEntry.Show
            Screen.MousePointer = 0
       Case 7
            Screen.MousePointer = 11
            JOURNALTYPE = "CRJ"
            On Error Resume Next
            Unload frmAMISJournalEntry
            frmAMISJournalEntry.Show
            Screen.MousePointer = 0
       Case 9
            Screen.MousePointer = 11
            frmAMISLEDGERAccounts.Show
            Screen.MousePointer = 0
       Case 10
            Screen.MousePointer = 11
            frmAMISLEDGERCustomers.Show
            Screen.MousePointer = 0
       Case 11
            Screen.MousePointer = 11
            frmAMISLEDGERVendors.Show
            Screen.MousePointer = 0
       Case 13
            Screen.MousePointer = 11
            REPORT_RANGETYPE = "APJ"
            frmAMISRangeWithSummary.Show
            frmAMISRangeWithSummary.Caption = "Accounts Payable Journal"
            Screen.MousePointer = 0
       Case 14
            Screen.MousePointer = 11
            REPORT_RANGETYPE = "CDJ"
            frmAMISRangeWithSummary.Show
            frmAMISRangeWithSummary.Caption = "Cash Disbursement Journal"
            Screen.MousePointer = 0
       Case 15
            Screen.MousePointer = 11
            REPORT_RANGETYPE = "SJ"
            frmAMISRangeWithSummary.Show
            frmAMISRangeWithSummary.Caption = "Sales Journal"
            Screen.MousePointer = 0
       Case 16
            Screen.MousePointer = 11
            REPORT_RANGETYPE = "CRJ"
            frmAMISRangeWithSummary.Show
            frmAMISRangeWithSummary.Caption = "Cash Receipts Journal"
            Screen.MousePointer = 0
       Case 17
            Screen.MousePointer = 11
            REPORT_RANGETYPE = "JVS"
            frmAMISRange.Show
            Screen.MousePointer = 0
       Case 19
            Screen.MousePointer = 11
            frmAMISTrialBalance.Show
            Screen.MousePointer = 0
       Case 20
            Screen.MousePointer = 11
            frmAMISWorkSheet.Show
            Screen.MousePointer = 0
       Case 21
            Screen.MousePointer = 11
            frmAMISFinancialStatements.Show
            Screen.MousePointer = 0
       Case 23
            Screen.MousePointer = 11
            frmAbout.Show
            Screen.MousePointer = 0
       Case 24
            If MsgBox("Exit AMIS? Are you Sure?", vbQuestion + vbYesNo, "Exit Application") = vbYes Then
               End
            End If
End Select
End Sub








Sub ApplySkin()
Dim SkinPath As String
    'SkinPath = GetSetting("DMIS", "SKIN", "Location")
Dim skinType As Integer

    skinType = 0
 SkinPath = "E:\HMI\REFERENCES\skin"
    Select Case skinType
        Case 0:
            SkinFramework1.LoadSkin SkinPath & "\Vista.cjstyles", "NormalBlack.ini"
        Case 1:
        Case 2:
        Case 3:
        Case 4:
    End Select
        SkinFramework1.ApplyWindow App.hInstance
        SkinFramework1.AutoApplyNewThreads = True
        SkinFramework1.AutoApplyNewWindows = True
        SkinFramework1.ApplyOptions = xtpSkinApplyFrame Or xtpSkinApplyColors Or xtpSkinApplyMetrics
        SkinFramework1.EnableThemeDialogTexture Picture2.hWnd, 0
End Sub

