VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{7E0552E8-E2C9-4C9E-BC75-F34D871C75F4}#2.0#0"; "wizEncrypt.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Accounting Management Information System"
   ClientHeight    =   6060
   ClientLeft      =   3000
   ClientTop       =   2295
   ClientWidth     =   9255
   Icon            =   "AMIS_Main.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "AMIS_Main.frx":0442
   WindowState     =   2  'Maximized
   Begin wizEncrypt.wizEnc wizEnc1 
      Left            =   360
      Top             =   2850
      _ExtentX        =   3969
      _ExtentY        =   3969
   End
   Begin VB.PictureBox picBars 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   9225
      TabIndex        =   20
      Top             =   5745
      Width           =   9255
      Begin VB.CommandButton barTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8610
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   3885
      End
      Begin VB.CommandButton barDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4740
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   3885
      End
      Begin VB.CommandButton barUserLevel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   2355
      End
      Begin VB.CommandButton BarUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         Width           =   2355
      End
   End
   Begin VB.PictureBox picToolBars 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   9225
      TabIndex        =   19
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmdTool15 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   10920
         MaskColor       =   &H0000FFFF&
         Picture         =   "AMIS_Main.frx":112806
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Exit"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdTool14 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   10140
         MaskColor       =   &H0000FFFF&
         Picture         =   "AMIS_Main.frx":112B10
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "About the Author"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdTool13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   9360
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "BIR Alphalist Processing"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdTool12 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   8580
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Year-To-Date Processing"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdTool11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7800
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "BIR Alphalist Processing"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdTool10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7020
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Year-To-Date Processing"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdTool9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6240
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Generate Payroll"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdTool8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5460
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Ledger"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdTool7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4680
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print ATM Advice"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdTool6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3900
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Adjustment"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdTool5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3120
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "ATM Entry"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdTool4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2340
         MaskColor       =   &H0000FFFF&
         OLEDropMode     =   1  'Manual
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Commission"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdTool3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1560
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Deductions"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdTool2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   780
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Overtime"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdTool1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   0
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Active Employees"
         Top             =   0
         Width           =   795
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AMIS_Main.frx":112E1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AMIS_Main.frx":113134
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AMIS_Main.frx":11344E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AMIS_Main.frx":113768
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AMIS_Main.frx":113A82
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AMIS_Main.frx":113D9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AMIS_Main.frx":1140B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AMIS_Main.frx":1143D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AMIS_Main.frx":1146EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AMIS_Main.frx":114A04
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AMIS_Main.frx":114D1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AMIS_Main.frx":115038
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AMIS_Main.frx":115352
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "&Files"
      Begin VB.Menu mnuChartOfAccount 
         Caption         =   "&Chart of Accounts"
      End
      Begin VB.Menu mnuAssetsMaster 
         Caption         =   "&Assets Registry List"
      End
      Begin VB.Menu mnuMasterList 
         Caption         =   "&Master List"
         Begin VB.Menu mnuVendorMaster 
            Caption         =   "&Vendors"
         End
         Begin VB.Menu mnuCustomerMaster 
            Caption         =   "C&ustomers"
         End
         Begin VB.Menu mnuPayeeMaster 
            Caption         =   "&Payee"
         End
         Begin VB.Menu mnuBanksMaster 
            Caption         =   "&Banks"
         End
         Begin VB.Menu mnuIncomeMaster 
            Caption         =   "&Incomes"
         End
         Begin VB.Menu mnuChargeMaster 
            Caption         =   "Char&ges"
         End
         Begin VB.Menu mnuExpenseMaster 
            Caption         =   "&Expenses"
         End
      End
      Begin VB.Menu mnuF_Line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccountType 
         Caption         =   "Account &Type"
      End
      Begin VB.Menu mnuDepartment 
         Caption         =   "&Department"
      End
      Begin VB.Menu mnuSignatories 
         Caption         =   "&Signatories"
      End
   End
   Begin VB.Menu mnuJournals 
      Caption         =   "&Journals"
      Begin VB.Menu mnuOpeningBalance 
         Caption         =   "&Opening Balances"
      End
      Begin VB.Menu mnuJ_Line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGeneralJournal 
         Caption         =   "General &Journal"
      End
      Begin VB.Menu mnuPurchaseJournal 
         Caption         =   "&Purchases Journal"
      End
      Begin VB.Menu mnuCashPayments 
         Caption         =   "&Cash Payments Journal"
      End
      Begin VB.Menu mnuSalesJournal 
         Caption         =   "&Sales Journal"
      End
      Begin VB.Menu mnuCashReceipts 
         Caption         =   "Cash &Receipts Journal"
      End
   End
   Begin VB.Menu mnuLedger 
      Caption         =   "&Ledgers"
      Begin VB.Menu mnuGeneralLedger 
         Caption         =   "&General Ledger"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuAccounts 
         Caption         =   "&Accounts"
         Begin VB.Menu mnuR_ChartOfAccounts 
            Caption         =   "&Chart of Accounts"
         End
         Begin VB.Menu mnuVendorList 
            Caption         =   "&Vendor List"
         End
         Begin VB.Menu mnuCustomerList 
            Caption         =   "&Customer List"
         End
         Begin VB.Menu mnuR_PayeeList 
            Caption         =   "&Payee List"
         End
      End
      Begin VB.Menu mnuR_Journals 
         Caption         =   "&Journals"
         Begin VB.Menu mnuR_OpeningBalance 
            Caption         =   "&Opening Balances"
         End
         Begin VB.Menu mnuR_GeneralJournal 
            Caption         =   "&General Journal"
         End
         Begin VB.Menu mnuR_PurchasesJournal 
            Caption         =   "&Purchases Journal"
         End
         Begin VB.Menu mnuCashPayJournal 
            Caption         =   "&Cash Payments Journal"
         End
         Begin VB.Menu mnuR_SalesJournal 
            Caption         =   "&Sales Journal"
         End
         Begin VB.Menu mnuR_CashReceiptJournal 
            Caption         =   "Cash &Receipt Journal"
         End
      End
      Begin VB.Menu mnuLedgers 
         Caption         =   "&Ledgers"
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
         Begin VB.Menu mnuR_ScheduleofAccountsReceivable 
            Caption         =   "Schedule of Accounts &Receivable"
         End
         Begin VB.Menu mnuR_AccountsReceivableLedger 
            Caption         =   "Accounts Receivable &Ledger"
         End
      End
      Begin VB.Menu mnuFinancianStatements 
         Caption         =   "&Financial Statements"
         Begin VB.Menu mnuFS_IncomeStatement 
            Caption         =   "&Income Statement"
         End
         Begin VB.Menu mnuFS_BalanceSheet 
            Caption         =   "&Balance Sheet"
         End
         Begin VB.Menu mnuFS_RetainedEarnings 
            Caption         =   "&Retained Earnings Statement"
         End
      End
      Begin VB.Menu mnuFinancialAnalysis 
         Caption         =   "Fi&nancial Analysis"
         Begin VB.Menu mnuStatementofCashFlows 
            Caption         =   "&Statement of Cash Flows"
         End
      End
   End
   Begin VB.Menu mnuMaintenance 
      Caption         =   "&Maintenance"
      Begin VB.Menu mnuCompany 
         Caption         =   "&Company Profile"
      End
      Begin VB.Menu mnuPassword 
         Caption         =   "&Password Maintenance"
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

Private Sub barDate_Click()
If MsgBox("Change Current Date? Are you Sure?", vbQuestion + vbYesNo, "System Date") = vbYes Then
   If MsgBox("Warning: You will need to Re-Login Again to Successfully Change the Date, Do you want to Proceed?", vbQuestion + vbYesNo, "Warning") = vbYes Then
      frmSetDate.Show vbModal
   End If
End If
End Sub

Private Sub barTime_Click()
If MsgBox("Change Current Time? Are you Sure?", vbQuestion + vbYesNo, "System Time") = vbYes Then
   If MsgBox("Warning: You will need to Re-Login Again to Successfully Change the Time, Do you want to Proceed?", vbQuestion + vbYesNo, "Warning") = vbYes Then
      frmSetTime.Show vbModal
   End If
End If
End Sub

Private Sub BarUserName_Click()
If MsgBox("Do you want to re-Login?", vbQuestion + vbYesNo, "Re-Login") = vbYes Then
   frmSecurity.Show vbModal
End If
picBars.SetFocus
End Sub

Private Sub cmdTool10_Click()
Screen.MousePointer = 11
frmProfile.Show
Screen.MousePointer = 0
End Sub

Private Sub cmdTool11_Click()
Screen.MousePointer = 11
frmAccMaintenance.Show
Screen.MousePointer = 0
End Sub

Private Sub cmdTool14_Click()
Screen.MousePointer = 11
frmAbout.Show
Screen.MousePointer = 0
End Sub

Private Sub cmdTool15_Click()
If MsgBox("Exit AMIS? Are you Sure?", vbQuestion + vbYesNo, "Exit Application") = vbYes Then
   End
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If MsgBox("Exit AMIS? Are you Sure?", vbQuestion + vbYesNo, "Exit Application") = vbYes Then
   End
End If
End Sub

Private Sub mnuAbout_Click()
Screen.MousePointer = 11
frmAbout.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuByEmployee_Click()
Screen.MousePointer = 11
frmPrintByEmployee.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuAccountType_Click()
Screen.MousePointer = 11
frmAccType.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuAssetsRegistry_Click()
Screen.MousePointer = 11
frmAssets.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuBanksMaster_Click()
Screen.MousePointer = 11
frmBanks.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuCashPayments_Click()
Screen.MousePointer = 11
JOURNALTYPE = "CPJ"
On Error Resume Next
Unload frmJournalEntry
frmJournalEntry.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuCashReceipts_Click()
Screen.MousePointer = 11
JOURNALTYPE = "CRJ"
On Error Resume Next
Unload frmJournalEntry
frmJournalEntry.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuChargeMaster_Click()
Screen.MousePointer = 11
frmCharges.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuChartOfAccount_Click()
Screen.MousePointer = 11
frmChartOfAccount.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuCompany_Click()
Screen.MousePointer = 11
frmProfile.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuDDDEntry_Click()
Screen.MousePointer = 11
frmNDD.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuCustomerMaster_Click()
Screen.MousePointer = 11
frmCustomer.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuDepartment_Click()
Screen.MousePointer = 11
frmDepartment.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuExit_Click()
Unload Me
End
End Sub

Private Sub mnuExpenseMaster_Click()
Screen.MousePointer = 11
frmExpenses.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuGeneralJournal_Click()
Screen.MousePointer = 11
JOURNALTYPE = "GJ"
On Error Resume Next
Unload frmJournalEntry
frmJournalEntry.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuIncomeMaster_Click()
Screen.MousePointer = 11
frmIncomes.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuOpeningBalance_Click()
Screen.MousePointer = 11
JOURNALTYPE = "OB"
On Error Resume Next
Unload frmJournalEntry
frmJournalEntry.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuPassword_Click()
Screen.MousePointer = 11
frmAccMaintenance.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuPayeeMaster_Click()
Screen.MousePointer = 11
frmPayee.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuPurchaseJournal_Click()
Screen.MousePointer = 11
JOURNALTYPE = "PJ"
On Error Resume Next
Unload frmJournalEntry
frmJournalEntry.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuSalesJournal_Click()
Screen.MousePointer = 11
JOURNALTYPE = "SJ"
On Error Resume Next
Unload frmJournalEntry
frmJournalEntry.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuSignatories_Click()
Screen.MousePointer = 11
frmSignatories.Show
Screen.MousePointer = 0
End Sub

Private Sub mnuVendorMaster_Click()
Screen.MousePointer = 11
frmVendor.Show
Screen.MousePointer = 0
End Sub
