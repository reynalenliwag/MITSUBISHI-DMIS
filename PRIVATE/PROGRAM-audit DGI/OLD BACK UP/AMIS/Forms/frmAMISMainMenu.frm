VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AMIS Main Menu"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "frmAMISMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6075
   ScaleWidth      =   10230
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   7335
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   11235
      _Version        =   655364
      _ExtentX        =   19817
      _ExtentY        =   12938
      _StockProps     =   64
      Appearance      =   2
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   120
      PaintManager.MinTabWidth=   100
      ItemCount       =   5
      Item(0).Caption =   "Main Modules"
      Item(0).ControlCount=   18
      Item(0).Control(0)=   "cmdJournal_AP"
      Item(0).Control(1)=   "cmdJournal_Sales"
      Item(0).Control(2)=   "cmdJournal_General"
      Item(0).Control(3)=   "cmdJournal_CRJ"
      Item(0).Control(4)=   "cmdLedger_Account"
      Item(0).Control(5)=   "cmdLedger_Customer"
      Item(0).Control(6)=   "cmdLedger_VendorSubisdy"
      Item(0).Control(7)=   "cmdJournalDRJ"
      Item(0).Control(8)=   "Label22"
      Item(0).Control(9)=   "Label1"
      Item(0).Control(10)=   "Label2"
      Item(0).Control(11)=   "Label3"
      Item(0).Control(12)=   "Label4"
      Item(0).Control(13)=   "Label5"
      Item(0).Control(14)=   "Label8"
      Item(0).Control(15)=   "Label9"
      Item(0).Control(16)=   "Label28"
      Item(0).Control(17)=   "cmdJournal_CashDisburshment"
      Item(1).Caption =   "Tables"
      Item(1).ControlCount=   12
      Item(1).Control(0)=   "cmdTable_ATCCode"
      Item(1).Control(1)=   "cmdTables_InvoiceTypes"
      Item(1).Control(2)=   "cmdTables_TermsOfPayment"
      Item(1).Control(3)=   "cmdTables_Bank"
      Item(1).Control(4)=   "cmdVendor"
      Item(1).Control(5)=   "cmdTable_Customer"
      Item(1).Control(6)=   "Label52"
      Item(1).Control(7)=   "Label53"
      Item(1).Control(8)=   "Label55"
      Item(1).Control(9)=   "Label56"
      Item(1).Control(10)=   "Label57"
      Item(1).Control(11)=   "Label66"
      Item(2).Caption =   "File Maintenance"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabControl4"
      Item(3).Caption =   "Reports"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "TabControl2"
      Item(4).Caption =   "Other Setups"
      Item(4).ControlCount=   19
      Item(4).Control(0)=   "Command10"
      Item(4).Control(1)=   "Command27"
      Item(4).Control(2)=   "cmdAuditInquiry"
      Item(4).Control(3)=   "cmdAuditReport"
      Item(4).Control(4)=   "Command4"
      Item(4).Control(5)=   "Command1"
      Item(4).Control(6)=   "Command6"
      Item(4).Control(7)=   "Command9"
      Item(4).Control(8)=   "Command8"
      Item(4).Control(9)=   "Label11"
      Item(4).Control(10)=   "Label73"
      Item(4).Control(11)=   "Label14"
      Item(4).Control(12)=   "Label15"
      Item(4).Control(13)=   "Label45"
      Item(4).Control(14)=   "Label42"
      Item(4).Control(15)=   "Label20"
      Item(4).Control(16)=   "Label69"
      Item(4).Control(17)=   "Label68"
      Item(4).Control(18)=   "Picture2"
      Begin VB.CommandButton cmdJournalDRJ 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   480
         MouseIcon       =   "frmAMISMainMenu.frx":15162
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":152B4
         Style           =   1  'Graphical
         TabIndex        =   151
         Tag             =   "1047"
         ToolTipText     =   "View Cash Receipts Journal"
         Top             =   4320
         Width           =   720
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2325
         Left            =   -65230
         ScaleHeight     =   2325
         ScaleWidth      =   5385
         TabIndex        =   143
         Top             =   3210
         Visible         =   0   'False
         Width           =   5385
         Begin VB.CommandButton cmdTranType 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   1080
            MouseIcon       =   "frmAMISMainMenu.frx":15B1F
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":15C71
            Style           =   1  'Graphical
            TabIndex        =   150
            Tag             =   "1052"
            ToolTipText     =   "View Vendors Subsidiary Ledger"
            Top             =   1590
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton Command11 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   1080
            MouseIcon       =   "frmAMISMainMenu.frx":16CF3
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":16E45
            Style           =   1  'Graphical
            TabIndex        =   145
            Tag             =   "1052"
            ToolTipText     =   "View Vendors Subsidiary Ledger"
            Top             =   30
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton Command12 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   1080
            MouseIcon       =   "frmAMISMainMenu.frx":17775
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":178C7
            Style           =   1  'Graphical
            TabIndex        =   144
            Tag             =   "1052"
            ToolTipText     =   "View Vendors Subsidiary Ledger"
            Top             =   810
            Width           =   720
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Un-Imported Transactions"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1920
            TabIndex        =   147
            Top             =   240
            Visible         =   0   'False
            Width           =   2430
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Importing Templates"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1920
            TabIndex        =   146
            Top             =   1020
            Width           =   1965
         End
      End
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":181F7
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":18349
         Style           =   1  'Graphical
         TabIndex        =   131
         Tag             =   "1052"
         ToolTipText     =   "View Vendors Subsidiary Ledger"
         Top             =   4020
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Command9 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":18C79
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":18DCB
         Style           =   1  'Graphical
         TabIndex        =   130
         Tag             =   "1052"
         ToolTipText     =   "View Vendors Subsidiary Ledger"
         Top             =   4800
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -64150
         MouseIcon       =   "frmAMISMainMenu.frx":196FB
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1984D
         Style           =   1  'Graphical
         TabIndex        =   129
         Tag             =   "1052"
         ToolTipText     =   "View Data Tools"
         Top             =   2490
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Command1 
         Height          =   645
         Left            =   -64150
         MouseIcon       =   "frmAMISMainMenu.frx":1A17D
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1A2CF
         Style           =   1  'Graphical
         TabIndex        =   128
         ToolTipText     =   "View Unused O.R."
         Top             =   1680
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Command4 
         Height          =   645
         Left            =   -64150
         MouseIcon       =   "frmAMISMainMenu.frx":1AA9F
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1ABF1
         Style           =   1  'Graphical
         TabIndex        =   127
         ToolTipText     =   "View O.R. Register"
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdAuditReport 
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1B36F
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1B4C1
         Style           =   1  'Graphical
         TabIndex        =   126
         ToolTipText     =   "View Signatories and Headers"
         Top             =   3240
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdAuditInquiry 
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1B903
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1BA55
         Style           =   1  'Graphical
         TabIndex        =   125
         ToolTipText     =   "View Signatories and Headers"
         Top             =   2460
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Command27 
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1BE97
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1BFE9
         Style           =   1  'Graphical
         TabIndex        =   124
         ToolTipText     =   "View Signatories and Headers"
         Top             =   1680
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Command10 
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1C42B
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1C57D
         Style           =   1  'Graphical
         TabIndex        =   123
         Tag             =   "1102"
         ToolTipText     =   "View Reminders"
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTable_Customer 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1CDF8
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1CF4A
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "1027"
         ToolTipText     =   "View Customer Master Files"
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdVendor 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1D5B1
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1D703
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "1028"
         ToolTipText     =   "View Vendor Master Files"
         Top             =   1710
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTables_Bank 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1DE1A
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1DF6C
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "1029"
         ToolTipText     =   "View Bank Master Files"
         Top             =   2520
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTables_TermsOfPayment 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1E647
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1E799
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "1031"
         ToolTipText     =   "View Terms Of Payment Master File"
         Top             =   4125
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTables_InvoiceTypes 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1EE32
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1EF84
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "1030"
         ToolTipText     =   "View Invoice Type Master Files"
         Top             =   3315
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTable_ATCCode 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1F636
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1F788
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "1032"
         ToolTipText     =   "View ATC Code Master File"
         Top             =   4935
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdJournal_CashDisburshment 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   480
         MouseIcon       =   "frmAMISMainMenu.frx":1FE3B
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1FF8D
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "1045"
         ToolTipText     =   "View Cash Disbursement Journal"
         Top             =   1740
         Width           =   720
      End
      Begin VB.CommandButton cmdLedger_VendorSubisdy 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   5970
         MouseIcon       =   "frmAMISMainMenu.frx":209B8
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":20B0A
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "1052"
         ToolTipText     =   "View Vendors Subsidiary Ledger"
         Top             =   2580
         Width           =   720
      End
      Begin VB.CommandButton cmdLedger_Customer 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   5970
         MouseIcon       =   "frmAMISMainMenu.frx":2143A
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":2158C
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "1050"
         ToolTipText     =   "View Customers A/R Ledger"
         Top             =   1740
         Width           =   720
      End
      Begin VB.CommandButton cmdLedger_Account 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   5970
         MouseIcon       =   "frmAMISMainMenu.frx":21E9F
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":21FF1
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "1049"
         ToolTipText     =   "View Accounts General Ledger"
         Top             =   900
         Width           =   720
      End
      Begin VB.CommandButton cmdJournal_CRJ 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   480
         MouseIcon       =   "frmAMISMainMenu.frx":228D3
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":22A25
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "1047"
         ToolTipText     =   "View Cash Receipts Journal"
         Top             =   3420
         Width           =   720
      End
      Begin VB.CommandButton cmdJournal_General 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   480
         MouseIcon       =   "frmAMISMainMenu.frx":23290
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":233E2
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "1048"
         ToolTipText     =   "View General Journal"
         Top             =   5175
         Width           =   720
      End
      Begin VB.CommandButton cmdJournal_Sales 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   480
         MouseIcon       =   "frmAMISMainMenu.frx":23CE7
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":23E39
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "1046"
         ToolTipText     =   "View Sales Journal"
         Top             =   2580
         Width           =   720
      End
      Begin VB.CommandButton cmdJournal_AP 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   480
         MouseIcon       =   "frmAMISMainMenu.frx":24617
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":24769
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "1044"
         ToolTipText     =   "View Accounts Payable Journal"
         Top             =   900
         Width           =   720
      End
      Begin XtremeSuiteControls.TabControl TabControl4 
         Height          =   5640
         Left            =   -69970
         TabIndex        =   30
         Top             =   570
         Visible         =   0   'False
         Width           =   10680
         _Version        =   655364
         _ExtentX        =   18838
         _ExtentY        =   9948
         _StockProps     =   64
         Appearance      =   2
         Color           =   4
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.LargeIcons=   -1  'True
         PaintManager.FixedTabWidth=   120
         PaintManager.MinTabWidth=   100
         ItemCount       =   3
         Item(0).Caption =   "Accounts"
         Item(0).Tooltip =   "Accounts"
         Item(0).ControlCount=   16
         Item(0).Control(0)=   "cmdAccount_AccountEntriesTemplate"
         Item(0).Control(1)=   "cmdAccount_DeaprtmentCodes"
         Item(0).Control(2)=   "cmdAccount_ExtendedClassification"
         Item(0).Control(3)=   "cmdAccount_AccountSubTotals"
         Item(0).Control(4)=   "cmdAccount_AccountClassification"
         Item(0).Control(5)=   "cmdAccount_AccountTypes"
         Item(0).Control(6)=   "cmdAccount_ChartOfAccounts"
         Item(0).Control(7)=   "Label58"
         Item(0).Control(8)=   "Label59"
         Item(0).Control(9)=   "Label60"
         Item(0).Control(10)=   "Label61"
         Item(0).Control(11)=   "Label62"
         Item(0).Control(12)=   "Label63"
         Item(0).Control(13)=   "Label65"
         Item(0).Control(14)=   "Label72"
         Item(0).Control(15)=   "cmdVehicleSales"
         Item(1).Caption =   "Opening Balances"
         Item(1).Tooltip =   "Opening Balances"
         Item(1).ControlCount=   10
         Item(1).Control(0)=   "Command3"
         Item(1).Control(1)=   "cmdOpeningBalance_Accounts"
         Item(1).Control(2)=   "cmdOpeningBalance_Customer"
         Item(1).Control(3)=   "cmdOpeningBalance_Vendor"
         Item(1).Control(4)=   "Label24"
         Item(1).Control(5)=   "Label37"
         Item(1).Control(6)=   "Label47"
         Item(1).Control(7)=   "Label48"
         Item(1).Control(8)=   "Label54"
         Item(1).Control(9)=   "cmdOpeningReport"
         Item(2).Caption =   "Adjustments"
         Item(2).Tooltip =   "Adjustments"
         Item(2).ControlCount=   15
         Item(2).Control(0)=   "cmdVendorCreditMemo"
         Item(2).Control(1)=   "cmdCustomerCreditMemo"
         Item(2).Control(2)=   "cmdVendorDebitMemo"
         Item(2).Control(3)=   "cmdAdjustment_ClosingEntries"
         Item(2).Control(4)=   "cmdCustomerDebitMemo"
         Item(2).Control(5)=   "cmdAdjustment_Client"
         Item(2).Control(6)=   "cmdAdjustment_Proposed"
         Item(2).Control(7)=   "Label41"
         Item(2).Control(8)=   "Label7"
         Item(2).Control(9)=   "Label30"
         Item(2).Control(10)=   "Label43"
         Item(2).Control(11)=   "Label44"
         Item(2).Control(12)=   "Label49"
         Item(2).Control(13)=   "Label51"
         Item(2).Control(14)=   "Picture1"
         Begin VB.CommandButton cmdVehicleSales 
            Height          =   645
            Left            =   4740
            MouseIcon       =   "frmAMISMainMenu.frx":24F9D
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":250EF
            Style           =   1  'Graphical
            TabIndex        =   148
            Tag             =   "1140"
            ToolTipText     =   "Vehicle Sales - Account Code Set-up"
            Top             =   2670
            Width           =   720
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   165
            Left            =   -69970
            ScaleHeight     =   165
            ScaleWidth      =   10095
            TabIndex        =   50
            Top             =   3750
            Visible         =   0   'False
            Width           =   10095
            Begin VB.Line Line 
               BorderColor     =   &H00808080&
               X1              =   0
               X2              =   10650
               Y1              =   60
               Y2              =   60
            End
         End
         Begin VB.CommandButton cmdAdjustment_Proposed 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69580
            MouseIcon       =   "frmAMISMainMenu.frx":25775
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":258C7
            Style           =   1  'Graphical
            TabIndex        =   49
            Tag             =   "1037"
            ToolTipText     =   "View Proposed Adjusting Journal Entries"
            Top             =   4755
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdAdjustment_Client 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69580
            MouseIcon       =   "frmAMISMainMenu.frx":25FAE
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":26100
            Style           =   1  'Graphical
            TabIndex        =   48
            Tag             =   "1036"
            ToolTipText     =   "View Client Adjusting Journal Entries"
            Top             =   3960
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdCustomerDebitMemo 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69580
            MouseIcon       =   "frmAMISMainMenu.frx":267E3
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":26935
            Style           =   1  'Graphical
            TabIndex        =   47
            Tag             =   "1038"
            ToolTipText     =   "View Customer Adjustments"
            Top             =   750
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdAdjustment_ClosingEntries 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -63850
            MouseIcon       =   "frmAMISMainMenu.frx":26FBA
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2710C
            Style           =   1  'Graphical
            TabIndex        =   46
            Tag             =   "1040"
            ToolTipText     =   "View Closing Entries"
            Top             =   690
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdVendorDebitMemo 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69580
            MouseIcon       =   "frmAMISMainMenu.frx":2774F
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":278A1
            Style           =   1  'Graphical
            TabIndex        =   45
            Tag             =   "1039"
            ToolTipText     =   "View Vendor Adjustments"
            Top             =   2295
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdCustomerCreditMemo 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69580
            MouseIcon       =   "frmAMISMainMenu.frx":27F0C
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2805E
            Style           =   1  'Graphical
            TabIndex        =   44
            Tag             =   "1038"
            ToolTipText     =   "View Customer Adjustments"
            Top             =   1530
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdVendorCreditMemo 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69580
            MouseIcon       =   "frmAMISMainMenu.frx":286E3
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":28835
            Style           =   1  'Graphical
            TabIndex        =   43
            Tag             =   "1039"
            ToolTipText     =   "View Vendor Adjustments"
            Top             =   3030
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdOpeningBalance_Vendor 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69490
            MouseIcon       =   "frmAMISMainMenu.frx":28EA0
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":28FF2
            Style           =   1  'Graphical
            TabIndex        =   42
            Tag             =   "1035"
            ToolTipText     =   "View Vendor Opening Balance"
            Top             =   2625
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdOpeningBalance_Customer 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69490
            MouseIcon       =   "frmAMISMainMenu.frx":297A4
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":298F6
            Style           =   1  'Graphical
            TabIndex        =   41
            Tag             =   "1034"
            ToolTipText     =   "View Customer Opening Balance"
            Top             =   1740
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdOpeningBalance_Accounts 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69490
            MouseIcon       =   "frmAMISMainMenu.frx":2A0D1
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2A223
            Style           =   1  'Graphical
            TabIndex        =   40
            Tag             =   "1033"
            ToolTipText     =   "View Accounts Opening Balance"
            Top             =   900
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdOpeningReport 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69490
            MouseIcon       =   "frmAMISMainMenu.frx":2A914
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2AA66
            Style           =   1  'Graphical
            TabIndex        =   39
            Tag             =   "1034"
            ToolTipText     =   "View Customer Opening Balance"
            Top             =   3465
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69460
            MouseIcon       =   "frmAMISMainMenu.frx":2B241
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2B393
            Style           =   1  'Graphical
            TabIndex        =   38
            Tag             =   "1045"
            ToolTipText     =   "View Cash Disbursement Journal"
            Top             =   4275
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdAccount_ChartOfAccounts 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   480
            MouseIcon       =   "frmAMISMainMenu.frx":2BDBE
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2BF10
            Style           =   1  'Graphical
            TabIndex        =   37
            Tag             =   "1020"
            ToolTipText     =   "View Chart Of Accounts"
            Top             =   900
            Width           =   720
         End
         Begin VB.CommandButton cmdAccount_AccountTypes 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   480
            MouseIcon       =   "frmAMISMainMenu.frx":2C544
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2C696
            Style           =   1  'Graphical
            TabIndex        =   36
            Tag             =   "1021"
            ToolTipText     =   "View Account Types"
            Top             =   1785
            Width           =   720
         End
         Begin VB.CommandButton cmdAccount_AccountClassification 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   480
            MouseIcon       =   "frmAMISMainMenu.frx":2CD15
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2CE67
            Style           =   1  'Graphical
            TabIndex        =   35
            Tag             =   "1022"
            ToolTipText     =   "View Account Classification"
            Top             =   2670
            Width           =   720
         End
         Begin VB.CommandButton cmdAccount_AccountSubTotals 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   480
            MouseIcon       =   "frmAMISMainMenu.frx":2D4D3
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2D625
            Style           =   1  'Graphical
            TabIndex        =   34
            Tag             =   "1024"
            ToolTipText     =   "View Account Sub-Totals"
            Top             =   4425
            Width           =   720
         End
         Begin VB.CommandButton cmdAccount_ExtendedClassification 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   480
            MouseIcon       =   "frmAMISMainMenu.frx":2DCD1
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2DE23
            Style           =   1  'Graphical
            TabIndex        =   33
            Tag             =   "1023"
            ToolTipText     =   "View Extended Classification"
            Top             =   3540
            Width           =   720
         End
         Begin VB.CommandButton cmdAccount_DeaprtmentCodes 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   4725
            MouseIcon       =   "frmAMISMainMenu.frx":2E5B3
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2E705
            Style           =   1  'Graphical
            TabIndex        =   32
            Tag             =   "1025"
            ToolTipText     =   "View Department Codes"
            Top             =   900
            Width           =   720
         End
         Begin VB.CommandButton cmdAccount_AccountEntriesTemplate 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   4725
            MouseIcon       =   "frmAMISMainMenu.frx":2EE66
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2EFB8
            Style           =   1  'Graphical
            TabIndex        =   31
            Tag             =   "1026"
            ToolTipText     =   "View Account Entries Templates"
            Top             =   1785
            Width           =   720
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Sales - Account Code Set-up"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5580
            TabIndex        =   149
            Top             =   2880
            Width           =   3480
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proposed Adjusting Journal Entries"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -68740
            TabIndex        =   69
            Top             =   4965
            Visible         =   0   'False
            Width           =   3330
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client Adjusting Journal Entries"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -68740
            TabIndex        =   68
            Top             =   4200
            Visible         =   0   'False
            Width           =   2985
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Debit Memo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -68710
            TabIndex        =   67
            Top             =   2460
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Closing Entries"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -62995
            TabIndex        =   66
            Top             =   885
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Debit Memo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -68710
            TabIndex        =   65
            Top             =   990
            Visible         =   0   'False
            Width           =   2100
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Credit Memo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -68725
            TabIndex        =   64
            Top             =   1680
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Credit Memo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -68710
            TabIndex        =   63
            Top             =   3195
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Opening Balance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -68590
            TabIndex        =   62
            Top             =   2760
            Visible         =   0   'False
            Width           =   2400
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Opening Balance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -68605
            TabIndex        =   61
            Top             =   1905
            Visible         =   0   'False
            Width           =   2610
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Accounts Opening Balance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -68590
            TabIndex        =   60
            Top             =   1140
            Visible         =   0   'False
            Width           =   2565
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Opening Balance Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -68590
            TabIndex        =   59
            Top             =   3615
            Visible         =   0   'False
            Width           =   2340
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Opening Balance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -68560
            TabIndex        =   58
            Top             =   4425
            Visible         =   0   'False
            Width           =   2190
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Entries Templates"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5595
            TabIndex        =   57
            Top             =   1905
            Width           =   2520
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Department Codes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5595
            TabIndex        =   56
            Top             =   1035
            Width           =   1755
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Extended Classification"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1350
            TabIndex        =   55
            Top             =   3675
            Width           =   2220
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Classification"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1350
            TabIndex        =   54
            Top             =   2805
            Width           =   2085
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Sub-Totals"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1350
            TabIndex        =   53
            Top             =   4560
            Width           =   1830
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Types"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1350
            TabIndex        =   52
            Top             =   1950
            Width           =   1380
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chart of Accounts"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1320
            TabIndex        =   51
            Top             =   1155
            Width           =   1665
         End
      End
      Begin XtremeSuiteControls.TabControl TabControl2 
         Height          =   5730
         Left            =   -69970
         TabIndex        =   70
         Top             =   570
         Visible         =   0   'False
         Width           =   11490
         _Version        =   655364
         _ExtentX        =   20267
         _ExtentY        =   10107
         _StockProps     =   64
         Appearance      =   2
         Color           =   4
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.LargeIcons=   -1  'True
         PaintManager.FixedTabWidth=   120
         PaintManager.MinTabWidth=   150
         ItemCount       =   2
         Item(0).Caption =   "Journals"
         Item(0).ControlCount=   2
         Item(0).Control(0)=   "TabControl3"
         Item(0).Control(1)=   "Picture15"
         Item(1).Caption =   "Financial Statement"
         Item(1).ControlCount=   10
         Item(1).Control(0)=   "cmdTrialBalance"
         Item(1).Control(1)=   "cmdWork_Sheet"
         Item(1).Control(2)=   "cmdScheduleOfAdjustments"
         Item(1).Control(3)=   "cmdFinancialStatements"
         Item(1).Control(4)=   "cmdWorkSheet"
         Item(1).Control(5)=   "Label10"
         Item(1).Control(6)=   "Label16"
         Item(1).Control(7)=   "Label17"
         Item(1).Control(8)=   "Label50"
         Item(1).Control(9)=   "cmdWitholdingTax"
         Begin VB.CommandButton cmdWitholdingTax 
            Height          =   645
            Left            =   -69580
            MouseIcon       =   "frmAMISMainMenu.frx":2F673
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2F7C5
            Style           =   1  'Graphical
            TabIndex        =   141
            ToolTipText     =   "View Financial Statements"
            Top             =   4140
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdFinancialStatements 
            Height          =   645
            Left            =   -69610
            MouseIcon       =   "frmAMISMainMenu.frx":2FC0F
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2FD61
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "View Financial Statements"
            Top             =   3360
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdScheduleOfAdjustments 
            Height          =   645
            Left            =   -69595
            MouseIcon       =   "frmAMISMainMenu.frx":304B7
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":30609
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "View Schedule Of Adjustments"
            Top             =   2565
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdWork_Sheet 
            Height          =   645
            Left            =   -69595
            MouseIcon       =   "frmAMISMainMenu.frx":30DC2
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":30F14
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "View Work Sheet"
            Top             =   900
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdTrialBalance 
            Height          =   645
            Left            =   -69595
            MouseIcon       =   "frmAMISMainMenu.frx":31717
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":31869
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "View Trial Balance"
            Top             =   1725
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.PictureBox Picture15 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5840
            Left            =   -70000
            ScaleHeight     =   5835
            ScaleWidth      =   10800
            TabIndex        =   71
            Top             =   600
            Width           =   10800
         End
         Begin XtremeSuiteControls.TabControl TabControl3 
            Height          =   5175
            Left            =   30
            TabIndex        =   76
            Top             =   570
            Width           =   10800
            _Version        =   655364
            _ExtentX        =   19050
            _ExtentY        =   9128
            _StockProps     =   64
            Appearance      =   1
            Color           =   4
            PaintManager.BoldSelected=   -1  'True
            PaintManager.HotTracking=   -1  'True
            PaintManager.ShowIcons=   -1  'True
            PaintManager.FixedTabWidth=   110
            PaintManager.MinTabWidth=   85
            ItemCount       =   5
            Item(0).Caption =   "Accounts Payable"
            Item(0).ControlCount=   10
            Item(0).Control(0)=   "cmdAccountDetailbySupplier"
            Item(0).Control(1)=   "cmdAccountsPayableAgingReport"
            Item(0).Control(2)=   "cmdLedgerCodeRunningBalance_AP"
            Item(0).Control(3)=   "cmdAccountsPayableJournal"
            Item(0).Control(4)=   "cmdReceivingReportRegister"
            Item(0).Control(5)=   "Label6"
            Item(0).Control(6)=   "Label18"
            Item(0).Control(7)=   "Label19"
            Item(0).Control(8)=   "Label21"
            Item(0).Control(9)=   "Label23"
            Item(1).Caption =   "Cash Disbursement"
            Item(1).ControlCount=   6
            Item(1).Control(0)=   "cmdCheckRegister"
            Item(1).Control(1)=   "cmdLedgerCodeRunningBalance_CD"
            Item(1).Control(2)=   "cmdCashDisbursementJournal"
            Item(1).Control(3)=   "Label25"
            Item(1).Control(4)=   "Label26"
            Item(1).Control(5)=   "Label27"
            Item(2).Caption =   "Sales"
            Item(2).ControlCount=   14
            Item(2).Control(0)=   "cmdScheduleOfAccountsReceivable"
            Item(2).Control(1)=   "cmdLedgerCodeRunningBalance_LCRB"
            Item(2).Control(2)=   "cmdAccountDetailbyCustomer"
            Item(2).Control(3)=   "cmdSalesJournal"
            Item(2).Control(4)=   "cmdInvoiceRegister"
            Item(2).Control(5)=   "cmdSalesbyInvoiceType"
            Item(2).Control(6)=   "cmdUnusedInvoices"
            Item(2).Control(7)=   "Label29"
            Item(2).Control(8)=   "Label31"
            Item(2).Control(9)=   "Label32"
            Item(2).Control(10)=   "Label34"
            Item(2).Control(11)=   "Label35"
            Item(2).Control(12)=   "Label36"
            Item(2).Control(13)=   "Label38"
            Item(3).Caption =   "Cash Receipts"
            Item(3).ControlCount=   8
            Item(3).Control(0)=   "cmdUnused_OR"
            Item(3).Control(1)=   "cmdOR_Register"
            Item(3).Control(2)=   "cmdLedgerCodeRunningBalance_CR"
            Item(3).Control(3)=   "cmdCashReceiptsJournal"
            Item(3).Control(4)=   "Label39"
            Item(3).Control(5)=   "Label40"
            Item(3).Control(6)=   "Label46"
            Item(3).Control(7)=   "Label67"
            Item(4).Caption =   "General Journal"
            Item(4).ControlCount=   4
            Item(4).Control(0)=   "cmdJournalVoucherSummary"
            Item(4).Control(1)=   "cmdGJLedgerCodeRunningBalance"
            Item(4).Control(2)=   "Label13"
            Item(4).Control(3)=   "Label12"
            Begin VB.CommandButton cmdGJLedgerCodeRunningBalance 
               Height          =   645
               Left            =   -69580
               MouseIcon       =   "frmAMISMainMenu.frx":31DF7
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":31F49
               Style           =   1  'Graphical
               TabIndex        =   97
               ToolTipText     =   "View Ledger Code Running Balance"
               Top             =   1545
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdJournalVoucherSummary 
               Height          =   645
               Left            =   -69580
               MouseIcon       =   "frmAMISMainMenu.frx":32602
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":32754
               Style           =   1  'Graphical
               TabIndex        =   96
               ToolTipText     =   "View General Journal"
               Top             =   720
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdCashReceiptsJournal 
               Height          =   645
               Left            =   -69580
               MouseIcon       =   "frmAMISMainMenu.frx":32ED2
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":33024
               Style           =   1  'Graphical
               TabIndex        =   95
               ToolTipText     =   "View Cash Receipts Journal"
               Top             =   630
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdLedgerCodeRunningBalance_CR 
               Height          =   645
               Left            =   -69595
               MouseIcon       =   "frmAMISMainMenu.frx":337A2
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":338F4
               Style           =   1  'Graphical
               TabIndex        =   94
               ToolTipText     =   "View Ledger Code Running Balance"
               Top             =   1500
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdOR_Register 
               Height          =   645
               Left            =   -69595
               MouseIcon       =   "frmAMISMainMenu.frx":33FAD
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":340FF
               Style           =   1  'Graphical
               TabIndex        =   93
               ToolTipText     =   "View O.R. Register"
               Top             =   2340
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdUnused_OR 
               Height          =   645
               Left            =   -69610
               MouseIcon       =   "frmAMISMainMenu.frx":3487D
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":349CF
               Style           =   1  'Graphical
               TabIndex        =   92
               ToolTipText     =   "View Unused O.R."
               Top             =   3180
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdUnusedInvoices 
               Height          =   645
               Left            =   -69610
               MouseIcon       =   "frmAMISMainMenu.frx":3519F
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":352F1
               Style           =   1  'Graphical
               TabIndex        =   91
               ToolTipText     =   "View Unused Invoices"
               Top             =   3960
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdSalesbyInvoiceType 
               Height          =   645
               Left            =   -64435
               MouseIcon       =   "frmAMISMainMenu.frx":35A81
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":35BD3
               Style           =   1  'Graphical
               TabIndex        =   90
               ToolTipText     =   "View Sales by Invoice Type"
               Top             =   1485
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdInvoiceRegister 
               Height          =   645
               Left            =   -64420
               MouseIcon       =   "frmAMISMainMenu.frx":36312
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":36464
               Style           =   1  'Graphical
               TabIndex        =   89
               ToolTipText     =   "View Invoice Register"
               Top             =   630
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdSalesJournal 
               Height          =   645
               Left            =   -69580
               MouseIcon       =   "frmAMISMainMenu.frx":36AFE
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":36C50
               Style           =   1  'Graphical
               TabIndex        =   88
               ToolTipText     =   "View Sales Journal"
               Top             =   630
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdAccountDetailbyCustomer 
               Height          =   645
               Left            =   -69580
               MouseIcon       =   "frmAMISMainMenu.frx":373B7
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":37509
               Style           =   1  'Graphical
               TabIndex        =   87
               ToolTipText     =   "View Account Detail by Customer"
               Top             =   1470
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdLedgerCodeRunningBalance_LCRB 
               Height          =   645
               Left            =   -69580
               MouseIcon       =   "frmAMISMainMenu.frx":37BB1
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":37D03
               Style           =   1  'Graphical
               TabIndex        =   86
               ToolTipText     =   "View Ledger Code Running Balance"
               Top             =   2310
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdScheduleOfAccountsReceivable 
               Height          =   645
               Left            =   -69610
               MouseIcon       =   "frmAMISMainMenu.frx":383BC
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3850E
               Style           =   1  'Graphical
               TabIndex        =   85
               ToolTipText     =   "View Schedule Of Accounts Receivable"
               Top             =   3120
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdCashDisbursementJournal 
               Height          =   645
               Left            =   -69580
               MouseIcon       =   "frmAMISMainMenu.frx":38C46
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":38D98
               Style           =   1  'Graphical
               TabIndex        =   84
               ToolTipText     =   "View Cash Disbursement Journal"
               Top             =   690
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdLedgerCodeRunningBalance_CD 
               Height          =   645
               Left            =   -69580
               MouseIcon       =   "frmAMISMainMenu.frx":3956E
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":396C0
               Style           =   1  'Graphical
               TabIndex        =   83
               ToolTipText     =   "View Ledger Code Running Balance"
               Top             =   1530
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdCheckRegister 
               Height          =   645
               Left            =   -69580
               MouseIcon       =   "frmAMISMainMenu.frx":39D79
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":39ECB
               Style           =   1  'Graphical
               TabIndex        =   82
               ToolTipText     =   "View Check Register"
               Top             =   2385
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdReceivingReportRegister 
               Height          =   645
               Left            =   5385
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3A628
               Style           =   1  'Graphical
               TabIndex        =   81
               ToolTipText     =   "View Receiving Report Register"
               Top             =   675
               Width           =   720
            End
            Begin VB.CommandButton cmdAccountsPayableJournal 
               Height          =   645
               Left            =   420
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3ACA9
               Style           =   1  'Graphical
               TabIndex        =   80
               ToolTipText     =   "View Accounts Payable Journal"
               Top             =   660
               Width           =   720
            End
            Begin VB.CommandButton cmdLedgerCodeRunningBalance_AP 
               Height          =   645
               Left            =   435
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3B367
               Style           =   1  'Graphical
               TabIndex        =   79
               ToolTipText     =   "View Ledger Code Running Balance"
               Top             =   1500
               Width           =   720
            End
            Begin VB.CommandButton cmdAccountsPayableAgingReport 
               Height          =   645
               Left            =   420
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3BA20
               Style           =   1  'Graphical
               TabIndex        =   78
               ToolTipText     =   "View Accounts Payable Aging Report"
               Top             =   3120
               Width           =   720
            End
            Begin VB.CommandButton cmdAccountDetailbySupplier 
               Height          =   645
               Left            =   435
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3C14D
               Style           =   1  'Graphical
               TabIndex        =   77
               ToolTipText     =   "View Account Detail by Supplier"
               Top             =   2340
               Width           =   720
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Journal Voucher"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   -68710
               TabIndex        =   118
               Top             =   915
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ledger Code Running Balance"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   -68710
               TabIndex        =   117
               Top             =   1680
               Visible         =   0   'False
               Width           =   2925
            End
            Begin VB.Label Label67 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unused O.R."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   -68725
               TabIndex        =   116
               Top             =   3375
               Visible         =   0   'False
               Width           =   1170
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "O.R."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   -68725
               TabIndex        =   115
               Top             =   2490
               Visible         =   0   'False
               Width           =   1245
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ledger Code Running Balance"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   -68725
               TabIndex        =   114
               Top             =   1665
               Visible         =   0   'False
               Width           =   2925
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cash Receipts Journal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   -68725
               TabIndex        =   113
               Top             =   870
               Visible         =   0   'False
               Width           =   2115
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unused Invoices"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   -68710
               TabIndex        =   112
               Top             =   4170
               Visible         =   0   'False
               Width           =   1545
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Invoice Register"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   -63535
               TabIndex        =   111
               Top             =   765
               Visible         =   0   'False
               Width           =   1530
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales by Invoice Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   -63565
               TabIndex        =   110
               Top             =   1620
               Visible         =   0   'False
               Width           =   2085
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "A/R Aging and Schedule Reports"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   -68710
               TabIndex        =   109
               Top             =   3300
               Visible         =   0   'False
               Width           =   3105
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ledger Code Running Balance"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   -68710
               TabIndex        =   108
               Top             =   2550
               Visible         =   0   'False
               Width           =   2925
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Account Detail by Customer"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   -68710
               TabIndex        =   107
               Top             =   1725
               Visible         =   0   'False
               Width           =   2625
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Journal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   -68710
               TabIndex        =   106
               Top             =   870
               Visible         =   0   'False
               Width           =   1305
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Check Register"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   -68710
               TabIndex        =   105
               Top             =   2580
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ledger Code Running Balance"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   -68710
               TabIndex        =   104
               Top             =   1740
               Visible         =   0   'False
               Width           =   2925
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cash Disbursement Journal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   -68710
               TabIndex        =   103
               Top             =   900
               Visible         =   0   'False
               Width           =   2595
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Receiving Report Register"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   6270
               TabIndex        =   102
               Top             =   825
               Width           =   2475
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "A/P Aging and Schedule Reports"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   1305
               TabIndex        =   101
               Top             =   3315
               Width           =   3105
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Account Detail by Supplier"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   1305
               TabIndex        =   100
               Top             =   2550
               Width           =   2535
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ledger Code Running Balance"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   1305
               TabIndex        =   99
               Top             =   1680
               Width           =   2925
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Accounts Payable Journal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   1305
               TabIndex        =   98
               Top             =   870
               Width           =   2475
            End
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Creditable Withholding Tax"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -68740
            TabIndex        =   142
            Top             =   4290
            Visible         =   0   'False
            Width           =   2625
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Financial Statements"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -68740
            TabIndex        =   122
            Top             =   3510
            Visible         =   0   'False
            Width           =   2010
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule Of Adjustments"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -68740
            TabIndex        =   121
            Top             =   2715
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trial Balance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -68740
            TabIndex        =   120
            Top             =   1905
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label cmdWorkSheet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Work Sheet"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   -68740
            TabIndex        =   119
            Top             =   1080
            Visible         =   0   'False
            Width           =   1110
         End
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Un-Applied Payment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -68650
         TabIndex        =   140
         Top             =   4230
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Un-Imported Reports"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -68650
         TabIndex        =   139
         Top             =   4950
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Tools"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -63310
         TabIndex        =   138
         Top             =   2670
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Re Printing Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -63310
         TabIndex        =   137
         Top             =   1125
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelled Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -63310
         TabIndex        =   136
         Top             =   1860
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Audit Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -68650
         TabIndex        =   135
         Top             =   3390
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Audit Inquiry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -68650
         TabIndex        =   134
         Top             =   2610
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Signatories and Headers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -68650
         TabIndex        =   133
         Top             =   1860
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reminders"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -68635
         TabIndex        =   132
         Top             =   1125
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Master File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -68650
         TabIndex        =   29
         Top             =   1935
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ATC Code Master File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -68650
         TabIndex        =   28
         Top             =   5130
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terms of Payment Master File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -68650
         TabIndex        =   27
         Top             =   4335
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Master File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -68650
         TabIndex        =   26
         Top             =   2730
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Type Master File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -68650
         TabIndex        =   25
         Top             =   3525
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Master File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -68650
         TabIndex        =   24
         Top             =   1125
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Receipts Journal (Deposited OR's)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1305
         TabIndex        =   16
         Top             =   4530
         Width           =   3720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accounts General Ledger"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6810
         TabIndex        =   15
         Top             =   1140
         Width           =   2415
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customers A/R Ledger"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6810
         TabIndex        =   14
         Top             =   2025
         Width           =   2115
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendors A/P Ledger"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6810
         TabIndex        =   13
         Top             =   2775
         Width           =   1905
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "General Journal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1305
         TabIndex        =   12
         Top             =   5415
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Journal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1305
         TabIndex        =   11
         Top             =   2805
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Receipts Journal (Un-Deposited OR's)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1305
         TabIndex        =   10
         Top             =   3630
         Width           =   4035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Disbursement Journal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1305
         TabIndex        =   9
         Top             =   1935
         Width           =   2595
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accounts Payable Journal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1305
         TabIndex        =   8
         Top             =   1095
         Width           =   2475
      End
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frmNEWAMISRangeWithSummary As frmAMISRangeWithSummary

Private Sub cmdAccount_AccountEntriesTemplate_Click()
    If Module_Access(LOGID, "ACCOUNT ENTRIES TEMPLATES", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAMISMASTERFILESTemplates
End Sub

Private Sub cmdAdjustment_Client_Click()
    If Module_Access(LOGID, "CLIENT ADJUSTING JOURNAL ENTRIES", "TRANSACTION") = False Then Exit Sub
    'JOURNALTYPE = "ADJ"
    On Error Resume Next
    'Unload frmAMISJournalEntry
    Call frmAMISJournalEntry.LoadJournal("ADJ")
    FormExistsShow frmAMISJournalEntry
End Sub

Private Sub cmdAdjustment_ClosingEntries_Click()
    If Module_Access(LOGID, "CLOSING ENTRIES", "TRANSACTION") = False Then Exit Sub
    'JOURNALTYPE = "CLO"
    On Error Resume Next
    'Unload frmAMISJournalEntry
    Call frmAMISJournalEntry.LoadJournal("CLO")
    FormExistsShow frmAMISJournalEntry
End Sub

Private Sub cmdCustom_Click()
    If COMPANY_CODE = "HGC" Then

    ElseIf COMPANY_CODE = "HAI" Then

    ElseIf COMPANY_CODE = "HAS" Then

    ElseIf COMPANY_CODE = "HBK" Then

    ElseIf COMPANY_CODE = "HHM" Then

    ElseIf COMPANY_CODE = "HSB" Then

    End If
End Sub

Private Sub cmdCustomerDebitMemo_Click()
    If Module_Access(LOGID, "CUSTOMER ADJUSTMENT", "DATA ENTRY") = False Then Exit Sub
    'JOURNALTYPE = "CSJ"
    On Error Resume Next
    'Unload frmAMISJournalEntry
    Call frmAMISJournalEntry.LoadJournal("CSJ")
    FormExistsShow frmAMISJournalEntry
    'frmAMISCustomerAdjustment.Show
End Sub

Private Sub cmdAdjustment_Proposed_Click()
    If Module_Access(LOGID, "PROPOSED ADJUSTING JOURNAL ENTRIES", "TRANSACTION") = False Then Exit Sub
    'JOURNALTYPE = "PDJ"
    On Error Resume Next
    Call frmAMISJournalEntry.LoadJournal("PDJ")
    FormExistsShow frmAMISJournalEntry
End Sub

Private Sub cmdTranType_Click()
    frmTranTypeImportingSetup.Show
End Sub

Private Sub cmdVehicleSales_Click()
    If Module_Access(LOGID, "VEHICLE MODEL SET-UP", "SYSTEM") = False Then Exit Sub
    FormExistsShow frmVehicleSalesCodeSetup
End Sub

Private Sub cmdVendorCreditMemo_Click()
    If Module_Access(LOGID, "VENDOR ADJUSTMENT", "DATA ENTRY") = False Then Exit Sub
    'JOURNALTYPE = "VCJ"
    On Error Resume Next
    Call frmAMISJournalEntry.LoadJournal("VCJ")
    FormExistsShow frmAMISJournalEntry
    'frmAMISVendorAdjustment.Show
End Sub

Private Sub cmdVendorDebitMemo_Click()
    If Module_Access(LOGID, "VENDOR ADJUSTMENT", "DATA ENTRY") = False Then Exit Sub
    'JOURNALTYPE = "VDJ"
    On Error Resume Next
    Call frmAMISJournalEntry.LoadJournal("VDJ")
    FormExistsShow frmAMISJournalEntry
    'frmAMISVendorAdjustment.Show
End Sub

Private Sub cmdAuditInquiry_Click()
    FormExistsShow frmInquiry_Audit
End Sub

Private Sub cmdAuditReport_Click()
    FormExistsShow frmReportAuditReport
End Sub

Private Sub cmdGJLedgerCodeRunningBalance_Click()
    If Module_Access(LOGID, "GENERAL JOURNAL LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "GJ"
    Call frmAMISRangeWithAccountCode.LoadJournal("GJ")
    FormExistsShow frmAMISRangeWithAccountCode
    frmAMISRangeWithAccountCode.Caption = "Journal Voucher Ledger Code Running Balance"
End Sub

Private Sub cmdJournal_AP_Click()
    If Module_Access(LOGID, "ACCOUNTS PAYABLE JOURNAL", "TRANSACTION") = False Then Exit Sub
    'JOURNALTYPE = "APJ"
    On Error Resume Next
    '    Unload frmAMISJournalEntry
    '    frmAMISJournalEntry.Show
    Call frmAMISJournalEntry_APJ.LoadJournal("APJ")
    FormExistsShow frmAMISJournalEntry_APJ
End Sub

Private Sub cmdJournal_CashDisburshment_Click()
    If Module_Access(LOGID, "CASH DISBURSEMENT JOURNAL", "TRANSACTION") = False Then Exit Sub
    'JOURNALTYPE = "CDJ"
    On Error Resume Next
    '    Unload frmAMISJournalEntry
    '    frmAMISJournalEntry.Show
    Call frmAMISJournalEntry_CDJ.LoadJournal("CDJ")
    FormExistsShow frmAMISJournalEntry_CDJ
End Sub

Private Sub cmdJournal_CRJ_Click()
    If Module_Access(LOGID, "CASH RECEIPTS JOURNAL", "TRANSACTION") = False Then Exit Sub
    'JOURNALTYPE = "CRJ"
    On Error Resume Next
    'Unload frmAMISJournalEntry
    'frmAMISJournalEntry.Show
    Call frmAMISJournalEntry_CRJ.LoadJournal("CRJ")
    FormExistsShow frmAMISJournalEntry_CRJ
End Sub

Private Sub cmdJournal_General_Click()
    On Error Resume Next
    'If Module_Access(LOGID, "GENERAL JOURNAL", "TRANSACTION") = False Then Exit Sub
    '    JOURNALTYPE = "GJ"
    '    On Error Resume Next
    '    Unload frmAMISJournalEntry
    '    frmAMISJournalEntry.Show
    If Module_Access(LOGID, "GENERAL JOURNAL", "TRANSACTION") = False Then Exit Sub

    On Error Resume Next
    'Unload frmAMISJournalEntry

    Call frmAMISJournalEntry_GJ.LoadJournal("GJ")
    FormExistsShow frmAMISJournalEntry_GJ
End Sub

Private Sub cmdJournal_Sales_Click()
    If Module_Access(LOGID, "SALES JOURNAL", "TRANSACTION") = False Then Exit Sub
    'JOURNALTYPE = "SJ"
    On Error Resume Next
    '    Unload frmAMISJournalEntry
    '    frmAMISJournalEntry.Show
    Call frmAMISJournalEntry_SJ.LoadJournal("SJ")
    FormExistsShow frmAMISJournalEntry_SJ
End Sub

Private Sub cmdJournalVoucherSummary_Click()
    If Module_Access(LOGID, "GENERAL JOURNAL SUMMARY", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "GJ"
    'FormExistsShow frmAMISRange
    FormExistsShow frmAMISRangeWithSummary
    frmAMISRangeWithSummary.Caption = "General Journal"
End Sub

Private Sub cmdLedger_Account_Click()
    If Module_Access(LOGID, "ACCOUNT GENERAL LEDGER", "INQUIRY") = False Then Exit Sub
    FormExistsShow frmAMISLEDGERAccounts
End Sub

Private Sub cmdLedger_Customer_Click()
    If Module_Access(LOGID, "CUSTOMER A/R LEDGER", "INQUIRY") = False Then Exit Sub
    'CUST_LEDGER_TYPE = "ARLEDGER"
    'frmAMISLEDGERCustomers.Show
    FormExistsShow frmAMIS_ARLEDGER
End Sub

Private Sub cmdLedger_VendorSubisdy_Click()
    If Module_Access(LOGID, "VENDOR SUBSIDIARY LEDGER", "INQUIRY") = False Then Exit Sub
'    If COMPANY_CODE <> "HPI" Then
        FormExistsShow frmAMIS_APLEDGER
'    Else
        FormExistsShow frmAMISLEDGERVendors
'    End If
End Sub

Private Sub cmdOpeningBalance_Accounts_Click()
    If Module_Access(LOGID, "ACCOUNT OPENING BALANCE", "TRANSACTION") = False Then Exit Sub
    'JOURNALTYPE = "OPB"
    On Error Resume Next
    Call frmAMISJournalEntry_OPB.LoadJournal("OPB")
    FormExistsShow frmAMISJournalEntry_OPB
End Sub

Private Sub cmdJournalDRJ_Click()
    If Module_Access(LOGID, "DEPOSITED RECEIPTS JOURNAL", "TRANSACTION") = False Then Exit Sub
    'JOURNALTYPE = "DRJ"
    On Error Resume Next
    Call frmAMISJournalEntry_DRJ.LoadJournal("DRJ")
    FormExistsShow frmAMISJournalEntry_DRJ
End Sub

Private Sub cmdCashDisbursementJournal_Click()
    If Module_Access(LOGID, "CASH DISBURSEMENT JOURNAL", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "CDJ"
    FormExistsShow frmAMISRangeWithSummary
    frmAMISRangeWithSummary.Caption = "Cash Disbursement Journal"
End Sub

Private Sub cmdCashReceiptsJournal_Click()
    If Module_Access(LOGID, "CASH RECEIPTS JOURNAL", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "CRJ"
    FormExistsShow frmAMISRangeWithSummary
    frmAMISRangeWithSummary.Caption = "Cash Receipts Journal"
End Sub

Private Sub cmdCheckRegister_Click()
    If Module_Access(LOGID, "REGISTER REPORT", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "CHECK_REGISTER"
    FormExistsShow frmAMISRange
    frmAMISRange.Caption = "Check Registers"
    DoEvents
End Sub

Private Sub cmdOpeningBalance_Customer_Click()
'AXP-07082007-000001
    If Module_Access(LOGID, "CUSTOMER OPENING BALANCE", "DATA ENTRY") = False Then Exit Sub
    On Error Resume Next
    'JOURNALTYPE = "COB"
    Call frmAMISCustomerAROpening.LoadJournal("COB")
    FormExistsShow frmAMISCustomerAROpening
End Sub

Private Sub cmdAccount_AccountClassification_Click()
    If Module_Access(LOGID, "ACCOUNT CLASSIFICATION", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAMISFILESHeader
End Sub

Private Sub cmdAccount_AccountSubTotals_Click()
    If Module_Access(LOGID, "ACCOUNT SUB TOTALS", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAMISFILESTitleCode
End Sub

Private Sub cmdAccount_DeaprtmentCodes_Click()
    If Module_Access(LOGID, "DEPARTMENT CODES", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAMISFILESDepartment
End Sub

Private Sub cmdAccount_ExtendedClassification_Click()
    If Module_Access(LOGID, "EXTENDED CLASSIFICATION", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAMISFILESSubHeader
End Sub

Private Sub cmdAccount_ChartOfAccounts_Click()
    If Module_Access(LOGID, "CHART OF ACCOUNTS", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAMISFILESChartOfAccount
End Sub

Private Sub cmdAccount_AccountTypes_Click()
    If Module_Access(LOGID, "ACCOUNT TYPES", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAMISFILESAccType
End Sub

Private Sub cmdFinancialStatements_Click()
    If Module_Access(LOGID, "FINANCIAL STATEMENTS", "REPORTS") = False Then Exit Sub
    FormExistsShow frmAMISFinancialStatements
End Sub

Private Sub cmdInvoiceRegister_Click()
    If Module_Access(LOGID, "REGISTER REPORT", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "INV_REGISTER"
    FormExistsShow frmAMISRange
    frmAMISRange.Caption = "Invoices Registers"
    DoEvents
End Sub

Private Sub cmdLedgerCodeRunningBalance_AP_Click()
    If Module_Access(LOGID, "ACCOUNTS LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "APJ"
    Call frmAMISRangeWithAccountCode.LoadJournal("APJ")
    FormExistsShow frmAMISRangeWithAccountCode
    frmAMISRangeWithAccountCode.Caption = "ACCOUNTS Disbursement Ledger Code Running Balance"
End Sub

Private Sub cmdLedgerCodeRunningBalance_CD_Click()
    If Module_Access(LOGID, "CASH DISBURSEMENT LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "CDJ"
    Call frmAMISRangeWithAccountCode.LoadJournal("CDJ")
    FormExistsShow frmAMISRangeWithAccountCode
    frmAMISRangeWithAccountCode.Caption = "Cash Disbursement Ledger Code Running Balance"
End Sub

Private Sub cmdLedgerCodeRunningBalance_CR_Click()
    If Module_Access(LOGID, "CASH RECEIPTS LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "CRJ"
    Call frmAMISRangeWithAccountCode.LoadJournal("CRJ")
    FormExistsShow frmAMISRangeWithAccountCode
    frmAMISRangeWithAccountCode.Caption = "Cash Receipts Ledger Code Running Balance"
End Sub

Private Sub cmdLedgerCodeRunningBalance_LCRB_Click()
    If Module_Access(LOGID, "SALES LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "SJ"
    Call frmAMISRangeWithAccountCode.LoadJournal("SJ")
    FormExistsShow frmAMISRangeWithAccountCode
    frmAMISRangeWithAccountCode.Caption = "Sales Journal Ledger Code Running Balance"
End Sub

Private Sub cmdOR_Register_Click()
    If Module_Access(LOGID, "REGISTER REPORT", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "OR_REGISTER"
    FormExistsShow frmAMISRange
    frmAMISRange.Caption = "O.R. Registers"
    DoEvents
End Sub

Private Sub cmdReceivingReportRegister_Click()
    If Module_Access(LOGID, "RECEIVING REPORT REGISTER", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "REC_REGISTER"
    FormExistsShow frmAMISDetailBySupplierWithAccountCode
    frmAMISDetailBySupplierWithAccountCode.Caption = "Receiving Report Registers"
End Sub

Private Sub cmdSalesbyInvoiceType_Click()
    If Module_Access(LOGID, "SALES BY INVOICE TYPE", "REPORTS") = False Then Exit Sub
    FormExistsShow frmAMIS_SalesbyInvoiceType
End Sub

Private Sub cmdSalesJournal_Click()
    If Module_Access(LOGID, "SALES JOURNAL", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "SJ"
    FormExistsShow frmAMISRangeWithSummary
    frmAMISRangeWithSummary.Caption = "Sales Journal"
End Sub

Private Sub cmdScheduleOfAccountsPayable_Click()
'COMMENTED BY: ACL
'    If Module_Access(LOGID, "SCHEDULE OF ACCOUNTS PAYABLE", "REPORTS") = False Then Exit Sub
'    frmAMISAPSchedReport.Show
End Sub

Private Sub cmdScheduleOfAccountsReceivable_Click()
    If Module_Access(LOGID, "ACCOUNTS RECEIVABLE AGING REPORT", "REPORTS") = False Then Exit Sub
    Report_AR = "AGING"
    '    If COMPANY_CODE = "HGC" Then
    '        frmAMISARSchedReport.Show
    '    Else
    FormExistsShow frmNEW_ARSchedReport
    '    End If

End Sub

Private Sub cmdScheduleOfAdjustments_Click()
    If Module_Access(LOGID, "SCHEDULE OF ADJUSTMENTS", "REPORTS") = False Then Exit Sub
    FormExistsShow frmAMISSchedAdjust
End Sub

Private Sub cmdTable_ATCCode_Click()
    If Module_Access(LOGID, "ATC CODES", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAMISMASTERFILEATC
End Sub

Private Sub cmdTable_Customer_Click()
    If Module_Access(LOGID, "CUSTOMER", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAllCustomer
End Sub

Private Sub cmdTables_Bank_Click()
    If Module_Access(LOGID, "BANKS", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAMISMASTERFILEBanks
End Sub

Private Sub cmdTables_InvoiceTypes_Click()
    If Module_Access(LOGID, "INVOICE TYPES", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAMISMASTERFILEInvoiceType
End Sub

Private Sub cmdTables_TermsOfPayment_Click()
    If Module_Access(LOGID, "TERMS OF PAYMENT", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAMISMASTERFILEPayTerm
End Sub

Private Sub cmdTrialBalance_Click()
    If Module_Access(LOGID, "FINANCIAL STATMENT TRIAL BALANCE", "REPORTS") = False Then Exit Sub
    FormExistsShow frmAMISTrialBalance
End Sub

Private Sub cmdUnused_OR_Click()
    If Module_Access(LOGID, "UNUSED OR", "REPORTS") = False Then Exit Sub
    FormExistsShow frmAMISProcessUnusedOR
End Sub

Private Sub cmdUnusedInvoices_Click()
    If Module_Access(LOGID, "UNUSED INVOICES", "REPORTS") = False Then Exit Sub
    FormExistsShow frmAMISProcessUnusedInvoices
End Sub

Private Sub cmdVendor_Click()
    If Module_Access(LOGID, "VENDORS", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAMISMASTERFILEVendor
End Sub

Private Sub cmdOpeningBalance_Vendor_Click()
    If Module_Access(LOGID, "VENDOR OPENING BALANCE", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    'JOURNALTYPE = "VPJ"
    Call frmAMISVendorAPOpening.LoadJournal("VPJ")
    FormExistsShow frmAMISVendorAPOpening
End Sub

Private Sub cmdWitholdingTax_Click()
    If Module_Access(LOGID, "CREDITABLE WITHHOLDING TAX", "REPORTS") = False Then Exit Sub
    FormExistsShow frmWithholdingtax
End Sub

Private Sub cmdWork_Sheet_Click()
    If Module_Access(LOGID, "WORKSHEET", "REPORTS") = False Then Exit Sub
    FormExistsShow frmAMISWorkSheet
End Sub

Private Sub Command1_Click()
    FormExistsShow frmCancelledReport
End Sub

Private Sub Command10_Click()
    FormExistsShow frmSMIS_Log_Reminder
End Sub

Private Sub cmdAccountDetailbyCustomer_Click()
    If Module_Access(LOGID, "ACCOUNTS DETAIL BY CUSTOMER", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "SJ"
    FormExistsShow frmAMISDetailBySupplierWithAccountCode
    frmAMISDetailBySupplierWithAccountCode.Caption = "Accounts Detail Report By Customer"
End Sub

Private Sub cmdAccountDetailbySupplier_Click()
    If Module_Access(LOGID, "ACCOUNTS DETAIL BY SUPPLIERS", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "APJ"
    FormExistsShow frmAMISDetailBySupplierWithAccountCode
    frmAMISDetailBySupplierWithAccountCode.Caption = "Accounts Payable Detail Report By Supplier"
End Sub

Private Sub cmdAccountsPayableAgingReport_Click()
    If Module_Access(LOGID, "ACCOUNTS PAYABLE AGING REPORT", "REPORTS") = False Then Exit Sub
    REPORT_AP = "AGING"
    'frmAMISDueReport.Show
    'COMMENTED BY: JUN ORIGINAL AP FORM---------
    'frmAPschedulestandard.Show
    'COMMENTED BY: JUN ORIGINAL AP FORM---------
    '    If COMPANY_CODE = "HGC" Then
    '        FormExistsShow frmAMIS_AP_Process_old
    '    Else
    FormExistsShow frmAMIS_AP_Process
    '    FormExistsShow frmAMIS_AP_Process_old
    '    End If
End Sub

Private Sub cmdAccountsPayableDueReport_Click()
    If Module_Access(LOGID, "ACCOUNTS PAYABLE DUE REPORT", "REPORTS") = False Then Exit Sub
    REPORT_AP = "SCHED"
    frmAMISDueReport.Show
End Sub

Private Sub cmdAccountsPayableJournal_Click()
    If Module_Access(LOGID, "ACCOUNTS PAYABLE JOURNAL", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "APJ"
    FormExistsShow frmAMISRangeWithSummary
    frmAMISRangeWithSummary.Caption = "Accounts Payable Journal"
End Sub

Private Sub cmdAccountsReceivableAgingReport_Click()
    If Module_Access(LOGID, "ACCOUNTS RECEIVABLE AGING REPORT", "REPORTS") = False Then Exit Sub
    Report_AR = "AGING"
    'COMMENTED BY: ACL
    'frmAMISARSchedReport.Show
End Sub

'Private Sub Command11_Click()
'   If Module_Access(LOGID, "CREDITABLE WITHHOLDING TAX", "REPORTS") = False Then Exit Sub
'   frmWithholdingtax.Show
'End Sub

Private Sub cmdOpeningReport_Click()
    FormExistsShow frmOpeningBalanceReport
End Sub

Private Sub Command12_Click()
'    frmAMISARCheckingTool.Show
    'If Module_Access(LOGID, "IMPORTING TEMPLATE", "SYSTEM") = False Then Exit Sub
    frmAMISImporting_Template.Show
End Sub

Private Sub Command2_Click()
    frmAMISProfitbyVoucher.Show
End Sub

Private Sub Command27_Click()
    If Module_Access(LOGID, "SYSTEM SETUP", "SYSTEM") = False Then Exit Sub
    FormExistsShow frmAMISProfile
End Sub

Private Sub cmdCustomerCreditMemo_Click()
    If Module_Access(LOGID, "CUSTOMER ADJUSTMENT", "DATA ENTRY") = False Then Exit Sub
    'JOURNALTYPE = "CCM"
    On Error Resume Next
    Call frmAMISJournalEntry.LoadJournal("CCM")
    FormExistsShow frmAMISJournalEntry
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Call frmAMISbanksOpening.LoadJournal("BOB")
    FormExistsShow frmAMISbanksOpening
End Sub

Private Sub Command4_Click()
    FormExistsShow frmReprintReport
End Sub

Private Sub Command6_Click()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    FormExistsShow frmTransactionStatus
End Sub

Private Sub Command7_Click()
    If Module_Access(LOGID, "BANK RECONCILIATION", "DATA ENTRY") = False Then Exit Sub
    frmReconcileAccount.Show
End Sub

Private Sub Command8_Click()
    Screen.MousePointer = 11
    FormExistsShow frmAMIS_UNAPPLIED_PAYMENT
    Screen.MousePointer = 0
End Sub

Private Sub Command9_Click()
    Screen.MousePointer = 11
    FormExistsShow frmAMIS_UniportedReports
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    TabControl1.SelectedItem = 0
    TabControl2.SelectedItem = 0
    TabControl3.SelectedItem = 0
    TabControl4.SelectedItem = 0
    If COMPANY_CODE = "HBK" Then
        cmdJournalDRJ.Enabled = False
        Label28.Enabled = False
        Command4.Enabled = False
        Label42.Enabled = False
        Command1.Enabled = False
        Label45.Enabled = False
    End If
End Sub
Sub DisplayInfo(XXX As String)
    Dim RS                                        As New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_HD where jtype='APJ' and voucherno='" & XXX & "'")
    If Not RS.EOF And Not RS.BOF Then

    End If
    Set RS = Nothing
End Sub
Sub InitSalesJournal()

End Sub

Private Sub Label8_Click()
'frmAMIS_ARLEDGER.Show
End Sub

Private Sub TabControl1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If UCase(LOGNAME) = "NETSPEED" Then
        cmdTranType.Visible = True
    Else
        cmdTranType.Visible = False
    End If
End Sub

