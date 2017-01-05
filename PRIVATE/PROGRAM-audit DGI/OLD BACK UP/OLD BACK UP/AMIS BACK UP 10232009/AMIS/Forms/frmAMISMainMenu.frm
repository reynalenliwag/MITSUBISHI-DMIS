VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AMIS Main Menu"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   10635
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11235
      _Version        =   655364
      _ExtentX        =   19817
      _ExtentY        =   12938
      _StockProps     =   64
      Appearance      =   9
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   120
      PaintManager.MinTabWidth=   100
      ItemCount       =   5
      Item(0).Caption =   "Main Modules"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "Picture1"
      Item(1).Caption =   "Tables"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "Picture2"
      Item(2).Caption =   "File Maintenance"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "Picture3"
      Item(3).Caption =   "Reports"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "Picture4"
      Item(4).Caption =   "Other Setups"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "Picture5"
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   6705
         Left            =   -60
         ScaleHeight     =   6705
         ScaleWidth      =   10980
         TabIndex        =   70
         Top             =   570
         Width           =   10980
         Begin VB.CommandButton Command7 
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
            Left            =   5400
            MouseIcon       =   "frmAMISMainMenu.frx":15162
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":152B4
            Style           =   1  'Graphical
            TabIndex        =   156
            Tag             =   "1052"
            ToolTipText     =   "View Vendors Subsidiary Ledger"
            Top             =   2790
            Width           =   720
         End
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
            Left            =   360
            MouseIcon       =   "frmAMISMainMenu.frx":15BE4
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":15D36
            Style           =   1  'Graphical
            TabIndex        =   121
            Tag             =   "1047"
            ToolTipText     =   "View Cash Receipts Journal"
            Top             =   3750
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
            Left            =   5400
            MouseIcon       =   "frmAMISMainMenu.frx":165A1
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":166F3
            Style           =   1  'Graphical
            TabIndex        =   78
            Tag             =   "1052"
            ToolTipText     =   "View Vendors Subsidiary Ledger"
            Top             =   1980
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
            Left            =   5400
            MouseIcon       =   "frmAMISMainMenu.frx":17023
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":17175
            Style           =   1  'Graphical
            TabIndex        =   77
            Tag             =   "1050"
            ToolTipText     =   "View Customers A/R Ledger"
            Top             =   1110
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
            Left            =   5400
            MouseIcon       =   "frmAMISMainMenu.frx":17A88
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":17BDA
            Style           =   1  'Graphical
            TabIndex        =   76
            Tag             =   "1049"
            ToolTipText     =   "View Accounts General Ledger"
            Top             =   360
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
            Left            =   390
            MouseIcon       =   "frmAMISMainMenu.frx":184BC
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":1860E
            Style           =   1  'Graphical
            TabIndex        =   75
            Tag             =   "1047"
            ToolTipText     =   "View Cash Receipts Journal"
            Top             =   2850
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
            Left            =   360
            MouseIcon       =   "frmAMISMainMenu.frx":18E79
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":18FCB
            Style           =   1  'Graphical
            TabIndex        =   74
            Tag             =   "1048"
            ToolTipText     =   "View General Journal"
            Top             =   4605
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
            Left            =   360
            MouseIcon       =   "frmAMISMainMenu.frx":198D0
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":19A22
            Style           =   1  'Graphical
            TabIndex        =   73
            Tag             =   "1046"
            ToolTipText     =   "View Sales Journal"
            Top             =   2010
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
            Left            =   360
            MouseIcon       =   "frmAMISMainMenu.frx":1A200
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":1A352
            Style           =   1  'Graphical
            TabIndex        =   72
            Tag             =   "1045"
            ToolTipText     =   "View Cash Disbursement Journal"
            Top             =   1155
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
            Left            =   360
            MouseIcon       =   "frmAMISMainMenu.frx":1AD7D
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":1AECF
            Style           =   1  'Graphical
            TabIndex        =   71
            Tag             =   "1044"
            ToolTipText     =   "View Accounts Payable Journal"
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label64 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Reconcillation"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   6240
            TabIndex        =   157
            Top             =   2895
            Width           =   3945
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Receipts Journal (Deposited OR's)"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   870
            Left            =   1185
            TabIndex        =   122
            Top             =   3780
            Width           =   4065
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Accounts General Ledger"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   6300
            TabIndex        =   86
            Top             =   480
            Width           =   3810
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Customers A/R Ledger"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   6300
            TabIndex        =   85
            Top             =   1305
            Width           =   3810
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Vendors A/P Ledger"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   6300
            TabIndex        =   84
            Top             =   2085
            Width           =   3945
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "General Journal"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1185
            TabIndex        =   83
            Top             =   4755
            Width           =   3300
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Journal"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   1185
            TabIndex        =   82
            Top             =   2145
            Width           =   3105
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Receipts Journal (Un-Deposited OR's)"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   870
            Left            =   1185
            TabIndex        =   81
            Top             =   2850
            Width           =   4065
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Disbursement Journal"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   1185
            TabIndex        =   80
            Top             =   1275
            Width           =   3945
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Accounts Payable Journal"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Left            =   1185
            TabIndex        =   79
            Top             =   495
            Width           =   3810
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
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
         Height          =   6660
         Left            =   -70015
         ScaleHeight     =   6660
         ScaleWidth      =   11115
         TabIndex        =   4
         Top             =   570
         Visible         =   0   'False
         Width           =   11115
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
            Left            =   390
            MouseIcon       =   "frmAMISMainMenu.frx":1B703
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":1B855
            Style           =   1  'Graphical
            TabIndex        =   10
            Tag             =   "1027"
            ToolTipText     =   "View Customer Master Files"
            Top             =   390
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
            Left            =   390
            MouseIcon       =   "frmAMISMainMenu.frx":1BEBC
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":1C00E
            Style           =   1  'Graphical
            TabIndex        =   9
            Tag             =   "1028"
            ToolTipText     =   "View Vendor Master Files"
            Top             =   1197
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
            Left            =   390
            MouseIcon       =   "frmAMISMainMenu.frx":1C725
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":1C877
            Style           =   1  'Graphical
            TabIndex        =   8
            Tag             =   "1029"
            ToolTipText     =   "View Bank Master Files"
            Top             =   2004
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
            Left            =   390
            MouseIcon       =   "frmAMISMainMenu.frx":1CF52
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":1D0A4
            Style           =   1  'Graphical
            TabIndex        =   7
            Tag             =   "1031"
            ToolTipText     =   "View Terms Of Payment Master File"
            Top             =   3618
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
            Left            =   390
            MouseIcon       =   "frmAMISMainMenu.frx":1D73D
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":1D88F
            Style           =   1  'Graphical
            TabIndex        =   6
            Tag             =   "1030"
            ToolTipText     =   "View Invoice Type Master Files"
            Top             =   2811
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
            Left            =   390
            MouseIcon       =   "frmAMISMainMenu.frx":1DF41
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":1E093
            Style           =   1  'Graphical
            TabIndex        =   5
            Tag             =   "1032"
            ToolTipText     =   "View ATC Code Master File"
            Top             =   4425
            Width           =   720
         End
         Begin VB.Label Label66 
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Master Files"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   1260
            TabIndex        =   16
            Top             =   1365
            Width           =   5805
         End
         Begin VB.Label Label57 
            BackStyle       =   0  'Transparent
            Caption         =   "ATC Code Master File"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   600
            Left            =   1260
            TabIndex        =   15
            Top             =   4560
            Width           =   6600
         End
         Begin VB.Label Label56 
            BackStyle       =   0  'Transparent
            Caption         =   "Terms Of Payment Master File"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   1260
            TabIndex        =   14
            Top             =   3765
            Width           =   6510
         End
         Begin VB.Label Label55 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Master Files"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   525
            Left            =   1260
            TabIndex        =   13
            Top             =   2160
            Width           =   6405
         End
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Type Master Files"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   450
            Left            =   1260
            TabIndex        =   12
            Top             =   2955
            Width           =   6090
         End
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Master Files"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   645
            Left            =   1260
            TabIndex        =   11
            Top             =   555
            Width           =   5715
         End
      End
      Begin VB.PictureBox Picture3 
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
         Height          =   6705
         Left            =   -70000
         ScaleHeight     =   6705
         ScaleWidth      =   11115
         TabIndex        =   3
         Top             =   570
         Visible         =   0   'False
         Width           =   11115
         Begin XtremeSuiteControls.TabControl TabControl4 
            Height          =   5640
            Left            =   0
            TabIndex        =   87
            Top             =   0
            Width           =   10680
            _Version        =   655364
            _ExtentX        =   18838
            _ExtentY        =   9948
            _StockProps     =   64
            Appearance      =   9
            Color           =   4
            PaintManager.BoldSelected=   -1  'True
            PaintManager.DisableLunaColors=   0   'False
            PaintManager.OneNoteColors=   -1  'True
            PaintManager.HotTracking=   -1  'True
            PaintManager.ShowIcons=   -1  'True
            PaintManager.LargeIcons=   -1  'True
            PaintManager.FixedTabWidth=   120
            PaintManager.MinTabWidth=   100
            ItemCount       =   3
            Item(0).Caption =   "Accounts"
            Item(0).Tooltip =   "Accounts"
            Item(0).ControlCount=   1
            Item(0).Control(0)=   "Picture11"
            Item(1).Caption =   "Opening Balances"
            Item(1).Tooltip =   "Opening Balances"
            Item(1).ControlCount=   1
            Item(1).Control(0)=   "Picture12"
            Item(2).Caption =   "Adjustments"
            Item(2).Tooltip =   "Adjustments"
            Item(2).ControlCount=   1
            Item(2).Control(0)=   "Picture13"
            Begin VB.PictureBox Picture13 
               BackColor       =   &H00E0E0E0&
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
               Height          =   6705
               Left            =   -70090
               ScaleHeight     =   6705
               ScaleWidth      =   11025
               TabIndex        =   90
               Top             =   510
               Visible         =   0   'False
               Width           =   11025
               Begin VB.CommandButton cmdVendorCreditMemo 
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
                  MouseIcon       =   "frmAMISMainMenu.frx":1E746
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":1E898
                  Style           =   1  'Graphical
                  TabIndex        =   145
                  Tag             =   "1039"
                  ToolTipText     =   "View Vendor Adjustments"
                  Top             =   2610
                  Width           =   720
               End
               Begin VB.CommandButton cmdCustomerCreditMemo 
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
                  MouseIcon       =   "frmAMISMainMenu.frx":1EF03
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":1F055
                  Style           =   1  'Graphical
                  TabIndex        =   143
                  Tag             =   "1038"
                  ToolTipText     =   "View Customer Adjustments"
                  Top             =   1110
                  Width           =   720
               End
               Begin VB.CommandButton cmdVendorDebitMemo 
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
                  MouseIcon       =   "frmAMISMainMenu.frx":1F6DA
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":1F82C
                  Style           =   1  'Graphical
                  TabIndex        =   115
                  Tag             =   "1039"
                  ToolTipText     =   "View Vendor Adjustments"
                  Top             =   1875
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
                  Left            =   7380
                  MouseIcon       =   "frmAMISMainMenu.frx":1FE97
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":1FFE9
                  Style           =   1  'Graphical
                  TabIndex        =   114
                  Tag             =   "1040"
                  ToolTipText     =   "View Closing Entries"
                  Top             =   300
                  Width           =   720
               End
               Begin VB.CommandButton cmdCustomerDebitMemo 
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
                  MouseIcon       =   "frmAMISMainMenu.frx":2062C
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":2077E
                  Style           =   1  'Graphical
                  TabIndex        =   113
                  Tag             =   "1038"
                  ToolTipText     =   "View Customer Adjustments"
                  Top             =   330
                  Width           =   720
               End
               Begin VB.CommandButton cmdAdjustment_Client 
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
                  MouseIcon       =   "frmAMISMainMenu.frx":20E03
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":20F55
                  Style           =   1  'Graphical
                  TabIndex        =   112
                  Tag             =   "1036"
                  ToolTipText     =   "View Client Adjusting Journal Entries"
                  Top             =   3540
                  Width           =   720
               End
               Begin VB.CommandButton cmdAdjustment_Proposed 
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
                  MouseIcon       =   "frmAMISMainMenu.frx":21638
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":2178A
                  Style           =   1  'Graphical
                  TabIndex        =   111
                  Tag             =   "1037"
                  ToolTipText     =   "View Proposed Adjusting Journal Entries"
                  Top             =   4335
                  Width           =   720
               End
               Begin VB.Line Line 
                  BorderColor     =   &H00808080&
                  X1              =   90
                  X2              =   10740
                  Y1              =   3420
                  Y2              =   3420
               End
               Begin VB.Label Label41 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Vendor Credit Memo"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   495
                  Left            =   1350
                  TabIndex        =   146
                  Top             =   2775
                  Width           =   7005
               End
               Begin VB.Label Label7 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Customer Credit Memo"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   615
                  Left            =   1335
                  TabIndex        =   144
                  Top             =   1260
                  Width           =   6720
               End
               Begin VB.Label Label30 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Customer Debit Memo"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   615
                  Left            =   1335
                  TabIndex        =   120
                  Top             =   480
                  Width           =   6720
               End
               Begin VB.Label Label43 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Closing Entries"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   615
                  Left            =   8235
                  TabIndex        =   119
                  Top             =   435
                  Width           =   6960
               End
               Begin VB.Label Label44 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Vendor Debit Memo"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   495
                  Left            =   1350
                  TabIndex        =   118
                  Top             =   2040
                  Width           =   7005
               End
               Begin VB.Label Label49 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Client Adjusting Journal Entries"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   435
                  Left            =   1320
                  TabIndex        =   117
                  Top             =   3690
                  Width           =   5910
               End
               Begin VB.Label Label51 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Proposed Adjusting Journal Entries"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   465
                  Left            =   1320
                  TabIndex        =   116
                  Top             =   4485
                  Width           =   6960
               End
            End
            Begin VB.PictureBox Picture12 
               BackColor       =   &H00E0E0E0&
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
               Height          =   6705
               Left            =   -70000
               ScaleHeight     =   6705
               ScaleWidth      =   11025
               TabIndex        =   89
               Top             =   510
               Visible         =   0   'False
               Width           =   11025
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
                  Left            =   420
                  MouseIcon       =   "frmAMISMainMenu.frx":21E71
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":21FC3
                  Style           =   1  'Graphical
                  TabIndex        =   147
                  Tag             =   "1045"
                  ToolTipText     =   "View Cash Disbursement Journal"
                  Top             =   3720
                  Width           =   720
               End
               Begin VB.CommandButton Command2 
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
                  Left            =   390
                  MouseIcon       =   "frmAMISMainMenu.frx":229EE
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":22B40
                  Style           =   1  'Graphical
                  TabIndex        =   128
                  Tag             =   "1034"
                  ToolTipText     =   "View Customer Opening Balance"
                  Top             =   2910
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
                  Left            =   390
                  MouseIcon       =   "frmAMISMainMenu.frx":2331B
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":2346D
                  Style           =   1  'Graphical
                  TabIndex        =   107
                  Tag             =   "1033"
                  ToolTipText     =   "View Accounts Opening Balance"
                  Top             =   345
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
                  Left            =   390
                  MouseIcon       =   "frmAMISMainMenu.frx":23B5E
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":23CB0
                  Style           =   1  'Graphical
                  TabIndex        =   106
                  Tag             =   "1034"
                  ToolTipText     =   "View Customer Opening Balance"
                  Top             =   1207
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
                  Left            =   390
                  MouseIcon       =   "frmAMISMainMenu.frx":2448B
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":245DD
                  Style           =   1  'Graphical
                  TabIndex        =   105
                  Tag             =   "1035"
                  ToolTipText     =   "View Vendor Opening Balance"
                  Top             =   2070
                  Width           =   720
               End
               Begin VB.Label Label24 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Bank Opening Balance"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   420
                  Left            =   1320
                  TabIndex        =   148
                  Top             =   3870
                  Width           =   3855
               End
               Begin VB.Label Label37 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Opening Balance Report"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   390
                  Left            =   1290
                  TabIndex        =   129
                  Top             =   3060
                  Width           =   4950
               End
               Begin VB.Label Label47 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Accounts Opening Balance"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   630
                  Left            =   1290
                  TabIndex        =   110
                  Top             =   495
                  Width           =   5280
               End
               Begin VB.Label Label48 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Customer Opening Balance"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   570
                  Left            =   1275
                  TabIndex        =   109
                  Top             =   1350
                  Width           =   4410
               End
               Begin VB.Label Label54 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Vendor Opening Balance"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   765
                  Left            =   1290
                  TabIndex        =   108
                  Top             =   2205
                  Width           =   4365
               End
            End
            Begin VB.PictureBox Picture11 
               BackColor       =   &H00E0E0E0&
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
               Height          =   6705
               Left            =   -75
               ScaleHeight     =   6705
               ScaleWidth      =   11025
               TabIndex        =   88
               Top             =   510
               Width           =   11025
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
                  Left            =   4275
                  MouseIcon       =   "frmAMISMainMenu.frx":24D8F
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":24EE1
                  Style           =   1  'Graphical
                  TabIndex        =   97
                  Tag             =   "1026"
                  ToolTipText     =   "View Account Entries Templates"
                  Top             =   1260
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
                  Left            =   4275
                  MouseIcon       =   "frmAMISMainMenu.frx":2559C
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":256EE
                  Style           =   1  'Graphical
                  TabIndex        =   96
                  Tag             =   "1025"
                  ToolTipText     =   "View Department Codes"
                  Top             =   375
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
                  Left            =   450
                  MouseIcon       =   "frmAMISMainMenu.frx":25E4F
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":25FA1
                  Style           =   1  'Graphical
                  TabIndex        =   95
                  Tag             =   "1023"
                  ToolTipText     =   "View Extended Classification"
                  Top             =   3020
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
                  Left            =   450
                  MouseIcon       =   "frmAMISMainMenu.frx":26731
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":26883
                  Style           =   1  'Graphical
                  TabIndex        =   94
                  Tag             =   "1024"
                  ToolTipText     =   "View Account Sub-Totals"
                  Top             =   3900
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
                  Left            =   450
                  MouseIcon       =   "frmAMISMainMenu.frx":26F2F
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":27081
                  Style           =   1  'Graphical
                  TabIndex        =   93
                  Tag             =   "1022"
                  ToolTipText     =   "View Account Classification"
                  Top             =   2141
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
                  Left            =   450
                  MouseIcon       =   "frmAMISMainMenu.frx":276ED
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":2783F
                  Style           =   1  'Graphical
                  TabIndex        =   92
                  Tag             =   "1021"
                  ToolTipText     =   "View Account Types"
                  Top             =   1267
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
                  Left            =   450
                  MouseIcon       =   "frmAMISMainMenu.frx":27EBE
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":28010
                  Style           =   1  'Graphical
                  TabIndex        =   91
                  Tag             =   "1020"
                  ToolTipText     =   "View Chart Of Accounts"
                  Top             =   375
                  Width           =   720
               End
               Begin VB.Label Label58 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Chart Of Accounts"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   435
                  Left            =   1320
                  TabIndex        =   104
                  Top             =   540
                  Width           =   4035
               End
               Begin VB.Label Label59 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Account Types"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   540
                  Left            =   1320
                  TabIndex        =   103
                  Top             =   1425
                  Width           =   4215
               End
               Begin VB.Label Label60 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Account Sub-Totals"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   495
                  Left            =   1320
                  TabIndex        =   102
                  Top             =   4035
                  Width           =   4680
               End
               Begin VB.Label Label61 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Account Classification"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   480
                  Left            =   1320
                  TabIndex        =   101
                  Top             =   2280
                  Width           =   4500
               End
               Begin VB.Label Label62 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Extended Classification"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   405
                  Left            =   1320
                  TabIndex        =   100
                  Top             =   3150
                  Width           =   4965
               End
               Begin VB.Label Label63 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Department Codes"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   450
                  Left            =   5145
                  TabIndex        =   99
                  Top             =   510
                  Width           =   4590
               End
               Begin VB.Label Label65 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Account Entries Templates"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   495
                  Left            =   5145
                  TabIndex        =   98
                  Top             =   1380
                  Width           =   6270
               End
            End
         End
      End
      Begin VB.PictureBox Picture4 
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
         Height          =   6705
         Left            =   -70015
         ScaleHeight     =   6705
         ScaleWidth      =   11070
         TabIndex        =   2
         Top             =   585
         Visible         =   0   'False
         Width           =   11070
         Begin XtremeSuiteControls.TabControl TabControl2 
            Height          =   5730
            Left            =   0
            TabIndex        =   17
            Top             =   15
            Width           =   11490
            _Version        =   655364
            _ExtentX        =   20267
            _ExtentY        =   10107
            _StockProps     =   64
            Appearance      =   9
            Color           =   4
            PaintManager.BoldSelected=   -1  'True
            PaintManager.DisableLunaColors=   0   'False
            PaintManager.OneNoteColors=   -1  'True
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
            Item(1).ControlCount=   1
            Item(1).Control(0)=   "Picture7"
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
               TabIndex        =   130
               Top             =   600
               Width           =   10800
            End
            Begin VB.PictureBox Picture7 
               BackColor       =   &H00E0E0E0&
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
               Height          =   6705
               Left            =   -70135
               ScaleHeight     =   6705
               ScaleWidth      =   11025
               TabIndex        =   61
               Top             =   540
               Visible         =   0   'False
               Width           =   11025
               Begin VB.CommandButton cmdTrialBalance 
                  Height          =   645
                  Left            =   450
                  MouseIcon       =   "frmAMISMainMenu.frx":28644
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":28796
                  Style           =   1  'Graphical
                  TabIndex        =   65
                  ToolTipText     =   "View Trial Balance"
                  Top             =   1035
                  Width           =   720
               End
               Begin VB.CommandButton cmdWork_Sheet 
                  Height          =   645
                  Left            =   450
                  MouseIcon       =   "frmAMISMainMenu.frx":28D24
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":28E76
                  Style           =   1  'Graphical
                  TabIndex        =   64
                  ToolTipText     =   "View Work Sheet"
                  Top             =   210
                  Width           =   720
               End
               Begin VB.CommandButton cmdScheduleOfAdjustments 
                  Height          =   645
                  Left            =   450
                  MouseIcon       =   "frmAMISMainMenu.frx":29679
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":297CB
                  Style           =   1  'Graphical
                  TabIndex        =   63
                  ToolTipText     =   "View Schedule Of Adjustments"
                  Top             =   1875
                  Width           =   720
               End
               Begin VB.CommandButton cmdFinancialStatements 
                  Height          =   645
                  Left            =   435
                  MouseIcon       =   "frmAMISMainMenu.frx":29F84
                  MousePointer    =   99  'Custom
                  Picture         =   "frmAMISMainMenu.frx":2A0D6
                  Style           =   1  'Graphical
                  TabIndex        =   62
                  ToolTipText     =   "View Financial Statements"
                  Top             =   2670
                  Width           =   720
               End
               Begin VB.Label cmdWorkSheet 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Work Sheet"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   420
                  Left            =   1305
                  TabIndex        =   69
                  Top             =   390
                  Width           =   2775
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Trial Balance"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   420
                  Left            =   1305
                  TabIndex        =   68
                  Top             =   1215
                  Width           =   2325
               End
               Begin VB.Label Label16 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Schedule Of Adjustments"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   420
                  Left            =   1305
                  TabIndex        =   67
                  Top             =   2025
                  Width           =   3900
               End
               Begin VB.Label Label17 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Financial Statements"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   420
                  Left            =   1305
                  TabIndex        =   66
                  Top             =   2820
                  Width           =   3855
               End
            End
            Begin XtremeSuiteControls.TabControl TabControl3 
               Height          =   5175
               Left            =   -75
               TabIndex        =   18
               Top             =   465
               Width           =   10800
               _Version        =   655364
               _ExtentX        =   19050
               _ExtentY        =   9128
               _StockProps     =   64
               Appearance      =   1
               Color           =   4
               PaintManager.BoldSelected=   -1  'True
               PaintManager.DisableLunaColors=   0   'False
               PaintManager.HotTracking=   -1  'True
               PaintManager.ShowIcons=   -1  'True
               PaintManager.FixedTabWidth=   110
               PaintManager.MinTabWidth=   85
               ItemCount       =   5
               Item(0).Caption =   "Accounts Payable"
               Item(0).ControlCount=   1
               Item(0).Control(0)=   "Picture6"
               Item(1).Caption =   "Cash Disbursement"
               Item(1).ControlCount=   1
               Item(1).Control(0)=   "Picture8"
               Item(2).Caption =   "Sales"
               Item(2).ControlCount=   1
               Item(2).Control(0)=   "Picture9"
               Item(3).Caption =   "Cash Receipts"
               Item(3).ControlCount=   1
               Item(3).Control(0)=   "Picture10"
               Item(4).Caption =   "General Journal"
               Item(4).ControlCount=   1
               Item(4).Control(0)=   "Picture14"
               Begin VB.PictureBox Picture14 
                  BackColor       =   &H00E0E0E0&
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
                  TabIndex        =   123
                  Top             =   330
                  Visible         =   0   'False
                  Width           =   10800
                  Begin VB.CommandButton cmdJournalVoucherSummary 
                     Height          =   645
                     Left            =   330
                     MouseIcon       =   "frmAMISMainMenu.frx":2A82C
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":2A97E
                     Style           =   1  'Graphical
                     TabIndex        =   125
                     ToolTipText     =   "View Cash Receipts Journal"
                     Top             =   345
                     Width           =   720
                  End
                  Begin VB.CommandButton cmdGJLedgerCodeRunningBalance 
                     Height          =   645
                     Left            =   330
                     MouseIcon       =   "frmAMISMainMenu.frx":2B0FC
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":2B24E
                     Style           =   1  'Graphical
                     TabIndex        =   124
                     ToolTipText     =   "View Ledger Code Running Balance"
                     Top             =   1185
                     Width           =   720
                  End
                  Begin VB.Label Label13 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Ledger Code Running Balance"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   480
                     Left            =   1200
                     TabIndex        =   127
                     Top             =   1320
                     Width           =   4935
                  End
                  Begin VB.Label Label12 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Journal Voucher Summary"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   420
                     Left            =   1200
                     TabIndex        =   126
                     Top             =   495
                     Width           =   3930
                  End
               End
               Begin VB.PictureBox Picture6 
                  BackColor       =   &H00E0E0E0&
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
                  Height          =   6165
                  Left            =   -15
                  ScaleHeight     =   6165
                  ScaleWidth      =   10770
                  TabIndex        =   50
                  Top             =   315
                  Width           =   10770
                  Begin VB.CommandButton cmdAccountDetailbySupplier 
                     Height          =   645
                     Left            =   285
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":2B907
                     Style           =   1  'Graphical
                     TabIndex        =   55
                     ToolTipText     =   "View Account Detail by Supplier"
                     Top             =   2010
                     Width           =   720
                  End
                  Begin VB.CommandButton cmdAccountsPayableAgingReport 
                     Height          =   645
                     Left            =   270
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":2BFCE
                     Style           =   1  'Graphical
                     TabIndex        =   54
                     ToolTipText     =   "View Accounts Payable Aging Report"
                     Top             =   2790
                     Width           =   720
                  End
                  Begin VB.CommandButton cmdLedgerCodeRunningBalance_AP 
                     Height          =   645
                     Left            =   285
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":2C6FB
                     Style           =   1  'Graphical
                     TabIndex        =   53
                     ToolTipText     =   "View Ledger Code Running Balance"
                     Top             =   1170
                     Width           =   720
                  End
                  Begin VB.CommandButton cmdAccountsPayableJournal 
                     Height          =   645
                     Left            =   285
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":2CDB4
                     Style           =   1  'Graphical
                     TabIndex        =   52
                     ToolTipText     =   "View Accounts Payable Journal"
                     Top             =   330
                     Width           =   720
                  End
                  Begin VB.CommandButton cmdReceivingReportRegister 
                     Height          =   645
                     Left            =   5355
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":2D472
                     Style           =   1  'Graphical
                     TabIndex        =   51
                     ToolTipText     =   "View Receiving Report Register"
                     Top             =   345
                     Width           =   720
                  End
                  Begin VB.Label Label6 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Accounts Payable Journal"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   420
                     Left            =   1155
                     TabIndex        =   60
                     Top             =   480
                     Width           =   3930
                  End
                  Begin VB.Label Label18 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Ledger Code Running Balance"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   675
                     Left            =   1155
                     TabIndex        =   59
                     Top             =   1170
                     Width           =   3225
                  End
                  Begin VB.Label Label19 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Account Detail by Supplier"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   420
                     Left            =   1155
                     TabIndex        =   58
                     Top             =   2160
                     Width           =   4515
                  End
                  Begin VB.Label Label21 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Accounts Payable Aging Report"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   420
                     Left            =   1155
                     TabIndex        =   57
                     Top             =   2925
                     Width           =   5430
                  End
                  Begin VB.Label Label23 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Receiving Report Register"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   615
                     Left            =   6240
                     TabIndex        =   56
                     Top             =   495
                     Width           =   3750
                  End
               End
               Begin VB.PictureBox Picture8 
                  BackColor       =   &H00E0E0E0&
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
                  Left            =   -70015
                  ScaleHeight     =   5835
                  ScaleWidth      =   10800
                  TabIndex        =   43
                  Top             =   330
                  Visible         =   0   'False
                  Width           =   10800
                  Begin VB.CommandButton cmdCheckRegister 
                     Height          =   645
                     Left            =   285
                     MouseIcon       =   "frmAMISMainMenu.frx":2DAF3
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":2DC45
                     Style           =   1  'Graphical
                     TabIndex        =   46
                     ToolTipText     =   "View Check Register"
                     Top             =   2010
                     Width           =   720
                  End
                  Begin VB.CommandButton cmdLedgerCodeRunningBalance_CD 
                     Height          =   645
                     Left            =   285
                     MouseIcon       =   "frmAMISMainMenu.frx":2E3A2
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":2E4F4
                     Style           =   1  'Graphical
                     TabIndex        =   45
                     ToolTipText     =   "View Ledger Code Running Balance"
                     Top             =   1162
                     Width           =   720
                  End
                  Begin VB.CommandButton cmdCashDisbursementJournal 
                     Height          =   645
                     Left            =   285
                     MouseIcon       =   "frmAMISMainMenu.frx":2EBAD
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":2ECFF
                     Style           =   1  'Graphical
                     TabIndex        =   44
                     ToolTipText     =   "View Cash Disbursement Journal"
                     Top             =   315
                     Width           =   720
                  End
                  Begin VB.Label Label25 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cash Disbursement Journal"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   420
                     Left            =   1155
                     TabIndex        =   49
                     Top             =   465
                     Width           =   5115
                  End
                  Begin VB.Label Label26 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Ledger Code Running Balance"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   450
                     Left            =   1155
                     TabIndex        =   48
                     Top             =   1305
                     Width           =   5490
                  End
                  Begin VB.Label Label27 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Check Register"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   420
                     Left            =   1155
                     TabIndex        =   47
                     Top             =   2145
                     Width           =   4515
                  End
               End
               Begin VB.PictureBox Picture9 
                  BackColor       =   &H00E0E0E0&
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
                  Height          =   5985
                  Left            =   -70000
                  ScaleHeight     =   5985
                  ScaleWidth      =   10800
                  TabIndex        =   28
                  Top             =   315
                  Visible         =   0   'False
                  Width           =   10800
                  Begin VB.CommandButton cmdScheduleOfAccountsReceivable 
                     Height          =   645
                     Left            =   240
                     MouseIcon       =   "frmAMISMainMenu.frx":2F4D5
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":2F627
                     Style           =   1  'Graphical
                     TabIndex        =   35
                     ToolTipText     =   "View Schedule Of Accounts Receivable"
                     Top             =   2820
                     Width           =   720
                  End
                  Begin VB.CommandButton cmdLedgerCodeRunningBalance_LCRB 
                     Height          =   645
                     Left            =   270
                     MouseIcon       =   "frmAMISMainMenu.frx":2FD5F
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":2FEB1
                     Style           =   1  'Graphical
                     TabIndex        =   34
                     ToolTipText     =   "View Ledger Code Running Balance"
                     Top             =   2010
                     Width           =   720
                  End
                  Begin VB.CommandButton cmdAccountDetailbyCustomer 
                     Height          =   645
                     Left            =   270
                     MouseIcon       =   "frmAMISMainMenu.frx":3056A
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":306BC
                     Style           =   1  'Graphical
                     TabIndex        =   33
                     ToolTipText     =   "View Account Detail by Customer"
                     Top             =   1170
                     Width           =   720
                  End
                  Begin VB.CommandButton cmdSalesJournal 
                     Height          =   645
                     Left            =   270
                     MouseIcon       =   "frmAMISMainMenu.frx":30D64
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":30EB6
                     Style           =   1  'Graphical
                     TabIndex        =   32
                     ToolTipText     =   "View Sales Journal"
                     Top             =   330
                     Width           =   720
                  End
                  Begin VB.CommandButton cmdInvoiceRegister 
                     Height          =   645
                     Left            =   5940
                     MouseIcon       =   "frmAMISMainMenu.frx":3161D
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":3176F
                     Style           =   1  'Graphical
                     TabIndex        =   31
                     ToolTipText     =   "View Invoice Register"
                     Top             =   330
                     Width           =   720
                  End
                  Begin VB.CommandButton cmdSalesbyInvoiceType 
                     Height          =   645
                     Left            =   5925
                     MouseIcon       =   "frmAMISMainMenu.frx":31E09
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":31F5B
                     Style           =   1  'Graphical
                     TabIndex        =   30
                     ToolTipText     =   "View Sales by Invoice Type"
                     Top             =   1185
                     Width           =   720
                  End
                  Begin VB.CommandButton cmdUnusedInvoices 
                     Height          =   645
                     Left            =   240
                     MouseIcon       =   "frmAMISMainMenu.frx":3269A
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":327EC
                     Style           =   1  'Graphical
                     TabIndex        =   29
                     ToolTipText     =   "View Unused Invoices"
                     Top             =   3660
                     Width           =   720
                  End
                  Begin VB.Label Label29 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Sales Journal"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   420
                     Left            =   1140
                     TabIndex        =   42
                     Top             =   480
                     Width           =   3930
                  End
                  Begin VB.Label Label31 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Account Detail by Customer"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   435
                     Left            =   1140
                     TabIndex        =   41
                     Top             =   1335
                     Width           =   4155
                  End
                  Begin VB.Label Label32 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Ledger Code Running Balance"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   420
                     Left            =   1140
                     TabIndex        =   40
                     Top             =   2160
                     Width           =   4515
                  End
                  Begin VB.Label Label34 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Accounts Receivable Report"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   420
                     Left            =   1140
                     TabIndex        =   39
                     Top             =   2940
                     Width           =   4110
                  End
                  Begin VB.Label Label35 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Sales by Invoice Type"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   615
                     Left            =   6795
                     TabIndex        =   38
                     Top             =   1320
                     Width           =   3750
                  End
                  Begin VB.Label Label36 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Invoice Register"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   405
                     Left            =   6825
                     TabIndex        =   37
                     Top             =   465
                     Width           =   3495
                  End
                  Begin VB.Label Label38 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Unused Invoices"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   615
                     Left            =   1140
                     TabIndex        =   36
                     Top             =   3840
                     Width           =   3750
                  End
               End
               Begin VB.PictureBox Picture10 
                  BackColor       =   &H00E0E0E0&
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
                  Left            =   -70015
                  ScaleHeight     =   5835
                  ScaleWidth      =   10800
                  TabIndex        =   19
                  Top             =   330
                  Visible         =   0   'False
                  Width           =   10800
                  Begin VB.CommandButton cmdUnused_OR 
                     Height          =   645
                     Left            =   315
                     MouseIcon       =   "frmAMISMainMenu.frx":32F7C
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":330CE
                     Style           =   1  'Graphical
                     TabIndex        =   23
                     ToolTipText     =   "View Unused O.R."
                     Top             =   2865
                     Width           =   720
                  End
                  Begin VB.CommandButton cmdOR_Register 
                     Height          =   645
                     Left            =   330
                     MouseIcon       =   "frmAMISMainMenu.frx":3389E
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":339F0
                     Style           =   1  'Graphical
                     TabIndex        =   22
                     ToolTipText     =   "View O.R. Register"
                     Top             =   2025
                     Width           =   720
                  End
                  Begin VB.CommandButton cmdLedgerCodeRunningBalance_CR 
                     Height          =   645
                     Left            =   330
                     MouseIcon       =   "frmAMISMainMenu.frx":3416E
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":342C0
                     Style           =   1  'Graphical
                     TabIndex        =   21
                     ToolTipText     =   "View Ledger Code Running Balance"
                     Top             =   1185
                     Width           =   720
                  End
                  Begin VB.CommandButton cmdCashReceiptsJournal 
                     Height          =   645
                     Left            =   330
                     MouseIcon       =   "frmAMISMainMenu.frx":34979
                     MousePointer    =   99  'Custom
                     Picture         =   "frmAMISMainMenu.frx":34ACB
                     Style           =   1  'Graphical
                     TabIndex        =   20
                     ToolTipText     =   "View Cash Receipts Journal"
                     Top             =   345
                     Width           =   720
                  End
                  Begin VB.Label Label39 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cash Receipts Journal"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   420
                     Left            =   1200
                     TabIndex        =   27
                     Top             =   495
                     Width           =   3930
                  End
                  Begin VB.Label Label40 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Ledger Code Running Balance"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   480
                     Left            =   1200
                     TabIndex        =   26
                     Top             =   1320
                     Width           =   4935
                  End
                  Begin VB.Label Label46 
                     BackStyle       =   0  'Transparent
                     Caption         =   "O.R. Register"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   420
                     Left            =   1200
                     TabIndex        =   25
                     Top             =   2175
                     Width           =   4515
                  End
                  Begin VB.Label Label67 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Unused O.R."
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   420
                     Left            =   1200
                     TabIndex        =   24
                     Top             =   3000
                     Width           =   4605
                  End
               End
            End
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00E0E0E0&
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
         Height          =   6705
         Left            =   -70090
         ScaleHeight     =   6705
         ScaleWidth      =   11025
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   11025
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
            Left            =   180
            MouseIcon       =   "frmAMISMainMenu.frx":35249
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":3539B
            Style           =   1  'Graphical
            TabIndex        =   153
            Tag             =   "1052"
            ToolTipText     =   "View Vendors Subsidiary Ledger"
            Top             =   3330
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
            Left            =   180
            MouseIcon       =   "frmAMISMainMenu.frx":35CCB
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":35E1D
            Style           =   1  'Graphical
            TabIndex        =   152
            Tag             =   "1052"
            ToolTipText     =   "View Vendors Subsidiary Ledger"
            Top             =   4050
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
            Left            =   5520
            MouseIcon       =   "frmAMISMainMenu.frx":3674D
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":3689F
            Style           =   1  'Graphical
            TabIndex        =   150
            Tag             =   "1052"
            ToolTipText     =   "View Vendors Subsidiary Ledger"
            Top             =   1800
            Width           =   720
         End
         Begin VB.CommandButton Command1 
            Height          =   645
            Left            =   5550
            MouseIcon       =   "frmAMISMainMenu.frx":371CF
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":37321
            Style           =   1  'Graphical
            TabIndex        =   137
            ToolTipText     =   "View Unused O.R."
            Top             =   990
            Width           =   720
         End
         Begin VB.CommandButton Command4 
            Height          =   645
            Left            =   5550
            MouseIcon       =   "frmAMISMainMenu.frx":37AF1
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":37C43
            Style           =   1  'Graphical
            TabIndex        =   136
            ToolTipText     =   "View O.R. Register"
            Top             =   210
            Width           =   720
         End
         Begin VB.CommandButton cmdAuditReport 
            Height          =   645
            Left            =   180
            MouseIcon       =   "frmAMISMainMenu.frx":383C1
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":38513
            Style           =   1  'Graphical
            TabIndex        =   135
            ToolTipText     =   "View Signatories and Headers"
            Top             =   2550
            Width           =   720
         End
         Begin VB.CommandButton cmdAuditInquiry 
            Height          =   645
            Left            =   180
            MouseIcon       =   "frmAMISMainMenu.frx":38955
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":38AA7
            Style           =   1  'Graphical
            TabIndex        =   134
            ToolTipText     =   "View Signatories and Headers"
            Top             =   1770
            Width           =   720
         End
         Begin VB.CommandButton Command27 
            Height          =   645
            Left            =   180
            MouseIcon       =   "frmAMISMainMenu.frx":38EE9
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":3903B
            Style           =   1  'Graphical
            TabIndex        =   133
            ToolTipText     =   "View Signatories and Headers"
            Top             =   990
            Width           =   720
         End
         Begin VB.CommandButton Command10 
            Height          =   645
            Left            =   180
            MouseIcon       =   "frmAMISMainMenu.frx":3947D
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":395CF
            Style           =   1  'Graphical
            TabIndex        =   131
            Tag             =   "1102"
            ToolTipText     =   "View Reminders"
            Top             =   210
            Width           =   720
         End
         Begin VB.Label Label68 
            BackStyle       =   0  'Transparent
            Caption         =   "Un-Applied Payment"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   1050
            TabIndex        =   155
            Top             =   3480
            Width           =   3855
         End
         Begin VB.Label Label69 
            BackStyle       =   0  'Transparent
            Caption         =   "Un-Imported Reports"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   1050
            TabIndex        =   154
            Top             =   4200
            Width           =   3855
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Tools"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   6360
            TabIndex        =   151
            Top             =   1920
            Width           =   3855
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Height          =   735
            Left            =   8100
            TabIndex        =   149
            Top             =   4920
            Width           =   2505
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "Re Printing Report"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   6420
            TabIndex        =   142
            Top             =   315
            Width           =   3855
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelled Report"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   6420
            TabIndex        =   141
            Top             =   1170
            Width           =   3855
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Audit Report"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   1050
            TabIndex        =   140
            Top             =   2700
            Width           =   1800
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Audit Inquiry"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   1050
            TabIndex        =   139
            Top             =   1920
            Width           =   1890
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Signatories and Headers"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   1050
            TabIndex        =   138
            Top             =   1170
            Width           =   5640
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Reminders"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1035
            TabIndex        =   132
            Top             =   315
            Width           =   3930
         End
      End
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccount_AccountEntriesTemplate_Click()
    If Module_Access(LOGID, "ACCOUNT ENTRIES TEMPLATES", "DATA ENTRY") = False Then Exit Sub
    frmAMISMASTERFILESTemplates.Show
End Sub

Private Sub cmdAdjustment_Client_Click()
    If Module_Access(LOGID, "CLIENT ADJUSTING JOURNAL ENTRIES", "TRANSACTION") = False Then Exit Sub
    JOURNALTYPE = "ADJ"
    On Error Resume Next
    Unload frmAMISJournalEntry
    frmAMISJournalEntry.Show
End Sub

Private Sub cmdAdjustment_ClosingEntries_Click()
    If Module_Access(LOGID, "CLOSING ENTRIES", "TRANSACTION") = False Then Exit Sub
    JOURNALTYPE = "CLO"
    On Error Resume Next
    Unload frmAMISJournalEntry
    frmAMISJournalEntry.Show
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
    JOURNALTYPE = "CSJ"
    On Error Resume Next
    Unload frmAMISJournalEntry
    frmAMISJournalEntry.Show
    'frmAMISCustomerAdjustment.Show
End Sub

Private Sub cmdAdjustment_Proposed_Click()
    If Module_Access(LOGID, "PROPOSED ADJUSTING JOURNAL ENTRIES", "TRANSACTION") = False Then Exit Sub
    JOURNALTYPE = "PDJ"
    On Error Resume Next
    Unload frmAMISJournalEntry
    frmAMISJournalEntry.Show
End Sub

Private Sub cmdVendorCreditMemo_Click()
    If Module_Access(LOGID, "VENDOR ADJUSTMENT", "DATA ENTRY") = False Then Exit Sub
    JOURNALTYPE = "VCJ"
    On Error Resume Next
    Unload frmAMISJournalEntry
    frmAMISJournalEntry.Show
    'frmAMISVendorAdjustment.Show
End Sub

Private Sub cmdVendorDebitMemo_Click()
    If Module_Access(LOGID, "VENDOR ADJUSTMENT", "DATA ENTRY") = False Then Exit Sub
    JOURNALTYPE = "VDJ"
    On Error Resume Next
    Unload frmAMISJournalEntry
    frmAMISJournalEntry.Show
    'frmAMISVendorAdjustment.Show
End Sub

Private Sub cmdAuditInquiry_Click()
    frmInquiry_Audit.Show
End Sub

Private Sub cmdAuditReport_Click()
    frmReportAuditReport.Show
End Sub

Private Sub cmdGJLedgerCodeRunningBalance_Click()
    If Module_Access(LOGID, "GENERAL JOURNAL LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "GJ"
    frmAMISRangeWithAccountCode.Show
    frmAMISRangeWithAccountCode.Caption = "Journal Voucher Ledger Code Running Balance"
End Sub

Private Sub cmdJournal_AP_Click()
    If Module_Access(LOGID, "ACCOUNTS PAYABLE JOURNAL", "TRANSACTION") = False Then Exit Sub
    JOURNALTYPE = "APJ"
    On Error Resume Next
    Unload frmAMISJournalEntry
    frmAMISJournalEntry.Show
End Sub

Private Sub cmdJournal_CashDisburshment_Click()
    If Module_Access(LOGID, "CASH DISBURSEMENT JOURNAL", "TRANSACTION") = False Then Exit Sub
    JOURNALTYPE = "CDJ"
    On Error Resume Next
    Unload frmAMISJournalEntry
    frmAMISJournalEntry.Show
End Sub

Private Sub cmdJournal_CRJ_Click()
    If Module_Access(LOGID, "CASH RECEIPTS JOURNAL", "TRANSACTION") = False Then Exit Sub
    JOURNALTYPE = "CRJ"
    On Error Resume Next
    Unload frmAMISJournalEntry
    frmAMISJournalEntry.Show
End Sub

Private Sub cmdJournal_General_Click()
    If Module_Access(LOGID, "GENERAL JOURNAL", "TRANSACTION") = False Then Exit Sub
    JOURNALTYPE = "GJ"
    On Error Resume Next
    Unload frmAMISJournalEntry
    frmAMISJournalEntry.Show
End Sub

Private Sub cmdJournal_Sales_Click()
    If Module_Access(LOGID, "SALES JOURNAL", "TRANSACTION") = False Then Exit Sub
    JOURNALTYPE = "SJ"
    On Error Resume Next
    Unload frmAMISJournalEntry
    frmAMISJournalEntry.Show
End Sub

Private Sub cmdJournalVoucherSummary_Click()
    If Module_Access(LOGID, "GENERAL JOURNAL SUMMARY", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "JVS"
    frmAMISRange.Show
End Sub

Private Sub cmdLedger_Account_Click()
    If Module_Access(LOGID, "ACCOUNT GENERAL LEDGER", "INQUIRY") = False Then Exit Sub
    frmAMISLEDGERAccounts.Show
End Sub

Private Sub cmdLedger_Customer_Click()
    If Module_Access(LOGID, "CUSTOMER A/R LEDGER", "INQUIRY") = False Then Exit Sub
    CUST_LEDGER_TYPE = "ARLEDGER"
    frmAMISLEDGERCustomers.Show
End Sub

Private Sub cmdLedger_VendorSubisdy_Click()
    If Module_Access(LOGID, "VENDOR SUBSIDIARY LEDGER", "INQUIRY") = False Then Exit Sub
    frmAMISLEDGERVendors.Show
End Sub

Private Sub cmdOpeningBalance_Accounts_Click()
    If Module_Access(LOGID, "ACCOUNT OPENING BALANCE", "TRANSACTION") = False Then Exit Sub
    JOURNALTYPE = "OPB"
    On Error Resume Next
    Unload frmAMISJournalEntry
    frmAMISJournalEntry.Show
End Sub

Private Sub cmdJournalDRJ_Click()
    If Module_Access(LOGID, "DEPOSITED RECEIPTS JOURNAL", "TRANSACTION") = False Then Exit Sub
    JOURNALTYPE = "DRJ"
    On Error Resume Next
    Unload frmAMISDRJJournalEntry
    frmAMISDRJJournalEntry.Show
End Sub

Private Sub cmdCashDisbursementJournal_Click()
    If Module_Access(LOGID, "CASH DISBURSEMENT JOURNAL", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "CDJ"
    frmAMISRangeWithSummary.Show
    frmAMISRangeWithSummary.Caption = "Cash Disbursement Journal"
End Sub

Private Sub cmdCashReceiptsJournal_Click()
    If Module_Access(LOGID, "CASH RECEIPTS JOURNAL", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "CRJ"
    frmAMISRangeWithSummary.Show
    frmAMISRangeWithSummary.Caption = "Cash Receipts Journal"
End Sub

Private Sub cmdCheckRegister_Click()
    If Module_Access(LOGID, "REGISTER REPORT", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "CHECK_REGISTER"
    frmAMISRange.Show
    frmAMISRange.Caption = "Check Registers"
    DoEvents
End Sub

Private Sub cmdOpeningBalance_Customer_Click()
    'AXP-07082007-000001
    If Module_Access(LOGID, "CUSTOMER OPENING BALANCE", "DATA ENTRY") = False Then Exit Sub
    On Error Resume Next
    JOURNALTYPE = "COB"
    frmAMISCustomerAROpening.Show

End Sub

Private Sub cmdAccount_AccountClassification_Click()
    If Module_Access(LOGID, "ACCOUNT CLASSIFICATION", "DATA ENTRY") = False Then Exit Sub
    frmAMISFILESHeader.Show
End Sub

Private Sub cmdAccount_AccountSubTotals_Click()
    If Module_Access(LOGID, "ACCOUNT SUB TOTALS", "DATA ENTRY") = False Then Exit Sub
    frmAMISFILESTitleCode.Show
End Sub

Private Sub cmdAccount_DeaprtmentCodes_Click()
    If Module_Access(LOGID, "DEPARTMENT CODES", "DATA ENTRY") = False Then Exit Sub
    frmAMISFILESDepartment.Show
End Sub

Private Sub cmdAccount_ExtendedClassification_Click()
    If Module_Access(LOGID, "EXTENDED CLASSIFICATION", "DATA ENTRY") = False Then Exit Sub
    frmAMISFILESSubHeader.Show
End Sub

Private Sub cmdAccount_ChartOfAccounts_Click()
    If Module_Access(LOGID, "CHART OF ACCOUNTS", "DATA ENTRY") = False Then Exit Sub
    frmAMISFILESChartOfAccount.Show
End Sub

Private Sub cmdAccount_AccountTypes_Click()
    If Module_Access(LOGID, "ACCOUNT TYPES", "DATA ENTRY") = False Then Exit Sub
    frmAMISFILESAccType.Show
End Sub

Private Sub cmdFinancialStatements_Click()
    If Module_Access(LOGID, "FINANCIAL STATEMENTS", "REPORTS") = False Then Exit Sub
    frmAMISFinancialStatements.Show
End Sub

Private Sub cmdInvoiceRegister_Click()
    If Module_Access(LOGID, "REGISTER REPORT", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "INV_REGISTER"
    frmAMISRange.Show
    frmAMISRange.Caption = "Invoices Registers"
    DoEvents
End Sub

Private Sub cmdLedgerCodeRunningBalance_AP_Click()
    'If Module_Access(LOGID, "ACCOUNTS LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "APJ"
    frmAMISRangeWithAccountCode.Show
    frmAMISRangeWithAccountCode.Caption = "ACCOUNTS Disbursement Ledger Code Running Balance"
End Sub

Private Sub cmdLedgerCodeRunningBalance_CD_Click()
    If Module_Access(LOGID, "CASH DISBURSEMENT LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "CDJ"
    frmAMISRangeWithAccountCode.Show
    frmAMISRangeWithAccountCode.Caption = "Cash Disbursement Ledger Code Running Balance"
End Sub

Private Sub cmdLedgerCodeRunningBalance_CR_Click()
    If Module_Access(LOGID, "CASH RECEIPTS LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "CRJ"
    frmAMISRangeWithAccountCode.Show
    frmAMISRangeWithAccountCode.Caption = "Cash Receipts Ledger Code Running Balance"
End Sub

Private Sub cmdLedgerCodeRunningBalance_LCRB_Click()
    If Module_Access(LOGID, "SALES LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "SJ"
    frmAMISRangeWithAccountCode.Show
    frmAMISRangeWithAccountCode.Caption = "Sales Journal Ledger Code Running Balance"
End Sub

Private Sub cmdOR_Register_Click()
    If Module_Access(LOGID, "REGISTER REPORT", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "OR_REGISTER"
    frmAMISRange.Show
    frmAMISRange.Caption = "O.R. Registers"
    DoEvents
End Sub

Private Sub cmdReceivingReportRegister_Click()
    If Module_Access(LOGID, "RECEIVING REPORT REGISTER", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "REC_REGISTER"
    frmAMISDetailBySupplierWithAccountCode.Show
    frmAMISDetailBySupplierWithAccountCode.Caption = "Receiving Report Registers"
End Sub

Private Sub cmdSalesbyInvoiceType_Click()
    If Module_Access(LOGID, "SALES BY INVOICE TYPE", "REPORTS") = False Then Exit Sub
    frmAMIS_SalesbyInvoiceType.Show
End Sub

Private Sub cmdSalesJournal_Click()
    If Module_Access(LOGID, "SALES JOURNAL", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "SJ"
    frmAMISRangeWithSummary.Show
    frmAMISRangeWithSummary.Caption = "Sales Journal"
End Sub

Private Sub cmdScheduleOfAccountsPayable_Click()
    If Module_Access(LOGID, "SCHEDULE OF ACCOUNTS PAYABLE", "REPORTS") = False Then Exit Sub
    frmAMISAPSchedReport.Show
End Sub

Private Sub cmdScheduleOfAccountsReceivable_Click()
      If Module_Access(LOGID, "ACCOUNTS RECEIVABLE AGING REPORT", "REPORTS") = False Then Exit Sub
    Report_AR = "AGING"
    frmAMISARSchedReport.Show
End Sub

Private Sub cmdScheduleOfAdjustments_Click()
    If Module_Access(LOGID, "SCHEDULE OF ADJUSTMENTS", "REPORTS") = False Then Exit Sub
    frmAMISSchedAdjust.Show
End Sub

Private Sub cmdTable_ATCCode_Click()
    If Module_Access(LOGID, "ATC CODES", "DATA ENTRY") = False Then Exit Sub
    frmAMISMASTERFILEATC.Show
End Sub

Private Sub cmdTable_Customer_Click()
    If Module_Access(LOGID, "CUSTOMER", "DATA ENTRY") = False Then Exit Sub
    frmAllCustomer.Show
End Sub

Private Sub cmdTables_Bank_Click()
    If Module_Access(LOGID, "BANKS", "DATA ENTRY") = False Then Exit Sub
    frmAMISMASTERFILEBanks.Show
End Sub

Private Sub cmdTables_InvoiceTypes_Click()
    If Module_Access(LOGID, "INVOICE TYPES", "DATA ENTRY") = False Then Exit Sub
    frmAMISMASTERFILEInvoiceType.Show
End Sub

Private Sub cmdTables_TermsOfPayment_Click()
    If Module_Access(LOGID, "TERMS OF PAYMENT", "DATA ENTRY") = False Then Exit Sub
    frmAMISMASTERFILEPayTerm.Show
End Sub

Private Sub cmdTrialBalance_Click()
    If Module_Access(LOGID, "FINANCIAL STATMENT TRIAL BALANCE", "REPORTS") = False Then Exit Sub
    frmAMISTrialBalance.Show
End Sub

Private Sub cmdUnused_OR_Click()
    If Module_Access(LOGID, "UNUSED OR", "REPORTS") = False Then Exit Sub
    frmAMISProcessUnusedOR.Show
End Sub

Private Sub cmdUnusedInvoices_Click()
    If Module_Access(LOGID, "UNUSED INVOICES", "REPORTS") = False Then Exit Sub
    frmAMISProcessUnusedInvoices.Show
End Sub

Private Sub cmdVendor_Click()


    If Module_Access(LOGID, "VENDORS", "DATA ENTRY") = False Then Exit Sub
    frmAMISMASTERFILEVendor.Show
End Sub

Private Sub cmdOpeningBalance_Vendor_Click()
    If Module_Access(LOGID, "VENDOR OPENING BALANCE", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    JOURNALTYPE = "VPJ"
    frmAMISVendorAPOpening.Show

End Sub

Private Sub cmdWork_Sheet_Click()
    If Module_Access(LOGID, "WORKSHEET", "REPORTS") = False Then Exit Sub
    frmAMISWorkSheet.Show
End Sub

Private Sub Command1_Click()

    frmCancelledReport.Show



End Sub

Private Sub Command10_Click()
    frmSMIS_Log_Reminder.Show
End Sub

Private Sub cmdAccountDetailbyCustomer_Click()
    If Module_Access(LOGID, "ACCOUNTS DETAIL BY CUSTOMER", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "SJ"
    frmAMISDetailBySupplierWithAccountCode.Show
    frmAMISDetailBySupplierWithAccountCode.Caption = "Accounts Detail Report By Customer"
End Sub

Private Sub cmdAccountDetailbySupplier_Click()
    If Module_Access(LOGID, "ACCOUNTS DETAIL BY SUPPLIERS", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "APJ"
    frmAMISDetailBySupplierWithAccountCode.Show
    frmAMISDetailBySupplierWithAccountCode.Caption = "Accounts Payable Detail Report By Supplier"
End Sub

Private Sub cmdAccountsPayableAgingReport_Click()
    If Module_Access(LOGID, "ACCOUNTS PAYABLE AGING REPORT", "REPORTS") = False Then Exit Sub
    REPORT_AP = "AGING"
    'frmAMISDueReport.Show
    frmAPschedulestandard.Show
End Sub

Private Sub cmdAccountsPayableDueReport_Click()
    If Module_Access(LOGID, "ACCOUNTS PAYABLE DUE REPORT", "REPORTS") = False Then Exit Sub
    REPORT_AP = "SCHED"
    frmAMISDueReport.Show
End Sub

Private Sub cmdAccountsPayableJournal_Click()
    If Module_Access(LOGID, "ACCOUNTS PAYABLE JOURNAL", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "APJ"
    frmAMISRangeWithSummary.Show
    frmAMISRangeWithSummary.Caption = "Accounts Payable Journal"
End Sub

Private Sub cmdAccountsReceivableAgingReport_Click()
    If Module_Access(LOGID, "ACCOUNTS RECEIVABLE AGING REPORT", "REPORTS") = False Then Exit Sub
    Report_AR = "AGING"
    frmAMISARSchedReport.Show
End Sub

Private Sub Command2_Click()
    frmOpeningBalanceReport.Show
End Sub

Private Sub Command27_Click()
    If Module_Access(LOGID, "SYSTEM SETUP", "SYSTEM") = False Then Exit Sub
    frmAMISProfile.Show
End Sub

Private Sub cmdCustomerCreditMemo_Click()
    If Module_Access(LOGID, "CUSTOMER ADJUSTMENT", "DATA ENTRY") = False Then Exit Sub
    JOURNALTYPE = "CCM"
    On Error Resume Next
    Unload frmAMISJournalEntry
    frmAMISJournalEntry.Show
End Sub

Private Sub Command3_Click()
   frmAMISbanksOpening.Show
End Sub

Private Sub Command4_Click()
    frmReprintReport.Show
End Sub



 
 
Private Sub Command5_Click()

End Sub

Private Sub Command6_Click()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    frmTransactionStatus.Show
End Sub

Private Sub Command7_Click()
If Module_Access(LOGID, "BANK RECONCILIATION", "DATA ENTRY") = False Then Exit Sub
            frmReconcileAccount.Show
End Sub

Private Sub Command8_Click()
    Screen.MousePointer = 11
        frmAMIS_UNAPPLIED_PAYMENT.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command9_Click()
    Screen.MousePointer = 11
        frmAMIS_UniportedReports.Show
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
    Dim rs As New ADODB.Recordset
    Set rs = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_HD where jtype='APJ' and voucherno='" & XXX & "'")
    If Not rs.EOF And Not rs.BOF Then
        
    End If
    Set rs = Nothing
End Sub
Sub InitSalesJournal()
    
End Sub

 
