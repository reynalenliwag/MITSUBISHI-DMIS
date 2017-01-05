VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO301B~1.OCX"
Begin VB.Form frmMainMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AMIS Main Menu"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
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
   ScaleWidth      =   10410
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
      Appearance      =   5
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   120
      PaintManager.MinTabWidth=   100
      ItemCount       =   6
      Item(0).Caption =   "Main Modules"
      Item(0).ControlCount=   19
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
      Item(0).Control(18)=   "picDeposit"
      Item(1).Caption =   "Tables"
      Item(1).ControlCount=   14
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
      Item(1).Control(12)=   "Label64"
      Item(1).Control(13)=   "cmdBankDeposits"
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
      Item(5).Caption =   "Dev Tools"
      Item(5).ControlCount=   8
      Item(5).Control(0)=   "cmdUTool(1)"
      Item(5).Control(1)=   "cmdTallyTool(0)"
      Item(5).Control(2)=   "Label77"
      Item(5).Control(3)=   "Label78"
      Item(5).Control(4)=   "Label79"
      Item(5).Control(5)=   "cmdTBUTool(0)"
      Item(5).Control(6)=   "cmdLogoTool(1)"
      Item(5).Control(7)=   "Label80"
      Begin VB.CommandButton cmdLogoTool 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Logo Tool"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   1
         Left            =   -64360
         MouseIcon       =   "frmAMISMainMenu.frx":15162
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":152B4
         Style           =   1  'Graphical
         TabIndex        =   175
         Tag             =   "1052"
         ToolTipText     =   "Logo Tool"
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTBUTool 
         BackColor       =   &H00FFFFFF&
         Caption         =   "U-Tool"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":159B6
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":15B08
         Style           =   1  'Graphical
         TabIndex        =   173
         Tag             =   "1052"
         ToolTipText     =   "Trial Balance Uploading Tool"
         Top             =   1740
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTallyTool 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tally Tool"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1623E
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":16390
         Style           =   1  'Graphical
         TabIndex        =   170
         Tag             =   "1052"
         ToolTipText     =   "Developer's Tally Tool"
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdUTool 
         BackColor       =   &H00FFFFFF&
         Caption         =   "U-Tool"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   1
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":16A92
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":16BE4
         Style           =   1  'Graphical
         TabIndex        =   169
         Tag             =   "1052"
         ToolTipText     =   "Schedule Uploading Tool"
         Top             =   2580
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdBankDeposits 
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
         Height          =   645
         Left            =   -65200
         MouseIcon       =   "frmAMISMainMenu.frx":1731A
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1746C
         Style           =   1  'Graphical
         TabIndex        =   156
         Tag             =   "1029"
         ToolTipText     =   "View Bank Master Files"
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox picDeposit 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   5940
         ScaleHeight     =   795
         ScaleWidth      =   3615
         TabIndex        =   153
         Top             =   3360
         Width           =   3615
         Begin VB.CommandButton cmdCustomersDeposit 
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
            Height          =   645
            Left            =   30
            MouseIcon       =   "frmAMISMainMenu.frx":17B47
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":17C99
            Style           =   1  'Graphical
            TabIndex        =   154
            Tag             =   "1027"
            ToolTipText     =   "Customer's Deposit Ledger"
            Top             =   30
            Width           =   720
         End
         Begin VB.Label lblDeposit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer's Deposit Ledger"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   870
            TabIndex        =   155
            Top             =   225
            Width           =   2445
         End
      End
      Begin VB.CommandButton cmdJournalDRJ 
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
         Height          =   645
         Left            =   480
         MouseIcon       =   "frmAMISMainMenu.frx":18300
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":18452
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "1047"
         ToolTipText     =   "View Cash Receipts Journal"
         Top             =   4320
         Width           =   720
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   2325
         Left            =   -65200
         ScaleHeight     =   2325
         ScaleWidth      =   5385
         TabIndex        =   142
         Top             =   3210
         Visible         =   0   'False
         Width           =   5385
         Begin VB.CommandButton cmdTranType 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Trantype"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   1080
            MouseIcon       =   "frmAMISMainMenu.frx":18CBD
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":18E0F
            Style           =   1  'Graphical
            TabIndex        =   149
            Tag             =   "1052"
            ToolTipText     =   "Trantype"
            Top             =   1590
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton Command11 
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
            Height          =   645
            Left            =   1080
            MouseIcon       =   "frmAMISMainMenu.frx":19E91
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":19FE3
            Style           =   1  'Graphical
            TabIndex        =   144
            Tag             =   "1052"
            ToolTipText     =   "View Vendors Subsidiary Ledger"
            Top             =   30
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton Command12 
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
            Height          =   645
            Left            =   1080
            MouseIcon       =   "frmAMISMainMenu.frx":1A913
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":1AA65
            Style           =   1  'Graphical
            TabIndex        =   143
            Tag             =   "1052"
            ToolTipText     =   "Importing Templates"
            Top             =   810
            Width           =   720
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Un-Imported Transactions"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   146
            Top             =   240
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Importing Templates"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   145
            Top             =   1020
            Width           =   1905
         End
      End
      Begin VB.CommandButton Command8 
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
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1B395
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1B4E7
         Style           =   1  'Graphical
         TabIndex        =   130
         Tag             =   "1052"
         ToolTipText     =   "View Vendors Subsidiary Ledger"
         Top             =   4020
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Command9 
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
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1BE17
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1BF69
         Style           =   1  'Graphical
         TabIndex        =   129
         Tag             =   "1052"
         ToolTipText     =   "View Vendors Subsidiary Ledger"
         Top             =   4800
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Command6 
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
         Height          =   645
         Left            =   -64150
         MouseIcon       =   "frmAMISMainMenu.frx":1C899
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1C9EB
         Style           =   1  'Graphical
         TabIndex        =   128
         Tag             =   "1052"
         ToolTipText     =   "View Data Tools"
         Top             =   2460
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   -64150
         MouseIcon       =   "frmAMISMainMenu.frx":1D31B
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1D46D
         Style           =   1  'Graphical
         TabIndex        =   127
         ToolTipText     =   "View Unused O.R."
         Top             =   1680
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   -64150
         MouseIcon       =   "frmAMISMainMenu.frx":1DC3D
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1DD8F
         Style           =   1  'Graphical
         TabIndex        =   126
         ToolTipText     =   "View O.R. Register"
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdAuditReport 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1E50D
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1E65F
         Style           =   1  'Graphical
         TabIndex        =   125
         ToolTipText     =   "View Signatories and Headers"
         Top             =   3240
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdAuditInquiry 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1EAA1
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1EBF3
         Style           =   1  'Graphical
         TabIndex        =   124
         ToolTipText     =   "View Signatories and Headers"
         Top             =   2460
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Command27 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1F035
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1F187
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "View Signatories and Headers"
         Top             =   1680
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1F5C9
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":1F71B
         Style           =   1  'Graphical
         TabIndex        =   122
         Tag             =   "1102"
         ToolTipText     =   "View Reminders"
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTable_Customer 
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
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":1FF96
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":200E8
         Style           =   1  'Graphical
         TabIndex        =   24
         Tag             =   "1027"
         ToolTipText     =   "View Customer Master Files"
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdVendor 
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
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":2074F
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":208A1
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "1028"
         ToolTipText     =   "View"
         Top             =   1710
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTables_Bank 
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
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":20FB8
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":2110A
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "1029"
         ToolTipText     =   "View Bank Master Files"
         Top             =   2520
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTables_TermsOfPayment 
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
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":217E5
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":21937
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "1031"
         ToolTipText     =   "View Terms Of Payment Master File"
         Top             =   4125
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTables_InvoiceTypes 
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
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":21FD0
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":22122
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "1030"
         ToolTipText     =   "View Invoice Type Master Files"
         Top             =   3315
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdTable_ATCCode 
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
         Height          =   645
         Left            =   -69520
         MouseIcon       =   "frmAMISMainMenu.frx":227D4
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":22926
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "1032"
         ToolTipText     =   "View ATC Code Master File"
         Top             =   4935
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdJournal_CashDisburshment 
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
         Height          =   645
         Left            =   480
         MouseIcon       =   "frmAMISMainMenu.frx":22FD9
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":2312B
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "1045"
         ToolTipText     =   "View Cash Disbursement Journal"
         Top             =   1740
         Width           =   720
      End
      Begin VB.CommandButton cmdLedger_VendorSubisdy 
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
         Height          =   645
         Left            =   5970
         MouseIcon       =   "frmAMISMainMenu.frx":23B56
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":23CA8
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "1052"
         ToolTipText     =   "View Vendors Subsidiary Ledger"
         Top             =   2580
         Width           =   720
      End
      Begin VB.CommandButton cmdLedger_Customer 
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
         Height          =   645
         Left            =   5970
         MouseIcon       =   "frmAMISMainMenu.frx":245D8
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":2472A
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "1050"
         ToolTipText     =   "View Customers A/R Ledger"
         Top             =   1740
         Width           =   720
      End
      Begin VB.CommandButton cmdLedger_Account 
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
         Height          =   645
         Left            =   6000
         MouseIcon       =   "frmAMISMainMenu.frx":2503D
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":2518F
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "1049"
         ToolTipText     =   "View Accounts General Ledger"
         Top             =   900
         Width           =   720
      End
      Begin VB.CommandButton cmdJournal_CRJ 
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
         Height          =   645
         Left            =   480
         MouseIcon       =   "frmAMISMainMenu.frx":25A71
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":25BC3
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "1047"
         ToolTipText     =   "View Cash Receipts Journal"
         Top             =   3420
         Width           =   720
      End
      Begin VB.CommandButton cmdJournal_General 
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
         Height          =   645
         Left            =   480
         MouseIcon       =   "frmAMISMainMenu.frx":2642E
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":26580
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "1048"
         ToolTipText     =   "View General Journal"
         Top             =   5160
         Width           =   720
      End
      Begin VB.CommandButton cmdJournal_Sales 
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
         Height          =   645
         Left            =   480
         MouseIcon       =   "frmAMISMainMenu.frx":26E85
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":26FD7
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "1046"
         ToolTipText     =   "View Sales Journal"
         Top             =   2580
         Width           =   720
      End
      Begin VB.CommandButton cmdJournal_AP 
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
         Height          =   645
         Left            =   480
         MouseIcon       =   "frmAMISMainMenu.frx":277B5
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISMainMenu.frx":27907
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "1044"
         ToolTipText     =   "View Accounts Payable Journal"
         Top             =   900
         Width           =   720
      End
      Begin XtremeSuiteControls.TabControl TabControl4 
         Height          =   5640
         Left            =   -69880
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   10680
         _Version        =   655364
         _ExtentX        =   18838
         _ExtentY        =   9948
         _StockProps     =   64
         Appearance      =   5
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
         Item(2).ControlCount=   16
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
         Item(2).Control(15)=   "picCreditMemo"
         Begin VB.PictureBox picCreditMemo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   -63880
            ScaleHeight     =   855
            ScaleWidth      =   3495
            TabIndex        =   164
            Top             =   1560
            Visible         =   0   'False
            Width           =   3495
            Begin VB.CommandButton cmdInvoiceCM 
               BackColor       =   &H00FFFFFF&
               Height          =   615
               Left            =   0
               MaskColor       =   &H00404040&
               MouseIcon       =   "frmAMISMainMenu.frx":2813B
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":2828D
               Style           =   1  'Graphical
               TabIndex        =   165
               ToolTipText     =   "Invoice Credit Memo"
               Top             =   0
               Width           =   735
            End
            Begin VB.Label Creditmemo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Credit Memo Register"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   915
               TabIndex        =   166
               Top             =   120
               Width           =   1980
            End
         End
         Begin VB.CommandButton cmdVehicleSales 
            BackColor       =   &H00FFFFFF&
            Height          =   645
            Left            =   4740
            MouseIcon       =   "frmAMISMainMenu.frx":28517
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":28669
            Style           =   1  'Graphical
            TabIndex        =   147
            Tag             =   "1140"
            ToolTipText     =   "Vehicle Sales - Account Code Set-up"
            Top             =   2670
            Width           =   720
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   165
            Left            =   -69970
            ScaleHeight     =   165
            ScaleWidth      =   10095
            TabIndex        =   51
            Top             =   3750
            Visible         =   0   'False
            Width           =   10095
            Begin VB.Line Line 
               BorderColor     =   &H00000000&
               X1              =   -240
               X2              =   10410
               Y1              =   120
               Y2              =   120
            End
         End
         Begin VB.CommandButton cmdAdjustment_Proposed 
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
            Height          =   645
            Left            =   -69580
            MouseIcon       =   "frmAMISMainMenu.frx":28CEF
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":28E41
            Style           =   1  'Graphical
            TabIndex        =   50
            Tag             =   "1037"
            ToolTipText     =   "View Proposed Adjusting Journal Entries"
            Top             =   4755
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdAdjustment_Client 
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
            Height          =   645
            Left            =   -69580
            MouseIcon       =   "frmAMISMainMenu.frx":29528
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2967A
            Style           =   1  'Graphical
            TabIndex        =   49
            Tag             =   "1036"
            ToolTipText     =   "View Client Adjusting Journal Entries"
            Top             =   3960
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdCustomerDebitMemo 
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
            Height          =   645
            Left            =   -69580
            MouseIcon       =   "frmAMISMainMenu.frx":29D5D
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":29EAF
            Style           =   1  'Graphical
            TabIndex        =   48
            Tag             =   "1038"
            ToolTipText     =   "View Customer Adjustments"
            Top             =   750
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdAdjustment_ClosingEntries 
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
            Height          =   645
            Left            =   -63850
            MouseIcon       =   "frmAMISMainMenu.frx":2A534
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2A686
            Style           =   1  'Graphical
            TabIndex        =   47
            Tag             =   "1040"
            ToolTipText     =   "View Closing Entries"
            Top             =   690
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdVendorDebitMemo 
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
            Height          =   645
            Left            =   -69580
            MouseIcon       =   "frmAMISMainMenu.frx":2ACC9
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2AE1B
            Style           =   1  'Graphical
            TabIndex        =   46
            Tag             =   "1039"
            ToolTipText     =   "View Vendor Adjustments"
            Top             =   2280
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdCustomerCreditMemo 
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
            Height          =   645
            Left            =   -69580
            MouseIcon       =   "frmAMISMainMenu.frx":2B486
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2B5D8
            Style           =   1  'Graphical
            TabIndex        =   45
            Tag             =   "1038"
            ToolTipText     =   "View Customer Adjustments"
            Top             =   1530
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdVendorCreditMemo 
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
            Height          =   645
            Left            =   -69580
            MouseIcon       =   "frmAMISMainMenu.frx":2BC5D
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2BDAF
            Style           =   1  'Graphical
            TabIndex        =   44
            Tag             =   "1039"
            ToolTipText     =   "View Vendor Adjustments"
            Top             =   3030
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdOpeningBalance_Vendor 
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
            Height          =   645
            Left            =   -69520
            MouseIcon       =   "frmAMISMainMenu.frx":2C41A
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2C56C
            Style           =   1  'Graphical
            TabIndex        =   43
            Tag             =   "1035"
            ToolTipText     =   "View Vendor Opening Balance"
            Top             =   2640
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdOpeningBalance_Customer 
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
            Height          =   645
            Left            =   -69520
            MouseIcon       =   "frmAMISMainMenu.frx":2CD1E
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2CE70
            Style           =   1  'Graphical
            TabIndex        =   42
            Tag             =   "1034"
            ToolTipText     =   "View Customer Opening Balance"
            Top             =   1740
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdOpeningBalance_Accounts 
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
            Height          =   645
            Left            =   -69490
            MouseIcon       =   "frmAMISMainMenu.frx":2D64B
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2D79D
            Style           =   1  'Graphical
            TabIndex        =   41
            Tag             =   "1033"
            ToolTipText     =   "View Accounts Opening Balance"
            Top             =   900
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdOpeningReport 
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
            Height          =   645
            Left            =   -69490
            MouseIcon       =   "frmAMISMainMenu.frx":2DE8E
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2DFE0
            Style           =   1  'Graphical
            TabIndex        =   40
            Tag             =   "1034"
            ToolTipText     =   "View Customer Opening Balance"
            Top             =   3465
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton Command3 
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
            Height          =   645
            Left            =   -69460
            MouseIcon       =   "frmAMISMainMenu.frx":2E7BB
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2E90D
            Style           =   1  'Graphical
            TabIndex        =   39
            Tag             =   "1045"
            ToolTipText     =   "View Cash Disbursement Journal"
            Top             =   4275
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdAccount_ChartOfAccounts 
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
            Height          =   645
            Left            =   480
            MouseIcon       =   "frmAMISMainMenu.frx":2F338
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2F48A
            Style           =   1  'Graphical
            TabIndex        =   38
            Tag             =   "1020"
            ToolTipText     =   "View Chart Of Accounts"
            Top             =   900
            Width           =   720
         End
         Begin VB.CommandButton cmdAccount_AccountTypes 
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
            Height          =   645
            Left            =   480
            MouseIcon       =   "frmAMISMainMenu.frx":2FABE
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":2FC10
            Style           =   1  'Graphical
            TabIndex        =   37
            Tag             =   "1021"
            ToolTipText     =   "View Account Types"
            Top             =   1785
            Width           =   720
         End
         Begin VB.CommandButton cmdAccount_AccountClassification 
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
            Height          =   645
            Left            =   480
            MouseIcon       =   "frmAMISMainMenu.frx":3028F
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":303E1
            Style           =   1  'Graphical
            TabIndex        =   36
            Tag             =   "1022"
            ToolTipText     =   "View Account Classification"
            Top             =   2670
            Width           =   720
         End
         Begin VB.CommandButton cmdAccount_AccountSubTotals 
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
            Height          =   645
            Left            =   480
            MouseIcon       =   "frmAMISMainMenu.frx":30A4D
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":30B9F
            Style           =   1  'Graphical
            TabIndex        =   35
            Tag             =   "1024"
            ToolTipText     =   "View Account Sub-Totals"
            Top             =   4425
            Width           =   720
         End
         Begin VB.CommandButton cmdAccount_ExtendedClassification 
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
            Height          =   645
            Left            =   480
            MouseIcon       =   "frmAMISMainMenu.frx":3124B
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":3139D
            Style           =   1  'Graphical
            TabIndex        =   34
            Tag             =   "1023"
            ToolTipText     =   "View Extended Classification"
            Top             =   3540
            Width           =   720
         End
         Begin VB.CommandButton cmdAccount_DeaprtmentCodes 
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
            Height          =   645
            Left            =   4725
            MouseIcon       =   "frmAMISMainMenu.frx":31B2D
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":31C7F
            Style           =   1  'Graphical
            TabIndex        =   33
            Tag             =   "1025"
            ToolTipText     =   "View Department Codes"
            Top             =   900
            Width           =   720
         End
         Begin VB.CommandButton cmdAccount_AccountEntriesTemplate 
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
            Height          =   645
            Left            =   4725
            MouseIcon       =   "frmAMISMainMenu.frx":323E0
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":32532
            Style           =   1  'Graphical
            TabIndex        =   32
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
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5580
            TabIndex        =   148
            Top             =   2880
            Width           =   3270
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proposed Adjusting Journal Entries"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -68740
            TabIndex        =   70
            Top             =   4965
            Visible         =   0   'False
            Width           =   3195
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Client Adjusting Journal Entries"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -68740
            TabIndex        =   69
            Top             =   4200
            Visible         =   0   'False
            Width           =   2835
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Debit Memo"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -68710
            TabIndex        =   68
            Top             =   2460
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Closing Entries"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -62995
            TabIndex        =   67
            Top             =   885
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Debit Memo"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -68710
            TabIndex        =   66
            Top             =   990
            Visible         =   0   'False
            Width           =   2070
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Credit Memo"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -68725
            TabIndex        =   65
            Top             =   1680
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Credit Memo"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -68710
            TabIndex        =   64
            Top             =   3195
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Opening Balance"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -68590
            TabIndex        =   63
            Top             =   2760
            Visible         =   0   'False
            Width           =   2250
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Opening Balance"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -68605
            TabIndex        =   62
            Top             =   1905
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Accounts Opening Balance"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -68590
            TabIndex        =   61
            Top             =   1140
            Visible         =   0   'False
            Width           =   2430
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Opening Balance Report"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -68590
            TabIndex        =   60
            Top             =   3615
            Visible         =   0   'False
            Width           =   2205
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Opening Balance"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -68560
            TabIndex        =   59
            Top             =   4425
            Visible         =   0   'False
            Width           =   2040
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Entries Templates"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5595
            TabIndex        =   58
            Top             =   1905
            Width           =   2415
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Department Codes"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5595
            TabIndex        =   57
            Top             =   1035
            Width           =   1710
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Extended Classification"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1350
            TabIndex        =   56
            Top             =   3675
            Width           =   2070
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Classification"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1350
            TabIndex        =   55
            Top             =   2805
            Width           =   1965
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Sub-Totals"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1350
            TabIndex        =   54
            Top             =   4560
            Width           =   1770
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Types"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1350
            TabIndex        =   53
            Top             =   1950
            Width           =   1335
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chart of Accounts"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1320
            TabIndex        =   52
            Top             =   1155
            Width           =   1635
         End
      End
      Begin XtremeSuiteControls.TabControl TabControl2 
         Height          =   5730
         Left            =   -69970
         TabIndex        =   71
         Top             =   570
         Visible         =   0   'False
         Width           =   11490
         _Version        =   655364
         _ExtentX        =   20267
         _ExtentY        =   10107
         _StockProps     =   64
         Appearance      =   5
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
         Item(1).ControlCount=   11
         Item(1).Control(0)=   "cmdTrialBalance"
         Item(1).Control(1)=   "cmdWork_Sheet"
         Item(1).Control(2)=   "cmdFinancialStatements"
         Item(1).Control(3)=   "cmdWorkSheet"
         Item(1).Control(4)=   "Label10"
         Item(1).Control(5)=   "Label17"
         Item(1).Control(6)=   "Label50"
         Item(1).Control(7)=   "cmdWitholdingTax"
         Item(1).Control(8)=   "Picture3"
         Item(1).Control(9)=   "cmdWTExpanded"
         Item(1).Control(10)=   "Label76"
         Begin VB.CommandButton cmdWTExpanded 
            BackColor       =   &H00FFFFFF&
            Height          =   645
            Left            =   -65200
            Picture         =   "frmAMISMainMenu.frx":32BED
            Style           =   1  'Graphical
            TabIndex        =   167
            Top             =   900
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   795
            Left            =   -69580
            ScaleHeight     =   795
            ScaleWidth      =   3285
            TabIndex        =   150
            Top             =   4170
            Visible         =   0   'False
            Width           =   3285
            Begin VB.CommandButton cmdScheduleOfAdjustments 
               Height          =   645
               Left            =   0
               MouseIcon       =   "frmAMISMainMenu.frx":32DB7
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":32F09
               Style           =   1  'Graphical
               TabIndex        =   151
               ToolTipText     =   "View Schedule Of Adjustments"
               Top             =   0
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Schedule Of Adjustments"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   855
               TabIndex        =   152
               Top             =   150
               Visible         =   0   'False
               Width           =   2295
            End
         End
         Begin VB.CommandButton cmdWitholdingTax 
            BackColor       =   &H00FFFFFF&
            Height          =   645
            Left            =   -69550
            MouseIcon       =   "frmAMISMainMenu.frx":336C2
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":33814
            Style           =   1  'Graphical
            TabIndex        =   140
            ToolTipText     =   "View Financial Statements"
            Top             =   3390
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdFinancialStatements 
            BackColor       =   &H00FFFFFF&
            Height          =   645
            Left            =   -69550
            MouseIcon       =   "frmAMISMainMenu.frx":33C5E
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":33DB0
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "View Financial Statements"
            Top             =   2580
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdWork_Sheet 
            BackColor       =   &H00FFFFFF&
            Height          =   645
            Left            =   -69550
            MouseIcon       =   "frmAMISMainMenu.frx":34506
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":34658
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "View Work Sheet"
            Top             =   900
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdTrialBalance 
            BackColor       =   &H00FFFFFF&
            Height          =   645
            Left            =   -69550
            MouseIcon       =   "frmAMISMainMenu.frx":34E5B
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISMainMenu.frx":34FAD
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "View Trial Balance"
            Top             =   1800
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
            TabIndex        =   72
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
            Appearance      =   4
            Color           =   2
            PaintManager.BoldSelected=   -1  'True
            PaintManager.HotTracking=   -1  'True
            PaintManager.ShowIcons=   -1  'True
            PaintManager.FixedTabWidth=   110
            PaintManager.MinTabWidth=   85
            ItemCount       =   5
            SelectedItem    =   2
            Item(0).Caption =   "Accounts Payable"
            Item(0).ControlCount=   12
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
            Item(0).Control(10)=   "cmdBillingDueReport"
            Item(0).Control(11)=   "Label70"
            Item(1).Caption =   "Cash Disbursement"
            Item(1).ControlCount=   6
            Item(1).Control(0)=   "cmdCheckRegister"
            Item(1).Control(1)=   "cmdLedgerCodeRunningBalance_CD"
            Item(1).Control(2)=   "cmdCashDisbursementJournal"
            Item(1).Control(3)=   "Label25"
            Item(1).Control(4)=   "Label26"
            Item(1).Control(5)=   "Label27"
            Item(2).Caption =   "Sales"
            Item(2).ControlCount=   18
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
            Item(2).Control(14)=   "cmdCollectionForecast"
            Item(2).Control(15)=   "Label74"
            Item(2).Control(16)=   "Label75"
            Item(2).Control(17)=   "cmdSalesAnalysis"
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
            Begin VB.CommandButton cmdInvoiceRegister 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   5580
               MouseIcon       =   "frmAMISMainMenu.frx":3553B
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3568D
               Style           =   1  'Graphical
               TabIndex        =   89
               ToolTipText     =   "View Invoice Register"
               Top             =   630
               Width           =   720
            End
            Begin VB.CommandButton cmdSalesAnalysis 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   5550
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":35D27
               Style           =   1  'Graphical
               TabIndex        =   162
               ToolTipText     =   "Vehicle Sales Analysis"
               Top             =   3120
               Width           =   720
            End
            Begin VB.CommandButton cmdCollectionForecast 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   5550
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":36DA9
               Style           =   1  'Graphical
               TabIndex        =   160
               ToolTipText     =   "Collection Forecast Report"
               Top             =   2310
               Width           =   720
            End
            Begin VB.CommandButton cmdBillingDueReport 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   -64600
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":37E2B
               Style           =   1  'Graphical
               TabIndex        =   158
               ToolTipText     =   "Billing Due Report"
               Top             =   1500
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdGJLedgerCodeRunningBalance 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   -69580
               MouseIcon       =   "frmAMISMainMenu.frx":38EAD
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":38FFF
               Style           =   1  'Graphical
               TabIndex        =   97
               ToolTipText     =   "View Ledger Code Running Balance"
               Top             =   1545
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdJournalVoucherSummary 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   -69580
               MouseIcon       =   "frmAMISMainMenu.frx":396B8
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3980A
               Style           =   1  'Graphical
               TabIndex        =   96
               ToolTipText     =   "View General Journal"
               Top             =   720
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdCashReceiptsJournal 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   -69580
               MouseIcon       =   "frmAMISMainMenu.frx":39F88
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3A0DA
               Style           =   1  'Graphical
               TabIndex        =   95
               ToolTipText     =   "View Cash Receipts Journal"
               Top             =   630
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdLedgerCodeRunningBalance_CR 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   -69595
               MouseIcon       =   "frmAMISMainMenu.frx":3A858
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3A9AA
               Style           =   1  'Graphical
               TabIndex        =   94
               ToolTipText     =   "View Ledger Code Running Balance"
               Top             =   1500
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdOR_Register 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   -69595
               MouseIcon       =   "frmAMISMainMenu.frx":3B063
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3B1B5
               Style           =   1  'Graphical
               TabIndex        =   93
               ToolTipText     =   "View O.R. Register"
               Top             =   2340
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdUnused_OR 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   -69610
               MouseIcon       =   "frmAMISMainMenu.frx":3B933
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3BA85
               Style           =   1  'Graphical
               TabIndex        =   92
               ToolTipText     =   "View Unused O.R."
               Top             =   3180
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdUnusedInvoices 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   390
               MouseIcon       =   "frmAMISMainMenu.frx":3C255
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3C3A7
               Style           =   1  'Graphical
               TabIndex        =   91
               ToolTipText     =   "View Unused Invoices"
               Top             =   3960
               Width           =   720
            End
            Begin VB.CommandButton cmdSalesbyInvoiceType 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   5565
               MouseIcon       =   "frmAMISMainMenu.frx":3CB37
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3CC89
               Style           =   1  'Graphical
               TabIndex        =   90
               ToolTipText     =   "View Sales by Invoice Type"
               Top             =   1470
               Width           =   720
            End
            Begin VB.CommandButton cmdSalesJournal 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   390
               MouseIcon       =   "frmAMISMainMenu.frx":3D3C8
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3D51A
               Style           =   1  'Graphical
               TabIndex        =   88
               ToolTipText     =   "View Sales Journal"
               Top             =   630
               Width           =   720
            End
            Begin VB.CommandButton cmdAccountDetailbyCustomer 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   390
               MouseIcon       =   "frmAMISMainMenu.frx":3DC81
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3DDD3
               Style           =   1  'Graphical
               TabIndex        =   87
               ToolTipText     =   "View Account Detail by Customer"
               Top             =   1470
               Width           =   720
            End
            Begin VB.CommandButton cmdLedgerCodeRunningBalance_LCRB 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   390
               MouseIcon       =   "frmAMISMainMenu.frx":3E47B
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3E5CD
               Style           =   1  'Graphical
               TabIndex        =   86
               ToolTipText     =   "View Ledger Code Running Balance"
               Top             =   2310
               Width           =   720
            End
            Begin VB.CommandButton cmdScheduleOfAccountsReceivable 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   390
               MouseIcon       =   "frmAMISMainMenu.frx":3EC86
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3EDD8
               Style           =   1  'Graphical
               TabIndex        =   85
               ToolTipText     =   "View Schedule Of Accounts Receivable"
               Top             =   3120
               Width           =   720
            End
            Begin VB.CommandButton cmdCashDisbursementJournal 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   -69580
               MouseIcon       =   "frmAMISMainMenu.frx":3F510
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3F662
               Style           =   1  'Graphical
               TabIndex        =   84
               ToolTipText     =   "View Cash Disbursement Journal"
               Top             =   690
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdLedgerCodeRunningBalance_CD 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   -69580
               MouseIcon       =   "frmAMISMainMenu.frx":3FE38
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":3FF8A
               Style           =   1  'Graphical
               TabIndex        =   83
               ToolTipText     =   "View Ledger Code Running Balance"
               Top             =   1530
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdCheckRegister 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   -69580
               MouseIcon       =   "frmAMISMainMenu.frx":40643
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":40795
               Style           =   1  'Graphical
               TabIndex        =   82
               ToolTipText     =   "View Check Register"
               Top             =   2385
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdReceivingReportRegister 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   -64615
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":40EF2
               Style           =   1  'Graphical
               TabIndex        =   81
               ToolTipText     =   "View Receiving Report Register"
               Top             =   675
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdAccountsPayableJournal 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   -69565
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":41573
               Style           =   1  'Graphical
               TabIndex        =   80
               ToolTipText     =   "View Accounts Payable Journal"
               Top             =   720
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdLedgerCodeRunningBalance_AP 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   -69565
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":41C31
               Style           =   1  'Graphical
               TabIndex        =   79
               ToolTipText     =   "View Ledger Code Running Balance"
               Top             =   1500
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdAccountsPayableAgingReport 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   -69580
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":422EA
               Style           =   1  'Graphical
               TabIndex        =   78
               ToolTipText     =   "View Accounts Payable Aging Report"
               Top             =   3120
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdAccountDetailbySupplier 
               BackColor       =   &H00FFFFFF&
               Height          =   645
               Left            =   -69565
               MousePointer    =   99  'Custom
               Picture         =   "frmAMISMainMenu.frx":42A17
               Style           =   1  'Graphical
               TabIndex        =   77
               ToolTipText     =   "View Account Detail by Supplier"
               Top             =   2340
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.Label Label75 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vehicle Sales Analysis"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   6435
               TabIndex        =   163
               Top             =   3300
               Width           =   1935
            End
            Begin VB.Label Label74 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Collection Forecast Report"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   6435
               TabIndex        =   161
               Top             =   2540
               Width           =   2385
            End
            Begin VB.Label Label70 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Billing Due Report"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   -63715
               TabIndex        =   159
               Top             =   1695
               Visible         =   0   'False
               Width           =   1635
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Journal Voucher"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   -68710
               TabIndex        =   118
               Top             =   915
               Visible         =   0   'False
               Width           =   1470
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ledger Code Running Balance"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   -68710
               TabIndex        =   117
               Top             =   1680
               Visible         =   0   'False
               Width           =   2715
            End
            Begin VB.Label Label67 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unused O.R."
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   -68725
               TabIndex        =   116
               Top             =   3375
               Visible         =   0   'False
               Width           =   1110
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "O.R."
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   -68725
               TabIndex        =   115
               Top             =   2490
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ledger Code Running Balance"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   -68725
               TabIndex        =   114
               Top             =   1665
               Visible         =   0   'False
               Width           =   2715
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cash Receipts Journal"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   -68725
               TabIndex        =   113
               Top             =   870
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unused Invoices"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1290
               TabIndex        =   112
               Top             =   4170
               Width           =   1485
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Invoice Register"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   6435
               TabIndex        =   111
               Top             =   765
               Width           =   1440
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales by Invoice Type"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   6435
               TabIndex        =   110
               Top             =   1700
               Width           =   1935
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "A/R Aging and Schedule Reports"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1290
               TabIndex        =   109
               Top             =   3300
               Width           =   2970
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ledger Code Running Balance"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1290
               TabIndex        =   108
               Top             =   2540
               Width           =   2715
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Account Detail by Customer"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1290
               TabIndex        =   107
               Top             =   1700
               Width           =   2550
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Journal"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   1290
               TabIndex        =   106
               Top             =   870
               Width           =   1170
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Check Register"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   -68710
               TabIndex        =   105
               Top             =   2580
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ledger Code Running Balance"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   -68710
               TabIndex        =   104
               Top             =   1740
               Visible         =   0   'False
               Width           =   2715
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cash Disbursement Journal"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   -68710
               TabIndex        =   103
               Top             =   900
               Visible         =   0   'False
               Width           =   2475
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Receiving Report Register"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   -63730
               TabIndex        =   102
               Top             =   825
               Visible         =   0   'False
               Width           =   2325
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "A/P Aging and Schedule Reports"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   -68695
               TabIndex        =   101
               Top             =   3315
               Visible         =   0   'False
               Width           =   2970
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Account Detail by Supplier"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   -68695
               TabIndex        =   100
               Top             =   2550
               Visible         =   0   'False
               Width           =   2400
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ledger Code Running Balance"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   -68695
               TabIndex        =   99
               Top             =   1680
               Visible         =   0   'False
               Width           =   2715
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Accounts Payable Journal"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   -68695
               TabIndex        =   98
               Top             =   870
               Visible         =   0   'False
               Width           =   2325
            End
         End
         Begin VB.Label Label76 
            BackStyle       =   0  'Transparent
            Caption         =   "Withholding Tax-Expanded"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   -64240
            TabIndex        =   168
            Top             =   1080
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tax Reports"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -68740
            TabIndex        =   141
            Top             =   3540
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Financial Statements"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -68740
            TabIndex        =   121
            Top             =   2730
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trial Balance"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -68740
            TabIndex        =   120
            Top             =   1905
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label cmdWorkSheet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Work Sheet"
            BeginProperty Font 
               Name            =   "Segoe UI Semibold"
               Size            =   9.75
               Charset         =   0
               Weight          =   600
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -68740
            TabIndex        =   119
            Top             =   1080
            Visible         =   0   'False
            Width           =   1065
         End
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Logo Tool"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -63520
         TabIndex        =   176
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trial Balance Uploading Tool"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68680
         TabIndex        =   174
         Top             =   1920
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule Uploading Tool"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68680
         TabIndex        =   172
         Top             =   2760
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GL vs. SL Tally Tool"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68680
         TabIndex        =   171
         Top             =   1080
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Master File - Deposits"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -64330
         TabIndex        =   157
         Top             =   1110
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Un-Applied Payment"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68650
         TabIndex        =   139
         Top             =   4230
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Un-Imported Reports"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68650
         TabIndex        =   138
         Top             =   4950
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Tools"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -63310
         TabIndex        =   137
         Top             =   2670
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Re Printing Report"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -63310
         TabIndex        =   136
         Top             =   1125
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelled Report"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -63310
         TabIndex        =   135
         Top             =   1860
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Audit Report"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68650
         TabIndex        =   134
         Top             =   3390
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Audit Inquiry"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68650
         TabIndex        =   133
         Top             =   2610
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Signatories and Headers"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68650
         TabIndex        =   132
         Top             =   1860
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reminders"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68650
         TabIndex        =   131
         Top             =   1125
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Master File"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68650
         TabIndex        =   30
         Top             =   1935
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ATC Code Master File"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68650
         TabIndex        =   29
         Top             =   5130
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terms of Payment Master File"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68650
         TabIndex        =   28
         Top             =   4335
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Master File - Checks"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68650
         TabIndex        =   27
         Top             =   2730
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Type Master File"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68650
         TabIndex        =   26
         Top             =   3525
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Master File"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -68650
         TabIndex        =   25
         Top             =   1125
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Receipts Journal (Deposited OR's)"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1305
         TabIndex        =   18
         Top             =   4530
         Width           =   3525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accounts General Ledger"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6810
         TabIndex        =   17
         Top             =   1140
         Width           =   2265
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/R Ledger"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6810
         TabIndex        =   16
         Top             =   1965
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/P Ledger"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6810
         TabIndex        =   15
         Top             =   2775
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "General Journal"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1305
         TabIndex        =   14
         Top             =   5415
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Journal"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1305
         TabIndex        =   13
         Top             =   2805
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Receipts Journal (Un-Deposited OR's)"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1305
         TabIndex        =   12
         Top             =   3630
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Disbursement Journal"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1305
         TabIndex        =   11
         Top             =   1935
         Width           =   2475
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accounts Payable Journal"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1305
         TabIndex        =   10
         Top             =   1095
         Width           =   2325
      End
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim frmNEWAMISRangeWithSummary                              As frmAMISRangeWithSummary

Private Sub cmdAccount_AccountEntriesTemplate_Click()
    If Module_Access(LOGID, "ACCOUNT ENTRIES TEMPLATES", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAMISMASTERFILESTemplates
End Sub

Private Sub cmdAdjustment_Client_Click()
    If Module_Access(LOGID, "CLIENT ADJUSTING JOURNAL ENTRIES", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Call frmAMISJouirnalEntry_CAJE.LOADJOURNAL("CAJ")
    FormExistsShow frmAMISJouirnalEntry_CAJE
End Sub

Private Sub cmdAdjustment_Proposed_Click()
    If Module_Access(LOGID, "PROPOSED ADJUSTING JOURNAL ENTRIES", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Call frmAMISJouirnalEntry_PAJE.LOADJOURNAL("PAJ")
    FormExistsShow frmAMISJouirnalEntry_PAJE
End Sub

Private Sub cmdAdjustment_ClosingEntries_Click()
    If Module_Access(LOGID, "CLOSING ENTRIES", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Call frmAMISJournalEntry.LOADJOURNAL("CLO")
    FormExistsShow frmAMISJournalEntry
End Sub

Private Sub cmdCustom_Click()
    If COMPANY_CODE = "HGC" Then

    ElseIf COMPANY_CODE = "HAI" Then

    ElseIf COMPANY_CODE = "HAS" Then

    ElseIf COMPANY_CODE = "HBK" Then

    ElseIf COMPANY_CODE = "HHM" Then

    ElseIf COMPANY_CODE = "HSB" Or COMPANY_CODE = "HCR" Then

    End If
End Sub

Private Sub cmdBankDeposits_Click()
    If Module_Access(LOGID, "BANK DEPOSITS", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAMISMASTERFILEBanks2
End Sub

Private Sub cmdBillingDueReport_Click()
    If Module_Access(LOGID, "BILLING DUE REPORT", "REPORTS") = False Then Exit Sub
    On Error Resume Next
    Call frmBillingDueReport.Report_Type("AP DUE")
    frmBillingDueReport.Caption = "BILLING DUE REPORT"
    FormExistsShow frmBillingDueReport
End Sub

Private Sub cmdCollectionForecast_Click()
    If Module_Access(LOGID, "COLLECTION FORECAST REPORT", "REPORTS") = False Then Exit Sub
    On Error Resume Next
    Call frmBillingDueReport.Report_Type("AR DUE")
    frmBillingDueReport.Caption = "COLLECTION FORECAST REPORT"
    FormExistsShow frmBillingDueReport
End Sub

Private Sub cmdCustomerDebitMemo_Click()
    If Module_Access(LOGID, "CUSTOMER DEBIT MEMO", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Call frmAMISJournalEntry_CDM.LOADJOURNAL("CDM")
    FormExistsShow frmAMISJournalEntry_CDM
End Sub

Private Sub cmdCustomersDeposit_Click()
    If Module_Access(LOGID, "CUSTOMERS DEPOSIT LEDGER", "INQUIRY") = False Then Exit Sub
    FormExistsShow frmAMIS_CustomerDepositLedger
End Sub

Private Sub cmdInvoiceCM_Click()
    If Module_Access(LOGID, "CREDITMEMO REPORT", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "CREDITMEMO_REGISTER"
    FormExistsShow frmAMISRange
    frmAMISRange.Caption = "Invoices Credit Memo"
    DoEvents
End Sub

Private Sub cmdSalesAnalysis_Click()
    If Module_Access(LOGID, "VEHICLE SALES ANALYSIS", "REPORTS") = False Then Exit Sub
    On Error Resume Next
    FormExistsShow frmVehicleSalesAnalysis
End Sub

Private Sub cmdTBUTool_Click(Index As Integer)
    frmTBUploadingTool.Show
End Sub

Private Sub cmdTranType_Click()
    frmTranTypeImportingSetup.Show
End Sub

Private Sub cmdTallyTool_Click(Index As Integer)
    frmTallyTool.Show
End Sub

Private Sub cmdUTool_Click(Index As Integer)
    frmUploadingTool.Show
End Sub

Private Sub cmdLogoTool_Click(Index As Integer)
    frmLogoTool.Show
End Sub

Private Sub cmdVehicleSales_Click()
    If Module_Access(LOGID, "VEHICLE MODEL SET-UP", "SYSTEM") = False Then Exit Sub
    FormExistsShow frmVehicleSalesCodeSetup
End Sub

Private Sub cmdVendorCreditMemo_Click()
    If Module_Access(LOGID, "VENDOR CREDIT MEMO", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Call frmAMISJournalEntry_VCM.LOADJOURNAL("VCM")
    FormExistsShow frmAMISJournalEntry_VCM
End Sub

Private Sub cmdVendorDebitMemo_Click()
    If Module_Access(LOGID, "VENDOR DEBIT MEMO", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Call frmAMISJournalEntry_VDM.LOADJOURNAL("VDM")
    FormExistsShow frmAMISJournalEntry_VDM
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
    Call frmAMISRangeWithAccountCode.LOADJOURNAL("GJ")
    FormExistsShow frmAMISRangeWithAccountCode
    frmAMISRangeWithAccountCode.Caption = "Journal Voucher Ledger Code Running Balance"
End Sub

Private Sub cmdJournal_AP_Click()
    If Module_Access(LOGID, "ACCOUNTS PAYABLE JOURNAL", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Call frmAMISJournalEntry_APJ.LOADJOURNAL("APJ")
    FormExistsShow frmAMISJournalEntry_APJ
End Sub

Private Sub cmdJournal_CashDisburshment_Click()
    If Module_Access(LOGID, "CASH DISBURSEMENT JOURNAL", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Call frmAMISJournalEntry_CDJ.LOADJOURNAL("CDJ")
    FormExistsShow frmAMISJournalEntry_CDJ
End Sub

Private Sub cmdJournal_CRJ_Click()
    If Module_Access(LOGID, "CASH RECEIPTS JOURNAL", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Call frmAMISJournalEntry_CRJ.LOADJOURNAL("CRJ")
    FormExistsShow frmAMISJournalEntry_CRJ
End Sub

Private Sub cmdJournal_General_Click()
    On Error Resume Next
    If Module_Access(LOGID, "GENERAL JOURNAL", "TRANSACTION") = False Then Exit Sub

    On Error Resume Next
    Call frmAMISJournalEntry_GJ.LOADJOURNAL("GJ")
    frmAMISJournalEntry_GJ.Show 1
End Sub

Private Sub cmdJournal_Sales_Click()
    If Module_Access(LOGID, "SALES JOURNAL", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Call frmAMISJournalEntry_SJ.LOADJOURNAL("SJ")
    FormExistsShow frmAMISJournalEntry_SJ
End Sub

Private Sub cmdJournalVoucherSummary_Click()
    If Module_Access(LOGID, "GENERAL JOURNAL SUMMARY", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "GJ"
    FormExistsShow frmAMISRangeWithSummary
    frmAMISRangeWithSummary.Caption = "General Journal"
End Sub

Private Sub cmdLedger_Account_Click()
    If Module_Access(LOGID, "ACCOUNT GENERAL LEDGER", "INQUIRY") = False Then Exit Sub
    FormExistsShow frmAMISLEDGERAccounts
End Sub

Private Sub cmdLedger_Customer_Click()
    If Module_Access(LOGID, "CUSTOMER A/R LEDGER", "INQUIRY") = False Then Exit Sub
    FormExistsShow frmAMIS_ARLEDGER
End Sub

Private Sub cmdLedger_VendorSubisdy_Click()
    If Module_Access(LOGID, "VENDOR SUBSIDIARY LEDGER", "INQUIRY") = False Then Exit Sub
    FormExistsShow frmAMIS_APLEDGER
End Sub

Private Sub cmdOpeningBalance_Accounts_Click()
    If Module_Access(LOGID, "ACCOUNT OPENING BALANCE", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Call frmAMISJournalEntry_OPB.LOADJOURNAL("OPB")
    FormExistsShow frmAMISJournalEntry_OPB
End Sub

Private Sub cmdJournalDRJ_Click()
    If Module_Access(LOGID, "DEPOSITED RECEIPTS JOURNAL", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    'SJR 082114
    If COMPANY_CODE = "HCA" Then
        Call frmAMISJournalEntry_DRJ_2.LOADJOURNAL("DRJ")
        FormExistsShow frmAMISJournalEntry_DRJ_2
    Else
        Call frmAMISJournalEntry_DRJ.LOADJOURNAL("DRJ")
        FormExistsShow frmAMISJournalEntry_DRJ
    End If
    'SJR 082114
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
    If Module_Access(LOGID, "CUSTOMER OPENING BALANCE", "DATA ENTRY") = False Then Exit Sub
    On Error Resume Next
    Call frmAMISCustomerAROpening.LOADJOURNAL("COB")
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
    If COMPANY_CODE = "HCA" Then
        FormExistsShow frmAMISFinancialStatementsHCA
    ElseIf COMPANY_CODE = "HSM" Then
        FormExistsShow frmAMISFinancialStatementsHSM
    ElseIf COMPANY_CODE = "HSB" Or COMPANY_CODE = "HCR" Or COMPANY_CODE = "HLB" Or COMPANY_CODE = "HBC" Then
        FormExistsShow frmAMISFinancialStatementsHSB
    Else
        FormExistsShow frmAMISFinancialStatements
    End If
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
    Call frmAMISRangeWithAccountCode.LOADJOURNAL("APJ")
    FormExistsShow frmAMISRangeWithAccountCode
    frmAMISRangeWithAccountCode.Caption = "ACCOUNTS Disbursement Ledger Code Running Balance"
End Sub

Private Sub cmdLedgerCodeRunningBalance_CD_Click()
    If Module_Access(LOGID, "CASH DISBURSEMENT LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "CDJ"
    Call frmAMISRangeWithAccountCode.LOADJOURNAL("CDJ")
    FormExistsShow frmAMISRangeWithAccountCode
    frmAMISRangeWithAccountCode.Caption = "Cash Disbursement Ledger Code Running Balance"
End Sub

Private Sub cmdLedgerCodeRunningBalance_CR_Click()
    If Module_Access(LOGID, "CASH RECEIPTS LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "CRJ"
    Call frmAMISRangeWithAccountCode.LOADJOURNAL("CRJ")
    FormExistsShow frmAMISRangeWithAccountCode
    frmAMISRangeWithAccountCode.Caption = "Cash Receipts Ledger Code Running Balance"
End Sub

Private Sub cmdLedgerCodeRunningBalance_LCRB_Click()
    If Module_Access(LOGID, "SALES LEDGER CODE RUNNING BALANCE", "REPORTS") = False Then Exit Sub
    REPORT_RANGETYPE = "SJ"
    Call frmAMISRangeWithAccountCode.LOADJOURNAL("SJ")
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

Private Sub cmdScheduleOfAccountsReceivable_Click()
    If Module_Access(LOGID, "ACCOUNTS RECEIVABLE AGING REPORT", "REPORTS") = False Then Exit Sub
    Report_AR = "AGING"

    Dim frmAMIS_AR                                          As frmNEW_ARSchedReport
    Set frmAMIS_AR = New frmNEW_ARSchedReport
    FormExistsShow frmAMIS_AR
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
    Call frmAMISVendorAPOpening.LOADJOURNAL("VPJ")
    FormExistsShow frmAMISVendorAPOpening
End Sub

Private Sub cmdWitholdingTax_Click()
    If Module_Access(LOGID, "TAX REPORTS", "REPORTS") = False Then Exit Sub
    FormExistsShow frmTaxReports
End Sub

Private Sub cmdWork_Sheet_Click()
    If Module_Access(LOGID, "WORKSHEET", "REPORTS") = False Then Exit Sub
    FormExistsShow frmAMISWorkSheet
End Sub

Private Sub cmdWTExpanded_Click()
    FormExistsShow frmWTExpanded
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
    Dim frmAMIS_AP                                          As frmAMIS_AP_Process
    Set frmAMIS_AP = New frmAMIS_AP_Process
    FormExistsShow frmAMIS_AP
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
End Sub

Private Sub cmdOpeningReport_Click()
    FormExistsShow frmOpeningBalanceReport
End Sub

Private Sub Command12_Click()
    If Module_Access(LOGID, "IMPORTING TEMPLATE", "SYSTEM") = False Then Exit Sub
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
    If Module_Access(LOGID, "CUSTOMER CREDIT MEMO", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Call frmAMISJournalEntry_CCM.LOADJOURNAL("CCM")
    FormExistsShow frmAMISJournalEntry_CCM
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Call frmAMISbanksOpening.LOADJOURNAL("BOB")
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

Private Sub Command9_Click()
    Screen.MousePointer = 11
    FormExistsShow frmAMIS_UniportedReports
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "]"
    TabControl1.SelectedItem = 0
    TabControl2.SelectedItem = 0
    TabControl3.SelectedItem = 0
    TabControl4.SelectedItem = 0
    
    If UCase(LOGNAME) <> "NETSPEED" Then
    TabControl1.Item(5).Visible = False
    End If
  
    If COMPANY_CODE = "DAI" Or COMPANY_CODE = "DPI" Or COMPANY_CODE = "DMI" Or COMPANY_CODE = "HCE" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "HCA" Or COMPANY_CODE = "HQA" Or COMPANY_CODE = "HMH" Then
        picDeposit.Visible = True
    Else
        picDeposit.Visible = False
    End If
End Sub

Sub DisplayInfo(XXX As String)
    Dim RS                                                  As New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_HD where jtype='APJ' and voucherno='" & XXX & "'")
    If Not RS.EOF And Not RS.BOF Then

    End If
    Set RS = Nothing
End Sub

Private Sub TabControl1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If UCase(LOGNAME) = "NETSPEED" Then
        cmdTranType.Visible = True
        cmdTallyTool(0).Visible = True
        cmdUTool(1).Visible = True
    Else
        cmdTranType.Visible = False
        cmdTallyTool(0).Visible = False
    End If
    
    If UCase(LOGNAME) = "NETSPEED" Then
        Command12.Visible = True
        cmdTallyTool(0).Visible = True
        cmdUTool(1).Visible = True
    Else
        Command12.Visible = False
        cmdTallyTool(0).Visible = False
    End If
End Sub

Private Sub TabControl4_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    'JULIE: ADD CREDIT_MEMO FIELD
    If COMPANY_CODE = "HMH" Then
        picCreditMemo.Visible = True
    Else
        picCreditMemo.Visible = False
    End If
End Sub

