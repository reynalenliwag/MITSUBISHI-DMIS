VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCRIS_Customer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Information"
   ClientHeight    =   7980
   ClientLeft      =   -1755
   ClientTop       =   2055
   ClientWidth     =   13245
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ProfilePersonal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   13245
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraSearch 
      Height          =   7215
      Left            =   0
      TabIndex        =   159
      Top             =   0
      Width           =   2475
      Begin VB.ComboBox cboSearchCustype 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   166
         Top             =   1500
         Width           =   2355
      End
      Begin VB.OptionButton optSearchKeyEmail 
         Caption         =   "Search By Email"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   165
         Top             =   1200
         Width           =   2295
      End
      Begin VB.OptionButton optSearchKeyAddress 
         Caption         =   "Search By Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   164
         Top             =   945
         Width           =   2295
      End
      Begin VB.OptionButton optSearchKeyAcctName 
         Caption         =   "Search By A/C Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   163
         Top             =   690
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton optSearchKeyCompany 
         Caption         =   "Search By Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   162
         Top             =   435
         Width           =   2295
      End
      Begin VB.OptionButton optSearchKeyLast 
         Caption         =   "Search By Last Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   161
         Top             =   180
         Width           =   2295
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   60
         MaxLength       =   35
         TabIndex        =   160
         TabStop         =   0   'False
         Top             =   1860
         Width           =   2295
      End
      Begin MSComctlLib.ListView lstCustomer 
         Height          =   4905
         Left            =   30
         TabIndex        =   167
         Top             =   2280
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   8652
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ProfilePersonal.frx":08CA
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CUSTOMER NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   11760
      MouseIcon       =   "ProfilePersonal.frx":0A2C
      MousePointer    =   99  'Custom
      Picture         =   "ProfilePersonal.frx":0B7E
      Style           =   1  'Graphical
      TabIndex        =   158
      Top             =   7260
      Width           =   645
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   12435
      MouseIcon       =   "ProfilePersonal.frx":0ECE
      MousePointer    =   99  'Custom
      Picture         =   "ProfilePersonal.frx":1020
      Style           =   1  'Graphical
      TabIndex        =   157
      Top             =   7260
      Width           =   645
   End
   Begin VB.PictureBox picForm 
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   2460
      ScaleHeight     =   7215
      ScaleWidth      =   10815
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin VB.TextBox txtAccName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   7650
         Locked          =   -1  'True
         TabIndex        =   2
         Tag             =   "@R"
         Text            =   "Text"
         ToolTipText     =   "Account Name"
         Top             =   90
         Width           =   3090
      End
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   6690
         Left            =   0
         TabIndex        =   3
         Top             =   540
         Width           =   10770
         _Version        =   655364
         _ExtentX        =   18997
         _ExtentY        =   11800
         _StockProps     =   64
         Appearance      =   1
         Color           =   4
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.FixedTabWidth=   130
         ItemCount       =   9
         Item(0).Caption =   "General"
         Item(0).ControlCount=   3
         Item(0).Control(0)=   "framePersonal"
         Item(0).Control(1)=   "Frame2"
         Item(0).Control(2)=   "Frame3"
         Item(1).Caption =   "Addresses/Family"
         Item(1).Tooltip =   "View/Add/Edit Mutliple Contact for This Account"
         Item(1).ControlCount=   2
         Item(1).Control(0)=   "Frame5"
         Item(1).Control(1)=   "Frame1"
         Item(2).Caption =   "Contacts"
         Item(2).ControlCount=   5
         Item(2).Control(0)=   "lvMultipleContacts"
         Item(2).Control(1)=   "cmdAddEntryContacts"
         Item(2).Control(2)=   "cmdCancelEntryContacts"
         Item(2).Control(3)=   "Text2"
         Item(2).Control(4)=   "lblCap(25)"
         Item(3).Caption =   "Vehicles"
         Item(3).ControlCount=   5
         Item(3).Control(0)=   "lvVehicles"
         Item(3).Control(1)=   "cmdCancelEntryVehicles"
         Item(3).Control(2)=   "cmdAddEntryVehicles"
         Item(3).Control(3)=   "Text3"
         Item(3).Control(4)=   "lblCap(26)"
         Item(4).Caption =   "History"
         Item(4).ControlCount=   1
         Item(4).Control(0)=   "TabControl2"
         Item(5).Caption =   "Letters"
         Item(5).ControlCount=   3
         Item(5).Control(0)=   "ReportControl1"
         Item(5).Control(1)=   "Command6"
         Item(5).Control(2)=   "Command7"
         Item(6).Caption =   "Complains"
         Item(6).ControlCount=   3
         Item(6).Control(0)=   "ReportControl2"
         Item(6).Control(1)=   "Command8"
         Item(6).Control(2)=   "Command9"
         Item(7).Caption =   "Invitations"
         Item(7).ControlCount=   3
         Item(7).Control(0)=   "Command10"
         Item(7).Control(1)=   "Command11"
         Item(7).Control(2)=   "ReportControl3"
         Item(8).Caption =   "Notes"
         Item(8).ControlCount=   0
         Begin XtremeReportControl.ReportControl ReportControl3 
            Height          =   5655
            Left            =   -69940
            TabIndex        =   156
            Top             =   930
            Visible         =   0   'False
            Width           =   10605
            _Version        =   655364
            _ExtentX        =   18706
            _ExtentY        =   9975
            _StockProps     =   64
            BorderStyle     =   2
            ShowFooter      =   -1  'True
         End
         Begin XtremeReportControl.ReportControl ReportControl1 
            Height          =   5775
            Left            =   -69970
            TabIndex        =   129
            Top             =   930
            Visible         =   0   'False
            Width           =   10665
            _Version        =   655364
            _ExtentX        =   18812
            _ExtentY        =   10186
            _StockProps     =   64
            BorderStyle     =   2
            ShowFooter      =   -1  'True
         End
         Begin XtremeReportControl.ReportControl lvVehicles 
            Height          =   5775
            Left            =   -69970
            TabIndex        =   37
            Top             =   930
            Visible         =   0   'False
            Width           =   10665
            _Version        =   655364
            _ExtentX        =   18812
            _ExtentY        =   10186
            _StockProps     =   64
            BorderStyle     =   2
            ShowFooter      =   -1  'True
         End
         Begin XtremeReportControl.ReportControl lvMultipleContacts 
            Height          =   5775
            Left            =   -69970
            TabIndex        =   38
            Top             =   930
            Visible         =   0   'False
            Width           =   10665
            _Version        =   655364
            _ExtentX        =   18812
            _ExtentY        =   10186
            _StockProps     =   64
            BorderStyle     =   2
            ShowFooter      =   -1  'True
         End
         Begin XtremeReportControl.ReportControl ReportControl2 
            Height          =   5655
            Left            =   -69940
            TabIndex        =   151
            Top             =   930
            Visible         =   0   'False
            Width           =   10605
            _Version        =   655364
            _ExtentX        =   18706
            _ExtentY        =   9975
            _StockProps     =   64
            BorderStyle     =   2
            ShowFooter      =   -1  'True
         End
         Begin VB.CommandButton Command11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -69475
            MaskColor       =   &H00000040&
            Picture         =   "ProfilePersonal.frx":135E
            Style           =   1  'Graphical
            TabIndex        =   155
            Top             =   480
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -69940
            MaskColor       =   &H00000040&
            Picture         =   "ProfilePersonal.frx":1568
            Style           =   1  'Graphical
            TabIndex        =   154
            Top             =   480
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.CommandButton Command9 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -69940
            MaskColor       =   &H00000040&
            Picture         =   "ProfilePersonal.frx":1C52
            Style           =   1  'Graphical
            TabIndex        =   153
            Top             =   480
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.CommandButton Command8 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -69475
            MaskColor       =   &H00000040&
            Picture         =   "ProfilePersonal.frx":233C
            Style           =   1  'Graphical
            TabIndex        =   152
            Top             =   480
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   -62710
            Locked          =   -1  'True
            TabIndex        =   145
            Tag             =   "@R"
            Text            =   "Text"
            ToolTipText     =   "Account Name"
            Top             =   540
            Visible         =   0   'False
            Width           =   3090
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   -62710
            Locked          =   -1  'True
            TabIndex        =   143
            Tag             =   "@R"
            Text            =   "Text"
            ToolTipText     =   "Account Name"
            Top             =   540
            Visible         =   0   'False
            Width           =   3090
         End
         Begin VB.CommandButton Command7 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -69445
            MaskColor       =   &H00000040&
            Picture         =   "ProfilePersonal.frx":2546
            Style           =   1  'Graphical
            TabIndex        =   131
            Top             =   480
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -69910
            MaskColor       =   &H00000040&
            Picture         =   "ProfilePersonal.frx":2750
            Style           =   1  'Graphical
            TabIndex        =   130
            Top             =   480
            Visible         =   0   'False
            Width           =   420
         End
         Begin XtremeSuiteControls.TabControl TabControl2 
            Height          =   6015
            Left            =   -69850
            TabIndex        =   12
            Top             =   480
            Visible         =   0   'False
            Width           =   10515
            _Version        =   655364
            _ExtentX        =   18547
            _ExtentY        =   10610
            _StockProps     =   64
            Appearance      =   1
            Color           =   4
            PaintManager.ShowIcons=   -1  'True
            ItemCount       =   5
            Item(0).Caption =   "Call History"
            Item(0).ControlCount=   0
            Item(1).Caption =   "Visit History"
            Item(1).ControlCount=   0
            Item(2).Caption =   "Transaction History"
            Item(2).ControlCount=   0
            Item(3).Caption =   "Service History"
            Item(3).ControlCount=   0
            Item(4).Caption =   "Vehicles Visits"
            Item(4).ControlCount=   0
         End
         Begin VB.Frame Frame1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6255
            Left            =   -69910
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   4995
            Begin XtremeReportControl.ReportControl lvChildrens 
               Height          =   4845
               Left            =   90
               TabIndex        =   11
               Top             =   1320
               Width           =   4815
               _Version        =   655364
               _ExtentX        =   8493
               _ExtentY        =   8546
               _StockProps     =   64
               BorderStyle     =   2
            End
            Begin VB.CommandButton Command5 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   3990
               MaskColor       =   &H00000040&
               Picture         =   "ProfilePersonal.frx":2E3A
               Style           =   1  'Graphical
               TabIndex        =   127
               Top             =   870
               Width           =   420
            End
            Begin VB.CommandButton Command4 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   4455
               MaskColor       =   &H00000040&
               Picture         =   "ProfilePersonal.frx":3524
               Style           =   1  'Graphical
               TabIndex        =   126
               Top             =   870
               Width           =   420
            End
            Begin VB.TextBox txtSpouseName 
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1305
               TabIndex        =   7
               Top             =   480
               Width           =   3645
            End
            Begin MSComCtl2.DTPicker dtAnniversary 
               Height          =   315
               Left            =   1305
               TabIndex        =   9
               Top             =   885
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarTitleBackColor=   8388608
               CalendarTitleForeColor=   16777215
               CheckBox        =   -1  'True
               Format          =   55377921
               CurrentDate     =   39135
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Spouse Name"
               ForeColor       =   &H8000000D&
               Height          =   210
               Left            =   90
               TabIndex        =   6
               Top             =   525
               Width           =   1005
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Anniversary"
               ForeColor       =   &H8000000D&
               Height          =   210
               Left            =   90
               TabIndex        =   8
               Top             =   930
               Width           =   900
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "#Children"
               ForeColor       =   &H8000000D&
               Height          =   210
               Left            =   2970
               TabIndex        =   10
               Top             =   930
               Width           =   675
            End
            Begin XtremeShortcutBar.ShortcutCaption CapInfo 
               Height          =   315
               Index           =   0
               Left            =   30
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   120
               Width           =   4920
               _Version        =   655364
               _ExtentX        =   8678
               _ExtentY        =   556
               _StockProps     =   14
               Caption         =   "Notes"
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SubItemCaption  =   -1  'True
               ForeColor       =   64
            End
         End
         Begin VB.Frame Frame5 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   6240
            Left            =   -64870
            TabIndex        =   17
            Top             =   360
            Visible         =   0   'False
            Width           =   5520
            Begin VB.TextBox txtBilling_CO 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1395
               TabIndex        =   29
               ToolTipText     =   "Middle Initial"
               Top             =   3180
               Width           =   3825
            End
            Begin VB.ComboBox cboBilling_City 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1380
               TabIndex        =   33
               Tag             =   "Positions"
               ToolTipText     =   "Positions"
               Top             =   4350
               Width           =   3900
            End
            Begin VB.ComboBox cboBilling_Province 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1380
               TabIndex        =   36
               Tag             =   "Positions"
               ToolTipText     =   "Positions"
               Top             =   4770
               Width           =   3900
            End
            Begin VB.TextBox txtBilling_Street 
               Height          =   780
               Left            =   1395
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   30
               Top             =   3540
               Width           =   3870
            End
            Begin VB.TextBox txtShipping_CO 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1395
               TabIndex        =   21
               ToolTipText     =   "Middle Initial"
               Top             =   780
               Width           =   3825
            End
            Begin VB.ComboBox cboShipping_City 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1380
               TabIndex        =   25
               Tag             =   "Positions"
               ToolTipText     =   "Positions"
               Top             =   1950
               Width           =   3900
            End
            Begin VB.ComboBox cboShipping_Province 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1380
               TabIndex        =   26
               Tag             =   "Positions"
               ToolTipText     =   "Positions"
               Top             =   2370
               Width           =   3900
            End
            Begin VB.TextBox txtShipping_Street 
               Height          =   780
               Left            =   1395
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   22
               Top             =   1140
               Width           =   3870
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "C/O"
               ForeColor       =   &H8000000D&
               Height          =   225
               Index           =   24
               Left            =   870
               TabIndex        =   35
               Top             =   3270
               Width           =   315
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Province/Sate"
               ForeColor       =   &H8000000D&
               Height          =   225
               Index           =   23
               Left            =   60
               TabIndex        =   34
               Top             =   4890
               Width           =   1125
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "City"
               ForeColor       =   &H8000000D&
               Height          =   225
               Index           =   21
               Left            =   885
               TabIndex        =   32
               Top             =   4380
               Width           =   300
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Street"
               ForeColor       =   &H8000000D&
               Height          =   225
               Index           =   18
               Left            =   705
               TabIndex        =   31
               Top             =   3600
               Width           =   480
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "C/O"
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   16
               Left            =   915
               TabIndex        =   19
               Top             =   870
               Width           =   270
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Province/Sate"
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   15
               Left            =   270
               TabIndex        =   28
               Top             =   2430
               Width           =   1065
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "City"
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   7
               Left            =   975
               TabIndex        =   24
               Top             =   1980
               Width           =   270
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               Caption         =   "Street"
               ForeColor       =   &H8000000D&
               Height          =   780
               Index           =   6
               Left            =   510
               TabIndex        =   23
               Top             =   1200
               Width           =   810
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Billing Address"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   14
               Left            =   270
               TabIndex        =   27
               Top             =   2850
               Width           =   1260
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Shipping Address"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   22
               Left            =   225
               TabIndex        =   20
               Top             =   480
               Width           =   1485
            End
            Begin XtremeShortcutBar.ShortcutCaption CapInfo 
               Height          =   315
               Index           =   5
               Left            =   15
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   120
               Width           =   5490
               _Version        =   655364
               _ExtentX        =   9684
               _ExtentY        =   556
               _StockProps     =   14
               Caption         =   "Other Address"
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SubItemCaption  =   -1  'True
               ForeColor       =   64
            End
         End
         Begin VB.CommandButton cmdCancelEntryContacts 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -69445
            MaskColor       =   &H00000040&
            Picture         =   "ProfilePersonal.frx":372E
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   480
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.CommandButton cmdAddEntryContacts 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -69910
            MaskColor       =   &H00000040&
            Picture         =   "ProfilePersonal.frx":3938
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   480
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.CommandButton cmdAddEntryVehicles 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -69910
            MaskColor       =   &H00000040&
            Picture         =   "ProfilePersonal.frx":4022
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   480
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.CommandButton cmdCancelEntryVehicles 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -69445
            MaskColor       =   &H00000040&
            Picture         =   "ProfilePersonal.frx":470C
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   480
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   4950
            Left            =   5490
            TabIndex        =   59
            Top             =   360
            Width           =   5190
            Begin VB.TextBox txtCrLimit 
               Height          =   375
               Left            =   1440
               TabIndex        =   148
               Top             =   4320
               Width           =   1395
            End
            Begin VB.TextBox txtCrDays 
               Height          =   375
               Left            =   3840
               TabIndex        =   147
               Top             =   4350
               Width           =   1215
            End
            Begin VB.ComboBox cboComp_City 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1425
               TabIndex        =   69
               Tag             =   "Positions"
               ToolTipText     =   "Positions"
               Top             =   2777
               Width           =   3660
            End
            Begin VB.ComboBox cboComp_Province 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1425
               TabIndex        =   72
               Tag             =   "Positions"
               ToolTipText     =   "Positions"
               Top             =   3175
               Width           =   3660
            End
            Begin VB.TextBox txtAsstPhone 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1425
               TabIndex        =   75
               ToolTipText     =   "Middle Initial"
               Top             =   3960
               Width           =   3660
            End
            Begin VB.TextBox txtAssistantName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1425
               TabIndex        =   74
               ToolTipText     =   "Middle Initial"
               Top             =   3573
               Width           =   3660
            End
            Begin VB.ComboBox cboIndustryType 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1425
               TabIndex        =   64
               Tag             =   "Industry Type"
               ToolTipText     =   "Industry Type"
               Top             =   893
               Width           =   3660
            End
            Begin VB.TextBox txtCompanyName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   600
               Left            =   1425
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   66
               Top             =   1291
               Width           =   3660
            End
            Begin VB.ComboBox cboJobTitle 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1425
               TabIndex        =   62
               Tag             =   "Positions"
               ToolTipText     =   "Positions"
               Top             =   495
               Width           =   3660
            End
            Begin VB.TextBox txtComp_Street 
               Height          =   780
               Left            =   1425
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   67
               Top             =   1944
               Width           =   3660
            End
            Begin VB.Label lblCreditLimit 
               AutoSize        =   -1  'True
               Caption         =   "CreditLimit"
               Height          =   210
               Index           =   0
               Left            =   360
               TabIndex        =   150
               Top             =   4350
               Width           =   975
            End
            Begin VB.Label lblCreditLimit 
               AutoSize        =   -1  'True
               Caption         =   "Credit Days"
               Height          =   210
               Index           =   1
               Left            =   2850
               TabIndex        =   149
               Top             =   4410
               Width           =   960
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Province/Sate"
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   2
               Left            =   180
               TabIndex        =   70
               Top             =   3240
               Width           =   1185
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "City"
               ForeColor       =   &H8000000D&
               Height          =   225
               Index           =   1
               Left            =   975
               TabIndex        =   68
               Top             =   2820
               Width           =   300
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Asst Number "
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   20
               Left            =   270
               TabIndex        =   76
               Top             =   3990
               Width           =   990
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Assistant "
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   19
               Left            =   615
               TabIndex        =   73
               Top             =   3630
               Width           =   735
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Industry Type"
               ForeColor       =   &H8000000D&
               Height          =   225
               Left            =   225
               TabIndex        =   63
               Top             =   960
               Width           =   1080
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Company Name"
               ForeColor       =   &H8000000D&
               Height          =   225
               Left            =   30
               TabIndex        =   65
               Top             =   1290
               Width           =   1350
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Job Title"
               ForeColor       =   &H8000000D&
               Height          =   225
               Left            =   630
               TabIndex        =   61
               Top             =   570
               Width           =   690
            End
            Begin XtremeShortcutBar.ShortcutCaption CapInfo 
               Height          =   315
               Index           =   4
               Left            =   30
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   120
               Width           =   5115
               _Version        =   655364
               _ExtentX        =   9022
               _ExtentY        =   556
               _StockProps     =   14
               Caption         =   "Commercial Address"
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.01
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SubItemCaption  =   -1  'True
               ForeColor       =   64
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               Caption         =   "Street"
               ForeColor       =   &H8000000D&
               Height          =   780
               Index           =   13
               Left            =   540
               TabIndex        =   71
               Top             =   1950
               Width           =   810
            End
         End
         Begin VB.Frame framePersonal 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6240
            Left            =   90
            TabIndex        =   39
            Top             =   360
            Width           =   5385
            Begin VB.TextBox txtLastName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1545
               TabIndex        =   142
               ToolTipText     =   "LastName"
               Top             =   866
               Width           =   3690
            End
            Begin VB.TextBox txtMI 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1545
               TabIndex        =   141
               ToolTipText     =   "Middle Initial"
               Top             =   1237
               Width           =   3690
            End
            Begin VB.TextBox txtHomePhone 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1545
               TabIndex        =   140
               ToolTipText     =   "Middle Initial"
               Top             =   3617
               Width           =   3690
            End
            Begin VB.TextBox txtBusinessPhone 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1545
               TabIndex        =   139
               ToolTipText     =   "Middle Initial"
               Top             =   3988
               Width           =   3690
            End
            Begin VB.TextBox txtCellphone 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1545
               TabIndex        =   138
               ToolTipText     =   "Middle Initial"
               Top             =   4730
               Width           =   3690
            End
            Begin VB.TextBox txtFax 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1545
               TabIndex        =   137
               ToolTipText     =   "Middle Initial"
               Top             =   5101
               Width           =   3690
            End
            Begin VB.TextBox txtEmail 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1545
               TabIndex        =   136
               ToolTipText     =   "Middle Initial"
               Top             =   5475
               Width           =   3690
            End
            Begin VB.TextBox txtOtherPhone 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1545
               TabIndex        =   135
               ToolTipText     =   "Middle Initial"
               Top             =   4359
               Width           =   3690
            End
            Begin VB.TextBox txtRes_street 
               Height          =   810
               Left            =   1545
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   134
               Top             =   1994
               Width           =   3690
            End
            Begin VB.ComboBox cboRes_ProvinceState 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1545
               TabIndex        =   133
               Tag             =   "Positions"
               ToolTipText     =   "Positions"
               Top             =   3231
               Width           =   3690
            End
            Begin VB.ComboBox cboRes_City 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1545
               TabIndex        =   132
               Tag             =   "Positions"
               ToolTipText     =   "Positions"
               Top             =   2845
               Width           =   3690
            End
            Begin VB.ComboBox cboSalutations 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   345
               ItemData        =   "ProfilePersonal.frx":4916
               Left            =   1545
               List            =   "ProfilePersonal.frx":4935
               TabIndex        =   42
               Tag             =   "@R"
               Text            =   "cboSalutations"
               ToolTipText     =   "Salutions"
               Top             =   480
               Width           =   750
            End
            Begin VB.TextBox txtFirstName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   2325
               TabIndex        =   43
               ToolTipText     =   "FirstName"
               Top             =   480
               Width           =   2910
            End
            Begin VB.ComboBox cboSex 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   345
               ItemData        =   "ProfilePersonal.frx":496F
               Left            =   1545
               List            =   "ProfilePersonal.frx":497C
               Style           =   2  'Dropdown List
               TabIndex        =   46
               Top             =   1608
               Width           =   1155
            End
            Begin MSComCtl2.DTPicker dtDateofBirth 
               Height          =   315
               Left            =   3735
               TabIndex        =   49
               Top             =   1620
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   556
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarTitleBackColor=   8388608
               CalendarTitleForeColor=   16777215
               CheckBox        =   -1  'True
               Format          =   55377921
               CurrentDate     =   39135
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Province/Sate"
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   5
               Left            =   405
               TabIndex        =   52
               Top             =   3330
               Width           =   1005
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "City"
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   4
               Left            =   1140
               TabIndex        =   51
               Top             =   2850
               Width           =   270
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Street"
               ForeColor       =   &H8000000D&
               Height          =   225
               Index           =   3
               Left            =   930
               TabIndex        =   50
               Top             =   1980
               Width           =   480
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Other Phone"
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   17
               Left            =   510
               TabIndex        =   58
               Top             =   4380
               Width           =   900
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Email"
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   12
               Left            =   1050
               TabIndex        =   57
               Top             =   5490
               Width           =   360
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Fax"
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   11
               Left            =   1140
               TabIndex        =   56
               Top             =   5115
               Width           =   270
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cell Phone"
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   10
               Left            =   660
               TabIndex        =   55
               Top             =   4755
               Width           =   750
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Business Phone"
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   9
               Left            =   240
               TabIndex        =   54
               Top             =   4020
               Width           =   1170
            End
            Begin VB.Label lblCap 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Home Phone"
               ForeColor       =   &H8000000D&
               Height          =   210
               Index           =   8
               Left            =   510
               TabIndex        =   53
               Top             =   3660
               Width           =   900
            End
            Begin XtremeShortcutBar.ShortcutCaption CapInfo 
               Height          =   315
               Index           =   2
               Left            =   45
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   120
               Width           =   5295
               _Version        =   655364
               _ExtentX        =   9340
               _ExtentY        =   556
               _StockProps     =   14
               Caption         =   "Personal Information"
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.01
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SubItemCaption  =   -1  'True
               ForeColor       =   64
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "MI"
               ForeColor       =   &H8000000D&
               Height          =   225
               Left            =   1020
               TabIndex        =   45
               Top             =   1290
               Width           =   180
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Last Name"
               ForeColor       =   &H8000000D&
               Height          =   210
               Left            =   300
               TabIndex        =   44
               Top             =   900
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "First Name"
               ForeColor       =   &H8000000D&
               Height          =   210
               Left            =   330
               TabIndex        =   41
               Top             =   510
               Width           =   915
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sex"
               ForeColor       =   &H8000000D&
               Height          =   210
               Left            =   945
               TabIndex        =   48
               Top             =   1710
               Width           =   285
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date of Birth"
               ForeColor       =   &H8000000D&
               Height          =   210
               Left            =   2745
               TabIndex        =   47
               Top             =   1665
               Width           =   915
            End
         End
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1305
            Left            =   5490
            TabIndex        =   77
            Top             =   5250
            Width           =   5235
            Begin VB.TextBox txtNotes 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1230
               Left            =   30
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   79
               Top             =   420
               Width           =   5160
            End
            Begin XtremeShortcutBar.ShortcutCaption CapInfo 
               Height          =   315
               Index           =   1
               Left            =   30
               TabIndex        =   78
               TabStop         =   0   'False
               Top             =   120
               Width           =   5160
               _Version        =   655364
               _ExtentX        =   9102
               _ExtentY        =   556
               _StockProps     =   14
               Caption         =   "Notes"
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.01
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               SubItemCaption  =   -1  'True
               ForeColor       =   64
            End
         End
         Begin VB.Label lblCap 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filter:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   225
            Index           =   26
            Left            =   -63220
            TabIndex        =   146
            Top             =   630
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lblCap 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filter:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   225
            Index           =   25
            Left            =   -63220
            TabIndex        =   144
            Top             =   630
            Visible         =   0   'False
            Width           =   465
         End
      End
      Begin VB.Label Label29 
         Caption         =   "Mr. Juan De La Curz Maria Esabel Granada Vitorez"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   128
         Top             =   90
         Width           =   6105
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Index           =   0
         Left            =   6330
         TabIndex        =   1
         Top             =   150
         Width           =   1185
      End
   End
   Begin VB.PictureBox picChildren 
      Appearance      =   0  'Flat
      BackColor       =   &H00EDDADC&
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   4800
      ScaleHeight     =   2055
      ScaleWidth      =   4395
      TabIndex        =   102
      Top             =   1770
      Visible         =   0   'False
      Width           =   4425
      Begin MSComCtl2.DTPicker dtChild_DOB 
         Height          =   345
         Left            =   1320
         TabIndex        =   109
         Top             =   1170
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   55377921
         CurrentDate     =   39156
      End
      Begin VB.ComboBox cboChild_Gender 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "ProfilePersonal.frx":499B
         Left            =   1320
         List            =   "ProfilePersonal.frx":49A8
         TabIndex        =   107
         Top             =   810
         Width           =   2925
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   3450
         TabIndex        =   111
         Top             =   1590
         Width           =   795
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   345
         Left            =   2610
         TabIndex        =   110
         Top             =   1590
         Width           =   795
      End
      Begin VB.TextBox txtChild_Name 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   105
         Top             =   420
         Width           =   2925
      End
      Begin XtremeShortcutBar.ShortcutCaption CapInfo 
         Height          =   315
         Index           =   3
         Left            =   0
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   0
         Width           =   4395
         _Version        =   655364
         _ExtentX        =   7752
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Childrens"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   64
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   660
         TabIndex        =   106
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   240
         TabIndex        =   108
         Top             =   1260
         Width           =   1050
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Child Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   345
         TabIndex        =   104
         Top             =   480
         Width           =   960
      End
   End
   Begin VB.PictureBox picMultipleContact 
      Appearance      =   0  'Flat
      BackColor       =   &H00C2FAE2&
      ForeColor       =   &H80000008&
      Height          =   2985
      Left            =   5100
      ScaleHeight     =   2955
      ScaleWidth      =   4845
      TabIndex        =   112
      Top             =   2100
      Visible         =   0   'False
      Width           =   4875
      Begin VB.CommandButton Command1 
         Caption         =   "NEW"
         Height          =   345
         Left            =   3690
         TabIndex        =   115
         Top             =   360
         Width           =   795
      End
      Begin VB.TextBox txtContact_Tel 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1740
         MaxLength       =   8
         TabIndex        =   121
         Top             =   1560
         Width           =   2925
      End
      Begin VB.ComboBox cboContactName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1740
         TabIndex        =   117
         Top             =   810
         Width           =   2925
      End
      Begin VB.CommandButton cmdCancelContact 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   3840
         TabIndex        =   125
         Top             =   2400
         Width           =   795
      End
      Begin VB.CommandButton cmdOkContact 
         Caption         =   "OK"
         Height          =   345
         Left            =   3030
         TabIndex        =   124
         Top             =   2400
         Width           =   795
      End
      Begin VB.ComboBox cboContact_Relation 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1740
         TabIndex        =   123
         Top             =   1980
         Width           =   2925
      End
      Begin VB.TextBox txtContact_Address 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1740
         MaxLength       =   8
         TabIndex        =   119
         Top             =   1200
         Width           =   2925
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Click on New If Contact is not on the List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   180
         TabIndex        =   114
         Top             =   390
         Width           =   3390
      End
      Begin XtremeShortcutBar.ShortcutCaption CapInfo 
         Height          =   315
         Index           =   7
         Left            =   0
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   0
         Width           =   7215
         _Version        =   655364
         _ExtentX        =   12726
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Contact Information"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   64
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Relation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   1005
         TabIndex        =   122
         Top             =   2010
         Width           =   690
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   90
         TabIndex        =   120
         Top             =   1650
         Width           =   1605
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   975
         TabIndex        =   118
         Top             =   1230
         Width           =   720
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   495
         TabIndex        =   116
         Top             =   840
         Width           =   1200
      End
   End
   Begin VB.PictureBox picVehicles 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDAE80&
      ForeColor       =   &H80000008&
      Height          =   4665
      Left            =   3900
      ScaleHeight     =   4635
      ScaleWidth      =   4965
      TabIndex        =   80
      Top             =   1320
      Visible         =   0   'False
      Width           =   4995
      Begin VB.TextBox txtVSerial 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1830
         MaxLength       =   18
         TabIndex        =   93
         Top             =   2310
         Width           =   2925
      End
      Begin VB.TextBox txtVWar_Cert 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1830
         MaxLength       =   15
         TabIndex        =   97
         Top             =   3090
         Width           =   2925
      End
      Begin VB.TextBox txtVTin_Number 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1830
         MaxLength       =   15
         TabIndex        =   99
         Top             =   3480
         Width           =   2925
      End
      Begin VB.TextBox dtVPurchased 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1830
         MaxLength       =   18
         TabIndex        =   95
         Top             =   2700
         Width           =   2925
      End
      Begin VB.TextBox txtVProdNo 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   83
         Top             =   390
         Width           =   2925
      End
      Begin VB.TextBox txtVCond_No 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1830
         MaxLength       =   8
         TabIndex        =   85
         Top             =   780
         Width           =   2925
      End
      Begin VB.TextBox txtVEngine 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1830
         MaxLength       =   18
         TabIndex        =   91
         Top             =   1920
         Width           =   2925
      End
      Begin VB.ComboBox cboVModel 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1830
         TabIndex        =   87
         Top             =   1170
         Width           =   2925
      End
      Begin VB.ComboBox cboColorVinf 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1830
         TabIndex        =   89
         Top             =   1560
         Width           =   2925
      End
      Begin VB.CommandButton cmdOkVehicles 
         Caption         =   "OK"
         Height          =   345
         Left            =   3120
         TabIndex        =   100
         Top             =   3990
         Width           =   795
      End
      Begin VB.CommandButton cmdCancelVehicles 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   3960
         TabIndex        =   101
         Top             =   3990
         Width           =   795
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   540
         TabIndex        =   92
         Top             =   2400
         Width           =   1245
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Purchased"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   390
         TabIndex        =   94
         Top             =   2730
         Width           =   1395
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Warranty Certificate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   30
         TabIndex        =   96
         Top             =   3120
         Width           =   1755
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "TIN Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   780
         TabIndex        =   98
         Top             =   3510
         Width           =   1005
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Product Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   360
         TabIndex        =   82
         Top             =   420
         Width           =   1425
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Conduction Sticker"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   150
         TabIndex        =   84
         Top             =   810
         Width           =   1635
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   1260
         TabIndex        =   86
         Top             =   1230
         Width           =   525
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   1320
         TabIndex        =   88
         Top             =   1620
         Width           =   465
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Engine Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   480
         TabIndex        =   90
         Top             =   1950
         Width           =   1305
      End
      Begin XtremeShortcutBar.ShortcutCaption CapInfo 
         Height          =   315
         Index           =   6
         Left            =   0
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   0
         Width           =   4965
         _Version        =   655364
         _ExtentX        =   8758
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Vehicles Information"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   64
      End
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuTriad 
         Caption         =   "Add"
         Index           =   0
      End
      Begin VB.Menu mnuTriad 
         Caption         =   "Edit"
         Index           =   1
      End
      Begin VB.Menu mnuTriad 
         Caption         =   "Delete"
         Index           =   2
      End
      Begin VB.Menu mnuTriad 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuTriad 
         Caption         =   "Make it Preferred @CON"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmCRIS_Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddEntryContacts_Click()
    frmCRIS_CustomerContact.Show 1
End Sub

Private Sub Command5_Click()
    frmCRIS_CustomerChild.Show 1
    
End Sub
