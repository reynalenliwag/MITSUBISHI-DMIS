VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmNew_CustomerLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers A/R Ledger"
   ClientHeight    =   8385
   ClientLeft      =   315
   ClientTop       =   435
   ClientWidth     =   11850
   ForeColor       =   &H00FFFFFF&
   Icon            =   "New_CustomersLedger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   11850
   Begin XtremeReportControl.ReportControl rptRO 
      Height          =   5895
      Left            =   2670
      TabIndex        =   42
      Top             =   1590
      Width           =   9135
      _Version        =   655364
      _ExtentX        =   16113
      _ExtentY        =   10398
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnReorder=   0   'False
      MultipleSelection=   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   3630
      ScaleHeight     =   1215
      ScaleWidth      =   7275
      TabIndex        =   43
      Top             =   3480
      Visible         =   0   'False
      Width           =   7305
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   405
         Left            =   30
         TabIndex        =   46
         Top             =   450
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   135
         Left            =   0
         TabIndex        =   48
         Top             =   1110
         Width           =   7275
         _Version        =   655364
         _ExtentX        =   12832
         _ExtentY        =   238
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         GradientColorLight=   16744576
         GradientColorDark=   0
         ForeColor       =   16777215
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   255
         Left            =   0
         TabIndex        =   47
         Top             =   -30
         Width           =   7275
         _Version        =   655364
         _ExtentX        =   12832
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "Processing AR..  Please Wait...."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         GradientColorLight=   0
         GradientColorDark=   16744576
         ForeColor       =   16777215
      End
      Begin VB.Label labPercent 
         BackStyle       =   0  'Transparent
         Caption         =   "labPercent"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   345
         Left            =   60
         TabIndex        =   45
         Top             =   240
         Width           =   4755
      End
      Begin VB.Label labVoucherno 
         BackStyle       =   0  'Transparent
         Caption         =   "labVoucherno"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   285
         Left            =   60
         TabIndex        =   44
         Top             =   900
         Width           =   4755
      End
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11280
      Picture         =   "New_CustomersLedger.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   60
      Width           =   495
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   375
      Left            =   8070
      TabIndex        =   29
      Top             =   60
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   55836673
      CurrentDate     =   39765
   End
   Begin VB.ComboBox cboAccountName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "New_CustomersLedger.frx":138C
      Left            =   1620
      List            =   "New_CustomersLedger.frx":138E
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   60
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   2655
      TabIndex        =   3
      Top             =   480
      Width           =   9135
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   5
         Top             =   180
         Width           =   1815
      End
      Begin VB.TextBox txtCode3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   2850
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "000"
         Top             =   180
         Width           =   435
      End
      Begin VB.TextBox txtCode2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "00"
         Top             =   180
         Width           =   345
      End
      Begin VB.TextBox txtCode1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "000"
         Top             =   180
         Width           =   435
      End
      Begin VB.TextBox txtCustName 
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   1620
         MaxLength       =   35
         TabIndex        =   14
         Top             =   570
         Width           =   7440
      End
      Begin VB.Label labID 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3810
         TabIndex        =   8
         Top             =   180
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         TabIndex        =   13
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2670
         TabIndex        =   11
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2070
         TabIndex        =   10
         Top             =   240
         Width           =   135
      End
      Begin VB.Label labIDprev 
         Caption         =   "IDprev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2220
         TabIndex        =   9
         Top             =   210
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         TabIndex        =   4
         Top             =   210
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7845
      Left            =   30
      TabIndex        =   0
      Top             =   480
      Width           =   2595
      Begin VB.OptionButton optVendor 
         Caption         =   "Vendor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   41
         Top             =   420
         Width           =   2415
      End
      Begin VB.OptionButton optCustomer 
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   40
         Top             =   120
         Width           =   2415
      End
      Begin VB.TextBox TextSearch 
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   90
         MaxLength       =   35
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   2415
      End
      Begin MSComctlLib.ListView lstCustomer 
         Height          =   6675
         Left            =   60
         TabIndex        =   2
         Top             =   1110
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   11774
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
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "New_CustomersLedger.frx":1390
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   11085
      MouseIcon       =   "New_CustomersLedger.frx":14F2
      MousePointer    =   99  'Custom
      Picture         =   "New_CustomersLedger.frx":1644
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Exit Window"
      Top             =   7515
      Width           =   705
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   10395
      MouseIcon       =   "New_CustomersLedger.frx":19AA
      MousePointer    =   99  'Custom
      Picture         =   "New_CustomersLedger.frx":1AFC
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Print this Record"
      Top             =   7515
      Width           =   705
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   9705
      MouseIcon       =   "New_CustomersLedger.frx":1E62
      MousePointer    =   99  'Custom
      Picture         =   "New_CustomersLedger.frx":1FB4
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Find a Record"
      Top             =   7515
      Width           =   705
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   9015
      MouseIcon       =   "New_CustomersLedger.frx":22AE
      MousePointer    =   99  'Custom
      Picture         =   "New_CustomersLedger.frx":2400
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Move to Next Record"
      Top             =   7515
      Width           =   705
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Prev"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   8325
      MouseIcon       =   "New_CustomersLedger.frx":2758
      MousePointer    =   99  'Custom
      Picture         =   "New_CustomersLedger.frx":28AA
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Move to Previous Record"
      Top             =   7515
      Width           =   705
   End
   Begin VB.Frame fraDetails 
      Height          =   5895
      Left            =   2670
      TabIndex        =   15
      Top             =   1515
      Width           =   9135
      Begin XtremeReportControl.ReportControl rptLedger 
         Height          =   4755
         Left            =   9240
         TabIndex        =   39
         Top             =   630
         Width           =   8955
         _Version        =   655364
         _ExtentX        =   15796
         _ExtentY        =   8387
         _StockProps     =   64
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   30
         ScaleHeight     =   525
         ScaleWidth      =   9075
         TabIndex        =   22
         Top             =   5340
         Width           =   9075
         Begin VB.TextBox txtTotalBalance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
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
            Height          =   360
            Left            =   7080
            MaxLength       =   20
            TabIndex        =   25
            Top             =   90
            Width           =   1785
         End
         Begin VB.TextBox txtTotalDebit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
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
            Height          =   360
            Left            =   4320
            MaxLength       =   20
            TabIndex        =   24
            Top             =   90
            Width           =   1395
         End
         Begin VB.TextBox txtTotalCredit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
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
            Height          =   360
            Left            =   5700
            MaxLength       =   20
            TabIndex        =   23
            Top             =   90
            Width           =   1395
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3210
            TabIndex        =   26
            Top             =   120
            Width           =   1395
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdAccountsLedger 
         Height          =   4635
         Left            =   60
         TabIndex        =   16
         Top             =   630
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   8176
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         ForeColor       =   0
         ForeColorFixed  =   0
         BackColorSel    =   16744448
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483633
         AllowBigSelection=   0   'False
         TextStyleFixed  =   3
         FocusRect       =   0
         HighLight       =   2
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "New_CustomersLedger.frx":2C09
      End
      Begin VB.TextBox txtSearch 
         Height          =   405
         Left            =   60
         TabIndex        =   36
         Top             =   630
         Width           =   6225
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   37
         Top             =   660
         Width           =   1065
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "New_CustomersLedger.frx":2F23
         Left            =   6360
         List            =   "New_CustomersLedger.frx":2F30
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   630
         Width           =   1545
      End
      Begin MSComctlLib.ListView lvwLedger 
         Height          =   5175
         Left            =   60
         TabIndex        =   31
         Top             =   120
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   9128
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "DOCDATE"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "REFERENCE"
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "INVOICE#/OR"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "DEBIT"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "CREDIT"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "BALANCE"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "JTYPE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   375
      Left            =   9840
      TabIndex        =   30
      Top             =   60
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   55836673
      CurrentDate     =   39765
   End
   Begin VB.Label Label8 
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9510
      TabIndex        =   34
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label7 
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7530
      TabIndex        =   33
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label3 
      Caption         =   "F3 - View Cash Receipts Voucher"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   2700
      TabIndex        =   32
      Top             =   8430
      Width           =   4365
   End
   Begin VB.Label Label 
      Caption         =   "Account Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   27
      Top             =   120
      Width           =   2625
   End
End
Attribute VB_Name = "frmNew_CustomerLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCustomer                                    As ADODB.Recordset
Dim rsJournal_HD                                  As ADODB.Recordset
Dim rsJournal_HDDet                               As ADODB.Recordset
Dim AddorEdit, ORDER_BY                           As String
Attribute ORDER_BY.VB_VarUserMemId = 1073938435
Dim TUTAL_DEBIT, TUTAL_CREDIT, TUTAL_BALANCE      As Double
Attribute TUTAL_DEBIT.VB_VarUserMemId = 1073938437
Attribute TUTAL_CREDIT.VB_VarUserMemId = 1073938437
Attribute TUTAL_BALANCE.VB_VarUserMemId = 1073938437
Dim LocalAcess                                    As String
Attribute LocalAcess.VB_VarUserMemId = 1073938440

Dim rsCUSTOMER_OPENING                            As ADODB.Recordset
Attribute rsCUSTOMER_OPENING.VB_VarUserMemId = 1073938441

'THIS IS FOR NEW CUSTOMER A/R LEDGER --------------------------------------------------
Dim REC                                           As XtremeReportControl.ReportRecord
Attribute REC.VB_VarUserMemId = 1073938442
Dim XX_DOCDATE                                    As String
Attribute XX_DOCDATE.VB_VarUserMemId = 1073938443
Dim XX_Reference                                  As String
Attribute XX_Reference.VB_VarUserMemId = 1073938444
Dim XX_INVOICE_NO                                 As String
Attribute XX_INVOICE_NO.VB_VarUserMemId = 1073938445
Dim XX_DEBIT                                      As Double
Attribute XX_DEBIT.VB_VarUserMemId = 1073938446
Dim XX_CREDIT                                     As Double
Attribute XX_CREDIT.VB_VarUserMemId = 1073938447
Dim XX_BALANCE                                    As Double
Attribute XX_BALANCE.VB_VarUserMemId = 1073938448
Dim XX_ID                                         As Long
Attribute XX_ID.VB_VarUserMemId = 1073938449
Dim G_TOTAL_DEBIT                                 As Double
Attribute G_TOTAL_DEBIT.VB_VarUserMemId = 1073938450
Dim G_TOTAL_CREDIT                                As Double
Attribute G_TOTAL_CREDIT.VB_VarUserMemId = 1073938451

Dim FORWARDED_BALANCE                             As Double
Attribute FORWARDED_BALANCE.VB_VarUserMemId = 1073938452
Dim POSI_OR_NEGA                                  As Boolean
Attribute POSI_OR_NEGA.VB_VarUserMemId = 1073938453

'THIS IS FOR NEW CUSTOMER A/R LEDGER --------------------------------------------------

Function SetCustomerName(VVV As Variant)
    Dim rsCustomerDup                             As ADODB.Recordset
    Set rsCustomerDup = New ADODB.Recordset
    rsCustomerDup.Open "Select CustCode,Custname from ALL_CUSTMASTER_AMIS where CustCode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCustomerDup.EOF And Not rsCustomerDup.BOF Then SetCustomerName = Null2String(rsCustomerDup!CUSTNAME) Else SetCustomerName = ""
End Function

Sub rsRefresh()
    Set rsCustomer = New ADODB.Recordset
    If cboAccountName.Text = "ALL ACCOUNTS" Then
        If optCustomer.Value = True Then
            'DESCRIPTION: THIS IS FOR CUSTOMER WITH AR ACCOUNT SCHEDULE
            'rsCUSTOMER.Open "SELECT DISTINCT dbo.ALL_Customer.ACCTNAME as CUSTNAME,dbo.ALL_Customer.ID,dbo.ALL_Customer.CUSCDE AS CUSTCODE FROM dbo.AMIS_Journal_HD left outer JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType INNER JOIN dbo.ALL_Customer ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer.CUSCDE WHERE ((LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') or (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03')) ORDER BY dbo.ALL_Customer.ACCTNAME", gconDMIS, adOpenKeyset
            rsCustomer.Open "SELECT DISTINCT CUST.ACCTNAME as CUSTNAME,CUST.ID,CUST.CUSCDE AS CUSTCODE " & _
                            "FROM AMIS_Journal_HD HD left outer JOIN AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo " & _
                            "AND HD.JType = DET.JType INNER JOIN ALL_Customer CUST ON " & _
                            "((HD.CustomerCode = CUST.CUSCDE) OR  ((RIGHT(DET.ENTITY,6)) = CUST.CUSCDE)) " & _
                            "WHERE ((LEFT(DET.Acct_Code, 5) = '11-02') OR (LEFT(DET.Acct_Code, 5) = '11-03')) ORDER BY CUST.ACCTNAME", gconDMIS, adOpenKeyset
        Else
            'DESCRIPTION: THIS IS FOR VENDOR WITH AR ACCOUNT SCHEDULE
            rsCustomer.Open "SELECT DISTINCT ALL_VENDOR.NameofVendor as VENDORNAME,ALL_VENDOR.ID as ID,ALL_VENDOR.CODE AS VENCODE " & _
                            "FROM dbo.AMIS_Journal_HD HD INNER JOIN dbo.AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo AND HD.JType = DET.JType " & _
                            "Left outer JOIN dbo.ALL_VENDOR_TABLE ALL_VENDOR ON HD.VENDORCODE = ALL_VENDOR.CODE WHERE  ALL_VENDOR.CODE IS NOT NULL  and " & _
                            "((LEFT(DET.Acct_Code, 5) = '11-02') OR (LEFT(DET.Acct_Code, 5) = '11-03')) AND ALL_VENDOR.CODE <> '999999'  ORDER BY ALL_VENDOR.VENDORNAME", gconDMIS, adOpenKeyset
        End If
    Else
        If optCustomer.Value = True Then
            'DESCRIPTION: THIS IS FOR CUSTOMER WITH AR ACCOUNT GROUP BY ACCOUNT CODE
            'rsCUSTOMER.Open "SELECT DISTINCT  dbo.ALL_Customer.ACCTNAME as CUSTNAME,dbo.ALL_Customer.ID,dbo.ALL_Customer.CUSCDE AS CUSTCODE FROM dbo.AMIS_Journal_HD left outer JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType INNER JOIN dbo.ALL_Customer ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer.CUSCDE WHERE ((LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') or (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03')) AND dbo.AMIS_Journal_Det.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "' ORDER BY dbo.ALL_Customer.ACCTNAME", gconDMIS, adOpenKeyset
            rsCustomer.Open "SELECT DISTINCT CUST.ACCTNAME as CUSTNAME,CUST.ID,CUST.CUSCDE AS CUSTCODE " & _
                            "FROM AMIS_Journal_HD HD left outer JOIN AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo " & _
                            "AND HD.JType = DET.JType INNER JOIN ALL_Customer CUST ON " & _
                            "((HD.CustomerCode = CUST.CUSCDE) OR  ((RIGHT(DET.ENTITY,6)) = CUST.CUSCDE)) " & _
                            "WHERE ((LEFT(DET.Acct_Code, 5) = '11-02') OR (LEFT(DET.Acct_Code, 5) = '11-03')) and DET.ACCT_CODE =  '" & Setacctcode(cboAccountName.Text) & "' ORDER BY CUST.ACCTNAME", gconDMIS, adOpenKeyset
        Else
            'DESCRIPTION: THIS IS FOR VENDOR WITH AR ACCOUNT GROUP BY ACCOUNT CODE
            rsCustomer.Open "SELECT DISTINCT ALL_VENDOR.NameofVendor as VENDORNAME,ALL_VENDOR.ID as ID,ALL_VENDOR.CODE AS VENCODE " & _
                            "FROM dbo.AMIS_Journal_HD HD INNER JOIN dbo.AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo AND HD.JType = DET.JType " & _
                            "Left outer JOIN dbo.ALL_VENDOR_TABLE ALL_VENDOR ON HD.VENDORCODE = ALL_VENDOR.CODE WHERE  ALL_VENDOR.CODE IS NOT NULL  and DET.ACCT_CODE = '" & Setacctcode(cboAccountName.Text) & "'  and " & _
                            "((LEFT(DET.Acct_Code, 5) = '11-02') OR (LEFT(DET.Acct_Code, 5) = '11-03')) AND ALL_VENDOR.CODE <> '999999' ORDER BY ALL_VENDOR.VENDORNAME", gconDMIS, adOpenKeyset
        End If
    End If
End Sub

Sub initMemvars()
    Frame1.Enabled = True
    txtCode.Text = "": txtCode1.Text = "": txtCode2.Text = "": txtCode3.Text = ""
    txtCustName.Text = "":
    txtTotalDebit.Text = ZERO: txtTotalCredit.Text = ZERO
    txtTotalBalance.Text = ZERO:
End Sub

Sub StoreMemVars()
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        Frame1.Enabled = False
        If optCustomer.Value = True Then
            labID.Caption = rsCustomer!ID
            txtCode.Text = Null2String(rsCustomer!CUSTCODE)
            txtCustName.Text = Null2String(rsCustomer!CUSTNAME)
        Else
            labID.Caption = rsCustomer!ID
            txtCode.Text = Null2String(rsCustomer!VENCODE)
            txtCustName.Text = Null2String(rsCustomer!VendorName)
        End If
        'UPDATED BY: JUN---------
        'DATE UPDATED: 06-10-2009
        'GET_BALANCE
        'UPDATED BY: JUN---------
        'FillGrids
        'FILL_LEDGER
        POPULATE_REPORT
    End If
End Sub

Sub InitGrid()
    With grdAccountsLedger
        .Rows = 2
        .ColWidth(0) = 1200: .ColWidth(1) = 1300: .ColWidth(2) = 2000
        .ColWidth(3) = 1400: .ColWidth(4) = 1400: .ColWidth(5) = 1400
        .ColWidth(6) = 1: .ColWidth(7) = 1: .Row = 0
        .Col = 0: .Text = "DOCDATE"
        .Col = 1: .Text = "REFERENCE"
        .Col = 2: .Text = "INVOICE#/OR'"
        .Col = 3: .Text = "DEBIT"
        .Col = 4: .Text = "CREDIT"
        .Col = 5: .Text = "BALANCE"
        .Col = 6: .Text = "ID"
        .Col = 7: .Text = "JTYPE"
    End With
End Sub
Sub GET_BALANCE()
'UPDATED BY: JUN
'DATE UPDATED: 06/09/2009
'DESCRIPTION: COMPUTING THE THE CUSTOMER OPENING BALANCE TO BE FORWARDED

    If cboAccountName.Text = "ALL ACCOUNTS" Then
        Dim xCOB                                  As Double
        Dim rsCHECK_COB                           As ADODB.Recordset
        Set rsCHECK_COB = New ADODB.Recordset
        Set rsCUSTOMER_OPENING = New ADODB.Recordset
        xCOB = 0
        xBALANCE = 0


        If optCustomer.Value = True Then
            'THIS IS FOR CUSTOMER
            rsCHECK_COB.Open "SELECT HD.JTYPE as xJTYPE,HD.INVOICEAMT as xINVOICE_AMT FROM AMIS_JOURNAL_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET DET ON DET.JNO  = HD.JNO AND DET.JTYPE = HD.JTYPE WHERE (HD.JDATE < '" & dtFrom & "'  AND HD.STATUS = 'P' AND HD.CUSTOMERCODE = '" & txtCode.Text & "') " & _
                             "AND ((LEFT(DET.ACCT_CODE,5) in ('11-02' , '11-03' ,'21-02' ,'21-07'))OR ((HD.JTYPE = 'CCM' AND (LEFT(DET.ACCT_CODE,5) in ('11-02','21-01'))) OR (HD.JTYPE = 'CSJ' AND (LEFT(DET.ACCT_CODE,5) in ( '11-03' , '21-07'))) ))", gconDMIS

            rsCUSTOMER_OPENING.Open "SELECT ROUND(SUM(DET.DEBIT) - SUM(DET.CREDIT),2) AS CUST_BALANCE FROM AMIS_JOURNAL_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET DET ON DET.JNO  = HD.JNO AND DET.JTYPE = HD.JTYPE WHERE (HD.JDATE < '" & dtFrom & "'  AND HD.STATUS = 'P' AND HD.JTYPE <> 'GJ' AND HD.CUSTOMERCODE = '" & txtCode.Text & "') " & _
                                    "AND ((LEFT(DET.ACCT_CODE,5) in ('11-02' , '11-03' ,'21-02' ,'21-07'))OR ((HD.JTYPE = 'CCM' AND (LEFT(DET.ACCT_CODE,5) in ('11-02','21-01'))) OR (HD.JTYPE = 'CSJ' AND (LEFT(DET.ACCT_CODE,5) in ( '11-03' , '21-07'))) ))", gconDMIS
        Else
            'THIS IS FOR VENDOR
            rsCHECK_COB.Open "SELECT HD.JTYPE as xJTYPE,HD.INVOICEAMT as xINVOICE_AMT FROM AMIS_JOURNAL_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET DET ON DET.JNO  = HD.JNO AND DET.JTYPE = HD.JTYPE WHERE (HD.JDATE < '" & dtFrom & "'  AND HD.STATUS = 'P' AND HD.VENDORCODE = '" & txtCode.Text & "') " & _
                             "AND ((LEFT(DET.ACCT_CODE,5) in ('11-02' , '11-03' ,'21-02' ,'21-07'))OR ((HD.JTYPE = 'CCM' AND (LEFT(DET.ACCT_CODE,5) in ('11-02','21-01'))) OR (HD.JTYPE = 'CSJ' AND (LEFT(DET.ACCT_CODE,5) in ( '11-03' , '21-07'))) ))", gconDMIS

            rsCUSTOMER_OPENING.Open "SELECT ROUND(SUM(DET.DEBIT) - SUM(DET.CREDIT),2) AS CUST_BALANCE FROM AMIS_JOURNAL_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET DET ON DET.JNO  = HD.JNO AND DET.JTYPE = HD.JTYPE WHERE (HD.JDATE < '" & dtFrom & "'  AND HD.STATUS = 'P' AND HD.JTYPE <> 'GJ' AND HD.VENDORCODE = '" & txtCode.Text & "') " & _
                                    "AND ((LEFT(DET.ACCT_CODE,5) in ('11-02' , '11-03' ,'21-02' ,'21-07'))OR ((HD.JTYPE = 'CCM' AND (LEFT(DET.ACCT_CODE,5) in ('11-02','21-01'))) OR (HD.JTYPE = 'CSJ' AND (LEFT(DET.ACCT_CODE,5) in ( '11-03' , '21-07'))) ))", gconDMIS
        End If

        If Not rsCHECK_COB.EOF And Not rsCHECK_COB.BOF Then
            Do While Not rsCHECK_COB.EOF
                If Null2String(rsCHECK_COB!xJType) = "COB" Then
                    xCOB = xCOB + ToDoubleNumber(rsCHECK_COB!xINVOICE_AMT)
                End If
                rsCHECK_COB.MoveNext
            Loop
        End If

        If Not rsCUSTOMER_OPENING.BOF And Not rsCUSTOMER_OPENING.EOF Then
            If Null2String(rsCUSTOMER_OPENING!CUST_BALANCE) = "" Then
                xBALANCE = ToDoubleNumber(0) + xCOB
            Else
                xBALANCE = ToDoubleNumber(rsCUSTOMER_OPENING!CUST_BALANCE) + xCOB
            End If
        Else
            xBALANCE = ToDoubleNumber(0)
        End If

        'DESCRIPITION: CHECK IF THERE IS AN ADJUSTMENT MADE IF FOUND DO THE COMPUTATION
        Dim rsADJ                                 As ADODB.Recordset
        Dim xADJ_AMOUNT                           As Double
        xADJ_AMOUNT = xBALANCE
        Set rsADJ = New ADODB.Recordset
        'rsADJ.Open "SELECT HD.ADJ_AMOUNT AS ADJ_AMOUNT, HD.ADJ_TYPE AS ADJ_TYPE " & _
         "FROM AMIS_JOURNAL_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET DET ON DET.VOUCHERNO  = HD.VOUCHERNO AND DET.JTYPE = HD.JTYPE " & _
         "WHERE (HD.JDATE < '" & dtFrom & "' AND HD.STATUS = 'P' AND HD.CUSTOMERCODE = '" & txtCode.Text & "') AND ((LEFT(DET.ACCT_CODE,5)in('11-02' , '11-03' ,'21-02' ,'21-07')) " & _
         "OR ((HD.JTYPE = 'GJ'AND(LEFT(DET.ACCT_CODE,5)in ('11-02','21-01','11-03','21-07')))))AND HD.JTYPE = 'GJ'", gconDMIS, adOpenKeyset
        rsADJ.Open "SELECT DET.DEBIT AS DET_DEBIT, DET.CREDIT AS DET_CREDIT FROM AMIS_JOURNAL_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET DET ON DET.VOUCHERNO  = HD.VOUCHERNO AND DET.JTYPE = HD.JTYPE " & _
                   "WHERE ((HD.JDATE < '" & dtFrom & "' AND HD.STATUS = 'P')  AND (HD.JTYPE = 'GJ'AND(LEFT(DET.ACCT_CODE,5)in ('11-02','21-01','11-03','21-07'))))AND HD.JTYPE = 'GJ' AND (RIGHT(DET.ENTITY,6) = '" & txtCode.Text & "') AND ADJ_JTYPE <> 'APJ'", gconDMIS, adOpenKeyset

        If Not rsADJ.EOF And Not rsADJ.BOF Then
            Do While Not rsADJ.EOF
                If NumericVal(rsADJ!DET_CREDIT) <> 0 Then
                    xADJ_AMOUNT = NumericVal(xADJ_AMOUNT) - NumericVal(rsADJ!DET_CREDIT)
                ElseIf NumericVal(rsADJ!DET_DEBIT) <> 0 Then
                    xADJ_AMOUNT = NumericVal(xADJ_AMOUNT) + NumericVal(rsADJ!DET_DEBIT)
                End If
                rsADJ.MoveNext
            Loop
        End If
        Set rsADJ = Nothing

        xBALANCE = NumericVal(xADJ_AMOUNT)
    Else
        Dim rsCUSTOMER_OPENING_ACCT               As ADODB.Recordset
        Dim rsCHECK_COB_ACCT                      As ADODB.Recordset
        Dim xCOB_ACCT                             As Double


        'DESCRIPTION: THIS IS FOR COMPUTING THE FOR THE BALANCE NOT EQUAL TO GJ AND COB
        Set rsCUSTOMER_OPENING_ACCT = New ADODB.Recordset
        rsCUSTOMER_OPENING_ACCT.Open "SELECT ROUND(SUM(DET.DEBIT) - SUM(DET.CREDIT),2) AS CUST_BALANCE FROM AMIS_JOURNAL_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET DET ON DET.VOUCHERNO  = HD.VOUCHERNO AND DET.JTYPE = HD.JTYPE WHERE (HD.JDATE < '" & dtFrom & "'  AND HD.STATUS = 'P' AND HD.JTYPE <> 'GJ' AND DET.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "'  AND HD.CUSTOMERCODE = '" & txtCode.Text & "') " & _
                                     "AND ((LEFT(DET.ACCT_CODE,5) in ('11-02' , '11-03' ,'21-02' ,'21-07'))OR ((HD.JTYPE = 'CCM' AND (LEFT(DET.ACCT_CODE,5) in ('11-02','21-01'))) OR (HD.JTYPE = 'CSJ' AND (LEFT(DET.ACCT_CODE,5) in ( '11-03' , '21-07'))) ))", gconDMIS

        xCOB_ACCT = 0
        xBALANCE = 0

        'DESCRIPTION: THIS IS FOR GETTING THE CUSTOMER OPENING BALANCE
        Set rsCHECK_COB_ACCT = New ADODB.Recordset
        rsCHECK_COB_ACCT.Open "SELECT HD.JTYPE as xJTYPE,HD.INVOICEAMT as xINVOICE_AMT FROM AMIS_JOURNAL_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET DET ON DET.JNO  = HD.JNO AND DET.JTYPE = HD.JTYPE WHERE (HD.JDATE < '" & dtFrom & "'  AND HD.STATUS = 'P'AND DET.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "' AND HD.CUSTOMERCODE = '" & txtCode.Text & "') " & _
                              "AND ((LEFT(DET.ACCT_CODE,5) in ('11-02' , '11-03' ,'21-02' ,'21-07'))OR ((HD.JTYPE = 'CCM' AND (LEFT(DET.ACCT_CODE,5) in ('11-02','21-01'))) OR (HD.JTYPE = 'CSJ' AND (LEFT(DET.ACCT_CODE,5) in ( '11-03' , '21-07'))) ))", gconDMIS


        If Not rsCHECK_COB_ACCT.EOF And Not rsCHECK_COB_ACCT.BOF Then
            Do While Not rsCHECK_COB_ACCT.EOF
                If Null2String(rsCHECK_COB_ACCT!xJType) = "COB" Then
                    xCOB_ACCT = xCOB_ACCT + ToDoubleNumber(rsCHECK_COB_ACCT!xINVOICE_AMT)
                End If
                rsCHECK_COB_ACCT.MoveNext
            Loop
        End If

        If Not rsCUSTOMER_OPENING_ACCT.BOF And Not rsCUSTOMER_OPENING_ACCT.EOF Then
            If Null2String(rsCUSTOMER_OPENING_ACCT!CUST_BALANCE) = "" Then
                xBALANCE = ToDoubleNumber(0) + xCOB_ACCT
            Else
                xBALANCE = ToDoubleNumber(rsCUSTOMER_OPENING_ACCT!CUST_BALANCE) + xCOB_ACCT
            End If
        Else
            xBALANCE = ToDoubleNumber(0)
        End If

        'DESCRIPTION: CHECK IF THERE IS AN ADJUSTMENT MADE GROUP BY ACCOUNT
        Dim rsADJ2                                As ADODB.Recordset
        Dim xADJ_AMOUNT2                          As Double
        xADJ_AMOUNT2 = xBALANCE
        Set rsADJ2 = New ADODB.Recordset
        'rsADJ2.Open "SELECT DET.CREDIT AS DET_CREDIT, DET.DEBIT AS DET_DEBIT ,HD.ADJ_AMOUNT AS ADJ_AMOUNT, HD.ADJ_TYPE AS ADJ_TYPE " & _
         "FROM AMIS_JOURNAL_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET DET ON DET.VOUCHERNO  = HD.VOUCHERNO AND DET.JTYPE = HD.JTYPE " & _
         "WHERE (HD.JDATE < '" & dtFrom & "' AND HD.STATUS = 'P' AND DET.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "' AND HD.CUSTOMERCODE = '" & txtCode.Text & "') AND ((LEFT(DET.ACCT_CODE,5)in('11-02' , '11-03' ,'21-02' ,'21-07')) " & _
         "OR ((HD.JTYPE = 'GJ'AND(LEFT(DET.ACCT_CODE,5)in ('11-02','21-01','11-03','21-07')))))AND HD.JTYPE = 'ADJ'", gconDMIS, adOpenKeyset
        rsADJ2.Open "SELECT DET.DEBIT AS DET_DEBIT, DET.CREDIT AS DET_CREDIT FROM AMIS_JOURNAL_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET DET ON DET.VOUCHERNO  = HD.VOUCHERNO AND DET.JTYPE = HD.JTYPE " & _
                    "WHERE ((HD.JDATE < '" & dtFrom & "' AND HD.STATUS = 'P')  AND (HD.JTYPE = 'GJ'AND(LEFT(DET.ACCT_CODE,5)in ('11-02','21-01','11-03','21-07'))))AND HD.JTYPE = 'GJ' AND (RIGHT(DET.ENTITY,6) = '" & txtCode.Text & "') AND ADJ_JTYPE <> 'APJ'", gconDMIS, adOpenKeyset


        If Not rsADJ2.EOF And Not rsADJ2.BOF Then
            Do While Not rsADJ2.EOF
                If NumericVal(rsADJ2!DET_CREDIT) <> 0 Then
                    xADJ_AMOUNT2 = NumericVal(xADJ_AMOUNT2) - NumericVal(rsADJ2!DET_CREDIT)
                ElseIf NumericVal(rsADJ2!DET_DEBIT) <> 0 Then
                    xADJ_AMOUNT2 = NumericVal(xADJ_AMOUNT2) + NumericVal(rsADJ2!DET_DEBIT)
                End If
                rsADJ2.MoveNext
            Loop
        End If
        Set rsADJ2 = Nothing

        xBALANCE = NumericVal(xADJ_AMOUNT2)

    End If
    Set rsCUSTOMER_OPENING = Nothing
    Set rsCHECK_COB = Nothing
End Sub


Sub FillGrids()
    Dim OUTBALANCE                                As Double
    Dim Reference                                 As String
    Dim cnt                                       As Integer
    Dim CREDIT                                    As Double
    Dim DEBIT                                     As Double
    Dim cnt_adjusment                             As Integer
    Dim tmp_voucher                               As String


    Set rsCUSTOMER_OPENING = New ADODB.Recordset

    cleargrid grdAccountsLedger: InitGrid
    TUTAL_BALANCE = 0: TUTAL_BALANCE = TUTAL_BALANCE: cnt = 0: TUTAL_DEBIT = 0: TUTAL_CREDIT = 0: OUTBALANCE = 0
    cnt_adjusment = 0
    Set rsJournal_HDDet = New ADODB.Recordset
    If cboAccountName.Text = "ALL ACCOUNTS" Then
        'rsJournal_HDDet.Open "select AMIS_Journal_Hd.Debit AS DM,AMIS_Journal_Hd.Credit AS CM,AMIS_Journal_Hd.AmountToPay,AMIS_Journal_Hd.InvoiceAmt,AMIS_Journal_Hd.InvoiceNo,AMIS_Journal_HD.ID,AMIS_Journal_Hd.JNo,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.JType,AMIS_Journal_Det.Debit,AMIS_Journal_Det.Credit,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.CheckNo,AMIS_Journal_Hd.InvoiceType,AMIS_Journal_Hd.VendorCode,AMIS_Journal_Hd.JNo from AMIS_Journal_HD left outer Join AMIS_Journal_det on AMIS_Journal_Det.JNo  = AMIS_Journal_Hd.JNo and AMIS_Journal_Det.jtype = AMIS_Journal_Hd.Jtype  where " & _
         '                     "(dbo.AMIS_Journal_HD.Jdate >= '" & dtFrom & "' and dbo.AMIS_Journal_HD.Jdate <= '" & dtTo & "') and (AMIS_Journal_Hd.Status = 'P' AND ((Left(AMIS_Journal_Det.Acct_Code,5) = '11-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '11-03' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-07') OR ((AMIS_Journal_HD.JTYPE = 'CCM' AND AMIS_Journal_HD.Debit = 0) OR (AMIS_Journal_HD.JTYPE = 'CSJ' AND AMIS_Journal_HD.Credit = 0)))) AND AMIS_Journal_Hd.CustomerCode = '" & txtCode.Text & "' order by AMIS_Journal_Hd.jdate asc,AMIS_Journal_Hd.id asc", gconDMIS
        rsJournal_HDDet.Open "select AMIS_Journal_Hd.Debit AS DM,AMIS_Journal_Hd.Credit AS CM,AMIS_Journal_Hd.AmountToPay,AMIS_Journal_Hd.InvoiceAmt,AMIS_Journal_Hd.InvoiceNo,AMIS_Journal_HD.ID,AMIS_Journal_Hd.JNo,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.JType,AMIS_Journal_Det.Debit,AMIS_Journal_Det.Credit,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.CheckNo,AMIS_Journal_Hd.InvoiceType,AMIS_Journal_Hd.VendorCode,AMIS_Journal_Hd.JNo from AMIS_Journal_HD left outer Join AMIS_Journal_det on AMIS_Journal_Det.JNo  = AMIS_Journal_Hd.JNo and AMIS_Journal_Det.jtype = AMIS_Journal_Hd.Jtype  where " & _
                             "(dbo.AMIS_Journal_HD.Jdate >= '" & dtFrom & "' and dbo.AMIS_Journal_HD.Jdate <= '" & dtTo & "') and (AMIS_Journal_Hd.Status = 'P' AND ((Left(AMIS_Journal_Det.Acct_Code,5) = '11-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '11-03' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-07') OR ((AMIS_Journal_HD.JTYPE = 'CCM' AND (left(AMIS_Journal_det.acct_code,5) = '11-02' or left(AMIS_Journal_det.acct_code,5) = '21-01')) OR (AMIS_Journal_HD.JTYPE = 'CSJ' AND (left(AMIS_Journal_det.acct_code,5) = '11-03' or left(AMIS_Journal_det.acct_code,5) = '21-07'))))) AND AMIS_Journal_Hd.CustomerCode = '" & txtCode.Text & "' order by AMIS_Journal_Hd.jdate asc,AMIS_Journal_Hd.id asc", gconDMIS

    Else
        rsJournal_HDDet.Open "select AMIS_Journal_Hd.Debit AS DM,AMIS_Journal_Hd.Credit AS CM,AMIS_Journal_Hd.AmountToPay,AMIS_Journal_Hd.InvoiceAmt,AMIS_Journal_Hd.InvoiceNo,AMIS_Journal_HD.ID,AMIS_Journal_Hd.JNo,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.JType,AMIS_Journal_Det.Debit,AMIS_Journal_Det.Credit,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.CheckNo,AMIS_Journal_Hd.InvoiceType,AMIS_Journal_Hd.VendorCode,AMIS_Journal_Hd.JNo from AMIS_Journal_HD left outer Join AMIS_Journal_det on AMIS_Journal_Det.JNo  = AMIS_Journal_Hd.JNo and AMIS_Journal_Det.jtype = AMIS_Journal_Hd.Jtype  where " & _
                             "(dbo.AMIS_Journal_HD.Jdate >= '" & dtFrom & "' and dbo.AMIS_Journal_HD.Jdate <= '" & dtTo & "') and (AMIS_Journal_Hd.Status = 'P' AND AMIS_Journal_Det.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "' AND ((Left(AMIS_Journal_Det.Acct_Code,5) = '11-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '11-03' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-07') OR ((AMIS_Journal_HD.JTYPE = 'CCM' AND (left(AMIS_Journal_det.acct_code,5) = '11-01' or left(AMIS_Journal_det.acct_code,5) = '21-01')) OR (AMIS_Journal_HD.JTYPE = 'CSJ' AND (left(AMIS_Journal_det.acct_code,5) = '11-03' or left(AMIS_Journal_det.acct_code,5) = '21-07'))))) AND AMIS_Journal_Hd.CustomerCode = '" & txtCode.Text & "' order by AMIS_Journal_Hd.jdate asc,AMIS_Journal_Hd.id asc", gconDMIS
        'UPDATED BY: JUN
        'DATE UPDATED: 06/09/2009
        'DESCRIPTION: CUSTOMER OPENING BALANCE
        'rsCUSTOMER_OPENING.Open "SELECT ROUND(SUM(DET.DEBIT) - SUM(DET.CREDIT),2) AS CUST_BALANCE FROM AMIS_JOURNAL_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET DET ON DET.JNO  = HD.JNO AND DET.JTYPE = HD.JTYPE WHERE (HD.JDATE < '" & dtFrom & "'  AND HD.STATUS = 'P' AND HD.CUSTOMERCODE = '" & txtCode.Text & "') " & _
         "AND ((LEFT(DET.ACCT_CODE,5) in ('11-02' , '11-03' ,'21-02' ,'21-07'))OR ((HD.JTYPE = 'CCM' AND (LEFT(DET.ACCT_CODE,5) in ('11-02','21-01'))) OR (HD.JTYPE = 'CSJ' AND (LEFT(DET.ACCT_CODE,5) in ( '11-03' , '21-07'))) ))", gconDMIS
    End If

    If Not rsJournal_HDDet.EOF And Not rsJournal_HDDet.BOF Then
        rsJournal_HDDet.MoveFirst
        Do While Not rsJournal_HDDet.EOF
            Reference = Null2String(rsJournal_HDDet!jtype) & "-" & Null2String(rsJournal_HDDet!VOUCHERNO)
            'REFERENCE = Null2String(rsJournal_HDDet!InvoiceType) & "-" & Null2String(rsJournal_HDDet!InvoiceNo)
            cnt = cnt + 1

            'UPDATED BY: JUN--------------------------------------------------------
            'DATE UPDATED: 06-10-2009
            'DESCRIPTION: SUMMATION OF PREVIOUS BALANCE AND CUSTOMER OPENING BALANCE
            If cnt = 1 Then
                OUTBALANCE = OUTBALANCE + xBALANCE
            End If
            'UPDATED BY: JUN--------------------------------------------------------

            If Null2String(rsJournal_HDDet!jtype) = "COB" Then
                OUTBALANCE = OUTBALANCE + N2Str2Zero(rsJournal_HDDet!InvoiceAmt)
            Else
                If Null2String(rsJournal_HDDet!jtype) = "CCM" Then
                    OUTBALANCE = N2Str2Zero(OUTBALANCE) - N2Str2Zero(rsJournal_HDDet!CM)

                ElseIf Null2String(rsJournal_HDDet!jtype) = "CSJ" Then
                    OUTBALANCE = N2Str2Zero(OUTBALANCE) + N2Str2Zero(rsJournal_HDDet!DM)
                Else
                    OUTBALANCE = N2Str2Zero(OUTBALANCE) + ((N2Str2Zero(rsJournal_HDDet!DEBIT) - N2Str2Zero(rsJournal_HDDet!CREDIT)))
                End If
            End If

            'UPDATED BY: JUN---------------------------------------------------------------------------------------------------------------
            'DATE UPDATED: 06-10-2009
            'DESCRIPTION: DISPLAYING CUSTOMER OPENING BALANCE
            If cnt = 1 Then
                grdAccountsLedger.AddItem dtFrom & Chr(9) & "COB" & Chr(9) & "" & Chr(9) & "0.00" & Chr(9) & "0.00" & Chr(9) & xBALANCE
            End If
            'UPDATED BY: JUN---------------------------------------------------------------------------------------------------------------


            If Null2String(rsJournal_HDDet!jtype) = "CCM" Then
                grdAccountsLedger.AddItem Null2String(rsJournal_HDDet!JDate) & Chr(9) & _
                                          Reference & Chr(9) & _
                                          " " & Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO) & Chr(9) & _
                                          "0.00" & Chr(9) & _
                                          ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!CM)) & Chr(9) & _
                                          ToDoubleNumber(OUTBALANCE) & Chr(9) & rsJournal_HDDet!ID & Chr(9) & Null2String(rsJournal_HDDet!jtype)

            ElseIf Null2String(rsJournal_HDDet!jtype) = "CSJ" Then
                grdAccountsLedger.AddItem Null2String(rsJournal_HDDet!JDate) & Chr(9) & _
                                          Reference & Chr(9) & _
                                          " " & Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO) & Chr(9) & _
                                          ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!DM)) & Chr(9) & _
                                          "0.00" & Chr(9) & _
                                          ToDoubleNumber(OUTBALANCE) & Chr(9) & rsJournal_HDDet!ID & Chr(9) & Null2String(rsJournal_HDDet!jtype)
            ElseIf Null2String(rsJournal_HDDet!jtype) = "COB" Then
                grdAccountsLedger.AddItem Null2String(rsJournal_HDDet!JDate) & Chr(9) & _
                                          Reference & Chr(9) & _
                                          " " & Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO) & Chr(9) & _
                                          ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!amounttopay)) & Chr(9) & _
                                          "0.00" & Chr(9) & _
                                          ToDoubleNumber(OUTBALANCE) & Chr(9) & rsJournal_HDDet!ID & Chr(9) & Null2String(rsJournal_HDDet!jtype)
            Else

                grdAccountsLedger.AddItem Null2String(rsJournal_HDDet!JDate) & Chr(9) & _
                                          Reference & Chr(9) & _
                                          " " & Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO) & Chr(9) & _
                                          ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!DEBIT)) & Chr(9) & _
                                          ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!CREDIT)) & Chr(9) & _
                                          ToDoubleNumber(OUTBALANCE) & Chr(9) & rsJournal_HDDet!ID & Chr(9) & Null2String(rsJournal_HDDet!jtype)

            End If

            If Null2String(rsJournal_HDDet!jtype) = "CCM" Then
                TUTAL_CREDIT = TUTAL_CREDIT + N2Str2Zero(rsJournal_HDDet!CM)
            ElseIf Null2String(rsJournal_HDDet!jtype) = "CSJ" Then
                TUTAL_DEBIT = TUTAL_DEBIT + N2Str2Zero(rsJournal_HDDet!DM)
            ElseIf Null2String(rsJournal_HDDet!jtype) = "COB" Then
                TUTAL_DEBIT = TUTAL_DEBIT + N2Str2Zero(rsJournal_HDDet!amounttopay)
            Else
                TUTAL_DEBIT = TUTAL_DEBIT + N2Str2Zero(rsJournal_HDDet!DEBIT)
                TUTAL_CREDIT = TUTAL_CREDIT + N2Str2Zero(rsJournal_HDDet!CREDIT)
            End If

            rsJournal_HDDet.MoveNext
        Loop
        If cnt > 0 Then grdAccountsLedger.RemoveItem 1
    End If

    txtTotalDebit.Text = ToDoubleNumber(TUTAL_DEBIT)
    txtTotalCredit.Text = ToDoubleNumber(TUTAL_CREDIT)
    txtTotalBalance.Text = ToDoubleNumber(TUTAL_BALANCE + N2Str2Zero(OUTBALANCE))
End Sub
Sub FillSearchGrid(XXX As String)
    Dim rsCustomers                               As ADODB.Recordset
    lstCustomer.Enabled = False
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set rsCustomers = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))

    'Set rsCustomers = gconDMIS.Execute("SELECT DISTINCT dbo.ALL_Customer.ACCTNAME as CUSTNAME,dbo.ALL_Customer.ID,dbo.ALL_Customer.CUSCDE AS CUSTCODE FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType INNER JOIN dbo.ALL_Customer ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer.CUSCDE WHERE ((LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') OR (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03')) AND dbo.ALL_Customer.ACCTNAME like '" & ReplaceQuote(XXX) & "%' ORDER BY dbo.ALL_Customer.ACCTNAME")
    If optCustomer.Value = True Then
        'Set rsCustomers = gconDMIS.Execute("SELECT DISTINCT dbo.ALL_Customer.ACCTNAME as CUSTNAME,dbo.ALL_Customer.ID,dbo.ALL_Customer.CUSCDE AS CUSTCODE FROM dbo.AMIS_Journal_HD left outer JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType INNER JOIN dbo.ALL_Customer ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer.CUSCDE WHERE ((LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') OR (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03')) AND dbo.ALL_Customer.ACCTNAME like '" & ReplaceQuote(XXX) & "%' ORDER BY dbo.ALL_Customer.ACCTNAME")
        Set rsCustomers = gconDMIS.Execute("SELECT DISTINCT CUST.ACCTNAME as CUSTNAME,CUST.ID,CUST.CUSCDE AS CUSTCODE " & _
                                           "FROM AMIS_Journal_HD HD left outer JOIN AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo " & _
                                           "AND HD.JType = DET.JType INNER JOIN ALL_Customer CUST ON " & _
                                           "((HD.CustomerCode = CUST.CUSCDE) OR  ((RIGHT(DET.ENTITY,6)) = CUST.CUSCDE)) " & _
                                           "WHERE ((LEFT(DET.Acct_Code, 5) = '11-02') OR (LEFT(DET.Acct_Code, 5) = '11-03')) " & _
                                           "AND CUST.ACCTNAME like '" & ReplaceQuote(XXX) & "%' ORDER BY CUST.ACCTNAME")
        If Not (rsCustomers.EOF And rsCustomers.BOF) Then
            Listview_Loadval Me.lstCustomer.ListItems, rsCustomers
            lstCustomer.Refresh
            lstCustomer.Enabled = True
            lstCustomer.Enabled = True
        Else
            lstCustomer.Enabled = False
        End If
    Else

        rsCustomers.Open "SELECT DISTINCT ALL_VENDOR.NameofVendor as VENDORNAME,ALL_VENDOR.ID as ID,ALL_VENDOR.CODE AS VENCODE " & _
                         "FROM dbo.AMIS_Journal_HD HD INNER JOIN dbo.AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo AND HD.JType = DET.JType " & _
                         "Left outer JOIN dbo.ALL_VENDOR_TABLE ALL_VENDOR ON HD.VENDORCODE = ALL_VENDOR.CODE WHERE  ALL_VENDOR.CODE IS NOT NULL  and " & _
                         "((LEFT(DET.Acct_Code, 5) = '11-02') OR (LEFT(DET.Acct_Code, 5) = '11-03')) AND ALL_VENDOR.CODE <> '999999' AND ALL_VENDOR.NameofVendor like '" & ReplaceQuote(XXX) & "%'  ORDER BY ALL_VENDOR.VENDORNAME", gconDMIS, adOpenKeyset
        If Not (rsCustomers.EOF And rsCustomers.BOF) Then
            Listview_Loadval Me.lstCustomer.ListItems, rsCustomers
            lstCustomer.Refresh
            lstCustomer.Enabled = True
            lstCustomer.Enabled = True
        Else
            lstCustomer.Enabled = False
        End If
    End If
End Sub

Private Sub cboAccountName_Click()
'FillGrids
'FILL_LEDGER
    POPULATE_REPORT
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Frame2.ZOrder 0
    On Error Resume Next
    TextSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    On Error GoTo Errorcode:

    rsCustomer.MoveNext
    If rsCustomer.EOF Then
        rsCustomer.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo Errorcode:

    rsCustomer.MovePrevious
    If rsCustomer.BOF Then
        rsCustomer.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPrint_Click()
'COMMENTED BY: JUN 11/5/2008 DUE CHANGES OF LIST VIEW TO REPORT CONTROL
'        If MsgBox("Print Customers Ledger for this Account?", vbYesNo + vbQuestion, "Print: " & txtCustName.Text) = vbYes Then
'            Dim filter
'
'            'UPDATED BY: JUN/ARNOLD-------
'            'DATE UPDATED: 06-11-2009
'             BEG_BALANCE_DATE = dtFrom
'            'UPDATED BY: JUN/ARNOLD-------
'
'            'filter = "{Journal_HD.Status} = 'P' and (left({Journal_Det.Acct_Code},5)='11-02' OR left({Journal_Det.Acct_Code},5)='11-03' OR {Journal_Det.Acct_Code}='21-02008-00') and ({Customer.CusCde})='" & txtCode.Text & "'"
'            If MsgBox("Generate for All Customer?", vbQuestion + vbYesNo, "Selecting No will generate only selected customer") = vbYes Then
'                'filter = "({Journal_Det.Jdate} >= Date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Det.Jdate} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")) and {Journal_HD.Status} = 'P' and (left({Journal_Det.Acct_Code},5)='11-02' OR left({Journal_Det.Acct_Code},5)='11-03') and {Journal_Det.Acct_Code} = '" & Setacctcode(cboAccountName.Text) & "'"
'                filter = "({Journal_Det.Jdate} >= Date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Det.Jdate} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")) and {Journal_HD.Status} = 'P' "
'            Else
'                'filter = "({Journal_Det.Jdate} >= Date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Det.Jdate} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")) and {Journal_HD.Status} = 'P' and (left({Journal_Det.Acct_Code},5)='11-02' OR left({Journal_Det.Acct_Code},5)='11-03') and {Journal_Det.Acct_Code} = '" & Setacctcode(cboAccountName.Text) & "' and ({Customer.CusCde})='" & txtCode.Text & "'"""
'                filter = "({Journal_Det.Jdate} >= Date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Det.Jdate} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")) and {Journal_HD.Status} = 'P' and ((left({Journal_Det.Acct_Code},5)='11-02' OR left({Journal_Det.Acct_Code},5)='11-03')) and ({Customer.CUSCDE})='" & txtCode.Text & "'"
'            End If
'            ShowReport "CustomersSubsidiaryLedger", _
             '                       "Ledgers", _
             '                       filter, "C U S T O M E R S  L E D G E R", "AS OF: " & LOGDATE, True
'        End If
'COMMENTED BY: JUN 11/5/2008 DUE CHANGES OF LIST VIEW TO REPORT CONTROL
    If rptRO.Records.Count <= 0 Then Exit Sub

    Set REC = rptRO.Records.Add
    REC.AddItem(Trim("")).Record.Visible = True
    REC.AddItem(Trim("")).Record.Visible = True
    REC.AddItem(Trim("TOTAL")).Record.Visible = True
    REC.AddItem(Trim(ToDoubleNumber(G_TOTAL_DEBIT))).Record.Visible = True
    REC.AddItem(Trim(ToDoubleNumber(G_TOTAL_CREDIT))).Record.Visible = True
    REC.AddItem(Trim(ToDoubleNumber(XX_BALANCE + FORWARDED_BALANCE))).Record.Visible = True
    REC.AddItem(Trim("")).Record.Visible = True
    REC.AddItem(Trim("")).Record.Visible = True
    rptRO.Populate
    Set REC = Nothing

    'rptRO.PrintOptions.Header.TextCenter = "AUDIT PRINT FOR " ''& Replace(frmALL_AuditInquiry.Caption, "Audit Inquiry", "")
    rptRO.PrintOptions.Header.TextLeft = "HYUNDAI GLOBAL CITY " & vbCrLf & "" & COMPANY_ADDRESS & "" & vbCrLf & vbCrLf & "CUSTOMER LEDGER " & vbCrLf & vbCrLf & "Period Covered From: " & "" & dtFrom.Value & " " & "To " & "" & dtTo.Value & "" & vbCrLf & vbCrLf & "Customer code: " & "" & txtCode.Text & "" & vbCrLf & "Customer Name: " & "" & txtCustName.Text & ""

    rptRO.PrintPreview True

    '    Dim I As Integer
    '    I = rptRO.Records.Count
    '    rptRO.Rows.FindRow(I).Record.Visible = False

    LogAudit "V", "CUSTOMERS A/R LEDGER", txtCode
End Sub

Private Sub Command1_Click()
'UPDATED BY: JUN
'DATE UPDATED: 06/22/2009

    If CDate(dtFrom.Value) > CDate(dtTo.Value) Then
        MessagePop InfoFriend, "INFORMATION", "DATE FROM is greater than to the DATE TO."
        Exit Sub
    End If

    rsCustomer.MoveFirst
    InitTotal
    rsRefresh
    rsCustomer.Find "ID =" & labID.Caption
    'rsCUSTOMER.Bookmark = rsFind(rsCUSTOMER.Clone, "ID", lstCustomer.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Sub InitTotal()
    txtTotalDebit.Text = ""
    txtTotalCredit.Text = ""
    txtTotalBalance.Text = ""
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Dim rsSetMinMaxDate                           As ADODB.Recordset
    Set rsSetMinMaxDate = New ADODB.Recordset
    Set rsSetMinMaxDate = gconDMIS.Execute("Select MIN(JDATE) AS STARTDATE,MAX(JDATE) AS ENDDATE from AMIS_Journal_Det where ((LEFT(AMIS_Journal_Det.Acct_Code, 5) = '11-02') or (LEFT(AMIS_Journal_Det.Acct_Code, 5) = '11-03'))")
    If Not rsSetMinMaxDate.EOF And Not rsSetMinMaxDate.BOF Then
        dtFrom = Null2Date(rsSetMinMaxDate!STARTDATE)
        dtTo = Null2Date(rsSetMinMaxDate!ENDDATE)
    Else
        dtFrom = LOGDATE
        dtTo = LOGDATE
    End If
    InitCbo
    optCustomer.Value = True

    rsRefresh
    TextSearch.Text = "": Frame2.ZOrder 1
    initMemvars

    StoreMemVars
    Screen.MousePointer = 0
End Sub

Function Setacctcode(XXX As String) As String
    Dim rsCOA                                     As ADODB.Recordset
    Set rsCOA = New ADODB.Recordset
    Set rsCOA = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where Description = '" & XXX & "'")
    If Not rsCOA.EOF And Not rsCOA.BOF Then
        Setacctcode = Null2String(rsCOA!ACCTCODE)
    End If
    Set rsCOA = Nothing
End Function

Sub InitCbo()
    Dim rsCOA                                     As ADODB.Recordset
    Set rsCOA = New ADODB.Recordset
    Set rsCOA = gconDMIS.Execute("Select Description from AMIS_ChartAccount Where Titles in('1102' ,'1103','1102','1204','2102','2107') order by acctcode asc")
    If Not rsCOA.EOF And Not rsCOA.BOF Then
        rsCOA.MoveFirst: cboAccountName.Clear: cboAccountName.AddItem "ALL ACCOUNTS"
        Do While Not rsCOA.EOF
            cboAccountName.AddItem Null2String(rsCOA!Description)
            rsCOA.MoveNext
        Loop
    End If
    cboAccountName.Text = "ALL ACCOUNTS": DoEvents
    Set rsCOA = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LocalAcess = ""
End Sub

Private Sub grdAccountsLedger_DblClick()
    grdAccountsLedger.Row = grdAccountsLedger.Row
    grdAccountsLedger.Col = 7
    Dim VARVOUCHERNO                              As String
    '    If Left(grdAccountsLedger.Text, 3) = "APJ" Then
    '        JOURNALTYPE = "APJ"
    '    ElseIf Left(grdAccountsLedger.Text, 3) = "CDJ" Then
    '        JOURNALTYPE = "CDJ"
    '    ElseIf Left(grdAccountsLedger.Text, 2) = "SJ" Then
    '        JOURNALTYPE = "SJ"
    '    ElseIf Left(grdAccountsLedger.Text, 3) = "CRJ" Then
    '        JOURNALTYPE = "CRJ"
    '    ElseIf Left(grdAccountsLedger.Text, 2) = "GJ" Then
    '        JOURNALTYPE = "GJ"
    '    ElseIf Left(grdAccountsLedger.Text, 3) = "OPB" Then
    '        MsgBox "Not Yet Implemented!"
    '        Exit Sub
    '    Else
    '        JOURNALTYPE = Left(grdAccountsLedger.Text, 3)
    '    End If
    JOURNALTYPE = grdAccountsLedger.Text
    grdAccountsLedger.Col = 6
    Dim RETURNVOUCHERNO                           As ADODB.Recordset
    Set RETURNVOUCHERNO = gconDMIS.Execute("Select VoucherNo from AMIS_Journal_HD Where ID = " & NumericVal(grdAccountsLedger.Text))
    If Not RETURNVOUCHERNO.EOF And Not RETURNVOUCHERNO.BOF Then
        VARVOUCHERNO = Null2String(RETURNVOUCHERNO!VOUCHERNO)    'Right(grdAccountsLedger.Text, 6)
        Screen.MousePointer = 11
        If JOURNALTYPE = "COB" Then
            On Error Resume Next
            Unload frmAMISCustomerAROpening
            frmAMISCustomerAROpening.Show
            frmAMISCustomerAROpening.StoreSearch (VARVOUCHERNO)
        Else
            On Error Resume Next
            Unload frmAMISJournalEntry
            frmAMISJournalEntry.Show
            frmAMISJournalEntry.StoreSearch (VARVOUCHERNO)
        End If
        Screen.MousePointer = 0
    Else
    End If
End Sub

Private Sub lstCustomer_GotFocus()
'rsCUSTOMER.Bookmark = rsFind(rsCUSTOMER.Clone, "ID", lstCustomer.SelectedItem.SubItems(1)).Bookmark
'rsRefresh
'rsCUSTOMER.Find "ID =" & labID.Caption
'StoreMemvars
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next

    If CDate(dtFrom.Value) > CDate(dtTo.Value) Then
        MessagePop InfoFriend, "INFORMATION", "DATE FROM is greater than to the DATE TO."
        Exit Sub
    End If


    'rsCUSTOMER.Bookmark = rsFind(rsCUSTOMER.Clone, "ID", lstCustomer.SelectedItem.SubItems(1)).Bookmark
    labID.Caption = lstCustomer.SelectedItem.SubItems(1)
    rsRefresh
    rsCustomer.Find "ID =" & labID.Caption
    StoreMemVars
End Sub

Private Sub lstCustomer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstCustomer
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        TextSearch.SetFocus
    End If
End Sub

Private Sub lvwLedger_DblClick()
    Dim VARVOUCHERNO                              As String
    Dim RETURNVOUCHERNO                           As ADODB.Recordset

    JOURNALTYPE = lvwLedger.SelectedItem.SubItems(7)
    Set RETURNVOUCHERNO = gconDMIS.Execute("Select VoucherNo from AMIS_Journal_HD Where ID = " & NumericVal(lvwLedger.SelectedItem.SubItems(6)))
    If Not RETURNVOUCHERNO.EOF And Not RETURNVOUCHERNO.BOF Then
        VARVOUCHERNO = Null2String(RETURNVOUCHERNO!VOUCHERNO)    'Right(grdAccountsLedger.Text, 6)
        Screen.MousePointer = 11
        If JOURNALTYPE = "COB" Then
            On Error Resume Next
            Unload frmAMISCustomerAROpening
            frmAMISCustomerAROpening.Show
            frmAMISCustomerAROpening.StoreSearch (VARVOUCHERNO)
        Else
            On Error Resume Next
            Unload frmAMISJournalEntry
            frmAMISJournalEntry.Show
            frmAMISJournalEntry.StoreSearch (VARVOUCHERNO)
        End If
        Screen.MousePointer = 0
    Else
    End If
End Sub

Private Sub lvwLedger_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        CUSCODE = txtCode
        INVOICENO = Right(lvwLedger.SelectedItem.SubItems(2), 6)
        InvoiceType = lvwLedger.SelectedItem.SubItems(8)
        frmAMISLedgerCRJ.Show
    Case Else
        MoveKeyPress KeyCode
    End Select
End Sub

Private Sub optCustomer_Click()
    FILL_CUST_VEN
End Sub

Private Sub optVendor_Click()
    FILL_CUST_VEN
End Sub

Private Sub rptRO_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If Row.Record(4).Value <> "0.00" And RTrim(LTrim(Row.Record(2).Value)) <> "TOTAL" Then
        Metrics.ForeColor = vbRed
    Else
        Metrics.ForeColor = &H808000
        Metrics.Font.Bold = True
    End If
End Sub

Private Sub rptRO_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)

    If Row.Record Is Nothing Then: Exit Sub

    JOURNALTYPE = Row.Record(7).Value

    If JOURNALTYPE = "COB" Then
        On Error Resume Next
        Unload frmAMISCustomerAROpening
        frmAMISCustomerAROpening.Show
        frmAMISCustomerAROpening.StoreSearch (Right(Row.Record(1).Value, 6))
    ElseIf JOURNALTYPE = "GJ" Then
        frmAMIS_GJ_JOURNAL_ENTRY.LoadJournal ("GJ")
        frmAMIS_GJ_JOURNAL_ENTRY.Show
        frmAMIS_GJ_JOURNAL_ENTRY.SearchVoucherNo (Right(Row.Record(1).Value, 6))
    Else
        On Error Resume Next
        Unload frmAMISJournalEntry
        frmAMISJournalEntry.Show
        frmAMISJournalEntry.StoreSearch (Right(Row.Record(1).Value, 6))
    End If
End Sub

Private Sub textSearch_Change()
    If Trim(TextSearch.Text) = "" Then
        FILL_CUST_VEN
        'FillGrid
    Else
        FillSearchGrid (TextSearch.Text)
    End If
End Sub

Private Sub FillGrid()
    Dim rsCustomers                               As ADODB.Recordset
    lstCustomer.Enabled = False
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set rsCustomers = New ADODB.Recordset
    Set rsCustomers = gconDMIS.Execute("SELECT DISTINCT dbo.ALL_Customer.ACCTNAME as CUSTNAME,dbo.ALL_Customer.ID,dbo.ALL_Customer.CUSCDE AS CUSTCODE FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType Left outer JOIN dbo.ALL_Customer ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer.CUSCDE WHERE ((LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') OR (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03')) ORDER BY dbo.ALL_Customer.ACCTNAME")
    If Not (rsCustomers.EOF And rsCustomers.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsCustomers
        lstCustomer.Refresh
        lstCustomer.Enabled = True
        lstCustomer.Enabled = True
    Else
        lstCustomer.Enabled = False
    End If
End Sub
Sub FILL_CUST_VEN()
    Dim rsFILL_CUST_VEN                           As ADODB.Recordset
    Dim rsFILL_Vendo                              As ADODB.Recordset

    Dim Item                                      As ListItem

    lstCustomer.ListItems.Clear

    If optCustomer.Value = True Then
        Set rsFILL_CUST_VEN = New ADODB.Recordset
        rsFILL_CUST_VEN.Open "SELECT DISTINCT TOP 27 dbo.ALL_Customer.ACCTNAME as CUSTNAME,dbo.ALL_Customer.ID as ID,dbo.ALL_Customer.CUSCDE AS CUSTCODE FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType Left outer JOIN dbo.ALL_Customer ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer.CUSCDE WHERE  dbo.ALL_Customer.CUSCDE IS NOT NULL  and   ((LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') OR (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03')) ORDER BY dbo.ALL_Customer.ACCTNAME", gconDMIS, adOpenKeyset
        If Not rsFILL_CUST_VEN.EOF And Not rsFILL_CUST_VEN.BOF Then
            Do While Not rsFILL_CUST_VEN.EOF
                Set Item = lstCustomer.ListItems.Add(, , Null2String(rsFILL_CUST_VEN!CUSTNAME))
                Item.SubItems(1) = rsFILL_CUST_VEN!ID
                rsFILL_CUST_VEN.MoveNext
            Loop
        Else
            MessagePop InfoFriend, "INFORMATION", "There is no such records"
            Exit Sub
        End If
    ElseIf optVendor.Value = True Then
        Set rsFILL_Vendo = New ADODB.Recordset
        rsFILL_Vendo.Open "SELECT DISTINCT ALL_VENDOR.NameofVendor as VENDORNAME,ALL_VENDOR.ID as ID,ALL_VENDOR.CODE AS VENCODE " & _
                          "FROM dbo.AMIS_Journal_HD HD INNER JOIN dbo.AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo AND HD.JType = DET.JType " & _
                          "Left outer JOIN dbo.ALL_VENDOR_TABLE ALL_VENDOR ON HD.VENDORCODE = ALL_VENDOR.CODE WHERE  ALL_VENDOR.CODE IS NOT NULL  and " & _
                          "((LEFT(DET.Acct_Code, 5) = '11-02') OR (LEFT(DET.Acct_Code, 5) = '11-03')) ORDER BY ALL_VENDOR.VENDORNAME", gconDMIS, adOpenKeyset
        If Not rsFILL_Vendo.EOF And Not rsFILL_Vendo.BOF Then
            Do While Not rsFILL_Vendo.EOF
                Set Item = lstCustomer.ListItems.Add(, , Null2String(rsFILL_Vendo!VendorName))
                Item.SubItems(1) = rsFILL_Vendo!ID
                rsFILL_Vendo.MoveNext
            Loop
        End If
    End If
    Set rsFILL_Vendo = Nothing
    Set rsFILL_CUST_VEN = Nothing
End Sub


Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Frame2.ZOrder 0
    If KeyCode = vbKeyDown Then
        If lstCustomer.ListItems.Count > 0 And lstCustomer.Enabled = True Then
            lstCustomer.SetFocus
        End If
    End If
End Sub
Function checkaccount(XXX As String, EVAN_PARAKITO As String)
    Dim RS                                        As New ADODB.Recordset
    Dim cnt                                       As Integer
    Dim Account_code                              As String
    Set RS = gconDMIS.Execute("select acct_code,voucherno from amis_journal_det where VoucherNo='" & XXX & "' and jtype='" & EVAN_PARAKITO & "'")
    If Not (RS.EOF And RS.BOF) Then
        RS.MoveFirst
        cnt = 0
        Do While Not RS.EOF
            Account_code = Null2String(RS!Acct_code)
            If Left(Account_code, 5) = "11-02" Or Left(Account_code, 5) = "11-03" Then
                cnt = cnt + 1
                checkaccount = Left(Account_code, 5)
            End If
            checkaccount = cnt
            RS.MoveNext
        Loop
    End If
    Set RS = Nothing
End Function


Sub FILL_LEDGER()
    Dim OUTBALANCE                                As Double
    Dim Reference                                 As String
    Dim cnt                                       As Integer
    Dim CREDIT                                    As Double
    Dim DEBIT                                     As Double
    Dim cnt_adjusment                             As Integer
    Dim tmp_voucher                               As String
    Dim lvw_COUNT                                 As Integer
    Dim Item                                      As ListItem

    lvw_COUNT = 1

    Set rsCUSTOMER_OPENING = New ADODB.Recordset

    cleargrid grdAccountsLedger: InitGrid
    TUTAL_BALANCE = 0: TUTAL_BALANCE = TUTAL_BALANCE: cnt = 0: TUTAL_DEBIT = 0: TUTAL_CREDIT = 0: OUTBALANCE = 0
    cnt_adjusment = 0
    Set rsJournal_HDDet = New ADODB.Recordset
    If cboAccountName.Text = "ALL ACCOUNTS" Then
        If optCustomer.Value = True Then
            rsJournal_HDDet.Open "SELECT HD_DET.ID AS DET_ID,HD.STATUS AS SS,HD.Debit AS DM,HD.Credit AS CM,HD.AmountToPay,HD.InvoiceAmt,HD.InvoiceNo,HD.ID,HD.JNo,HD.JDate,HD.JType,HD_DET.Debit,HD_DET.Credit, HD.VoucherNo,HD.CheckNo,HD.InvoiceType,HD.VendorCode,HD.JNo " & _
                                 "FROM AMIS_Journal_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET HD_DET ON HD_DET.VoucherNo  = HD.VoucherNo AND HD_DET.jtype = HD.Jtype where " & _
                                 "(HD.Jdate >= '" & dtFrom & "'and HD.Jdate <= '" & dtTo & "')AND (((HD.JTYPE = 'GJ' AND HD_DET.ADJ_JTYPE <> 'APJ'  AND right(HD_DET.ENTITY,6) = '" & txtCode.Text & "' AND (left(HD_DET.acct_code,5) IN ('11-02','21-01','11-03','21-07'))))OR((Left(HD_DET.Acct_Code,5) IN ('11-02','11-03','21-02','21-07')) AND ((HD.CustomerCode = '" & txtCode.Text & "') OR (HD.VENDORCODE = '" & txtCode.Text & "'))))and HD.Status = 'P' order by HD.jdate asc,HD.id asc", gconDMIS, adOpenKeyset
        Else
            rsJournal_HDDet.Open "SELECT HD_DET.ID AS DET_ID,HD.STATUS AS SS,HD.Debit AS DM,HD.Credit AS CM,HD.AmountToPay,HD.InvoiceAmt,HD.InvoiceNo,HD.ID,HD.JNo,HD.JDate,HD.JType,HD_DET.Debit,HD_DET.Credit, HD.VoucherNo,HD.CheckNo,HD.InvoiceType,HD.VendorCode,HD.JNo " & _
                                 "FROM AMIS_Journal_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET HD_DET ON HD_DET.VoucherNo  = HD.VoucherNo AND HD_DET.jtype = HD.Jtype where " & _
                                 "(HD.Jdate >= '" & dtFrom & "'and HD.Jdate <= '" & dtTo & "')AND (((HD.JTYPE = 'GJ' AND HD_DET.ADJ_JTYPE <> 'APJ'  AND right(HD_DET.ENTITY,6) = '" & txtCode.Text & "' AND (left(HD_DET.acct_code,5) IN ('11-02','21-01','11-03','21-07'))))OR((Left(HD_DET.Acct_Code,5) IN ('11-02','11-03','21-02','21-07')) AND ((HD.VendorCode = '" & txtCode.Text & "') OR (HD.VENDORCODE = '" & txtCode.Text & "'))))and HD.Status = 'P' order by HD.jdate asc,HD.id asc", gconDMIS, adOpenKeyset
        End If
    Else
        If cboAccountName.Text = "A/REC CREDIT CARD" Then
            If CheckIfBank(txtCode) = True Then
                rsJournal_HDDet.Open "select AMIS_Journal_Hd.Debit AS DM,AMIS_Journal_Hd.Credit AS CM,AMIS_Journal_Hd.AmountToPay,AMIS_Journal_Hd.InvoiceAmt,AMIS_Journal_Hd.InvoiceNo,AMIS_Journal_HD.ID,AMIS_Journal_Hd.JNo,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.JType,AMIS_Journal_Det.Debit,AMIS_Journal_Det.Credit,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.CheckNo,AMIS_Journal_Hd.InvoiceType,AMIS_Journal_Hd.VendorCode,AMIS_Journal_Hd.JNo,AMIS_Journal_Hd.Bank from AMIS_Journal_HD left outer Join AMIS_Journal_det on AMIS_Journal_Det.VoucherNo  = AMIS_Journal_Hd.VoucherNo and AMIS_Journal_Det.jtype = AMIS_Journal_Hd.Jtype  where " & _
                                     "(dbo.AMIS_Journal_HD.Jdate >= '" & dtFrom & "' and dbo.AMIS_Journal_HD.Jdate <= '" & dtTo & "') and (AMIS_Journal_Hd.Status = 'P' AND AMIS_Journal_Det.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "' AND ((Left(AMIS_Journal_Det.Acct_Code,5) = '11-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '11-03' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-07') OR ((AMIS_Journal_HD.JTYPE = 'CCM' AND (left(AMIS_Journal_det.acct_code,5) = '11-01' or left(AMIS_Journal_det.acct_code,5) = '21-01')) OR (AMIS_Journal_HD.JTYPE = 'CSJ' AND (left(AMIS_Journal_det.acct_code,5) = '11-03' or left(AMIS_Journal_det.acct_code,5) = '21-07'))))) AND AMIS_Journal_Hd.Bank = '" & txtCode.Text & "' order by AMIS_Journal_Hd.jdate asc,AMIS_Journal_Hd.id asc", gconDMIS
            Else
                rsJournal_HDDet.Open "SELECT AMIS_vw_JOURNAL_HD.DEBIT AS DM,AMIS_vw_JOURNAL_HD.CREDIT AS CM,AMIS_vw_JOURNAL_HD.AMOUNTTOPAY,AMIS_vw_JOURNAL_HD.INVOICEAMT,AMIS_vw_JOURNAL_HD.INVOICENO,AMIS_vw_JOURNAL_HD.ID,AMIS_vw_JOURNAL_HD.JNO,AMIS_vw_JOURNAL_HD.JDATE,AMIS_vw_JOURNAL_HD.JTYPE,AMIS_JOURNAL_DET.DEBIT,AMIS_JOURNAL_DET.CREDIT,AMIS_vw_JOURNAL_HD.VOUCHERNO,AMIS_vw_JOURNAL_HD.CHECKNO,AMIS_vw_JOURNAL_HD.INVOICETYPE,AMIS_vw_JOURNAL_HD.VENDORCODE,AMIS_vw_JOURNAL_HD.JNO,AMIS_vw_JOURNAL_HD.REFERENCENO FROM AMIS_vw_JOURNAL_HD LEFT OUTER JOIN AMIS_JOURNAL_DET ON AMIS_JOURNAL_DET.VOUCHERNO  = AMIS_vw_JOURNAL_HD.VOUCHERNO AND AMIS_JOURNAL_DET.JTYPE = AMIS_vw_JOURNAL_HD.JTYPE AND AMIS_vw_JOURNAL_HD.REFERENCENO=AMIS_JOURNAL_DET.REFERENCENO WHERE " & _
                                     "(dbo.AMIS_vw_Journal_HD.Jdate >= '" & dtFrom & "' and dbo.AMIS_vw_Journal_HD.Jdate <= '" & dtTo & "') and (AMIS_vw_Journal_Hd.Status = 'P' AND AMIS_Journal_Det.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "' AND ((Left(AMIS_Journal_Det.Acct_Code,5) = '11-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '11-03' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-07') OR ((AMIS_vw_Journal_HD.JTYPE = 'CCM' AND (left(AMIS_Journal_det.acct_code,5) in ('11-02','21-07'))) OR (AMIS_vw_Journal_HD.JTYPE = 'CSJ' AND (left(AMIS_Journal_det.acct_code,5) in ('11-03','21-07')))))) AND AMIS_vw_JOURNAL_HD.REFERENCENO = '" & txtCode.Text & "' order by AMIS_vw_Journal_Hd.jdate asc,AMIS_vw_Journal_Hd.id asc", gconDMIS
            End If
        Else
            If optCustomer.Value = True Then
                rsJournal_HDDet.Open "SELECT HD_DET.ID AS DET_ID,HD.STATUS AS SS,HD.Debit AS DM,HD.Credit AS CM,HD.AmountToPay,HD.InvoiceAmt,HD.InvoiceNo,HD.ID,HD.JNo,HD.JDate,HD.JType,HD_DET.Debit,HD_DET.Credit, HD.VoucherNo,HD.CheckNo,HD.InvoiceType,HD.VendorCode,HD.JNo " & _
                                     "FROM AMIS_Journal_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET HD_DET ON HD_DET.VoucherNo  = HD.VoucherNo AND HD_DET.jtype = HD.Jtype where " & _
                                     "(HD.Jdate >= '" & dtFrom & "'and HD.Jdate <= '" & dtTo & "')AND (((HD.JTYPE = 'GJ' AND HD_DET.ADJ_JTYPE <> 'APJ'  AND right(HD_DET.ENTITY,6) = '" & txtCode.Text & "' AND (left(HD_DET.acct_code,5) IN ('11-02','21-01','11-03','21-07'))))OR((Left(HD_DET.Acct_Code,5) IN ('11-02','11-03','21-02','21-07')) AND HD.CustomerCode = '" & txtCode.Text & "'))and HD.Status = 'P' and HD_DET.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "' order by HD.jdate asc,HD.id asc", gconDMIS, adOpenKeyset
            Else
                rsJournal_HDDet.Open "SELECT HD_DET.ID AS DET_ID,HD.STATUS AS SS,HD.Debit AS DM,HD.Credit AS CM,HD.AmountToPay,HD.InvoiceAmt,HD.InvoiceNo,HD.ID,HD.JNo,HD.JDate,HD.JType,HD_DET.Debit,HD_DET.Credit, HD.VoucherNo,HD.CheckNo,HD.InvoiceType,HD.VendorCode,HD.JNo " & _
                                     "FROM AMIS_Journal_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET HD_DET ON HD_DET.VoucherNo  = HD.VoucherNo AND HD_DET.jtype = HD.Jtype where " & _
                                     "(HD.Jdate >= '" & dtFrom & "'and HD.Jdate <= '" & dtTo & "')AND (((HD.JTYPE = 'GJ' AND HD_DET.ADJ_JTYPE <> 'APJ'  AND right(HD_DET.ENTITY,6) = '" & txtCode.Text & "' AND (left(HD_DET.acct_code,5) IN ('11-02','21-01','11-03','21-07'))))OR((Left(HD_DET.Acct_Code,5) IN ('11-02','11-03','21-02','21-07')) AND HD.VendorCode = '" & txtCode.Text & "'))and HD.Status = 'P' and HD_DET.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "' order by HD.jdate asc,HD.id asc", gconDMIS, adOpenKeyset
            End If
        End If
    End If

    Me.lvwLedger.ListItems.Clear: Me.lvwLedger.Enabled = False
    If Not rsJournal_HDDet.EOF And Not rsJournal_HDDet.BOF Then
        'Me.lvwLedger.Sorted = True:
        rsJournal_HDDet.MoveFirst
        Do While Not rsJournal_HDDet.EOF
            'If lvw_COUNT = 5 Then Stop
            Reference = Null2String(rsJournal_HDDet!jtype) & "-" & Null2String(rsJournal_HDDet!VOUCHERNO)

            'UPDATED BY: JUN--------------------------------------------------------
            'DATE UPDATED: 06-10-2009
            'DESCRIPTION: SUMMATION OF PREVIOUS BALANCE AND CUSTOMER OPENING BALANCE
            If lvw_COUNT = 1 Then
                OUTBALANCE = OUTBALANCE + xBALANCE
            End If
            'UPDATED BY: JUN--------------------------------------------------------

            If Null2String(rsJournal_HDDet!jtype) = "COB" Then
                OUTBALANCE = OUTBALANCE + N2Str2Zero(rsJournal_HDDet!InvoiceAmt)
            Else
                If Null2String(rsJournal_HDDet!jtype) = "GJ" Then
                    If NumericVal(rsJournal_HDDet!CREDIT) <> 0 Then
                        OUTBALANCE = N2Str2Zero(OUTBALANCE) - N2Str2Zero(rsJournal_HDDet!CREDIT)
                    ElseIf NumericVal(rsJournal_HDDet!DEBIT) <> 0 Then
                        OUTBALANCE = N2Str2Zero(OUTBALANCE) + N2Str2Zero(rsJournal_HDDet!DEBIT)
                    End If
                Else
                    If Null2String(rsJournal_HDDet!jtype) = "CRJ" Then
                        If CHK_PYMENT_DISPLAY(Null2String(rsJournal_HDDet!VOUCHERNO), Null2String(rsJournal_HDDet!jtype)) = True Then
                            OUTBALANCE = N2Str2Zero(OUTBALANCE) + ((N2Str2Zero(rsJournal_HDDet!DEBIT) - N2Str2Zero(rsJournal_HDDet!CREDIT)))
                        Else
                            'DON'T COMPUTE
                        End If
                    Else
                        OUTBALANCE = N2Str2Zero(OUTBALANCE) + ((N2Str2Zero(rsJournal_HDDet!DEBIT) - N2Str2Zero(rsJournal_HDDet!CREDIT)))
                    End If

                End If
            End If

            If lvw_COUNT = 1 Then
                Set Item = lvwLedger.ListItems.Add(, , dtFrom.Value)
                Item.SubItems(1) = "FWD BALANCE"
                Item.SubItems(2) = ""
                Item.SubItems(3) = "0.00"
                Item.SubItems(4) = "0.00"
                Item.SubItems(5) = ToDoubleNumber(xBALANCE)
                Item.SubItems(6) = ""
                Item.SubItems(7) = ""
                lvw_COUNT = lvw_COUNT + 1
            End If

            If Null2String(rsJournal_HDDet!jtype) = "GJ" Then
                If NumericVal(rsJournal_HDDet!CREDIT) <> 0 Then
                    Set Item = lvwLedger.ListItems.Add(, , Null2String(rsJournal_HDDet!JDate))
                    Item.SubItems(1) = Null2String(Reference)
                    Item.SubItems(2) = FIND_GJ_REFERENCE(rsJournal_HDDet!DET_ID)
                    Item.SubItems(3) = "0.00"
                    Item.SubItems(4) = ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!CREDIT))
                    Item.SubItems(5) = ToDoubleNumber(OUTBALANCE)
                    Item.SubItems(6) = rsJournal_HDDet!ID
                    Item.SubItems(7) = Null2String(rsJournal_HDDet!jtype)
                ElseIf NumericVal(rsJournal_HDDet!DEBIT) <> 0 Then
                    Set Item = lvwLedger.ListItems.Add(, , Null2String(rsJournal_HDDet!JDate))
                    Item.SubItems(1) = Null2String(Reference)
                    Item.SubItems(2) = FIND_GJ_REFERENCE(rsJournal_HDDet!DET_ID)
                    Item.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!DEBIT))
                    Item.SubItems(4) = "0.00"
                    Item.SubItems(5) = ToDoubleNumber(OUTBALANCE)
                    Item.SubItems(6) = rsJournal_HDDet!ID
                    Item.SubItems(7) = Null2String(rsJournal_HDDet!jtype)
                End If
            ElseIf Null2String(rsJournal_HDDet!jtype) = "COB" Then
                Set Item = lvwLedger.ListItems.Add(, , Null2String(rsJournal_HDDet!JDate))
                Item.SubItems(1) = Null2String(Reference)
                Item.SubItems(2) = Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO)
                Item.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!amounttopay))
                Item.SubItems(4) = "0.00"
                Item.SubItems(5) = ToDoubleNumber(OUTBALANCE)
                Item.SubItems(6) = rsJournal_HDDet!ID
                Item.SubItems(7) = Null2String(rsJournal_HDDet!jtype)
            Else

                If Null2String(rsJournal_HDDet!jtype) = "CRJ" Then
                    If CHK_PYMENT_DISPLAY(Null2String(rsJournal_HDDet!VOUCHERNO), Null2String(rsJournal_HDDet!jtype)) = True Then
                        Set Item = lvwLedger.ListItems.Add(, , Null2String(rsJournal_HDDet!JDate))
                        Item.SubItems(1) = Null2String(Reference)
                        Item.SubItems(2) = Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO)
                        Item.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!DEBIT))
                        Item.SubItems(4) = ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!CREDIT))
                        Item.SubItems(5) = ToDoubleNumber(OUTBALANCE)
                        Item.SubItems(6) = rsJournal_HDDet!ID
                        Item.SubItems(7) = Null2String(rsJournal_HDDet!jtype)
                        Item.SubItems(8) = Null2String(rsJournal_HDDet!InvoiceType)
                    Else
                        lvw_COUNT = lvw_COUNT - 1
                        'DONT DISPLAY BECAUSE CUSTOMER CODE FOR PAYMENT IS WRONG EVENTHOUGH THE INVOICETYPE AND INVOICE IS IN SALES JOURNAL
                    End If
                Else
                    Set Item = lvwLedger.ListItems.Add(, , Null2String(rsJournal_HDDet!JDate))
                    Item.SubItems(1) = Null2String(Reference)
                    Item.SubItems(2) = Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO)
                    Item.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!DEBIT))
                    Item.SubItems(4) = ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!CREDIT))
                    Item.SubItems(5) = ToDoubleNumber(OUTBALANCE)
                    Item.SubItems(6) = rsJournal_HDDet!ID
                    Item.SubItems(7) = Null2String(rsJournal_HDDet!jtype)
                    Item.SubItems(8) = Null2String(rsJournal_HDDet!InvoiceType)
                End If

                If Null2String(rsJournal_HDDet!jtype) = "SJ" Then
                    If CHECK_PAYMENT(Null2String(rsJournal_HDDet!VOUCHERNO), "SJ") = True Then
                        lvwLedger.ListItems(lvw_COUNT).ForeColor = vbRed
                        lvwLedger.ListItems(lvw_COUNT).ListSubItems.Item(1).ForeColor = vbRed
                        lvwLedger.ListItems(lvw_COUNT).ListSubItems.Item(2).ForeColor = vbRed
                        lvwLedger.ListItems(lvw_COUNT).ListSubItems.Item(3).ForeColor = vbRed
                        lvwLedger.ListItems(lvw_COUNT).ListSubItems.Item(4).ForeColor = vbRed
                        lvwLedger.ListItems(lvw_COUNT).ListSubItems.Item(5).ForeColor = vbRed
                        lvwLedger.ListItems(lvw_COUNT).ListSubItems.Item(6).ForeColor = vbRed
                        lvwLedger.ListItems(lvw_COUNT).ListSubItems.Item(7).ForeColor = vbRed
                    Else
                        'NO PAYMENT FOUND
                    End If
                End If
            End If

            If Null2String(rsJournal_HDDet!jtype) = "CCM" Then
                TUTAL_CREDIT = TUTAL_CREDIT + N2Str2Zero(rsJournal_HDDet!CM)
            ElseIf Null2String(rsJournal_HDDet!jtype) = "CSJ" Then
                TUTAL_DEBIT = TUTAL_DEBIT + N2Str2Zero(rsJournal_HDDet!DM)
            ElseIf Null2String(rsJournal_HDDet!jtype) = "COB" Then
                TUTAL_DEBIT = TUTAL_DEBIT + N2Str2Zero(rsJournal_HDDet!amounttopay)
            ElseIf Null2String(rsJournal_HDDet!jtype) = "GJ" Then
                If NumericVal(rsJournal_HDDet!CREDIT) <> 0 Then
                    TUTAL_CREDIT = TUTAL_CREDIT + N2Str2Zero(rsJournal_HDDet!CREDIT)
                ElseIf NumericVal(rsJournal_HDDet!DEBIT) <> 0 Then
                    TUTAL_DEBIT = TUTAL_DEBIT + N2Str2Zero(rsJournal_HDDet!DEBIT)
                End If
            Else
                If Null2String(rsJournal_HDDet!jtype) = "CRJ" Then
                    If CHK_PYMENT_DISPLAY(Null2String(rsJournal_HDDet!VOUCHERNO), Null2String(rsJournal_HDDet!jtype)) = True Then
                        TUTAL_DEBIT = TUTAL_DEBIT + N2Str2Zero(rsJournal_HDDet!DEBIT)
                        TUTAL_CREDIT = TUTAL_CREDIT + N2Str2Zero(rsJournal_HDDet!CREDIT)
                    Else
                        'DON'T PAYMENT
                    End If
                Else
                    TUTAL_DEBIT = TUTAL_DEBIT + N2Str2Zero(rsJournal_HDDet!DEBIT)
                    TUTAL_CREDIT = TUTAL_CREDIT + N2Str2Zero(rsJournal_HDDet!CREDIT)
                End If

            End If
            lvw_COUNT = lvw_COUNT + 1
            rsJournal_HDDet.MoveNext
        Loop

    Else
        Set Item = lvwLedger.ListItems.Add(, , dtFrom.Value)
        Item.SubItems(1) = "FWD BALANCE"
        Item.SubItems(2) = ""
        Item.SubItems(3) = "0.00"
        Item.SubItems(4) = "0.00"
        Item.SubItems(5) = ToDoubleNumber(xBALANCE)
        Item.SubItems(6) = ""
        Item.SubItems(7) = ""
    End If

    Me.lvwLedger.Enabled = True: Me.lvwLedger.Sorted = False: Me.lvwLedger.Refresh

    txtTotalDebit.Text = ToDoubleNumber(TUTAL_DEBIT)
    txtTotalCredit.Text = ToDoubleNumber(TUTAL_CREDIT)
    txtTotalBalance.Text = ToDoubleNumber(TUTAL_BALANCE + N2Str2Zero(OUTBALANCE))
End Sub

Function CHECK_PAYMENT(xVOUCHERNO As String, xJType As String) As Boolean
    Dim rsINV_INVTYPE                             As ADODB.Recordset
    Dim rsCHECK_PAYMENT                           As ADODB.Recordset
    Set rsINV_INVTYPE = gconDMIS.Execute("SELECT CUSTOMERCODE,INVOICENO,INVOICETYPE FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & xVOUCHERNO & "' AND JTYPE = '" & xJType & "'")
    If Not rsINV_INVTYPE.EOF And Not rsINV_INVTYPE.BOF Then
        Set rsCHECK_PAYMENT = gconDMIS.Execute("SELECT VOUCHERNO,INVOICENO,INVOICETYPE FROM AMIS_CRJ_DETAIL WHERE INVOICENO = '" & Null2String(rsINV_INVTYPE!INVOICENO) & "' AND INVOICETYPE = '" & Null2String(rsINV_INVTYPE!InvoiceType) & "'")
        If Not rsCHECK_PAYMENT.BOF And Not rsCHECK_PAYMENT.EOF Then
            Dim rsCHECK_CUS_CODE                  As ADODB.Recordset
            Set rsCHECK_CUS_CODE = New ADODB.Recordset
            rsCHECK_CUS_CODE.Open "Select CustomerCode from Amis_journal_hd where Voucherno = '" & Null2String(rsCHECK_PAYMENT!VOUCHERNO) & "' and CustomerCode = '" & Null2String(rsINV_INVTYPE!CustomerCode) & "' and JTYPE = 'CRJ'", gconDMIS, adOpenKeyset
            If Not rsCHECK_CUS_CODE.EOF And Not rsCHECK_CUS_CODE.BOF Then
                CHECK_PAYMENT = True
            Else
                CHECK_PAYMENT = False
            End If
            Set rsCHECK_CUS_CODE = Nothing
        End If
    Else
        'vocher number has no header
    End If
    Set rsINV_INVTYPE = Nothing
    Set rsCHECK_PAYMENT = Nothing
End Function

Function FIND_GJ_REFERENCE(xID As Variant) As String
    Dim rsFIND_GJ_REFERENCE                       As ADODB.Recordset
    Set rsFIND_GJ_REFERENCE = New ADODB.Recordset
    rsFIND_GJ_REFERENCE.Open "Select InvoiceNo,InvoiceType from Amis_Journal_det  where ID = '" & xID & "'", gconDMIS, adOpenKeyset
    If Not rsFIND_GJ_REFERENCE.EOF And Not rsFIND_GJ_REFERENCE.BOF Then
        FIND_GJ_REFERENCE = Null2String(rsFIND_GJ_REFERENCE!INVOICENO)
    End If
    Set rsFIND_GJ_REFERENCE = Nothing
End Function

Function CHK_PYMENT_DISPLAY(xVOUCHERNO As String, xJType As String) As Boolean
    Dim rsCHK_PYMENT_DISPLAY                      As ADODB.Recordset
    Dim rsCRJ_CODE                                As ADODB.Recordset
    Dim rsSJ_CODE                                 As ADODB.Recordset
    Set rsCHK_PYMENT_DISPLAY = New ADODB.Recordset
    rsCHK_PYMENT_DISPLAY.Open "Select InvoiceNo,InvoiceType from Amis_CRJ_detail where VoucherNo = '" & xVOUCHERNO & "' and Status = 'P'", gconDMIS, adOpenKeyset
    If Not rsCHK_PYMENT_DISPLAY.EOF And Not rsCHK_PYMENT_DISPLAY.BOF Then
        Set rsCRJ_CODE = New ADODB.Recordset
        rsCRJ_CODE.Open "Select CustomerCode From Amis_Journal_hd where VoucherNo = '" & xVOUCHERNO & "' and CustomerCode = '" & txtCode.Text & "' and Jtype = 'CRJ'", gconDMIS, adOpenKeyset
        If Not rsCRJ_CODE.EOF And Not rsCRJ_CODE.BOF Then
            Set rsSJ_CODE = New ADODB.Recordset
            rsSJ_CODE.Open "SELECT * FROM AMIS_JOURNAL_HD WHERE CUSTOMERCODE = '" & txtCode.Text & "' AND INVOICENO = '" & Null2String(rsCHK_PYMENT_DISPLAY!INVOICENO) & "' and INVOICETYPE = '" & Null2String(rsCHK_PYMENT_DISPLAY!InvoiceType) & "'", gconDMIS, adOpenKeyset
            If Not rsSJ_CODE.EOF And Not rsSJ_CODE.BOF Then
                CHK_PYMENT_DISPLAY = True
            Else
                CHK_PYMENT_DISPLAY = False
            End If
        End If
    End If
    Set rsCHK_PYMENT_DISPLAY = Nothing
    Set rsCRJ_CODE = Nothing
End Function

Function CheckIfBank(xCUSCDE As String) As Boolean
    Dim rsCheckCode                               As ADODB.Recordset
    Set rsCheckCode = New ADODB.Recordset
    rsCheckCode.Open "Select Cuscde from All_Customer_Table where CusCde = " & N2Str2Null(xCUSCDE) & "", gconDMIS, adOpenForwardOnly
    If Not rsCheckCode.EOF And Not rsCheckCode.BOF Then
        Do While Not rsCheckCode.EOF
            Dim rsCheckBank                       As ADODB.Recordset
            Set rsCheckBank = New ADODB.Recordset
            rsCheckBank.Open "Select CusCde from CMIS_CardBank where CusCde = " & N2Str2Null(rsCheckCode!CUSCDE) & "", gconDMIS, adOpenForwardOnly
            If Not rsCheckBank.EOF And Not rsCheckBank.BOF Then
                CheckIfBank = True
            Else
                CheckIfBank = False
            End If
            rsCheckCode.MoveNext
        Loop
    End If
    Set rsCheckCode = Nothing
    Set rsCheckBank = Nothing
End Function


Sub INITIALIZE_RCONTROL()
    With rptRO
        .Columns.DeleteAll
        .Columns.Add 0, "DOCDATE", 80, True: .Columns(0).Alignment = xtpAlignmentRight: .Columns(0).AllowRemove = False: .Columns(0).AutoSortWhenGrouped = True
        .Columns.Add 1, "REFERENCE", 110, True: .Columns(1).Alignment = xtpAlignmentCenter: .Columns(1).AllowRemove = False
        .Columns.Add 2, "INVOICE#/OR", 110, True: .Columns(2).Alignment = xtpAlignmentCenter: .Columns(2).AllowRemove = False
        .Columns.Add 3, "DEBIT", 110, True: .Columns(3).Alignment = xtpAlignmentRight: .Columns(3).AllowRemove = False
        .Columns.Add 4, "CREDIT", 110, True: .Columns(4).Alignment = xtpAlignmentRight: .Columns(4).AllowRemove = False
        .Columns.Add 5, "BALANCE", 80, True: .Columns(5).Alignment = xtpAlignmentRight: .Columns(5).AllowRemove = False
        .Columns.Add 6, "ID", 0, True: .Columns(6).Alignment = xtpAlignmentIconRight: .Columns(6).AllowRemove = False: .Columns(6).Visible = False
        .Columns.Add 7, "jtype", 0, True: .Columns(7).Alignment = xtpAlignmentIconRight: .Columns(7).AllowRemove = False: .Columns(7).Visible = False

        .PaintManager.HorizontalGridStyle = xtpGridSolid   ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSolid     ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
        .AllowColumnSort = False

        .ShowFooter = True

        .Columns(0).DrawFooterDivider = False
        .Columns(1).DrawFooterDivider = False
        .Columns(2).FooterText = "TOTAL : ": .Columns(2).FooterAlignment = xtpAlignmentCenter
        .Columns(3).FooterText = 0
        .Columns(4).FooterText = 0
        .Columns(5).FooterText = 0
        .Columns(3).FooterAlignment = xtpAlignmentRight
        .Columns(4).FooterAlignment = xtpAlignmentRight
        .Columns(5).FooterAlignment = xtpAlignmentRight
        .Columns(6).DrawFooterDivider = False
        .Columns(7).DrawFooterDivider = False
    End With
End Sub

Sub POPULATE_REPORT()
    Dim rsLEGDER                                  As ADODB.Recordset
    Set rsLEGDER = New ADODB.Recordset

    INITIALIZE_RCONTROL

    FORWARDED_BALANCES


    Picture2.Visible = True
    Picture2.ZOrder 0

    XX_BALANCE = 0
    G_TOTAL_DEBIT = 0
    G_TOTAL_CREDIT = 0

    If cboAccountName.Text = "ALL ACCOUNTS" Then
        If optCustomer.Value = True Then
            rsLEGDER.Open "SELECT DISTINCT HD.VOUCHERNO,HD.INVOICENO,HD.BANK,HD.ID,HD.JDATE,HD.JTYPE,HD.VOUCHERNO,HD.INVOICETYPE,HD.VENDORCODE,DET.ACCT_CODE, " & _
                          "DET.INVOICENO AS GJ_INVOICE,DET.INVOICETYPE AS GJ_INVOICETYPE,DET.ADJ_VOUCHERNO AS ADJ_VOUCHERNO,DET.ADJ_JTYPE AS ADJ_JTYPE,DET.IS_OTHERS FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                          "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                          "WHERE (LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND CUSTOMERCODE = " & N2Str2Null(txtCode.Text) & " AND HD.JTYPE IN ('COB','SJ','CRJ') AND HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "') " & _
                          "OR ((RIGHT(ENTITY,6) = " & N2Str2Null(txtCode.Text) & ") AND LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND DET.DEBIT <> 0) ORDER BY HD.JDATE ASC", gconDMIS, adOpenKeyset
        Else
            rsLEGDER.Open "SELECT DISTINCT HD.VOUCHERNO,HD.INVOICENO,HD.ID,HD.JDATE,HD.JTYPE,HD.VOUCHERNO,HD.INVOICETYPE,HD.VENDORCODE,DET.ACCT_CODE, " & _
                          "DET.INVOICENO AS GJ_INVOICE,DET.INVOICETYPE AS GJ_INVOICETYPE,DET.ADJ_VOUCHERNO AS ADJ_VOUCHERNO,DET.ADJ_JTYPE AS ADJ_JTYPE,DET.IS_OTHERS FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                          "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                          "WHERE (LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND HD.VENDORCODE = " & N2Str2Null(txtCode.Text) & " AND HD.JTYPE IN ('APJ','CDJ','GJ') AND HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "') " & _
                          "OR ((RIGHT(ENTITY,6) = " & N2Str2Null(txtCode.Text) & ") AND LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND DET.DEBIT <> 0) ORDER BY HD.JDATE ASC", gconDMIS, adOpenKeyset
        End If
    Else
        If optCustomer.Value = True Then
            rsLEGDER.Open "SELECT DISTINCT HD.VOUCHERNO,HD.INVOICENO,HD.BANK,HD.ID,HD.JDATE,HD.JTYPE,HD.VOUCHERNO,HD.INVOICETYPE,HD.VENDORCODE,DET.ACCT_CODE, " & _
                          "DET.INVOICENO AS GJ_INVOICE,DET.INVOICETYPE AS GJ_INVOICETYPE,DET.ADJ_VOUCHERNO AS ADJ_VOUCHERNO,DET.ADJ_JTYPE AS ADJ_JTYPE,DET.IS_OTHERS FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                          "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                          "WHERE (DET.ACCT_CODE = " & N2Str2Null(Setacctcode(cboAccountName)) & " AND CUSTOMERCODE = " & N2Str2Null(txtCode.Text) & " AND HD.JTYPE IN ('COB','SJ','CRJ','GJ') AND HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "') " & _
                          "OR ((RIGHT(ENTITY,6) = " & N2Str2Null(txtCode.Text) & ") AND LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND DET.DEBIT <> 0) ORDER BY HD.JDATE ASC", gconDMIS, adOpenKeyset
        Else
            rsLEGDER.Open "DISTINCT HD.VOUCHERNO,HD.INVOICENO,HD.ID,HD.JDATE,HD.JTYPE,HD.VOUCHERNO,HD.INVOICETYPE,HD.VENDORCODE,DET.ACCT_CODE, " & _
                          "DET.INVOICENO AS GJ_INVOICE,DET.INVOICETYPE AS GJ_INVOICETYPE,DET.ADJ_VOUCHERNO AS ADJ_VOUCHERNO,DET.ADJ_JTYPE AS ADJ_JTYPE,DET.IS_OTHERS FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                          "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                          "WHERE (LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND HD.VENDORCODE = " & N2Str2Null(txtCode.Text) & " AND HD.JTYPE IN ('APJ','CDJ','GJ') AND HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "') " & _
                          "OR ((RIGHT(ENTITY,6) = " & N2Str2Null(txtCode.Text) & ") AND LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND DET.DEBIT <> 0) ORDER BY HD.JDATE ASC", gconDMIS, adOpenKeyset
        End If
    End If

    rptRO.Records.DeleteAll
    Set REC = rptRO.Records.Add

    REC.AddItem (Trim(dtFrom.Value))
    REC.AddItem (Trim("FWD BALANCE"))
    REC.AddItem (Trim(""))
    REC.AddItem (Trim(ToDoubleNumber(0)))
    REC.AddItem (Trim(ToDoubleNumber(0)))
    REC.AddItem (Trim(ToDoubleNumber(FORWARDED_BALANCE)))
    REC.AddItem ("")

    rptRO.Populate
    Set REC = Nothing
    If Not rsLEGDER.EOF And Not rsLEGDER.BOF Then
        ProgressBar2.Value = 0
        ProgressBar2.Max = rsLEGDER.RecordCount

        Do While Not rsLEGDER.EOF
            If Null2String(rsLEGDER!jtype) = "COB" Then
                Set REC = rptRO.Records.Add
                REC.AddItem (Trim(Null2Date(rsLEGDER!JDate)))
                REC.AddItem (Trim(Null2String(rsLEGDER!jtype) & "-" & Null2String(rsLEGDER!VOUCHERNO)))
                REC.AddItem (Trim(Null2String(rsLEGDER!InvoiceType)) & "-" & Null2String(rsLEGDER!INVOICENO))

                If POSITIVE_NEGATIVE(Null2String(rsLEGDER!VOUCHERNO), Null2String(rsLEGDER!jtype)) = True Then
                    REC.AddItem (Trim(ToDoubleNumber(SUM_COB_DEBIT(Null2String(rsLEGDER!VOUCHERNO), Null2String(rsLEGDER!jtype)))))
                    REC.AddItem (Trim(ToDoubleNumber(0)))
                Else
                    REC.AddItem (Trim(ToDoubleNumber(0)))
                    REC.AddItem (Trim(ToDoubleNumber(SUM_COB_DEBIT(Null2String(rsLEGDER!VOUCHERNO), Null2String(rsLEGDER!jtype)))))
                End If
                REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))

                REC.AddItem (rsLEGDER!ID)
                REC.AddItem (rsLEGDER!jtype)

                rptRO.Populate
                Set REC = Nothing

                Call FIND_CRJ(Null2String(rsLEGDER!INVOICENO), Null2String(rsLEGDER!InvoiceType), txtCode.Text, Null2String(rsLEGDER!Acct_code))
                Call FIND_VOUCHERNO_ADJUSTMENT(Null2String(rsLEGDER!VOUCHERNO), Null2String(rsLEGDER!jtype), txtCode.Text, Null2String(rsLEGDER!Acct_code))

            ElseIf Null2String(rsLEGDER!jtype) = "SJ" Then
                Set REC = rptRO.Records.Add

                REC.AddItem (Trim(rsLEGDER!JDate))
                REC.AddItem (Trim(rsLEGDER!jtype & "-" & rsLEGDER!VOUCHERNO))
                REC.AddItem (Trim(rsLEGDER!InvoiceType) & "-" & rsLEGDER!INVOICENO)
                REC.AddItem (Trim(ToDoubleNumber(SUM_SJ_DEBIT(Null2String(rsLEGDER!VOUCHERNO), Null2String(rsLEGDER!jtype), Null2String(rsLEGDER!Acct_code)))))
                REC.AddItem (Trim(ToDoubleNumber(0)))
                REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))
                REC.AddItem (rsLEGDER!ID)
                REC.AddItem (rsLEGDER!jtype)

                rptRO.Populate
                Set REC = Nothing

                Call FIND_CRJ(Null2String(rsLEGDER!INVOICENO), Null2String(rsLEGDER!InvoiceType), txtCode.Text, Null2String(rsLEGDER!Acct_code))
                Call FIND_VOUCHERNO_ADJUSTMENT(Null2String(rsLEGDER!VOUCHERNO), Null2String(rsLEGDER!jtype), txtCode.Text, Null2String(rsLEGDER!Acct_code))

                'THIS IS INVOICE TO INVOICE ADJUSTMENT
                Call GJ_INVOICE_TO_INVOICE(Null2String(rsLEGDER!INVOICENO), Null2String(rsLEGDER!InvoiceType), Null2String(rsLEGDER!Acct_code), txtCode.Text, Null2String(rsLEGDER!VOUCHERNO), Null2String(rsLEGDER!jtype))

            ElseIf Null2String(rsLEGDER!jtype) = "APJ" Then
                Call GET_APJ_AMOUNT(Null2String(rsLEGDER!VOUCHERNO), Null2String(rsLEGDER!jtype), Null2String(rsLEGDER!Acct_code), Null2String(rsLEGDER!VendorCode))
                Call FIND_VOUCHERNO_ADJUSTMENT(Null2String(rsLEGDER!VOUCHERNO), Null2String(rsLEGDER!jtype), txtCode.Text, Null2String(rsLEGDER!Acct_code))

                'THIS IS FOR APJ NO TO APJ NO ADJUSTMENT
                Call GJ_INVOICE_TO_INVOICE(Null2String(rsLEGDER!INVOICENO), Null2String(rsLEGDER!InvoiceType), Null2String(rsLEGDER!Acct_code), txtCode.Text, Null2String(rsLEGDER!VOUCHERNO), Null2String(rsLEGDER!jtype))

            ElseIf Null2String(rsLEGDER!jtype) = "CDJ" Then
                Call GET_CDJ_AMOUNT(Null2String(rsLEGDER!VOUCHERNO), Null2String(rsLEGDER!jtype), txtCode.Text, Null2String(rsLEGDER!Acct_code))
                Call FIND_VOUCHERNO_ADJUSTMENT(Null2String(rsLEGDER!VOUCHERNO), Null2String(rsLEGDER!jtype), txtCode.Text, Null2String(rsLEGDER!Acct_code))

            ElseIf Null2String(rsLEGDER!jtype) = "CRJ" Then
                If COMPANY_CODE = "HGC" And Null2String(rsLEGDER!Acct_code) = "11-02002-00" Then
                    If IS_CRJ_AR(Null2String(rsLEGDER!VOUCHERNO), RTrim(LTrim(txtCode.Text)), Null2String(rsLEGDER!Acct_code)) = True Then
                        Set REC = rptRO.Records.Add
                        REC.AddItem (Trim(rsLEGDER!JDate))
                        REC.AddItem (Trim(rsLEGDER!jtype & "-" & rsLEGDER!VOUCHERNO))
                        REC.AddItem (Trim(rsLEGDER!InvoiceType) & "-" & rsLEGDER!INVOICENO)
                        REC.AddItem (Trim(ToDoubleNumber(GET_AR_CRJ(Null2String(rsLEGDER!VOUCHERNO), Null2String(rsLEGDER!jtype), Null2String(rsLEGDER!Acct_code)))))
                        REC.AddItem (Trim(ToDoubleNumber(0)))
                        REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))
                        REC.AddItem (rsLEGDER!ID)
                        REC.AddItem (rsLEGDER!jtype)

                        rptRO.Populate
                        Set REC = Nothing

                        Call FIND_VOUCHERNO_ADJUSTMENT(Null2String(rsLEGDER!VOUCHERNO), Null2String(rsLEGDER!jtype), txtCode.Text, Null2String(rsLEGDER!Acct_code))
                    End If
                End If
            ElseIf Null2String(rsLEGDER!jtype) = "GJ" Then
                'DET.INVOICENO AS GJ_INVOICE,DET.ADJ_JTYPE AS ADJ_JTYPE
                If IF_HAS_GJ_AR(Null2String(rsLEGDER!ADJ_VOUCHERNO), Null2String(rsLEGDER!ADJ_JTYPE), Null2String(rsLEGDER!Acct_code), rsLEGDER!IS_OTHERS) = True Then
                    Set REC = rptRO.Records.Add
                    REC.AddItem (Trim(rsLEGDER!JDate))
                    REC.AddItem (Trim(rsLEGDER!jtype & "-" & rsLEGDER!VOUCHERNO))
                    REC.AddItem (Trim(rsLEGDER!ADJ_JTYPE) & "-" & rsLEGDER!ADJ_VOUCHERNO)
                    REC.AddItem (Trim(ToDoubleNumber(GET_GJ_AR(Null2String(rsLEGDER!Acct_code), txtCode.Text, Null2String(rsLEGDER!ADJ_VOUCHERNO), Null2String(rsLEGDER!ADJ_JTYPE), rsLEGDER!IS_OTHERS))))
                    REC.AddItem (Trim(ToDoubleNumber(0)))
                    REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))
                    REC.AddItem (rsLEGDER!ID)
                    REC.AddItem (rsLEGDER!jtype)

                    rptRO.Populate
                    Set REC = Nothing

                    Call FIND_CONTROL_NUMBER_ADJUSTMENT(Null2String(rsLEGDER!ADJ_VOUCHERNO), txtCode.Text, Null2String(rsLEGDER!Acct_code))
                Else
                    Call ADJ_AGAINTS_NO_AR_ACCOUNT(Null2String(rsLEGDER!VOUCHERNO), Null2String(rsLEGDER!jtype), Null2String(rsLEGDER!ADJ_VOUCHERNO), Null2String(rsLEGDER!ADJ_JTYPE), txtCode.Text)
                End If
            End If
            labvoucherno.Caption = Null2String(rsLEGDER!jtype) & "-" & Null2String(rsLEGDER!VOUCHERNO)
            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round((ProgressBar2.Value / ProgressBar2.Max) * 100, 0) & "%" & " Completed"
            DoEvents
            rsLEGDER.MoveNext
        Loop
    End If

    'THIS IS FOR CRJ_PAYMENT WHICH HAS NO REFERENCE IN SALES JOURNAL
    Call CRJ_NO_SJ

    'THIS IS FOR GJ ADJUSTENT  WHERE GJ JDATE > = DATE FROM
    Call FWD_ADVANCE_GJ

    'THIS IS FOR PAYMENT CRJ WITH SJ WHERE CRJ JDATE >= DATE FROM AND SJ JDATE < DATE FROM
    Call ADVANCE_CRJ_PAYMENT

    'THIS IS FOR DIPLAYING THE TOTAL DEBIT AND CREDIT
    rptRO.Columns(3).FooterText = ToDoubleNumber(G_TOTAL_DEBIT)
    rptRO.Columns(4).FooterText = ToDoubleNumber(G_TOTAL_CREDIT)
    rptRO.Columns(5).FooterText = ToDoubleNumber(XX_BALANCE + FORWARDED_BALANCE)

    Picture2.Visible = False
    Picture2.ZOrder 1
End Sub

Function IF_HAS_GJ_AR(xADJ_VOUCHERNO As String, xADJ_JTYPE As String, xACCT_CODE As String, xIS_OTHERS As Boolean) As Boolean
'DESCRIPTION: THIS IS TO CHECK IF IN GJ HAS AN AR ACCOUNT
    Dim rsGET_GJ_AR                               As ADODB.Recordset
    Set rsGET_GJ_AR = New ADODB.Recordset

    rsGET_GJ_AR.Open "SELECT * " & _
                     "FROM AMIS_JOURNAL_DET WHERE ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND INVOICENO IS NULL  AND INVOICETYPE IS NULL AND ADJ_VOUCHERNO = " & N2Str2Null(xADJ_VOUCHERNO) & " AND ADJ_JTYPE = " & N2Str2Null(xADJ_JTYPE) & " AND RIGHT(ENTITY,6) = '" & txtCode.Text & "' AND STATUS = 'P' and JDATE >= '" & dtFrom & "' AND JDATE <= '" & dtTo & "' AND DEBIT <> 0 AND IS_OTHERS = 1", gconDMIS, adOpenKeyset

    If Not rsGET_GJ_AR.EOF And Not rsGET_GJ_AR.BOF Then
        IF_HAS_GJ_AR = True
    Else
        IF_HAS_GJ_AR = False
    End If
    Set rsGET_GJ_AR = Nothing
End Function


Function GET_GJ_AR(xACCT_CODE As String, xCUST_CODE As String, xADJ_VOUCHERNO As String, xADJ_JTYPE As String, xIS_OTHERS As Boolean) As Double
'DESCRIPTION: THIS IS TO GET THE AR IN GJ
    Dim rsGET_GJ_AR                               As ADODB.Recordset
    Set rsGET_GJ_AR = New ADODB.Recordset
    rsGET_GJ_AR.Open "SELECT ROUND(SUM(DEBIT),2) AS SUM_GJ_DEBIT " & _
                     "FROM AMIS_JOURNAL_DET WHERE ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND INVOICENO IS NULL AND INVOICETYPE IS NULL AND ADJ_VOUCHERNO = " & xADJ_VOUCHERNO & "  AND  ADJ_JTYPE = " & N2Str2Null(xADJ_JTYPE) & " AND RIGHT(ENTITY,6) = '" & txtCode.Text & "' AND STATUS = 'P' and JDATE >= '" & dtFrom & "' AND JDATE <= '" & dtTo & "' AND DEBIT <> 0 AND IS_OTHERS = 1", gconDMIS, adOpenKeyset

    If Not rsGET_GJ_AR.EOF And Not rsGET_GJ_AR.BOF Then
        GET_GJ_AR = NumericVal(rsGET_GJ_AR!SUM_GJ_DEBIT)

        G_TOTAL_DEBIT = Round((G_TOTAL_DEBIT + NumericVal(rsGET_GJ_AR!SUM_GJ_DEBIT)), 2)
        XX_BALANCE = Round((XX_BALANCE + NumericVal(rsGET_GJ_AR!SUM_GJ_DEBIT)), 2)
    Else
        GET_GJ_AR = NumericVal(0)
    End If
    Set rsGET_GJ_AR = Nothing
End Function

Sub FWD_GET_GJ_AR(xACCT_CODE As String, xCUST_CODE As String, xADJ_VOUCHERNO As String, xADJ_JTYPE As String, xIS_OTHERS As Boolean)
'DESCRIPTION: THIS IS FORWARD THE AR AMOUNT IN GJ
    Dim rsGET_GJ_AR                               As ADODB.Recordset
    Set rsGET_GJ_AR = New ADODB.Recordset
    rsGET_GJ_AR.Open "SELECT ROUND(SUM(DEBIT),2) AS SUM_GJ_DEBIT " & _
                     "FROM AMIS_JOURNAL_DET WHERE ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND INVOICENO IS NULL AND INVOICETYPE IS NULL AND ADJ_VOUCHERNO = " & xADJ_VOUCHERNO & "  AND  ADJ_JTYPE = " & N2Str2Null(xADJ_JTYPE) & " AND RIGHT(ENTITY,6) = '" & txtCode.Text & "' AND STATUS = 'P' AND JDATE < '" & dtFrom & "' AND DEBIT <> 0 AND IS_OTHERS = 1", gconDMIS, adOpenKeyset

    If Not rsGET_GJ_AR.EOF And Not rsGET_GJ_AR.BOF Then
        FORWARDED_BALANCE = Round((FORWARDED_BALANCE + NumericVal(rsGET_GJ_AR!SUM_GJ_DEBIT)), 2)
    Else
        FORWARDED_BALANCE = FORWARDED_BALANCE
    End If
    Set rsGET_GJ_AR = Nothing
End Sub

Function GET_GJ_PAYMENT(xINVOICENO As String, xADJ_JTYPE As String, xACCT_CODE As String) As Double
'DESCRIPTION: THIS IS TO GET THE GJ PAYMENT WITH THE REFERENCE INVOICE
    Dim rsGET_GJ_PAYMENT                          As ADODB.Recordset
    Set rsGET_GJ_PAYMENT = New ADODB.Recordset
    rsGET_GJ_PAYMENT.Open "SELECT DISTINCT HD.VOUCHERNO,HD.JDATE,HD.JTYPE,HD.VOUCHERNO,DET.INVOICENO,DET.ADJ_JTYPE,DET.INVOICETYPE,HD.ID,HD.JTYPE,DET.CREDIT " & _
                          "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE WHERE DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND DET.INVOICENO = " & N2Str2Null(xINVOICENO) & " AND DET.ADJ_JTYPE = " & N2Str2Null(xADJ_JTYPE) & " AND RIGHT(DET.ENTITY,6) = '" & txtCode.Text & "' AND HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "' AND DET.CREDIT <> 0", gconDMIS, adOpenKeyset
    If Not rsGET_GJ_PAYMENT.EOF And Not rsGET_GJ_PAYMENT.BOF Then
        Do While Not rsGET_GJ_PAYMENT.EOF
            Set REC = rptRO.Records.Add
            REC.AddItem (Trim(rsGET_GJ_PAYMENT!JDate))
            REC.AddItem (Trim(rsGET_GJ_PAYMENT!jtype & "-" & rsGET_GJ_PAYMENT!VOUCHERNO))
            REC.AddItem (Trim(rsGET_GJ_PAYMENT!ADJ_JTYPE) & "-" & rsGET_GJ_PAYMENT!INVOICENO)
            REC.AddItem (Trim(ToDoubleNumber(0)))
            REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsGET_GJ_PAYMENT!CREDIT))))

            G_TOTAL_CREDIT = Round((G_TOTAL_CREDIT + NumericVal(rsGET_GJ_PAYMENT!CREDIT)), 2)
            XX_BALANCE = Round((XX_BALANCE - NumericVal(rsGET_GJ_PAYMENT!CREDIT)), 2)

            REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))

            REC.AddItem (rsGET_GJ_PAYMENT!ID)
            REC.AddItem (rsGET_GJ_PAYMENT!jtype)

            rptRO.Populate
            Set REC = Nothing
            rsGET_GJ_PAYMENT.MoveNext
        Loop
    End If
    Set rsGET_GJ_PAYMENT = Nothing
End Function

Sub FWD_GET_GJ_PAYMENT(xINVOICENO As String, xADJ_JTYPE As String, xACCT_CODE As String)
'DESCRIPTION: THIS IS TO FORWARD THE GJ PAYMENT
    Dim rsGET_GJ_PAYMENT                          As ADODB.Recordset
    Set rsGET_GJ_PAYMENT = New ADODB.Recordset
    rsGET_GJ_PAYMENT.Open "SELECT DISTINCT HD.VOUCHERNO,HD.JDATE,HD.JTYPE,HD.VOUCHERNO,DET.INVOICENO,DET.ADJ_JTYPE,DET.INVOICETYPE,HD.ID,HD.JTYPE,DET.CREDIT " & _
                          "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE WHERE DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND DET.INVOICENO = " & N2Str2Null(xINVOICENO) & " AND DET.ADJ_JTYPE = " & N2Str2Null(xADJ_JTYPE) & " AND RIGHT(DET.ENTITY,6) = '" & txtCode.Text & "' AND HD.STATUS = 'P' AND HD.JDATE < '" & dtFrom & "' AND DET.CREDIT <> 0", gconDMIS, adOpenKeyset
    If Not rsGET_GJ_PAYMENT.EOF And Not rsGET_GJ_PAYMENT.BOF Then
        Do While Not rsGET_GJ_PAYMENT.EOF
            FORWARDED_BALANCE = Round((FORWARDED_BALANCE - NumericVal(rsGET_GJ_PAYMENT!CREDIT)), 2)
            rsGET_GJ_PAYMENT.MoveNext
        Loop
    End If
    Set rsGET_GJ_PAYMENT = Nothing
End Sub

Function IS_CRJ_AR(xVOUCHERNO As String, xBANK As String, xACCT_CODE As String) As Boolean
'DESCRIPTION: THIS TO CHECK IF THERE IS AN AR ACCOUNT IN CRJ MOSTLY ARE ACCOUNT RECEIVABLE CREDIT CARD
    Dim rsIS_CRJ_AR                               As ADODB.Recordset
    Set rsIS_CRJ_AR = New ADODB.Recordset
    rsIS_CRJ_AR.Open "SELECT * FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                     "WHERE HD.VOUCHERNO = '" & xVOUCHERNO & "' AND HD.JTYPE = 'CRJ' AND HD.BANK = '" & xBANK & "' AND DET.ACCT_CODE = '" & xACCT_CODE & "' AND DET.DEBIT <> 0 AND HD.STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsIS_CRJ_AR.EOF And Not rsIS_CRJ_AR.BOF Then
        IS_CRJ_AR = True
    Else
        IS_CRJ_AR = False
    End If
    Set rsIS_CRJ_AR = Nothing
End Function

Function POSITIVE_NEGATIVE(xVOUCHERNO As String, xJType As String) As Boolean
'DESCRIPTION: THIS IS DETERMINE THE OPENING BALANCE IF THE USER INPUTTED NEGATIVE AS BALANCE FOR THE PURPOSE OF DISPLAYING IN AR CUSTOMER LEDGER
    Dim rsPOSITIVE_NEGATIVE                       As ADODB.Recordset
    Set rsPOSITIVE_NEGATIVE = New ADODB.Recordset
    rsPOSITIVE_NEGATIVE.Open "SELECT ROUND(SUM(INVOICEAMT),2) AS SUM_COB_DEBIT FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & " AND JTYPE = " & N2Str2Null(xJType) & " AND JDATE >= '" & dtFrom & "' AND JDATE <= '" & dtTo & "' AND STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsPOSITIVE_NEGATIVE.EOF And Not rsPOSITIVE_NEGATIVE.BOF Then
        If NumericVal(rsPOSITIVE_NEGATIVE!SUM_COB_DEBIT) > 0 Then
            POSITIVE_NEGATIVE = True
        Else
            POSITIVE_NEGATIVE = False
        End If
    End If
    Set rsPOSITIVE_NEGATIVE = Nothing
End Function

Function GET_AR_CRJ(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
'DESCRIPTION: THIS IS TO GET THE AMOUNT OF IN CRJ MOST OF THE AR IN CRJ IS AR CREDIT CARD
    Dim rsGET_AR_CRJ                              As ADODB.Recordset
    Set rsGET_AR_CRJ = New ADODB.Recordset
    rsGET_AR_CRJ.Open "SELECT ROUND(SUM(DET.DEBIT),2) AS SUM_CRJ_AR FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE WHERE HD.VOUCHERNO = '" & xVOUCHERNO & "' AND HD.JTYPE = '" & xJType & "' AND DET.ACCT_CODE = '" & xACCT_CODE & "' AND HD.BANK = '" & RTrim(LTrim(txtCode.Text)) & "' AND HD.STATUS = 'P' AND DET.DEBIT <> 0", gconDMIS, adOpenKeyset
    If Not rsGET_AR_CRJ.EOF And Not rsGET_AR_CRJ.BOF Then
        GET_AR_CRJ = ToDoubleNumber(NumericVal(rsGET_AR_CRJ!SUM_CRJ_AR))
        G_TOTAL_DEBIT = Round((GET_AR_CRJ + NumericVal(rsGET_AR_CRJ!SUM_CRJ_AR)), 2)
        XX_BALANCE = Round((XX_BALANCE + NumericVal(rsGET_AR_CRJ!SUM_CRJ_AR)), 2)
    Else
        GET_AR_CRJ = ToDoubleNumber(0)
    End If
    Set rsGET_AR_CRJ = Nothing
End Function

Sub GET_CDJ_AMOUNT(xVOUCHERNO As String, xJType As String, xVENDORCODE As String, xACCT_CODE As String)
'DESCRIPTION: THIS IS TO GET THE CDJ AMOUNT
    Dim rsGET_CDJ_AMOUNT                          As ADODB.Recordset
    Set rsGET_CDJ_AMOUNT = New ADODB.Recordset
    rsGET_CDJ_AMOUNT.Open "SELECT DISTINCT HD.VOUCHERNO,HD.JDATE,HD.JTYPE,HD.VENDORCODE,DET.ACCT_CODE,DET.DEBIT,DET.CREDIT,HD.ID " & _
                          "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                          "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                          "WHERE DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND " & _
                          "HD.VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & " AND HD.JTYPE = " & N2Str2Null(xJType) & " " & _
                          "AND HD.VENDORCODE = " & N2Str2Null(xVENDORCODE) & "", gconDMIS, adOpenKeyset
    If Not rsGET_CDJ_AMOUNT.EOF And Not rsGET_CDJ_AMOUNT.BOF Then
        Do While Not rsGET_CDJ_AMOUNT.EOF
            Set REC = rptRO.Records.Add
            REC.AddItem (Trim(rsGET_CDJ_AMOUNT!JDate))
            REC.AddItem (Trim(rsGET_CDJ_AMOUNT!jtype & "-" & rsGET_CDJ_AMOUNT!VOUCHERNO))
            REC.AddItem (Trim(""))
            If NumericVal(rsGET_CDJ_AMOUNT!DEBIT) <> 0 Then
                REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsGET_CDJ_AMOUNT!DEBIT))))
                REC.AddItem (Trim(ToDoubleNumber(0)))

                XX_BALANCE = Round((XX_BALANCE + NumericVal(rsGET_CDJ_AMOUNT!DEBIT)), 2)
                G_TOTAL_DEBIT = Round((G_TOTAL_DEBIT + NumericVal(rsGET_CDJ_AMOUNT!DEBIT)), 2)

                'THIS IS FOR GETTING THE PAYMENT OR ADJUSTMENT IN GJ
                'Call GET_CDJ_PAYMENT_CREDIT(Null2String(rsGET_CDJ_AMOUNT!VOUCHERNO), Null2String(rsGET_CDJ_AMOUNT!jtype), Null2String(rsGET_CDJ_AMOUNT!VendorCode), Null2String(rsGET_CDJ_AMOUNT!ACCT_CODE))
            Else
                REC.AddItem (Trim(ToDoubleNumber(0)))
                REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsGET_CDJ_AMOUNT!CREDIT))))

                XX_BALANCE = Round((XX_BALANCE - NumericVal(rsGET_CDJ_AMOUNT!CREDIT)), 2)
                G_TOTAL_CREDIT = Round((G_TOTAL_CREDIT + NumericVal(rsGET_CDJ_AMOUNT!CREDIT)), 2)

                'Call GET_CDJ_PAYMENT_DEBIT(Null2String(rsGET_CDJ_AMOUNT!VOUCHERNO), Null2String(rsGET_CDJ_AMOUNT!jtype), Null2String(rsGET_CDJ_AMOUNT!VendorCode), Null2String(rsGET_CDJ_AMOUNT!ACCT_CODE))
            End If
            REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))
            REC.AddItem (rsGET_CDJ_AMOUNT!ID)
            REC.AddItem (rsGET_CDJ_AMOUNT!jtype)
            rptRO.Populate

            Set REC = Nothing

            If NumericVal(rsGET_CDJ_AMOUNT!DEBIT) <> 0 Then
                Call GET_CDJ_PAYMENT_CREDIT(Null2String(rsGET_CDJ_AMOUNT!VOUCHERNO), Null2String(rsGET_CDJ_AMOUNT!jtype), Null2String(rsGET_CDJ_AMOUNT!VendorCode), Null2String(rsGET_CDJ_AMOUNT!Acct_code))
            ElseIf NumericVal(rsGET_CDJ_AMOUNT!CREDIT) <> 0 Then
                Call GET_CDJ_PAYMENT_DEBIT(Null2String(rsGET_CDJ_AMOUNT!VOUCHERNO), Null2String(rsGET_CDJ_AMOUNT!jtype), Null2String(rsGET_CDJ_AMOUNT!VendorCode), Null2String(rsGET_CDJ_AMOUNT!Acct_code))
            End If

            Set REC = Nothing
            rsGET_CDJ_AMOUNT.MoveNext
        Loop
    End If
    Set rsGET_CDJ_AMOUNT = Nothing
End Sub

Sub FWD_GET_CDJ_AMOUNT(xVOUCHERNO As String, xJType As String, xVENDORCODE As String, xACCT_CODE As String)
'DESCRIPTION: THIS IS TO FORWARD THE BALANCE OF CDJ AMOUNT
    Dim rsGET_CDJ_AMOUNT                          As ADODB.Recordset
    Set rsGET_CDJ_AMOUNT = New ADODB.Recordset
    rsGET_CDJ_AMOUNT.Open "SELECT DISTINCT HD.VOUCHERNO,HD.JDATE,HD.JTYPE,HD.VENDORCODE,DET.ACCT_CODE,DET.DEBIT,DET.CREDIT,HD.ID " & _
                          "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                          "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                          "WHERE DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND " & _
                          "HD.VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & " AND HD.JTYPE = " & N2Str2Null(xJType) & " " & _
                          "AND HD.VENDORCODE = " & N2Str2Null(xVENDORCODE) & " AND HD.JDATE < '" & dtFrom.Value & "'", gconDMIS, adOpenKeyset
    If Not rsGET_CDJ_AMOUNT.EOF And Not rsGET_CDJ_AMOUNT.BOF Then
        Do While Not rsGET_CDJ_AMOUNT.EOF
            If NumericVal(rsGET_CDJ_AMOUNT!DEBIT) <> 0 Then
                FORWARDED_BALANCE = Round((FORWARDED_BALANCE + NumericVal(rsGET_CDJ_AMOUNT!DEBIT)), 2)
            Else
                FORWARDED_BALANCE = Round((FORWARDED_BALANCE - NumericVal(rsGET_CDJ_AMOUNT!CREDIT)), 2)
            End If

            If NumericVal(rsGET_CDJ_AMOUNT!DEBIT) <> 0 Then
                Call FWD_GET_CDJ_PAYMENT_CREDIT(Null2String(rsGET_CDJ_AMOUNT!VOUCHERNO), Null2String(rsGET_CDJ_AMOUNT!jtype), Null2String(rsGET_CDJ_AMOUNT!VendorCode), Null2String(rsGET_CDJ_AMOUNT!Acct_code))
            ElseIf NumericVal(rsGET_CDJ_AMOUNT!CREDIT) <> 0 Then
                Call FWD_GET_CDJ_PAYMENT_DEBIT(Null2String(rsGET_CDJ_AMOUNT!VOUCHERNO), Null2String(rsGET_CDJ_AMOUNT!jtype), Null2String(rsGET_CDJ_AMOUNT!VendorCode), Null2String(rsGET_CDJ_AMOUNT!Acct_code))
            End If
            rsGET_CDJ_AMOUNT.MoveNext
        Loop
    End If
    Set rsGET_CDJ_AMOUNT = Nothing
End Sub

Sub FWD_GET_CDJ_PAYMENT_CREDIT(xVOUCHERNO As String, xJType As String, xVENDORCODE As String, xACCT_CODE As String)
'THIS IS FOR PAYMENT OF CDJ
    Dim rsGET_CDJ_PAYMENT                         As ADODB.Recordset
    Set rsGET_CDJ_PAYMENT = New ADODB.Recordset
    rsGET_CDJ_PAYMENT.Open "SELECT DISTINCT HD.VOUCHERNO,HD.JDATE,HD.JTYPE,DET.INVOICENO,DET.INVOICETYPE,DET.CREDIT,HD.ID " & _
                           "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                           "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                           "WHERE DET.INVOICENO = '" & xVOUCHERNO & "' AND DET.ADJ_JTYPE = '" & xJType & "' " & _
                           "AND RIGHT(DET.ENTITY,6) = '" & xVENDORCODE & "' AND DET.ACCT_CODE = '" & xACCT_CODE & "' AND DET.CREDIT <> 0 AND HD.STATUS = 'P' AND HD.JDATE < '" & dtFrom & "'", gconDMIS, adOpenKeyset
    If Not rsGET_CDJ_PAYMENT.EOF And Not rsGET_CDJ_PAYMENT.BOF Then
        Do While Not rsGET_CDJ_PAYMENT.EOF
            FORWARDED_BALANCE = (Trim(ToDoubleNumber(Round((FORWARDED_BALANCE - NumericVal(rsGET_CDJ_PAYMENT!CREDIT)), 2))))
            rsGET_CDJ_PAYMENT.MoveNext
        Loop
    End If
    Set rsGET_CDJ_PAYMENT = Nothing
End Sub



Sub GET_CDJ_PAYMENT_CREDIT(xVOUCHERNO As String, xJType As String, xVENDORCODE As String, xACCT_CODE As String)
'THIS IS FOR PAYMENT OF CDJ
    Dim rsGET_CDJ_PAYMENT                         As ADODB.Recordset
    Set rsGET_CDJ_PAYMENT = New ADODB.Recordset
    rsGET_CDJ_PAYMENT.Open "SELECT DISTINCT HD.VOUCHERNO,HD.JDATE,HD.JTYPE,DET.INVOICENO,DET.INVOICETYPE,DET.CREDIT,HD.ID,DET.ADJ_JTYPE " & _
                           "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                           "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                           "WHERE DET.INVOICENO = '" & xVOUCHERNO & "' AND DET.ADJ_JTYPE = '" & xJType & "' " & _
                           "AND RIGHT(DET.ENTITY,6) = '" & xVENDORCODE & "' AND DET.ACCT_CODE = '" & xACCT_CODE & "' AND DET.CREDIT <> 0 AND HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "'", gconDMIS, adOpenKeyset
    If Not rsGET_CDJ_PAYMENT.EOF And Not rsGET_CDJ_PAYMENT.BOF Then
        Do While Not rsGET_CDJ_PAYMENT.EOF
            Set REC = rptRO.Records.Add
            REC.AddItem (Trim(rsGET_CDJ_PAYMENT!JDate))
            REC.AddItem (Trim(rsGET_CDJ_PAYMENT!jtype & "-" & rsGET_CDJ_PAYMENT!VOUCHERNO))
            REC.AddItem (Trim(rsGET_CDJ_PAYMENT!ADJ_JTYPE) & "-" & rsGET_CDJ_PAYMENT!INVOICENO)
            REC.AddItem (Trim(ToDoubleNumber(0)))
            REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsGET_CDJ_PAYMENT!CREDIT))))

            XX_BALANCE = (Trim(ToDoubleNumber(Round((XX_BALANCE - NumericVal(rsGET_CDJ_PAYMENT!CREDIT)), 2))))
            G_TOTAL_CREDIT = (Trim(ToDoubleNumber(Round((G_TOTAL_CREDIT + NumericVal(rsGET_CDJ_PAYMENT!CREDIT)), 2))))

            REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))

            REC.AddItem (Trim((rsGET_CDJ_PAYMENT!ID)))
            REC.AddItem (Trim((rsGET_CDJ_PAYMENT!jtype)))
            rptRO.Populate
            Set REC = Nothing
            rsGET_CDJ_PAYMENT.MoveNext
        Loop
    End If
    Set rsGET_CDJ_PAYMENT = Nothing
End Sub
Sub FWD_GET_CDJ_PAYMENT_DEBIT(xVOUCHERNO As String, xJType As String, xVENDORCODE As String, xACCT_CODE As String)
'THIS IS FOR PAYMENT OF CDJ CLOSING CREDIT ENTRY
    Dim rsGET_CDJ_PAYMENT                         As ADODB.Recordset
    Set rsGET_CDJ_PAYMENT = New ADODB.Recordset
    rsGET_CDJ_PAYMENT.Open "SELECT HD.JDATE,HD.JTYPE,HD.VOUCHERNO,DET.INVOICENO,DET.INVOICETYPE,DET.DEBIT,HD.ID " & _
                           "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                           "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                           "WHERE DET.INVOICENO = '" & xVOUCHERNO & "' AND DET.ADJ_JTYPE = '" & xJType & "' " & _
                           "AND RIGHT(DET.ENTITY,6) = '" & xVENDORCODE & "' AND DET.ACCT_CODE = '" & xACCT_CODE & "' AND DET.DEBIT <> 0 AND HD.STATUS = 'P' AND HD.JDATE < '" & dtFrom & "'", gconDMIS, adOpenKeyset
    If Not rsGET_CDJ_PAYMENT.EOF And Not rsGET_CDJ_PAYMENT.BOF Then
        Do While Not rsGET_CDJ_PAYMENT.EOF
            Set REC = rptRO.Records.Add
            FORWARDED_BALANCE = (Trim(ToDoubleNumber(Round((FORWARDED_BALANCE + NumericVal(rsGET_CDJ_PAYMENT!DEBIT)), 2))))
            rsGET_CDJ_PAYMENT.MoveNext
        Loop
    End If
    Set rsGET_CDJ_PAYMENT = Nothing
End Sub

Sub GET_CDJ_PAYMENT_DEBIT(xVOUCHERNO As String, xJType As String, xVENDORCODE As String, xACCT_CODE As String)
'THIS IS FOR PAYMENT OF CDJ CLOSING CREDIT ENTRY
    Dim rsGET_CDJ_PAYMENT                         As ADODB.Recordset
    Set rsGET_CDJ_PAYMENT = New ADODB.Recordset
    rsGET_CDJ_PAYMENT.Open "SELECT HD.JDATE,HD.JTYPE,HD.VOUCHERNO,DET.INVOICENO,DET.INVOICETYPE,DET.DEBIT,HD.ID,DET.ADJ_JTYPE " & _
                           "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                           "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                           "WHERE DET.INVOICENO = '" & xVOUCHERNO & "' AND DET.ADJ_JTYPE = '" & xJType & "' " & _
                           "AND RIGHT(DET.ENTITY,6) = '" & xVENDORCODE & "' AND DET.ACCT_CODE = '" & xACCT_CODE & "' AND DET.DEBIT <> 0 AND HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "'", gconDMIS, adOpenKeyset
    If Not rsGET_CDJ_PAYMENT.EOF And Not rsGET_CDJ_PAYMENT.BOF Then
        Do While Not rsGET_CDJ_PAYMENT.EOF
            Set REC = rptRO.Records.Add
            REC.AddItem (Trim(rsGET_CDJ_PAYMENT!JDate))
            REC.AddItem (Trim(rsGET_CDJ_PAYMENT!jtype & "-" & rsGET_CDJ_PAYMENT!VOUCHERNO))
            REC.AddItem (Trim(rsGET_CDJ_PAYMENT!ADJ_JTYPE) & "-" & rsGET_CDJ_PAYMENT!INVOICENO)
            REC.AddItem (Trim(ToDoubleNumber(rsGET_CDJ_PAYMENT!DEBIT)))
            REC.AddItem (Trim(ToDoubleNumber(0)))

            XX_BALANCE = (Trim(ToDoubleNumber(Round((XX_BALANCE + NumericVal(rsGET_CDJ_PAYMENT!DEBIT)), 2))))
            G_TOTAL_DEBIT = (Trim(ToDoubleNumber(Round((G_TOTAL_DEBIT + NumericVal(rsGET_CDJ_PAYMENT!DEBIT)), 2))))

            REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))
            REC.AddItem (Trim((rsGET_CDJ_PAYMENT!ID)))
            REC.AddItem (Trim((rsGET_CDJ_PAYMENT!jtype)))
            rptRO.Populate
            Set REC = Nothing
            rsGET_CDJ_PAYMENT.MoveNext
        Loop
    End If
    Set rsGET_CDJ_PAYMENT = Nothing
End Sub

Sub FWD_ADVANCE_GJ()
'DESCRIPTION: THIS IS TO FORWARD THE AMOUNT OF ADVANCE GJ
    Dim rsADVANCE_GJ                              As ADODB.Recordset
    Set rsADVANCE_GJ = New ADODB.Recordset
    'THIS IS FOR GJ ADJUSTENT  WHERE GJ JDATE > = DATE FROM
    If RTrim(LTrim(cboAccountName.Text)) = "ALL ACCOUNTS" Then
        rsADVANCE_GJ.Open "SELECT X.INVNO_INVTYPE,X.CUST_CODE,X.JDATE,X.VOUCHERNO,X.JTYPE,X.INVOICENO,X.INVOICETYPE,X.DEBIT,X.CREDIT,X.ID FROM " & _
                          "( " & _
                          "SELECT DISTINCT DET.INVOICETYPE + '-' + DET.INVOICENO AS INVNO_INVTYPE, RIGHT(DET.ENTITY,6) AS CUST_CODE, HD.JDATE AS JDATE,HD.ID AS ID, " & _
                          "HD.VOUCHERNO As VOUCHERNO, HD.JTYPE As JTYPE, DET.INVOICENO As INVOICENO, DET.INVOICETYPE As INVOICETYPE, DET.DEBIT, DET.CREDIT " & _
                          "FROM AMIS_JOURNAL_DET DET INNER JOIN AMIS_JOURNAL_HD HD ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE  WHERE HD.JDATE >= '" & dtFrom.Value & "' AND RIGHT(DET.ENTITY,6) = '" & txtCode.Text & "' AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03') AND HD.STATUS = 'P' AND HD.JTYPE = 'GJ' " & _
                          ") " & _
                          "X WHERE X.INVNO_INVTYPE IN (SELECT HD.INVOICETYPE + '-' + HD.INVOICENO FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE WHERE HD.JDATE < '" & dtFrom.Value & "' AND HD.CUSTOMERCODE = X.CUST_CODE AND LEFT(ACCT_CODE,5) IN ('11-02','11-03') AND HD.STATUS = 'P' AND HD.JTYPE IN ('SJ','COB'))", gconDMIS, adOpenKeyset
    Else
        rsADVANCE_GJ.Open "SELECT X.INVNO_INVTYPE,X.CUST_CODE,X.JDATE,X.VOUCHERNO,X.JTYPE,X.INVOICENO,X.INVOICETYPE,X.DEBIT,X.CREDIT,X.ID FROM " & _
                          "( " & _
                          "SELECT DISTINCT DET.INVOICETYPE + '-' + DET.INVOICENO AS INVNO_INVTYPE, RIGHT(DET.ENTITY,6) AS CUST_CODE, HD.JDATE AS JDATE,HD.ID AS ID, " & _
                          "HD.VOUCHERNO As VOUCHERNO, HD.JTYPE As JTYPE, DET.INVOICENO As INVOICENO, DET.INVOICETYPE As INVOICETYPE, DET.DEBIT, DET.CREDIT " & _
                          "FROM AMIS_JOURNAL_DET DET INNER JOIN AMIS_JOURNAL_HD HD ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE  WHERE HD.JDATE >= '" & dtFrom.Value & "' AND RIGHT(DET.ENTITY,6) = '" & txtCode.Text & "' AND DET.ACCT_CODE = '" & Setacctcode(cboAccountName.Text) & "' AND HD.STATUS = 'P' AND HD.JTYPE = 'GJ' " & _
                          ") " & _
                          "X WHERE X.INVNO_INVTYPE IN (SELECT HD.INVOICETYPE + '-' + HD.INVOICENO FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE WHERE HD.JDATE < '" & dtFrom.Value & "' AND HD.CUSTOMERCODE = X.CUST_CODE AND DET.ACCT_CODE = '" & Setacctcode(cboAccountName.Text) & "' AND HD.STATUS = 'P' AND HD.JTYPE IN ('SJ','COB'))", gconDMIS, adOpenKeyset
    End If

    If Not rsADVANCE_GJ.EOF And Not rsADVANCE_GJ.BOF Then
        Do While Not rsADVANCE_GJ.EOF
            Set REC = rptRO.Records.Add
            REC.AddItem (Trim(Null2String(rsADVANCE_GJ!JDate)))
            REC.AddItem (Trim(Null2String(rsADVANCE_GJ!jtype) & "-" & Null2String(rsADVANCE_GJ!VOUCHERNO)))
            REC.AddItem (Trim(Null2String(rsADVANCE_GJ!InvoiceType)) & "-" & Null2String(rsADVANCE_GJ!INVOICENO))
            If NumericVal(rsADVANCE_GJ!DEBIT) <> 0 Then
                REC.AddItem (Trim(rsADVANCE_GJ!DEBIT))
                REC.AddItem (Trim(ToDoubleNumber(0)))

                XX_BALANCE = ToDoubleNumber(Round((XX_BALANCE + NumericVal(rsADVANCE_GJ!DEBIT)), 2))
                G_TOTAL_DEBIT = ToDoubleNumber(Round((G_TOTAL_DEBIT + NumericVal(rsADVANCE_GJ!DEBIT)), 2))
            Else
                REC.AddItem (Trim(ToDoubleNumber(0)))
                REC.AddItem (Trim(rsADVANCE_GJ!CREDIT))

                XX_BALANCE = ToDoubleNumber(Round((XX_BALANCE - NumericVal(rsADVANCE_GJ!CREDIT)), 2))
                G_TOTAL_CREDIT = Round((G_TOTAL_CREDIT + NumericVal(rsADVANCE_GJ!CREDIT)), 2)
            End If
            REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))
            REC.AddItem (rsADVANCE_GJ!ID)
            REC.AddItem (rsADVANCE_GJ!jtype)
            rptRO.Populate
            Set REC = Nothing
            rsADVANCE_GJ.MoveNext
        Loop
    End If
    Set rsADVANCE_GJ = Nothing
End Sub

Sub GET_APJ_AMOUNT(xVOUCHERNO As String, xJType As String, xACCT_CODE As String, xVENDORCODE As String)
'DESCRIPTION: THIS IS TO GET THE AMOUNT OF APJ
    Dim rsSUM_APJ_DEBIT                           As ADODB.Recordset
    Set rsSUM_APJ_DEBIT = New ADODB.Recordset
    rsSUM_APJ_DEBIT.Open "SELECT HD.JDATE,HD.JTYPE,HD.VOUCHERNO,DET.DEBIT,DET.CREDIT,HD.ID " & _
                         "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                         "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                         "WHERE DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND " & _
                         "HD.VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & " AND HD.JTYPE = " & N2Str2Null(xJType) & " " & _
                         "AND HD.VENDORCODE = " & N2Str2Null(xVENDORCODE) & "", gconDMIS, adOpenKeyset
    If Not rsSUM_APJ_DEBIT.EOF And Not rsSUM_APJ_DEBIT.BOF Then
        Do While Not rsSUM_APJ_DEBIT.EOF
            Set REC = rptRO.Records.Add
            'If Null2String(rsSUM_APJ_DEBIT!VOUCHERNO) = "000191" Then Stop
            REC.AddItem (Trim(rsSUM_APJ_DEBIT!JDate))
            REC.AddItem (Trim(rsSUM_APJ_DEBIT!jtype & "-" & rsSUM_APJ_DEBIT!VOUCHERNO))
            REC.AddItem (Trim(""))
            If NumericVal(rsSUM_APJ_DEBIT!DEBIT) <> 0 Then
                REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsSUM_APJ_DEBIT!DEBIT))))
                REC.AddItem (Trim(ToDoubleNumber(0)))

                XX_BALANCE = Round((XX_BALANCE + NumericVal(rsSUM_APJ_DEBIT!DEBIT)), 2)
                G_TOTAL_DEBIT = Round((G_TOTAL_DEBIT + NumericVal(rsSUM_APJ_DEBIT!DEBIT)), 2)
            Else
                REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsSUM_APJ_DEBIT!CREDIT))))
                REC.AddItem (Trim(ToDoubleNumber(0)))

                XX_BALANCE = Round((XX_BALANCE - NumericVal(rsSUM_APJ_DEBIT!CREDIT)), 2)
                G_TOTAL_CREDIT = Round((G_TOTAL_CREDIT + NumericVal(rsSUM_APJ_DEBIT!CREDIT)), 2)
            End If
            REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))
            REC.AddItem (rsSUM_APJ_DEBIT!ID)
            REC.AddItem (rsSUM_APJ_DEBIT!jtype)
            rptRO.Populate
            Set REC = Nothing
            rsSUM_APJ_DEBIT.MoveNext
        Loop
    End If
    Set rsSUM_APJ_DEBIT = Nothing
End Sub

Sub FWD_GET_APJ_AMOUNT(xVOUCHERNO As String, xJType As String, xACCT_CODE As String, xVENDORCODE As String)
'DESCRIPTION: THIS IS TO FORWARD THE BALANCE APJ AMOUNT
    Dim rsSUM_APJ_DEBIT                           As ADODB.Recordset
    Set rsSUM_APJ_DEBIT = New ADODB.Recordset
    rsSUM_APJ_DEBIT.Open "SELECT HD.JDATE,HD.JTYPE,HD.VOUCHERNO,DET.DEBIT,DET.CREDIT,HD.ID " & _
                         "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                         "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                         "WHERE DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND " & _
                         "HD.VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & " AND HD.JTYPE = " & N2Str2Null(xJType) & " " & _
                         "AND HD.VENDORCODE = " & N2Str2Null(xVENDORCODE) & " AND HD.JDATE < '" & dtFrom.Value & "'", gconDMIS, adOpenKeyset
    If Not rsSUM_APJ_DEBIT.EOF And Not rsSUM_APJ_DEBIT.BOF Then
        Do While Not rsSUM_APJ_DEBIT.EOF
            If NumericVal(rsSUM_APJ_DEBIT!DEBIT) <> 0 Then
                FORWARDED_BALANCE = Round((FORWARDED_BALANCE + NumericVal(rsSUM_APJ_DEBIT!DEBIT)), 2)
            ElseIf NumericVal(rsSUM_APJ_DEBIT!CREDIT) <> 0 Then
                FORWARDED_BALANCE = Round((FORWARDED_BALANCE + NumericVal(rsSUM_APJ_DEBIT!CREDIT)), 2)
            End If
            rsSUM_APJ_DEBIT.MoveNext
        Loop
    End If
    Set rsSUM_APJ_DEBIT = Nothing
End Sub

Sub CRJ_NO_SJ()
'DESCRIPTION:THIS IS CRJ WITH NOT REFERENCE SJ LIKE INVALID REFERENCE WAS USE DURING THE POSTING OF CRJ TRANSACTION
    Dim rsCRJ_NO_SJ                               As ADODB.Recordset
    Set rsCRJ_NO_SJ = New ADODB.Recordset
    If RTrim(LTrim(cboAccountName.Text)) = "ALL ACCOUNTS" Then
        rsCRJ_NO_SJ.Open "SELECT X.INV,X.CUST_CODE,X.JDATE,X.JTYPE,X.VOUCHERNO,X.INVOICETYPE,X.INVOICENO,X.INV_AMT,X.XID,X.ACCT_CODE FROM " & _
                         "( " & _
                         "SELECT DISTINCT CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS INV,HD.CUSTOMERCODE AS CUST_CODE,HD.JDATE AS JDATE,HD.JTYPE AS JTYPE, HD.VOUCHERNO AS VOUCHERNO, CRJ.INVOICENO AS INVOICENO, CRJ.INVOICETYPE AS INVOICETYPE,CRJ.INVOICEAMOUNT AS INV_AMT,HD.ID AS XID, DET.ACCT_CODE AS ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                         "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CRJ_DETAIL CRJ " & _
                         "ON HD.VOUCHERNO = CRJ.VOUCHERNO AND HD.JTYPE = CRJ.CR_TYPE " & _
                         "WHERE HD.JDATE >= '" & dtFrom.Value & "' AND HD.JDATE <= '" & dtTo.Value & "' AND HD.CUSTOMERCODE = '" & txtCode.Text & "' AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03') AND HD.STATUS = 'P' AND HD.JTYPE = 'CRJ'" & _
                         ") X WHERE X.INV NOT IN " & _
                         "( " & _
                         "SELECT HD.INVOICETYPE + '-' + HD.INVOICENO FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                         "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                         "WHERE HD.CUSTOMERCODE = X.CUST_CODE AND HD.JDATE >= '" & dtFrom.Value & "' AND HD.JDATE <= '" & dtTo.Value & "' AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03') AND HD.STATUS = 'P' AND HD.JTYPE IN ('SJ','COB') " & _
                         ")", gconDMIS, adOpenKeyset
    Else
        rsCRJ_NO_SJ.Open "SELECT X.INV,X.CUST_CODE,X.JDATE,X.JTYPE,X.VOUCHERNO,X.INVOICETYPE,X.INVOICENO,X.INV_AMT,X.XID,X.ACCT_CODE FROM " & _
                         "( " & _
                         "SELECT DISTINCT CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS INV,HD.CUSTOMERCODE AS CUST_CODE,HD.JDATE AS JDATE,HD.JTYPE AS JTYPE, HD.VOUCHERNO AS VOUCHERNO, CRJ.INVOICENO AS INVOICENO, CRJ.INVOICETYPE AS INVOICETYPE,CRJ.INVOICEAMOUNT AS INV_AMT,HD.ID AS XID, DET.ACCT_CODE AS ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                         "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CRJ_DETAIL CRJ " & _
                         "ON HD.VOUCHERNO = CRJ.VOUCHERNO AND HD.JTYPE = CRJ.CR_TYPE " & _
                         "WHERE HD.JDATE >= '" & dtFrom.Value & "' AND HD.JDATE <= '" & dtTo.Value & "' AND HD.CUSTOMERCODE = '" & txtCode.Text & "' AND DET.ACCT_CODE = '" & Setacctcode(cboAccountName.Text) & "' AND HD.STATUS = 'P' AND HD.JTYPE = 'CRJ' " & _
                         ") X WHERE X.INV NOT IN " & _
                         "( " & _
                         "SELECT HD.INVOICETYPE + '-' + HD.INVOICENO FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                         "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                         "WHERE HD.CUSTOMERCODE = X.CUST_CODE AND HD.JDATE >= '" & dtFrom.Value & "' AND HD.JDATE <= '" & dtTo.Value & "' AND DET.ACCT_CODE = '" & Setacctcode(cboAccountName.Text) & "' AND HD.STATUS = 'P' AND HD.JTYPE IN ('SJ','COB') " & _
                         ")", gconDMIS, adOpenKeyset
    End If
    If Not rsCRJ_NO_SJ.EOF And Not rsCRJ_NO_SJ.BOF Then

        Do While Not rsCRJ_NO_SJ.EOF
            If COMPANY_CODE = "HGC" And Null2String(rsCRJ_NO_SJ!Acct_code) <> "11-02002-00" Then
                Set REC = rptRO.Records.Add
                REC.AddItem (Trim(rsCRJ_NO_SJ!JDate))
                REC.AddItem (Trim(rsCRJ_NO_SJ!jtype & "-" & rsCRJ_NO_SJ!VOUCHERNO))
                REC.AddItem (Trim(rsCRJ_NO_SJ!InvoiceType) & "-" & rsCRJ_NO_SJ!INVOICENO)
                REC.AddItem (Trim(ToDoubleNumber(0)))
                REC.AddItem (Trim(rsCRJ_NO_SJ!INV_AMT))
                XX_BALANCE = (Trim(ToDoubleNumber(Round((XX_BALANCE - NumericVal(rsCRJ_NO_SJ!INV_AMT)), 2))))

                G_TOTAL_CREDIT = Round((G_TOTAL_CREDIT + NumericVal(rsCRJ_NO_SJ!INV_AMT)), 2)
                REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))

                REC.AddItem (Trim(rsCRJ_NO_SJ!xID))
                REC.AddItem (Trim(rsCRJ_NO_SJ!jtype))
                rptRO.Populate
                Set REC = Nothing

                'THIS IS VOUCHER TO VOUCHER WHERE IN VOUCHER AND JOURNALTYPE WAS USE AS A REFERENCE
                Call FIND_VOUCHERNO_ADJUSTMENT(Null2String(rsCRJ_NO_SJ!VOUCHERNO), Null2String(rsCRJ_NO_SJ!jtype), txtCode.Text, Null2String(rsCRJ_NO_SJ!Acct_code))

                'THIS IS ADJUSTMENT WHERE IN INVOICE NO AND INVOICE TYPE WAS USE AS REFERENCE
                Call FIND_INVOICEDETAIL_ADJUSTMENT(Null2String(rsCRJ_NO_SJ!INVOICENO), Null2String(rsCRJ_NO_SJ!INVOICENO), Null2String(rsCRJ_NO_SJ!Acct_code), txtCode.Text)
            End If
            rsCRJ_NO_SJ.MoveNext
        Loop
    End If
    Set rsCRJ_NO_SJ = Nothing
End Sub

Sub FIND_GJ_ADJUSTMENT(xINVOICENO As String, xINVOICETYPE As String, XCustomerCode As String, xACCT_CODE As String)
'DESCRIPTION: THIS IS TO FIND THE GJ ADJUSTMENT WITH THE REFERENCE INVOICENO,INVOICETYPE,CUSTOMERCODE, AND ACCT_CODE
    Dim rsFIND_GJ_ADJUSTMENT                      As ADODB.Recordset
    Set rsFIND_GJ_ADJUSTMENT = New ADODB.Recordset
    rsFIND_GJ_ADJUSTMENT.Open "SELECT HD.ID,HD.JDATE,HD.VOUCHERNO,HD.JTYPE,DET.INVOICENO,DET.INVOICETYPE,DET.CREDIT,DET.DEBIT " & _
                              "FROM AMIS_JOURNAL_DET DET INNER JOIN AMIS_JOURNAL_HD HD " & _
                              "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                              "WHERE DET.INVOICENO = " & N2Str2Null(xINVOICENO) & " AND DET.INVOICETYPE = " & N2Str2Null(xINVOICETYPE) & " " & _
                              "AND RIGHT(DET.ENTITY,6) = " & N2Str2Null(XCustomerCode) & " AND DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & "", gconDMIS, adOpenKeyset
    'If xINVOICENO = "010929" Then Stop

    If Not rsFIND_GJ_ADJUSTMENT.EOF And Not rsFIND_GJ_ADJUSTMENT.BOF Then
        Do While Not rsFIND_GJ_ADJUSTMENT.EOF
            Set REC = rptRO.Records.Add
            REC.AddItem (Trim(rsFIND_GJ_ADJUSTMENT!JDate))
            REC.AddItem (Trim(rsFIND_GJ_ADJUSTMENT!jtype & "-" & rsFIND_GJ_ADJUSTMENT!VOUCHERNO))
            REC.AddItem (Trim(rsFIND_GJ_ADJUSTMENT!InvoiceType) & "-" & rsFIND_GJ_ADJUSTMENT!INVOICENO)
            If NumericVal(rsFIND_GJ_ADJUSTMENT!CREDIT) <> 0 Then
                REC.AddItem (Trim(ToDoubleNumber(0)))
                REC.AddItem (Trim(ToDoubleNumber(rsFIND_GJ_ADJUSTMENT!CREDIT)))

                G_TOTAL_CREDIT = Round((G_TOTAL_CREDIT + NumericVal(rsFIND_GJ_ADJUSTMENT!CREDIT)), 2)
                XX_BALANCE = (Trim(ToDoubleNumber(Round((XX_BALANCE - NumericVal(rsFIND_GJ_ADJUSTMENT!CREDIT)), 2))))
            Else
                REC.AddItem (Trim(ToDoubleNumber(rsFIND_GJ_ADJUSTMENT!DEBIT)))
                REC.AddItem (Trim(ToDoubleNumber(0)))

                G_TOTAL_DEBIT = ToDoubleNumber(Round((G_TOTAL_DEBIT + NumericVal(rsFIND_GJ_ADJUSTMENT!DEBIT)), 2))
                XX_BALANCE = (Trim(ToDoubleNumber(Round((XX_BALANCE + NumericVal(rsFIND_GJ_ADJUSTMENT!DEBIT)), 2))))
            End If
            REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))
            REC.AddItem (Trim(rsFIND_GJ_ADJUSTMENT!ID))
            REC.AddItem (Trim(rsFIND_GJ_ADJUSTMENT!jtype))
            rptRO.Populate
            Set REC = Nothing
            rsFIND_GJ_ADJUSTMENT.MoveNext
        Loop
    End If
    Set rsFIND_GJ_ADJUSTMENT = Nothing
End Sub

Sub FIND_CRJ(xINVOICENO As String, xINVOICETYPE As String, XCustomerCode As String, xACCT_CODE As String)
'DESCRIPTION: THIS IS TO GET THE PAYMENT WITH THE REFERENCE INVOICENO,INVOICETYPE,ACCT_CODE AND CUSTOMER CODE
    Dim rsFIND_CRJ                                As ADODB.Recordset
    Set rsFIND_CRJ = New ADODB.Recordset
    rsFIND_CRJ.Open "SELECT DISTINCT CRJ.INVOICENO + '-' + CRJ.INVOICETYPE,HD.JTYPE,HD.VOUCHERNO,HD.ID,HD.JDATE,CRJ.INVOICENO,CRJ.INVOICETYPE,CRJ.INVOICEAMOUNT,DET.ACCT_CODE " & _
                    "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                    "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                    "INNER JOIN AMIS_CRJ_DETAIL CRJ ON HD.VOUCHERNO = CRJ.VOUCHERNO AND HD.JTYPE = CRJ.CR_TYPE " & _
                    "WHERE DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND HD.CUSTOMERCODE = " & N2Str2Null(XCustomerCode) & " AND CRJ.INVOICENO = " & N2Str2Null(xINVOICENO) & " AND CRJ.INVOICETYPE = " & N2Str2Null(xINVOICETYPE) & " AND HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "'", gconDMIS, adOpenKeyset
    If Not rsFIND_CRJ.EOF And Not rsFIND_CRJ.BOF Then
        Do While Not rsFIND_CRJ.EOF
            Set REC = rptRO.Records.Add
            REC.AddItem (Trim(Null2Date(rsFIND_CRJ!JDate)))
            REC.AddItem (Trim(Null2String(rsFIND_CRJ!jtype) & "-" & Null2String(rsFIND_CRJ!VOUCHERNO)))
            REC.AddItem (Trim(Null2String(rsFIND_CRJ!InvoiceType)) & "-" & Null2String(rsFIND_CRJ!INVOICENO))
            REC.AddItem (Trim(ToDoubleNumber(0)))
            REC.AddItem (Trim(ToDoubleNumber(rsFIND_CRJ!invoiceamount)))

            XX_BALANCE = Round((XX_BALANCE - ToDoubleNumber(rsFIND_CRJ!invoiceamount)), 2)
            G_TOTAL_CREDIT = Round((G_TOTAL_CREDIT + NumericVal(rsFIND_CRJ!invoiceamount)), 2)

            REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))
            REC.AddItem (rsFIND_CRJ!ID)
            REC.AddItem (rsFIND_CRJ!jtype)
            rptRO.Populate

            Call FIND_INVOICEDETAIL_ADJUSTMENT(Null2String(rsFIND_CRJ!INVOICENO), Null2String(rsFIND_CRJ!InvoiceType), Null2String(rsFIND_CRJ!Acct_code), txtCode.Text)
            Call FIND_VOUCHERNO_ADJUSTMENT(Null2String(rsFIND_CRJ!VOUCHERNO), Null2String(rsFIND_CRJ!jtype), txtCode.Text, Null2String(rsFIND_CRJ!Acct_code))
            rsFIND_CRJ.MoveNext
            Set REC = Nothing
        Loop
    Else
        'NO PAYMENTS
    End If
    Set rsFIND_CRJ = Nothing
End Sub

Function SUM_SJ_DEBIT(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
    Dim rsSUM_SJ_DEBIT                            As ADODB.Recordset
    Set rsSUM_SJ_DEBIT = New ADODB.Recordset
    rsSUM_SJ_DEBIT.Open "SELECT ROUND(SUM(DEBIT),2) AS SUM_DEBIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & " AND " & _
                        "JTYPE = " & N2Str2Null(xJType) & " AND ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND JDATE >= '" & dtFrom & "' AND JDATE <= '" & dtTo & "' AND STATUS = 'P' AND DEBIT <> 0", gconDMIS, adOpenKeyset
    If Not rsSUM_SJ_DEBIT.EOF And Not rsSUM_SJ_DEBIT.BOF Then
        SUM_SJ_DEBIT = ToDoubleNumber(NumericVal(rsSUM_SJ_DEBIT!SUM_DEBIT))

        G_TOTAL_DEBIT = ToDoubleNumber(Round((G_TOTAL_DEBIT + SUM_SJ_DEBIT), 2))
        XX_BALANCE = ToDoubleNumber(Round((XX_BALANCE + SUM_SJ_DEBIT), 2))
    Else
        SUM_SJ_DEBIT = ToDoubleNumber(0)
    End If
    Set rsSUM_SJ_DEBIT = Nothing
End Function

Function SUM_COB_DEBIT(xVOUCHERNO As String, xJType As String) As Double
    Dim rsSUM_COB_DEBIT                           As ADODB.Recordset
    Set rsSUM_COB_DEBIT = New ADODB.Recordset
    rsSUM_COB_DEBIT.Open "SELECT ROUND(SUM(INVOICEAMT),2) AS SUM_COB_DEBIT FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & " AND JTYPE = " & N2Str2Null(xJType) & " AND JDATE >= '" & dtFrom & "' AND JDATE <= '" & dtTo & "' AND STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsSUM_COB_DEBIT.EOF And Not rsSUM_COB_DEBIT.BOF Then
        SUM_COB_DEBIT = ToDoubleNumber(rsSUM_COB_DEBIT!SUM_COB_DEBIT)
        'SUM_COB_DEBIT = ToDoubleNumber(0)
        If SUM_COB_DEBIT > 0 Then

            G_TOTAL_DEBIT = ToDoubleNumber(Round((G_TOTAL_DEBIT + SUM_COB_DEBIT), 2))
            XX_BALANCE = ToDoubleNumber((Round((XX_BALANCE + NumericVal(rsSUM_COB_DEBIT!SUM_COB_DEBIT)), 2)))
        Else
            SUM_COB_DEBIT = ToDoubleNumber(Abs(rsSUM_COB_DEBIT!SUM_COB_DEBIT))
            G_TOTAL_CREDIT = ToDoubleNumber(Round((Abs(G_TOTAL_CREDIT + SUM_COB_DEBIT)), 2))
            XX_BALANCE = ToDoubleNumber((Round((XX_BALANCE - NumericVal(Abs(rsSUM_COB_DEBIT!SUM_COB_DEBIT))), 2)))
        End If
    Else
        SUM_COB_DEBIT = ToDoubleNumber(0)
    End If
    Set rsSUM_COB_DEBIT = Nothing
End Function

Sub FORWARDED_BALANCES()
    Dim rsFORWARDED_BALANCE                       As ADODB.Recordset
    Set rsFORWARDED_BALANCE = New ADODB.Recordset
    If cboAccountName.Text = "ALL ACCOUNTS" Then
        If optCustomer.Value = True Then
            rsFORWARDED_BALANCE.Open "SELECT DISTINCT HD.VOUCHERNO,HD.INVOICENO,HD.BANK,HD.ID,HD.JDATE,HD.JTYPE,HD.VOUCHERNO,HD.INVOICETYPE,HD.VENDORCODE,DET.ACCT_CODE, " & _
                                     "DET.INVOICENO AS GJ_INVOICE,DET.INVOICETYPE AS GJ_INVOICETYPE,DET.ADJ_VOUCHERNO AS ADJ_VOUCHERNO,DET.ADJ_JTYPE AS ADJ_JTYPE,DET.IS_OTHERS FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                                     "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                                     "WHERE (LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND CUSTOMERCODE = " & N2Str2Null(txtCode.Text) & " AND HD.JTYPE IN ('COB','SJ','CRJ') AND HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "') " & _
                                     "OR ((RIGHT(ENTITY,6) = " & N2Str2Null(txtCode.Text) & ") AND LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND DET.DEBIT <> 0) ORDER BY HD.JDATE ASC", gconDMIS, adOpenKeyset
        Else
            rsFORWARDED_BALANCE.Open "SELECT DISTINCT HD.VOUCHERNO,HD.INVOICENO,HD.ID,HD.JDATE,HD.JTYPE,HD.VOUCHERNO,HD.INVOICETYPE,HD.VENDORCODE,DET.ACCT_CODE, " & _
                                     "DET.INVOICENO AS GJ_INVOICE,DET.INVOICETYPE AS GJ_INVOICETYPE,DET.ADJ_VOUCHERNO AS ADJ_VOUCHERNO,DET.ADJ_JTYPE AS ADJ_JTYPE,DET.IS_OTHERS FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                                     "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                                     "WHERE (LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND HD.VENDORCODE = " & N2Str2Null(txtCode.Text) & " AND HD.JTYPE IN ('APJ','CDJ') AND HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "') " & _
                                     "OR ((RIGHT(ENTITY,6) = " & N2Str2Null(txtCode.Text) & ") AND LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND DET.DEBIT <> 0) ORDER BY HD.JDATE ASC", gconDMIS, adOpenKeyset
        End If
    Else
        If optCustomer.Value = True Then
            rsFORWARDED_BALANCE.Open "SELECT DISTINCT HD.VOUCHERNO,HD.INVOICENO,HD.BANK,HD.ID,HD.JDATE,HD.JTYPE,HD.VOUCHERNO,HD.INVOICETYPE,HD.VENDORCODE,DET.ACCT_CODE, " & _
                                     "DET.INVOICENO AS GJ_INVOICE,DET.INVOICETYPE AS GJ_INVOICETYPE,DET.ADJ_VOUCHERNO AS ADJ_VOUCHERNO,DET.ADJ_JTYPE AS ADJ_JTYPE,DET.IS_OTHERS FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                                     "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                                     "WHERE (DET.ACCT_CODE = " & N2Str2Null(Setacctcode(cboAccountName)) & " AND CUSTOMERCODE = " & N2Str2Null(txtCode.Text) & " AND HD.JTYPE IN ('COB','SJ','CRJ') AND HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "') " & _
                                     "OR ((RIGHT(ENTITY,6) = " & N2Str2Null(txtCode.Text) & ") AND LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND DET.DEBIT <> 0) ORDER BY HD.JDATE ASC", gconDMIS, adOpenKeyset
        Else
            rsFORWARDED_BALANCE.Open "DISTINCT HD.VOUCHERNO,HD.INVOICENO,HD.ID,HD.JDATE,HD.JTYPE,HD.VOUCHERNO,HD.INVOICETYPE,HD.VENDORCODE,DET.ACCT_CODE, " & _
                                     "DET.INVOICENO AS GJ_INVOICE,DET.INVOICETYPE AS GJ_INVOICETYPE,DET.ADJ_VOUCHERNO AS ADJ_VOUCHERNO,DET.ADJ_JTYPE AS ADJ_JTYPE,DET.IS_OTHERS FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                                     "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                                     "WHERE (LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND HD.VENDORCODE = " & N2Str2Null(txtCode.Text) & " AND HD.JTYPE IN ('APJ','CDJ') AND HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "') " & _
                                     "OR ((RIGHT(ENTITY,6) = " & N2Str2Null(txtCode.Text) & ") AND LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND DET.DEBIT <> 0) ORDER BY HD.JDATE ASC", gconDMIS, adOpenKeyset
        End If
    End If

    FORWARDED_BALANCE = 0

    If Not rsFORWARDED_BALANCE.EOF And Not rsFORWARDED_BALANCE.BOF Then
        Do While Not rsFORWARDED_BALANCE.EOF
            If Null2String(rsFORWARDED_BALANCE!jtype) = "COB" Then
                Call FWD_SUM_COB_DEBIT(Null2String(rsFORWARDED_BALANCE!VOUCHERNO), Null2String(rsFORWARDED_BALANCE!jtype))
                Call FWD_FIND_CRJ(Null2String(rsFORWARDED_BALANCE!INVOICENO), Null2String(rsFORWARDED_BALANCE!InvoiceType), RTrim(LTrim(txtCode.Text)), Null2String(rsFORWARDED_BALANCE!Acct_code))
                'Call FWD_FIND_GJ_ADJUSTMENT(Null2String(rsFORWARDED_BALANCE!INVOICENO), Null2String(rsFORWARDED_BALANCE!INVOICETYPE), txtCode.Text, Null2String(rsFORWARDED_BALANCE!ACCT_CODE))
                Call FWD_FIND_VOUCHERNO_ADJUSTMENT(Null2String(rsFORWARDED_BALANCE!VOUCHERNO), Null2String(rsFORWARDED_BALANCE!jtype), txtCode.Text, Null2String(rsFORWARDED_BALANCE!Acct_code))
            ElseIf Null2String(rsFORWARDED_BALANCE!jtype) = "SJ" Then
                Call FWD_SUM_SJ_DEBIT(Null2String(rsFORWARDED_BALANCE!VOUCHERNO), Null2String(rsFORWARDED_BALANCE!jtype), Null2String(rsFORWARDED_BALANCE!Acct_code))
                Call FWD_FIND_CRJ(Null2String(rsFORWARDED_BALANCE!INVOICENO), Null2String(rsFORWARDED_BALANCE!InvoiceType), txtCode.Text, Null2String(rsFORWARDED_BALANCE!Acct_code))
                'Call FWD_FIND_GJ_ADJUSTMENT(Null2String(rsFORWARDED_BALANCE!INVOICENO), Null2String(rsFORWARDED_BALANCE!INVOICETYPE), txtCode.Text, Null2String(rsFORWARDED_BALANCE!ACCT_CODE))
                Call FWD_FIND_VOUCHERNO_ADJUSTMENT(Null2String(rsFORWARDED_BALANCE!VOUCHERNO), Null2String(rsFORWARDED_BALANCE!jtype), txtCode.Text, Null2String(rsFORWARDED_BALANCE!Acct_code))
            ElseIf Null2String(rsFORWARDED_BALANCE!jtype) = "APJ" Then
                Call FWD_GET_APJ_AMOUNT(Null2String(rsFORWARDED_BALANCE!VOUCHERNO), Null2String(rsFORWARDED_BALANCE!jtype), Null2String(rsFORWARDED_BALANCE!Acct_code), RTrim(LTrim(txtCode.Text)))
                Call FWD_FIND_VOUCHERNO_ADJUSTMENT(Null2String(rsFORWARDED_BALANCE!VOUCHERNO), Null2String(rsFORWARDED_BALANCE!jtype), txtCode.Text, Null2String(rsFORWARDED_BALANCE!Acct_code))
            ElseIf Null2String(rsFORWARDED_BALANCE!jtype) = "CDJ" Then
                Call FWD_GET_CDJ_AMOUNT(Null2String(rsFORWARDED_BALANCE!VOUCHERNO), Null2String(rsFORWARDED_BALANCE!jtype), RTrim(LTrim(txtCode.Text)), Null2String(rsFORWARDED_BALANCE!Acct_code))
                Call FWD_FIND_VOUCHERNO_ADJUSTMENT(Null2String(rsFORWARDED_BALANCE!VOUCHERNO), Null2String(rsFORWARDED_BALANCE!jtype), txtCode.Text, Null2String(rsFORWARDED_BALANCE!Acct_code))
            ElseIf Null2String(rsFORWARDED_BALANCE!jtype) = "GJ" Then
                Call FWD_GET_GJ_AR(Null2String(rsFORWARDED_BALANCE!Acct_code), txtCode.Text, Null2String(rsFORWARDED_BALANCE!ADJ_VOUCHERNO), Null2String(rsFORWARDED_BALANCE!ADJ_JTYPE), rsFORWARDED_BALANCE!IS_OTHERS)
                Call FWD_FIND_CONTROL_NUMBER_ADJUSTMENT(Null2String(rsFORWARDED_BALANCE!ADJ_VOUCHERNO), txtCode.Text, Null2String(rsFORWARDED_BALANCE!Acct_code))
                Call FWD_ADJ_AGAINTS_NO_AR_ACCOUNT(Null2String(rsFORWARDED_BALANCE!VOUCHERNO), Null2String(rsFORWARDED_BALANCE!jtype), Null2String(rsFORWARDED_BALANCE!ADJ_VOUCHERNO), Null2String(rsFORWARDED_BALANCE!ADJ_JTYPE), txtCode.Text)
            End If
            rsFORWARDED_BALANCE.MoveNext
        Loop
    End If

    'THIS IS FOR FORWARDING OF BALANCE CRJ WITH NO SJ
    Call FWD_CRJ_NO_SJ

    Set rsFORWARDED_BALANCE = Nothing
End Sub

Function FWD_SUM_SJ_DEBIT(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
    Dim rsSUM_SJ_DEBIT                            As ADODB.Recordset
    Set rsSUM_SJ_DEBIT = New ADODB.Recordset
    rsSUM_SJ_DEBIT.Open "SELECT ROUND(SUM(DEBIT),2) AS SUM_DEBIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & " AND " & _
                        "JTYPE = " & N2Str2Null(xJType) & " AND ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND JDATE < '" & dtFrom & "' AND STATUS = 'P' AND DEBIT <> 0", gconDMIS, adOpenKeyset
    If Not rsSUM_SJ_DEBIT.EOF And Not rsSUM_SJ_DEBIT.BOF Then
        FWD_SUM_SJ_DEBIT = ToDoubleNumber(NumericVal(rsSUM_SJ_DEBIT!SUM_DEBIT))
        FORWARDED_BALANCE = ToDoubleNumber(Round((FORWARDED_BALANCE + NumericVal(FWD_SUM_SJ_DEBIT)), 2))
    Else
        FWD_SUM_SJ_DEBIT = ToDoubleNumber(0)
    End If
    Set rsSUM_SJ_DEBIT = Nothing
End Function

Function FWD_SUM_COB_DEBIT(xVOUCHERNO As String, xJType As String) As Double
    Dim rsSUM_COB_DEBIT                           As ADODB.Recordset
    Set rsSUM_COB_DEBIT = New ADODB.Recordset
    rsSUM_COB_DEBIT.Open "SELECT ROUND(SUM(INVOICEAMT),2) AS SUM_COB_DEBIT FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & " AND JTYPE = " & N2Str2Null(xJType) & " AND JDATE < '" & dtFrom & "' AND STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsSUM_COB_DEBIT.EOF And Not rsSUM_COB_DEBIT.BOF Then
        FWD_SUM_COB_DEBIT = ToDoubleNumber(NumericVal(rsSUM_COB_DEBIT!SUM_COB_DEBIT))
        FORWARDED_BALANCE = ToDoubleNumber((Round((FORWARDED_BALANCE + FWD_SUM_COB_DEBIT), 2)))
    Else
        FWD_SUM_COB_DEBIT = ToDoubleNumber(0)
    End If
    Set rsSUM_COB_DEBIT = Nothing
End Function

Sub FWD_FIND_CRJ(xINVOICENO As String, xINVOICETYPE As String, XCustomerCode As String, xACCT_CODE As String)
    Dim rsFIND_CRJ                                As ADODB.Recordset
    Set rsFIND_CRJ = New ADODB.Recordset
    rsFIND_CRJ.Open "SELECT DISTINCT CRJ.INVOICENO + '-' + CRJ.INVOICETYPE,HD.JTYPE,HD.VOUCHERNO,HD.ID,HD.JDATE,CRJ.INVOICENO,CRJ.INVOICETYPE,CRJ.INVOICEAMOUNT " & _
                    "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                    "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                    "INNER JOIN AMIS_CRJ_DETAIL CRJ ON HD.VOUCHERNO = CRJ.VOUCHERNO AND HD.JTYPE = CRJ.CR_TYPE " & _
                    "WHERE DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND HD.CUSTOMERCODE = " & N2Str2Null(XCustomerCode) & " AND CRJ.INVOICENO = " & N2Str2Null(xINVOICENO) & " AND CRJ.INVOICETYPE = " & N2Str2Null(xINVOICETYPE) & " AND HD.STATUS = 'P' AND HD.JDATE < '" & dtFrom & "'", gconDMIS, adOpenKeyset
    If Not rsFIND_CRJ.EOF And Not rsFIND_CRJ.BOF Then
        Do While Not rsFIND_CRJ.EOF
            FORWARDED_BALANCE = ToDoubleNumber(Round((FORWARDED_BALANCE - NumericVal(rsFIND_CRJ!invoiceamount)), 2))

            Call FWD_FIND_INVOICEDETAIL_ADJUSTMENT(Null2String(rsFIND_CRJ!INVOICENO), Null2String(rsFIND_CRJ!InvoiceType), Null2String(rsFIND_CRJ!Acct_code), txtCode.Text)
            Call FWD_FIND_VOUCHERNO_ADJUSTMENT(Null2String(rsFIND_CRJ!VOUCHERNO), Null2String(rsFIND_CRJ!jtype), txtCode.Text, Null2String(rsFIND_CRJ!Acct_code))
            rsFIND_CRJ.MoveNext
            Set REC = Nothing
        Loop
    Else
        'NO PAYMENTS
    End If
    Set rsFIND_CRJ = Nothing
End Sub

Sub FWD_FIND_GJ_ADJUSTMENT(xINVOICENO As String, xINVOICETYPE As String, XCustomerCode As String, xACCT_CODE As String)
    Dim rsFIND_GJ_ADJUSTMENT                      As ADODB.Recordset
    Set rsFIND_GJ_ADJUSTMENT = New ADODB.Recordset
    rsFIND_GJ_ADJUSTMENT.Open "SELECT HD.ID,HD.JDATE,HD.VOUCHERNO,HD.JTYPE,DET.INVOICENO,DET.INVOICETYPE,DET.CREDIT,DET.DEBIT " & _
                              "FROM AMIS_JOURNAL_DET DET INNER JOIN AMIS_JOURNAL_HD HD " & _
                              "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                              "WHERE DET.INVOICENO = " & N2Str2Null(xINVOICENO) & " AND DET.INVOICETYPE = " & N2Str2Null(xINVOICETYPE) & " " & _
                              "AND RIGHT(DET.ENTITY,6) = " & N2Str2Null(XCustomerCode) & " AND DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND HD.JDATE < '" & dtFrom.Value & "'", gconDMIS, adOpenKeyset

    If Not rsFIND_GJ_ADJUSTMENT.EOF And Not rsFIND_GJ_ADJUSTMENT.BOF Then
        Do While Not rsFIND_GJ_ADJUSTMENT.EOF
            If NumericVal(rsFIND_GJ_ADJUSTMENT!CREDIT) <> 0 Then
                FORWARDED_BALANCE = ToDoubleNumber(Round((FORWARDED_BALANCE - NumericVal(rsFIND_GJ_ADJUSTMENT!CREDIT)), 2))
            Else
                FORWARDED_BALANCE = ToDoubleNumber(Round((FORWARDED_BALANCE + NumericVal(rsFIND_GJ_ADJUSTMENT!DEBIT)), 2))
            End If
            rsFIND_GJ_ADJUSTMENT.MoveNext
        Loop
        Set REC = Nothing
    End If
    Set rsFIND_GJ_ADJUSTMENT = Nothing
End Sub

Sub FWD_CRJ_NO_SJ()
'DECSRIPTION: THIS IS FOR FORWARDED BALANCE
    Dim rsCRJ_NO_SJ                               As ADODB.Recordset
    Set rsCRJ_NO_SJ = New ADODB.Recordset
    rsCRJ_NO_SJ.Open "SELECT X.INV,X.CUST_CODE,X.JDATE,X.JTYPE,X.VOUCHERNO,X.INVOICETYPE,X.INVOICENO,X.INV_AMT,X.XID,X.ACCT_CODE FROM " & _
                     "( " & _
                     "SELECT DISTINCT CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS INV,HD.CUSTOMERCODE AS CUST_CODE,HD.JDATE AS JDATE,HD.JTYPE AS JTYPE, HD.VOUCHERNO AS VOUCHERNO, CRJ.INVOICENO AS INVOICENO, CRJ.INVOICETYPE AS INVOICETYPE,CRJ.INVOICEAMOUNT AS INV_AMT,HD.ID AS XID,DET.ACCT_CODE AS ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                     "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CRJ_DETAIL CRJ " & _
                     "ON HD.VOUCHERNO = CRJ.VOUCHERNO AND HD.JTYPE = CRJ.CR_TYPE " & _
                     "WHERE HD.JDATE < '" & dtFrom & "' AND HD.CUSTOMERCODE = '" & txtCode.Text & "' AND LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND HD.STATUS = 'P' AND HD.JTYPE = 'CRJ' " & _
                     ") X WHERE X.INV NOT IN " & _
                     "( " & _
                     "SELECT HD.INVOICETYPE + '-' + HD.INVOICENO FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                     "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                     "WHERE HD.CUSTOMERCODE = X.CUST_CODE AND HD.JDATE < '" & dtFrom & "' AND LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND HD.STATUS = 'P' AND HD.JTYPE IN ('SJ','COB') " & _
                     ")", gconDMIS, adOpenKeyset
    If Not rsCRJ_NO_SJ.EOF And Not rsCRJ_NO_SJ.BOF Then
        Do While Not rsCRJ_NO_SJ.EOF
            If Null2String(rsCRJ_NO_SJ!Acct_code) <> "11-02002-00" And COMPANY_CODE = "HGC" Then
                FORWARDED_BALANCE = ToDoubleNumber(Round((FORWARDED_BALANCE - NumericVal(rsCRJ_NO_SJ!INV_AMT)), 2))
                Call FWD_FIND_VOUCHERNO_ADJUSTMENT(Null2String(rsCRJ_NO_SJ!VOUCHERNO), Null2String(rsCRJ_NO_SJ!jtype), txtCode.Text, Null2String(rsCRJ_NO_SJ!Acct_code))
                Call FWD_FIND_INVOICEDETAIL_ADJUSTMENT(Null2String(rsCRJ_NO_SJ!INVOICENO), Null2String(rsCRJ_NO_SJ!InvoiceType), Null2String(rsCRJ_NO_SJ!Acct_code), txtCode.Text)
            End If
            rsCRJ_NO_SJ.MoveNext
        Loop
        Set REC = Nothing
    End If
    Set rsCRJ_NO_SJ = Nothing
End Sub

Sub ADVANCE_CRJ_PAYMENT()
'THIS IS FOR PAYMENT CRJ WITH SJ WHERE CRJ JDATE >= DATE FROM AND SJ JDATE < DATE FROM
    Dim rsADVANCE_CRJ_PAYMENT                     As ADODB.Recordset
    Set rsADVANCE_CRJ_PAYMENT = New ADODB.Recordset
    If RTrim(LTrim(cboAccountName.Text)) = "ALL ACCOUNTS" Then
        rsADVANCE_CRJ_PAYMENT.Open "SELECT X.INV,X.CUST_CODE,X.JDATE,X.JTYPE,X.VOUCHERNO,X.INVOICETYPE,X.INVOICENO,X.INV_AMT,X.XID FROM " & _
                                   "( " & _
                                   "SELECT DISTINCT CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS INV,HD.CUSTOMERCODE AS CUST_CODE,HD.JDATE AS JDATE, " & _
                                   "HD.JTYPE AS JTYPE, HD.VOUCHERNO AS VOUCHERNO, CRJ.INVOICENO AS INVOICENO, CRJ.INVOICETYPE AS INVOICETYPE, " & _
                                   "CRJ.INVOICEAMOUNT AS INV_AMT,HD.ID AS XID FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                                   "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CRJ_DETAIL CRJ ON HD.VOUCHERNO = CRJ.VOUCHERNO AND HD.JTYPE = CRJ.CR_TYPE " & _
                                   "WHERE HD.JDATE >= '" & dtFrom.Value & "' AND HD.CUSTOMERCODE = '" & txtCode.Text & "' AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03') AND HD.STATUS = 'P' AND HD.JTYPE = 'CRJ' " & _
                                   ") " & _
                                   "X WHERE X.INV IN( SELECT HD.INVOICETYPE + '-' + HD.INVOICENO FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                                   "WHERE HD.CUSTOMERCODE = X.CUST_CODE AND HD.JDATE < '" & dtFrom.Value & "' AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03') AND HD.STATUS = 'P' AND HD.JTYPE IN ('SJ','COB') )", gconDMIS, adOpenKeyset
    Else
        rsADVANCE_CRJ_PAYMENT.Open "SELECT X.INV,X.CUST_CODE,X.JDATE,X.JTYPE,X.VOUCHERNO,X.INVOICETYPE,X.INVOICENO,X.INV_AMT,X.XID FROM " & _
                                   "( " & _
                                   "SELECT DISTINCT CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS INV,HD.CUSTOMERCODE AS CUST_CODE,HD.JDATE AS JDATE, " & _
                                   "HD.JTYPE AS JTYPE, HD.VOUCHERNO AS VOUCHERNO, CRJ.INVOICENO AS INVOICENO, CRJ.INVOICETYPE AS INVOICETYPE, " & _
                                   "CRJ.INVOICEAMOUNT AS INV_AMT,HD.ID AS XID FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                                   "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CRJ_DETAIL CRJ ON HD.VOUCHERNO = CRJ.VOUCHERNO AND HD.JTYPE = CRJ.CR_TYPE " & _
                                   "WHERE HD.JDATE >= '" & dtFrom.Value & "' AND HD.CUSTOMERCODE = '" & txtCode.Text & "' AND DET.ACCT_CODE = '" & Setacctcode(cboAccountName.Text) & "' AND HD.STATUS = 'P' AND HD.JTYPE = 'CRJ' " & _
                                   ") " & _
                                   "X WHERE X.INV IN( SELECT HD.INVOICETYPE + '-' + HD.INVOICENO FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                                   "WHERE HD.CUSTOMERCODE = X.CUST_CODE AND HD.JDATE < '" & dtFrom.Value & "' AND DET.ACCT_CODE = '" & Setacctcode(cboAccountName.Text) & "' AND HD.STATUS = 'P' AND HD.JTYPE IN ('SJ','COB') )", gconDMIS, adOpenKeyset
    End If
    If Not rsADVANCE_CRJ_PAYMENT.EOF And Not rsADVANCE_CRJ_PAYMENT.BOF Then
        Do While Not rsADVANCE_CRJ_PAYMENT.EOF
            Set REC = rptRO.Records.Add

            REC.AddItem (Trim(rsADVANCE_CRJ_PAYMENT!JDate))
            REC.AddItem (Trim(rsADVANCE_CRJ_PAYMENT!jtype & "-" & rsADVANCE_CRJ_PAYMENT!VOUCHERNO))
            REC.AddItem (Trim(rsADVANCE_CRJ_PAYMENT!InvoiceType) & "-" & rsADVANCE_CRJ_PAYMENT!INVOICENO)
            REC.AddItem (Trim(ToDoubleNumber(0)))
            REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsADVANCE_CRJ_PAYMENT!INV_AMT))))

            XX_BALANCE = ToDoubleNumber(Round((XX_BALANCE - NumericVal(rsADVANCE_CRJ_PAYMENT!INV_AMT)), 2))
            G_TOTAL_CREDIT = Round((G_TOTAL_CREDIT + NumericVal(rsADVANCE_CRJ_PAYMENT!INV_AMT)), 2)

            REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))
            REC.AddItem (rsADVANCE_CRJ_PAYMENT!ID)
            REC.AddItem (rsADVANCE_CRJ_PAYMENT!jtype)
            rptRO.Populate
            Set REC = Nothing
            rsADVANCE_CRJ_PAYMENT.MoveNext
        Loop
    End If
    Set rsADVANCE_CRJ_PAYMENT = Nothing
End Sub

Sub ADVANCE_CRJ_NO_SJ()
'THIS IS FOR CRJ WITH NO SJ WHERE CRJ JDATE >= DATE FROM AND SJ JDATE IS < DATE FROM
    Dim rsADVANCE_CRJ_NO_SJ                       As ADODB.Recordset
    Set rsADVANCE_CRJ_NO_SJ = New ADODB.Recordset
    rsADVANCE_CRJ_NO_SJ.Open "SELECT X.INV,X.CUST_CODE,X.JDATE,X.JTYPE,X.VOUCHERNO,X.INVOICETYPE,X.INVOICENO,X.INV_AMT,X.XID FROM " & _
                             "( " & _
                             "SELECT DISTINCT CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS INV,HD.CUSTOMERCODE AS CUST_CODE,HD.JDATE AS JDATE, " & _
                             "HD.JTYPE AS JTYPE, HD.VOUCHERNO AS VOUCHERNO, CRJ.INVOICENO AS INVOICENO, CRJ.INVOICETYPE AS INVOICETYPE, " & _
                             "CRJ.INVOICEAMOUNT AS INV_AMT,HD.ID AS XID FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                             "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CRJ_DETAIL CRJ ON HD.VOUCHERNO = CRJ.VOUCHERNO AND HD.JTYPE = CRJ.CR_TYPE " & _
                             "WHERE HD.JDATE >= '" & dtFrom.Value & "' AND HD.CUSTOMERCODE = '" & txtCode.Text & "' AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03') AND HD.STATUS = 'P' AND HD.JTYPE = 'CRJ' " & _
                             ")X WHERE X.INV NOT IN " & _
                             "( SELECT HD.INVOICETYPE + '-' + HD.INVOICENO FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE  " & _
                             "WHERE HD.CUSTOMERCODE = X.CUST_CODE AND HD.JDATE < '" & dtFrom.Value & "' AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03') AND HD.STATUS = 'P' AND HD.JTYPE IN ('SJ','COB') )"", gconDMIS, adOpenKeyset"
    If Not rsADVANCE_CRJ_NO_SJ.EOF And Not rsADVANCE_CRJ_NO_SJ.BOF Then
        Do While Not rsADVANCE_CRJ_NO_SJ.EOF
            Set REC = rptRO.Records.Add

            REC.AddItem (Trim(rsADVANCE_CRJ_NO_SJ!JDate))
            REC.AddItem (Trim(rsADVANCE_CRJ_NO_SJ!jtype & "-" & rsADVANCE_CRJ_NO_SJ!VOUCHERNO))
            REC.AddItem (Trim(rsADVANCE_CRJ_NO_SJ!InvoiceType) & "-" & rsADVANCE_CRJ_NO_SJ!INVOICENO)
            REC.AddItem (Trim(ToDoubleNumber(0)))
            REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsADVANCE_CRJ_NO_SJ!INV_AMT))))

            XX_BALANCE = ToDoubleNumber(Round((XX_BALANCE - NumericVal(rsADVANCE_CRJ_NO_SJ!INV_AMT)), 2))
            G_TOTAL_CREDIT = Round((G_TOTAL_CREDIT + NumericVal(rsADVANCE_CRJ_NO_SJ!INV_AMT)), 2)

            REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))
            REC.AddItem (rsADVANCE_CRJ_NO_SJ!ID)
            REC.AddItem (rsADVANCE_CRJ_NO_SJ!jtype)
            rptRO.Populate

            rsADVANCE_CRJ_NO_SJ.MoveNext
        Loop
    End If
    Set rsADVANCE_CRJ_NO_SJ = Nothing
End Sub

Sub FIND_VOUCHERNO_ADJUSTMENT(xVOUCHERNO As String, xJType As String, xEntity As String, xACCT_CODE As String)
'DESCRIPTION:THIS IS AN ADJUSTMENT WHERE IN VOUCHERNO AN JOURNAL TYPE WAS USE AS A REFERENCE IN AN ADJUSTMENT
    Dim rsAdjustment                              As ADODB.Recordset
    Set rsAdjustment = New ADODB.Recordset
    rsAdjustment.Open "SELECT HD.JDATE,HD.JTYPE,HD.VOUCHERNO,DET.ADJ_VOUCHERNO,DET.ADJ_JTYPE,DET.CREDIT,DET.DEBIT,HD.ID " & _
                      "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                      "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                      "WHERE DET.ADJ_VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & " AND DET.ADJ_JTYPE = " & N2Str2Null(xJType) & " AND DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND RIGHT(DET.ENTITY,6) = " & N2Str2Null(xEntity) & " AND DET.STATUS = 'P' and HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "' AND DET.INVOICENO IS NULL AND DET.INVOICETYPE IS NULL", gconDMIS, adOpenKeyset
    If Not rsAdjustment.EOF And Not rsAdjustment.BOF Then
        Do While Not rsAdjustment.EOF
            Set REC = rptRO.Records.Add

            REC.AddItem (Trim(rsAdjustment!JDate))
            REC.AddItem (Trim(rsAdjustment!jtype & "-" & rsAdjustment!VOUCHERNO))
            REC.AddItem (Trim(rsAdjustment!ADJ_JTYPE) & "-" & rsAdjustment!ADJ_VOUCHERNO)

            If NumericVal(rsAdjustment!DEBIT) <> 0 Then
                REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsAdjustment!DEBIT))))
                REC.AddItem (Trim(ToDoubleNumber(0)))

                XX_BALANCE = ToDoubleNumber(Round((XX_BALANCE + NumericVal(rsAdjustment!DEBIT)), 2))
                G_TOTAL_DEBIT = Round((G_TOTAL_DEBIT + NumericVal(rsAdjustment!DEBIT)), 2)

            ElseIf NumericVal(rsAdjustment!CREDIT) <> 0 Then
                REC.AddItem (Trim(ToDoubleNumber(0)))
                REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsAdjustment!CREDIT))))

                XX_BALANCE = ToDoubleNumber(Round((XX_BALANCE - NumericVal(rsAdjustment!CREDIT)), 2))
                G_TOTAL_CREDIT = Round((G_TOTAL_CREDIT + NumericVal(rsAdjustment!CREDIT)), 2)
            End If

            REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))
            REC.AddItem (rsAdjustment!ID)
            REC.AddItem (rsAdjustment!jtype)

            rptRO.Populate
            Set REC = Nothing
            rsAdjustment.MoveNext
        Loop
    End If
    Set rsAdjustment = Nothing
End Sub
Sub FWD_FIND_VOUCHERNO_ADJUSTMENT(xVOUCHERNO As String, xJType As String, xEntity As String, xACCT_CODE As String)
'DESCRIPTION:THIS IS AN ADJUSTMENT WHERE IN VOUCHERNO AN JOURNAL TYPE WAS USE AS A REFERENCE IN AN ADJUSTMENT
    Dim rsAdjustment                              As ADODB.Recordset
    Set rsAdjustment = New ADODB.Recordset
    rsAdjustment.Open "SELECT HD.JDATE,HD.JTYPE,HD.VOUCHERNO,DET.ADJ_VOUCHERNO,DET.ADJ_JTYPE,DET.CREDIT,DET.DEBIT,HD.ID " & _
                      "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                      "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                      "WHERE DET.ADJ_VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & " AND DET.ADJ_JTYPE = " & N2Str2Null(xJType) & " AND DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND RIGHT(DET.ENTITY,6) = " & N2Str2Null(xEntity) & " AND DET.STATUS = 'P' and HD.JDATE < '" & dtFrom & "' AND DET.INVOICENO IS NULL AND DET.INVOICETYPE IS NULL", gconDMIS, adOpenKeyset
    If Not rsAdjustment.EOF And Not rsAdjustment.BOF Then
        Do While Not rsAdjustment.EOF
            If NumericVal(rsAdjustment!DEBIT) <> 0 Then
                FORWARDED_BALANCE = ToDoubleNumber(Round((FORWARDED_BALANCE + NumericVal(rsAdjustment!DEBIT)), 2))
            ElseIf NumericVal(rsAdjustment!CREDIT) <> 0 Then
                FORWARDED_BALANCE = ToDoubleNumber(Round((FORWARDED_BALANCE - NumericVal(rsAdjustment!CREDIT)), 2))
            End If
            rsAdjustment.MoveNext
        Loop
    End If
    Set rsAdjustment = Nothing
End Sub

'Sub FWD_FIND_VOUCHERNO_ADJUSTMENT(xVOUCHERNO As String, xJTYPE As String, xEntity As String, xACCT_CODE As String)
'    'DESCRIPTION:THIS IS AN ADJUSTMENT WHERE IN VOUCHERNO AN JOURNAL TYPE WAS USE AS A REFERENCE IN AN ADJUSTMENT
'    Dim rsAdjustment As ADODB.Recordset
'        Set rsAdjustment = New ADODB.Recordset
'            rsAdjustment.Open "SELECT HD.JDATE,HD.JTYPE,HD.VOUCHERNO,DET.ADJ_VOUCHERNO,DET.ADJ_JTYPE,DET.CREDIT,DET.DEBIT,HD.ID " & _
             '                              "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
             '                              "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
             '                              "WHERE DET.ADJ_VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & " AND DET.ADJ_JTYPE = " & N2Str2Null(xJTYPE) & " AND DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND RIGHT(DET.ENTITY,6) = " & N2Str2Null(xEntity) & " AND DET.STATUS = 'P' and HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "' AND DET.INVOICENO IS NULL AND DET.INVOICETYPE IS NULL", gconDMIS, adOpenKeyset
'            If Not rsAdjustment.EOF And Not rsAdjustment.BOF Then
'                Do While Not rsAdjustment.EOF
'                   If NumericVal(rsAdjustment!DEBIT) <> 0 Then
'                       FORWARDED_BALANCE = ToDoubleNumber(Round((FORWARDED_BALANCE + NumericVal(rsAdjustment!DEBIT)), 2))
'                   ElseIf NumericVal(rsAdjustment!CREDIT) <> 0 Then
'                       FORWARDED_BALANCE = ToDoubleNumber(Round((FORWARDED_BALANCE - NumericVal(rsAdjustment!CREDIT)), 2))
'                   End If
'                   rsAdjustment.MoveNext
'                Loop
'            End If
'    Set rsAdjustment = Nothing
'End Sub

Sub FIND_CONTROL_NUMBER_ADJUSTMENT(xCON_NUMBER As String, xEntity As String, xACCT_CODE As String)
'DESCRIPTION: THIS IS AN ADJUSTMENT WHERE IN CONTROL NUMBER GENERATED IN GJ OTHERS WAS USE AS A REFERENCE IN AN ADJUSTMENT
    Dim rsCON_NUM                                 As ADODB.Recordset
    Set rsCON_NUM = New ADODB.Recordset
    rsCON_NUM.Open "SELECT HD.ID, HD.JDATE, HD.JTYPE,HD.VOUCHERNO,DET.DEBIT,DET.CREDIT,DET.ADJ_JTYPE,DET.ADJ_VOUCHERNO " & _
                   "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                   "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                   "WHERE DET.ADJ_VOUCHERNO = " & N2Str2Null(xCON_NUMBER) & " AND DET.ADJ_JTYPE = 'OTH' AND DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND RIGHT(DET.ENTITY,6) = " & N2Str2Null(xEntity) & " AND DET.STATUS = 'P' and HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "' AND DET.CREDIT <> 0", gconDMIS, adOpenKeyset
    If Not rsCON_NUM.EOF And Not rsCON_NUM.BOF Then
        Do While Not rsCON_NUM.EOF
            Set REC = rptRO.Records.Add
            REC.AddItem (Trim(rsCON_NUM!JDate))
            REC.AddItem (Trim(rsCON_NUM!jtype & "-" & rsCON_NUM!VOUCHERNO))
            REC.AddItem (Trim(rsCON_NUM!ADJ_JTYPE) & "-" & rsCON_NUM!ADJ_VOUCHERNO)

            If NumericVal(rsCON_NUM!DEBIT) <> 0 Then
                REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsCON_NUM!DEBIT))))
                REC.AddItem (Trim(ToDoubleNumber(0)))

                XX_BALANCE = ToDoubleNumber(Round((XX_BALANCE + NumericVal(rsCON_NUM!DEBIT)), 2))
                G_TOTAL_DEBIT = Round((G_TOTAL_DEBIT + NumericVal(rsCON_NUM!DEBIT)), 2)

            ElseIf NumericVal(rsCON_NUM!CREDIT) <> 0 Then
                REC.AddItem (Trim(ToDoubleNumber(0)))
                REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsCON_NUM!CREDIT))))

                XX_BALANCE = ToDoubleNumber(Round((XX_BALANCE - NumericVal(rsCON_NUM!CREDIT)), 2))
                G_TOTAL_CREDIT = Round((G_TOTAL_CREDIT + NumericVal(rsCON_NUM!CREDIT)), 2)
            End If

            REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))
            REC.AddItem (rsCON_NUM!ID)
            REC.AddItem (rsCON_NUM!jtype)

            rptRO.Populate
            Set REC = Nothing
            rsCON_NUM.MoveNext
        Loop
    End If
    Set rsCON_NUM = Nothing
End Sub

Sub FWD_FIND_CONTROL_NUMBER_ADJUSTMENT(xCON_NUMBER As String, xEntity As String, xACCT_CODE As String)
'DESCRIPTION: THIS IS AN ADJUSTMENT WHERE IN CONTROL NUMBER GENERATED IN GJ OTHERS WAS USE AS A REFERENCE IN AN ADJUSTMENT
    Dim rsCON_NUM                                 As ADODB.Recordset
    Set rsCON_NUM = New ADODB.Recordset
    rsCON_NUM.Open "SELECT HD.ID, HD.JDATE, HD.JTYPE,HD.VOUCHERNO,DET.DEBIT,DET.CREDIT,DET.ADJ_JTYPE,DET.ADJ_VOUCHERNO " & _
                   "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                   "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                   "WHERE DET.ADJ_VOUCHERNO = " & N2Str2Null(xCON_NUMBER) & " AND DET.ADJ_JTYPE = 'OTH' AND DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND RIGHT(DET.ENTITY,6) = " & N2Str2Null(xEntity) & " AND DET.STATUS = 'P' and HD.JDATE < '" & dtFrom & "' AND DET.CREDIT <> 0 ", gconDMIS, adOpenKeyset
    If Not rsCON_NUM.EOF And Not rsCON_NUM.BOF Then
        Do While Not rsCON_NUM.EOF
            If NumericVal(rsCON_NUM!DEBIT) <> 0 Then
                FORWARDED_BALANCE = ToDoubleNumber(Round((FORWARDED_BALANCE + NumericVal(rsCON_NUM!DEBIT)), 2))
            ElseIf NumericVal(rsCON_NUM!CREDIT) <> 0 Then
                FORWARDED_BALANCE = ToDoubleNumber(Round((FORWARDED_BALANCE - NumericVal(rsCON_NUM!CREDIT)), 2))
            End If
            rsCON_NUM.MoveNext
        Loop
    End If
    Set rsCON_NUM = Nothing
End Sub

Sub FIND_INVOICEDETAIL_ADJUSTMENT(xINVOICENO As String, InvoiceType As String, xACCT_CODE As String, xEntity As String)
'DESCRIPTION: THIS IS AN ADJUSMENT WHERE IN INVOICENO AND INVOICE TYPE WAS USE IN ADJUSTMENT
    Dim rsINV_DETAIL                              As ADODB.Recordset
    Set rsINV_DETAIL = New ADODB.Recordset
    rsINV_DETAIL.Open "SELECT HD.ID, HD.JDATE, HD.JTYPE,HD.VOUCHERNO,DET.DEBIT,DET.CREDIT,DET.ADJ_JTYPE " & _
                      "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                      "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                      "WHERE RIGHT(DET.ENTITY,6) = " & N2Str2Null(xEntity) & " AND DET.STATUS = 'P' AND DET.INVOICENO = " & N2Str2Null(xINVOICENO) & " AND DET.INVOICETYPE = " & N2Str2Null(xINVOICENO) & " and HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "'", gconDMIS, adOpenKeyset
    If Not rsINV_DETAIL.EOF And Not rsINV_DETAIL.BOF Then
        Do While Not rsINV_DETAIL.EOF
            Set REC = rptRO.Records.Add
            REC.AddItem (Trim(rsINV_DETAIL!JDate))
            REC.AddItem (Trim(rsINV_DETAIL!jtype & "-" & rsINV_DETAIL!VOUCHERNO))
            REC.AddItem (Trim(rsINV_DETAIL!ADJ_JTYPE) & "-" & rsINV_DETAIL!ADJ_VOUCHERNO)

            If NumericVal(rsINV_DETAIL!DEBIT) <> 0 Then
                REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsINV_DETAIL!DEBIT))))
                REC.AddItem (Trim(ToDoubleNumber(0)))

                XX_BALANCE = ToDoubleNumber(Round((XX_BALANCE + NumericVal(rsINV_DETAIL!DEBIT)), 2))
                G_TOTAL_DEBIT = Round((G_TOTAL_DEBIT + NumericVal(rsINV_DETAIL!DEBIT)), 2)

            ElseIf NumericVal(rsINV_DETAIL!CREDIT) <> 0 Then
                REC.AddItem (Trim(ToDoubleNumber(0)))
                REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsINV_DETAIL!CREDIT))))

                XX_BALANCE = ToDoubleNumber(Round((XX_BALANCE - NumericVal(rsINV_DETAIL!CREDIT)), 2))
                G_TOTAL_CREDIT = Round((G_TOTAL_CREDIT + NumericVal(rsINV_DETAIL!CREDIT)), 2)
            End If

            REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))
            REC.AddItem (rsINV_DETAIL!ID)
            REC.AddItem (rsINV_DETAIL!jtype)
            rptRO.Populate
            Set REC = Nothing
            rsINV_DETAIL.MoveNext
        Loop
    End If
    Set rsINV_DETAIL = Nothing
End Sub

Sub FWD_FIND_INVOICEDETAIL_ADJUSTMENT(xINVOICENO As String, InvoiceType As String, xACCT_CODE As String, xEntity As String)
'DESCRIPTION: THIS IS AN ADJUSMENT WHERE IN INVOICENO AND INVOICE TYPE WAS USE IN ADJUSTMENT
    Dim rsINV_DETAIL                              As ADODB.Recordset
    Set rsINV_DETAIL = New ADODB.Recordset
    rsINV_DETAIL.Open "SELECT HD.ID, HD.JDATE, HD.JTYPE,HD.VOUCHERNO,DET.DEBIT,DET.CREDIT,DET.ADJ_JTYPE " & _
                      "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                      "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                      "WHERE RIGHT(DET.ENTITY,6) = " & N2Str2Null(xEntity) & " AND DET.STATUS = 'P' AND DET.INVOICENO = " & N2Str2Null(xINVOICENO) & " AND DET.INVOICETYPE = " & N2Str2Null(xINVOICENO) & " and HD.JDATE < '" & dtFrom & "'", gconDMIS, adOpenKeyset
    If Not rsINV_DETAIL.EOF And Not rsINV_DETAIL.BOF Then
        Do While Not rsINV_DETAIL.EOF
            If NumericVal(rsINV_DETAIL!DEBIT) <> 0 Then
                FORWARDED_BALANCE = ToDoubleNumber(Round((FORWARDED_BALANCE + NumericVal(rsINV_DETAIL!DEBIT)), 2))
            ElseIf NumericVal(rsINV_DETAIL!CREDIT) <> 0 Then
                FORWARDED_BALANCE = ToDoubleNumber(Round((FORWARDED_BALANCE - NumericVal(rsINV_DETAIL!CREDIT)), 2))
            End If
            rsINV_DETAIL.MoveNext
        Loop
    End If
    Set rsINV_DETAIL = Nothing
End Sub

Sub GJ_INVOICE_TO_INVOICE(xINVOICENO As String, xINVOICETYPE As String, xACCT_CODE As String, xCUST_CODE As String, xADJ_VOUCHERNO As String, xADJ_JTYPE As String)
'DESCRIPTION: THIS IS TO FIND THE ADJUSTMENT FOR
    Dim rsINVOICE                                 As ADODB.Recordset
    Set rsINVOICE = New ADODB.Recordset

    If xINVOICENO <> "" And xINVOICETYPE <> "" And xADJ_VOUCHERNO <> "" And xADJ_JTYPE <> "" Then
        rsINVOICE.Open "SELECT HD.JDATE,HD.VOUCHERNO,HD.JTYPE,DET.DEBIT,DET.CREDIT,HD.ID FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                       "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE WHERE DET.INVOICENO = " & N2Str2Null(xINVOICENO) & " AND DET.INVOICETYPE = " & N2Str2Null(xINVOICETYPE) & " " & _
                       "AND ADJ_VOUCHERNO = " & N2Str2Null(xADJ_VOUCHERNO) & " AND ADJ_JTYPE = " & N2Str2Null(xADJ_JTYPE) & " AND HD.STATUS = 'P' AND  HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "'", gconDMIS, adOpenKeyset
    ElseIf xINVOICENO = "" And xINVOICETYPE = "" And xADJ_VOUCHERNO <> "" And xADJ_JTYPE <> "" Then
        rsINVOICE.Open "SELECT HD.JDATE,HD.VOUCHERNO,HD.JTYPE,DET.DEBIT,DET.CREDIT,HD.ID FROM AMIS_JOURNAL_HD HD INNER JOIND AMIS_JOURNAL_DET DET " & _
                       "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE WHERE DET.INVOICENO = " & N2Str2Null(xINVOICENO) & " AND DET.INVOICETYPE = " & N2Str2Null(xINVOICETYPE) & " " & _
                       "AND ADJ_VOUCHERNO = " & N2Str2Null(xADJ_VOUCHERNO) & " AND ADJ_JTYPE = " & N2Str2Null(xADJ_JTYPE) & " AND HD.STATUS = 'P' AND  HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "'", gconDMIS, adOpenKeyset
    End If

    If Not rsINVOICE.EOF And Not rsINVOICE.BOF Then
        Do While Not rsINVOICE.EOF
            Set REC = rptRO.Records.Add
            REC.AddItem (Trim(rsINVOICE!JDate))
            REC.AddItem (Trim(rsINVOICE!jtype & "-" & rsINVOICE!VOUCHERNO))
            REC.AddItem (Trim(rsINVOICE!INVOICENO) & "-" & rsINVOICE!InvoiceType)

            If NumericVal(rsINVOICE!DEBIT) <> 0 Then
                REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsINVOICE!DEBIT))))
                REC.AddItem (Trim(ToDoubleNumber(0)))

                XX_BALANCE = ToDoubleNumber(Round((XX_BALANCE + NumericVal(rsINVOICE!DEBIT)), 2))
                G_TOTAL_DEBIT = Round((G_TOTAL_DEBIT + NumericVal(rsINVOICE!DEBIT)), 2)

            ElseIf NumericVal(rsINVOICE!CREDIT) <> 0 Then
                REC.AddItem (Trim(ToDoubleNumber(0)))
                REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsINVOICE!CREDIT))))

                XX_BALANCE = ToDoubleNumber(Round((XX_BALANCE - NumericVal(rsINVOICE!CREDIT)), 2))
                G_TOTAL_CREDIT = Round((G_TOTAL_CREDIT + NumericVal(rsINVOICE!CREDIT)), 2)
            End If

            REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))
            REC.AddItem (rsINVOICE!ID)
            REC.AddItem (rsINVOICE!jtype)
            rptRO.Populate
            Set REC = Nothing
            rsINVOICE.MoveNext
        Loop
    End If
    Set rsINVOICE = Nothing
End Sub

Sub ADJ_AGAINTS_NO_AR_ACCOUNT(xVOUCHERNO As String, xJType As String, xADJ_VOUCHERNO As String, xADJ_JTYPE As String, xCUST_CODE As String)
'DESCRIPTION: THIS IS FOR GETTING THE AR IN GJ WHICH IS SALES JOURNAL HAS NO AR ACCOUNT SCHEDULE EX. IT'S A CASH SALES CLEARING TRANSACTON
    Dim rsNO_AR                                   As ADODB.Recordset
    Dim rsGJ                                      As ADODB.Recordset

    Set rsNO_AR = New ADODB.Recordset
    rsNO_AR.Open "SELECT * FROM  AMIS_JOURNAL_DET  " & _
                 "WHERE VOUCHERNO = " & N2Str2Null(xADJ_VOUCHERNO) & " AND JTYPE = " & N2Str2Null(xADJ_JTYPE) & " AND " & _
                 "STATUS = 'P' AND LEFT(ACCT_CODE,5) IN ('11-02','11-03')", gconDMIS, adOpenKeyset

    If Not rsNO_AR.EOF And Not rsNO_AR.BOF Then
        'IT HAS AN SCHEDULE ACCT CODE
    Else
        Set rsGJ = New ADODB.Recordset

        rsGJ.Open "SELECT HD.ID,HD.JDATE,HD.JTYPE,HD.VOUCHERNO,DET.ADJ_VOUCHERNO,DET.ADJ_JTYPE,DET.ACCT_CODE,DET.DEBIT FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                  "ON HD.VOUCHERNO =DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE WHERE HD.VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & " AND HD.JTYPE = " & N2Str2Null(xJType) & " AND " & _
                  "RIGHT(DET.ENTITY,6) = " & N2Str2Null(xCUST_CODE) & " AND HD.STATUS = 'P' AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03') AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "'", gconDMIS, adOpenKeyset
        If Not rsGJ.EOF And Not rsGJ.BOF Then
            Set REC = rptRO.Records.Add
            REC.AddItem (Trim(rsGJ!JDate))
            REC.AddItem (Trim(rsGJ!jtype & "-" & rsGJ!VOUCHERNO))
            REC.AddItem (Trim(rsGJ!ADJ_JTYPE) & "-" & rsGJ!ADJ_VOUCHERNO)
            REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsGJ!DEBIT))))
            REC.AddItem (Trim(ToDoubleNumber(0)))
            XX_BALANCE = ToDoubleNumber(Round((XX_BALANCE + NumericVal(rsGJ!DEBIT)), 2))
            G_TOTAL_DEBIT = Round((G_TOTAL_DEBIT + NumericVal(rsGJ!DEBIT)), 2)
            REC.AddItem (Trim(ToDoubleNumber(XX_BALANCE)))
            REC.AddItem (rsGJ!ID)
            REC.AddItem (rsGJ!jtype)
            rptRO.Populate

            Call FIND_VOUCHERNO_ADJUSTMENT(Null2String(rsGJ!VOUCHERNO), Null2String(rsGJ!jtype), txtCode.Text, Null2String(rsGJ!Acct_code))
            Set REC = Nothing
        End If
        Set rsGJ = Nothing
    End If
    Set rsNO_AR = Nothing
End Sub

Sub FWD_ADJ_AGAINTS_NO_AR_ACCOUNT(xVOUCHERNO As String, xJType As String, xADJ_VOUCHERNO As String, xADJ_JTYPE As String, xCUST_CODE As String)
'DESCRIPTION: THIS IS FORWARDING BALANCE FOR ADJUSTMENT FOR VOUCHER WHICH HAS NO AR ACCOUNT CODE
    Dim rsNO_AR                                   As ADODB.Recordset
    Dim rsGJ                                      As ADODB.Recordset

    Set rsNO_AR = New ADODB.Recordset
    rsNO_AR.Open "SELECT * FROM  AMIS_JOURNAL_DET  " & _
                 "WHERE VOUCHERNO = " & N2Str2Null(xADJ_VOUCHERNO) & " AND JTYPE = " & N2Str2Null(xADJ_JTYPE) & " AND " & _
                 "STATUS = 'P' AND LEFT(ACCT_CODE,5) IN ('11-02','11-03')", gconDMIS, adOpenKeyset

    If Not rsNO_AR.EOF And Not rsNO_AR.BOF Then
        'IT HAS AN SCHEDULE ACCT CODE
    Else
        Set rsGJ = New ADODB.Recordset
        rsGJ.Open "SELECT HD.ID,HD.JDATE,HD.JTYPE,HD.VOUCHERNO,DET.ADJ_VOUCHERNO,DET.ADJ_JTYPE,DET.ACCT_CODE,DET.DEBIT FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                  "ON HD.VOUCHERNO =DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE WHERE HD.VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & " AND HD.JTYPE = " & N2Str2Null(xJType) & " AND " & _
                  "RIGHT(DET.ENTITY,6) = " & N2Str2Null(xCUST_CODE) & " AND HD.STATUS = 'P' AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03') AND HD.JDATE < '" & dtFrom & "'", gconDMIS, adOpenKeyset
        If Not rsGJ.EOF And Not rsGJ.BOF Then
            FORWARDED_BALANCE = ToDoubleNumber(Round((FORWARDED_BALANCE + NumericVal(rsGJ!DEBIT)), 2))
            Call FWD_FIND_VOUCHERNO_ADJUSTMENT(Null2String(rsGJ!VOUCHERNO), Null2String(rsGJ!jtype), txtCode.Text, Null2String(rsGJ!Acct_code))
        End If
        Set rsGJ = Nothing
    End If
    Set rsNO_AR = Nothing
End Sub




