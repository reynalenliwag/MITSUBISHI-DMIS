VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmAMISLEDGERCustomers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers A/R Ledger"
   ClientHeight    =   8430
   ClientLeft      =   315
   ClientTop       =   435
   ClientWidth     =   11850
   ForeColor       =   &H00FFFFFF&
   Icon            =   "CustomersLedger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   11850
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
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
      Left            =   11280
      TabIndex        =   35
      Top             =   60
      Width           =   525
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   345
      Left            =   8190
      TabIndex        =   29
      Top             =   60
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      _Version        =   393216
      Format          =   51904513
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
      ItemData        =   "CustomersLedger.frx":030A
      Left            =   1860
      List            =   "CustomersLedger.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   60
      Width           =   5685
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   2655
      TabIndex        =   3
      Top             =   480
      Width           =   9135
      Begin VB.TextBox txtCode 
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
         Left            =   1650
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
         Left            =   1650
         MaxLength       =   35
         TabIndex        =   14
         Top             =   570
         Width           =   7320
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
            Size            =   9.75
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
            Size            =   9.75
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
         Height          =   255
         Left            =   90
         TabIndex        =   41
         Top             =   450
         Width           =   2415
      End
      Begin VB.OptionButton optCustomer 
         Caption         =   "Customer"
         Height          =   345
         Left            =   90
         TabIndex        =   40
         Top             =   150
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
         Left            =   120
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
         MouseIcon       =   "CustomersLedger.frx":030E
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CUSTOMER NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   882
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
      Left            =   10965
      MouseIcon       =   "CustomersLedger.frx":0470
      MousePointer    =   99  'Custom
      Picture         =   "CustomersLedger.frx":05C2
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
      Left            =   10275
      MouseIcon       =   "CustomersLedger.frx":0928
      MousePointer    =   99  'Custom
      Picture         =   "CustomersLedger.frx":0A7A
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
      Left            =   9585
      MouseIcon       =   "CustomersLedger.frx":0DE0
      MousePointer    =   99  'Custom
      Picture         =   "CustomersLedger.frx":0F32
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
      Left            =   8895
      MouseIcon       =   "CustomersLedger.frx":122C
      MousePointer    =   99  'Custom
      Picture         =   "CustomersLedger.frx":137E
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
      Left            =   8205
      MouseIcon       =   "CustomersLedger.frx":16D6
      MousePointer    =   99  'Custom
      Picture         =   "CustomersLedger.frx":1828
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
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   90
         ScaleHeight     =   525
         ScaleWidth      =   8925
         TabIndex        =   22
         Top             =   5340
         Width           =   8925
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
         MouseIcon       =   "CustomersLedger.frx":1B87
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
         ItemData        =   "CustomersLedger.frx":1EA1
         Left            =   6360
         List            =   "CustomersLedger.frx":1EAE
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   630
         Width           =   1545
      End
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   345
      Left            =   9960
      TabIndex        =   30
      Top             =   60
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      _Version        =   393216
      Format          =   51904513
      CurrentDate     =   39765
   End
   Begin VB.Label Label8 
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9540
      TabIndex        =   34
      Top             =   60
      Width           =   675
   End
   Begin VB.Label Label7 
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7590
      TabIndex        =   33
      Top             =   90
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
      Top             =   7470
      Width           =   4365
   End
   Begin VB.Label Label 
      Caption         =   "Account Name :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
Attribute VB_Name = "frmAMISLEDGERCustomers"
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

Dim rsCUSTOMER_OPENING                            As ADODB.Recordset
'Dim xBALANCE        As Double

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
            'rsCUSTOMER.Open "SELECT DISTINCT ALL_VENDOR.NameofVendor as VENDORNAME,ALL_VENDOR.ID as ID,ALL_VENDOR.CODE AS VENCODE " & _
             "FROM dbo.AMIS_Journal_HD HD INNER JOIN dbo.AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo AND HD.JType = DET.JType " & _
             "RIGHT OUTER JOIN dbo.ALL_VENDOR ALL_VENDOR ON ((HD.VENDORCODE = ALL_VENDOR.CODE) OR  (RIGHT(DET.ENTITY,6) = ALL_VENDOR.CODE))  WHERE  ALL_VENDOR.CODE IS NOT NULL  and " & _
             "((LEFT(DET.Acct_Code, 5) = '11-02') OR (LEFT(DET.Acct_Code, 5) = '11-03')) ORDER BY ALL_VENDOR.NameofVendor", gconDMIS, adOpenKeyset
            rsCustomer.Open "SELECT DISTINCT ALL_VENDOR.NameofVendor as VENDORNAME,ALL_VENDOR.ID as ID,ALL_VENDOR.CODE AS VENCODE " & _
                            "FROM dbo.AMIS_Journal_HD HD INNER JOIN dbo.AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo AND HD.JType = DET.JType " & _
                            "RIGHT OUTER JOIN dbo.ALL_VENDOR ALL_VENDOR ON HD.VENDORCODE = ALL_VENDOR.CODE", gconDMIS, adOpenKeyset
        End If
    Else
        If optCustomer.Value = True Then
            'DESCRIPTION: THIS IS FOR CUSTOMER WITH AR ACCOUNT GROUP BY ACCOUNT CODE
            'rsCUSTOMER.Open "SELECT DISTINCT dbo.ALL_Customer.ACCTNAME as CUSTNAME,dbo.ALL_Customer.ID,dbo.ALL_Customer.CUSCDE AS CUSTCODE FROM dbo.AMIS_Journal_HD left outer JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType INNER JOIN dbo.ALL_Customer ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer.CUSCDE WHERE ((LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') or (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03')) AND dbo.AMIS_Journal_Det.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "' ORDER BY dbo.ALL_Customer.ACCTNAME", gconDMIS, adOpenKeyset
            rsCustomer.Open "SELECT DISTINCT CUST.ACCTNAME as CUSTNAME,CUST.ID,CUST.CUSCDE AS CUSTCODE " & _
                            "FROM AMIS_Journal_HD HD left outer JOIN AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo " & _
                            "AND HD.JType = DET.JType INNER JOIN ALL_Customer CUST ON " & _
                            "((HD.CustomerCode = CUST.CUSCDE) OR  ((RIGHT(DET.ENTITY,6)) = CUST.CUSCDE)) " & _
                            "WHERE ((LEFT(DET.Acct_Code, 5) = '11-02') OR (LEFT(DET.Acct_Code, 5) = '11-03')) and DET.ACCT_CODE =  '" & Setacctcode(cboAccountName.Text) & "' ORDER BY CUST.ACCTNAME", gconDMIS, adOpenKeyset
        Else
            'DESCRIPTION: THIS IS FOR VENDOR WITH AR ACCOUNT GROUP BY ACCOUNT CODE
            rsCustomer.Open "SELECT DISTINCT ALL_VENDOR.NameofVendor as VENDORNAME,ALL_VENDOR.ID as ID,ALL_VENDOR.CODE AS VENCODE " & _
                            "FROM dbo.AMIS_Journal_HD HD INNER JOIN dbo.AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo AND HD.JType = DET.JType " & _
                            "RIGHT OUTER JOIN dbo.ALL_VENDOR ALL_VENDOR ON HD.VENDORCODE = ALL_VENDOR.CODE WHERE  ALL_VENDOR.CODE IS NOT NULL  and DET.ACCT_CODE = '" & Setacctcode(cboAccountName.Text) & "'  and " & _
                            "((LEFT(DET.Acct_Code, 5) = '11-02') OR (LEFT(DET.Acct_Code, 5) = '11-03')) ORDER BY ALL_VENDOR.NameofVendor", gconDMIS, adOpenKeyset
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
        GET_BALANCE
        'UPDATED BY: JUN---------
        'FillGrids
        FILL_LEDGER
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
    Else
        rsCustomers.Open "SELECT DISTINCT ALL_VENDOR.NameofVendor as VENDORNAME,ALL_VENDOR.ID as ID,ALL_VENDOR.CODE AS VENCODE " & _
                         "FROM dbo.AMIS_Journal_HD HD INNER JOIN dbo.AMIS_Journal_Det DET ON HD.VoucherNo = DET.VoucherNo AND HD.JType = DET.JType " & _
                         "RIGHT OUTER JOIN dbo.ALL_VENDOR ALL_VENDOR ON HD.VENDORCODE = ALL_VENDOR.CODE WHERE  ALL_VENDOR.NAMEOFVENDOR like '" & ReplaceQuote(XXX) & "%'", gconDMIS, adOpenKeyset
    End If

    If Not (rsCustomers.EOF And rsCustomers.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsCustomers
        lstCustomer.Refresh
        lstCustomer.Enabled = True
        lstCustomer.Enabled = True
    Else
        'lstCustomer.Enabled = False
    End If
End Sub

Private Sub cboAccountName_Click()
'FillGrids
    FILL_LEDGER
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
    If MsgBox("Print Customers Ledger for this Account?", vbYesNo + vbQuestion, "Print: " & txtCustName.Text) = vbYes Then
        Dim filter

        'UPDATED BY: JUN/ARNOLD-------
        'DATE UPDATED: 06-11-2009
        BEG_BALANCE_DATE = dtFrom
        'UPDATED BY: JUN/ARNOLD-------

        'filter = "{Journal_HD.Status} = 'P' and (left({Journal_Det.Acct_Code},5)='11-02' OR left({Journal_Det.Acct_Code},5)='11-03' OR {Journal_Det.Acct_Code}='21-02008-00') and ({Customer.CusCde})='" & txtCode.Text & "'"
        If MsgBox("Generate for All Customer?", vbQuestion + vbYesNo, "Selecting No will generate only selected customer") = vbYes Then
            'filter = "({Journal_Det.Jdate} >= Date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Det.Jdate} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")) and {Journal_HD.Status} = 'P' and (left({Journal_Det.Acct_Code},5)='11-02' OR left({Journal_Det.Acct_Code},5)='11-03') and {Journal_Det.Acct_Code} = '" & Setacctcode(cboAccountName.Text) & "'"
            filter = "({Journal_Det.Jdate} >= Date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Det.Jdate} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")) and {Journal_HD.Status} = 'P' "
        Else
            'filter = "({Journal_Det.Jdate} >= Date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Det.Jdate} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")) and {Journal_HD.Status} = 'P' and (left({Journal_Det.Acct_Code},5)='11-02' OR left({Journal_Det.Acct_Code},5)='11-03') and {Journal_Det.Acct_Code} = '" & Setacctcode(cboAccountName.Text) & "' and ({Customer.CusCde})='" & txtCode.Text & "'"""
            filter = "({Journal_Det.Jdate} >= Date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Det.Jdate} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")) and {Journal_HD.Status} = 'P' and ((left({Journal_Det.Acct_Code},5)='11-02' OR left({Journal_Det.Acct_Code},5)='11-03')) and ({Customer.CUSCDE})='" & txtCode.Text & "'"
        End If
        ShowReport "CustomersSubsidiaryLedger", _
                   "Ledgers", _
                   filter, "C U S T O M E R S  L E D G E R", "AS OF: " & LOGDATE, True
    End If
    LogAudit "V", "CUSTOMERS A/R LEDGER", txtCode
End Sub

Private Sub Command1_Click()
'UPDATED BY: JUN
'DATE UPDATED: 06/22/2009
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

Private Sub Command2_Click()
    Dim j                                         As Integer

    FILL_LEDGER
    For j = 1 To lvwLedger.ListItems.Count
        If Combo1.Text = "SJ" Then
            If "SJ" & "-" & Format(RTrim(Trim(txtSearch.Text)), "000000") = lvwLedger.ListItems(j).ListSubItems.Item(1) Then
                lvwLedger.ListItems(j).ListSubItems.Item(1).ForeColor = vbBlue
                Exit For
            Else
            End If
        ElseIf Combo1.Text = "CRJ" Then
            If "CRJ" & "-" & Format(RTrim(Trim(txtSearch.Text)), "000000") = lvwLedger.ListItems(j).ListSubItems.Item(1) Then
                lvwLedger.ListItems(j).ListSubItems.Item(1).ForeColor = vbBlue
                Exit For
            Else
                'not found
            End If
        ElseIf Combo1.Text = "ADJ" Then
            If "ADJ" & "-" & Format(RTrim(Trim(txtSearch.Text)), "000000") = lvwLedger.ListItems(j).ListSubItems.Item(1) Then
                lvwLedger.ListItems(j).ListSubItems.Item(1).ForeColor = vbBlue
                Exit For
            Else
                'not found
            End If
        End If
        DoEvents
        lvwLedger.Refresh
    Next j
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

    labID.Caption = lstCustomer.SelectedItem.SubItems(1)
    rsRefresh
    rsCustomer.Find "ID =" & labID.Caption
    StoreMemVars
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    labID.Caption = lstCustomer.SelectedItem.SubItems(1)
    rsRefresh
    rsCustomer.Find "ID =" & labID.Caption
    'rsCUSTOMER.Bookmark = rsFind(rsCUSTOMER.Clone, "ID", lstCustomer.SelectedItem.SubItems(1)).Bookmark
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
        ElseIf JOURNALTYPE = "GJ" Then
            On Error Resume Next
            frmAMIS_GJ_JOURNAL_ENTRY.LoadJournal ("GJ")
            frmAMIS_GJ_JOURNAL_ENTRY.Show
            frmAMIS_GJ_JOURNAL_ENTRY.SearchVoucherNo (VARVOUCHERNO)
        ElseIf JOURNALTYPE = "APJ" Then
            On Error Resume Next
            frmAMISJournalEntry_APJ.LoadJournal ("APJ")
            frmAMISJournalEntry_APJ.Show
            frmAMISJournalEntry_APJ.SearchVoucherNo (VARVOUCHERNO)
        ElseIf JOURNALTYPE = "CDJ" Then
            On Error Resume Next
            frmAMISJournalEntry_APJ.LoadJournal ("CDJ")
            frmAMISJournalEntry_APJ.Show
            frmAMISJournalEntry_APJ.SearchVoucherNo (VARVOUCHERNO)
        ElseIf JOURNALTYPE = "SJ" Then
            On Error Resume Next
            frmAMISJournalEntry_SJ.LoadJournal ("SJ")
            frmAMISJournalEntry_SJ.Show
            frmAMISJournalEntry_SJ.StoreSearch (VARVOUCHERNO)
        ElseIf JOURNALTYPE = "CRJ" Then
            On Error Resume Next
            frmAMISJournalEntry_CRJ.LoadJournal ("CRJ")
            frmAMISJournalEntry_CRJ.Show
            frmAMISJournalEntry_CRJ.StoreSearch (VARVOUCHERNO)
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

Private Sub textSearch_Change()
    If Trim(TextSearch.Text) = "" Then
        FILL_CUST_VEN
        'FillGrid
    Else
        FillSearchGrid (TextSearch.Text)
        'FILL_CUST_VEN_SEARCH (TextSearch.Text)
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
Sub FILL_CUST_VEN_SEARCH(xCUSNAME As String)
    Dim rsFILL_CUST_VEN                           As ADODB.Recordset
    Dim rsFILL_Vendo                              As ADODB.Recordset

    Dim Item                                      As ListItem

    lstCustomer.ListItems.Clear

    If optCustomer.Value = True Then
        Set rsFILL_CUST_VEN = New ADODB.Recordset
        rsFILL_CUST_VEN.Open "SELECT DISTINCT dbo.ALL_Customer.ACCTNAME as CUSTNAME,dbo.ALL_Customer.ID as ID,dbo.ALL_Customer.CUSCDE AS CUSTCODE FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType RIGHT OUTER JOIN dbo.ALL_Customer ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer.CUSCDE WHERE  dbo.ALL_Customer.CUSCDE IS NOT NULL AND dbo.ALL_Customer.ACCTNAME LIKE '" & xCUSNAME & "%' ORDER BY dbo.ALL_Customer.ACCTNAME", gconDMIS, adOpenKeyset
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
                          "RIGHT OUTER JOIN dbo.ALL_Vendor ALL_VENDOR ON HD.VENDORCODE = ALL_VENDOR.CODE WHERE " & _
                          "ALL_VENDOR.NameofVendor LIKE '" & xCUSNAME & "%' ORDER BY ALL_VENDOR.VENDORNAME", gconDMIS, adOpenKeyset

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

Sub FILL_CUST_VEN()
    Dim rsFILL_CUST_VEN                           As ADODB.Recordset
    Dim rsFILL_Vendo                              As ADODB.Recordset

    Dim Item                                      As ListItem

    lstCustomer.ListItems.Clear

    If optCustomer.Value = True Then
        Set rsFILL_CUST_VEN = New ADODB.Recordset
        rsFILL_CUST_VEN.Open "SELECT DISTINCT dbo.ALL_Customer.ACCTNAME as CUSTNAME,dbo.ALL_Customer.ID as ID,dbo.ALL_Customer.CUSCDE AS CUSTCODE FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType Left outer JOIN dbo.ALL_Customer ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer.CUSCDE WHERE  dbo.ALL_Customer.CUSCDE IS NOT NULL  and   ((LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') OR (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03')) ORDER BY dbo.ALL_Customer.ACCTNAME", gconDMIS, adOpenKeyset
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
        'rsJournal_HDDet.Open "select AMIS_Journal_Hd.Debit AS DM,AMIS_Journal_Hd.Credit AS CM,AMIS_Journal_Hd.AmountToPay,AMIS_Journal_Hd.InvoiceAmt,AMIS_Journal_Hd.InvoiceNo,AMIS_Journal_HD.ID,AMIS_Journal_Hd.JNo,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.JType,AMIS_Journal_Det.Debit,AMIS_Journal_Det.Credit,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.CheckNo,AMIS_Journal_Hd.InvoiceType,AMIS_Journal_Hd.VendorCode,AMIS_Journal_Hd.JNo from AMIS_Journal_HD left outer Join AMIS_Journal_det on AMIS_Journal_Det.JNo  = AMIS_Journal_Hd.JNo and AMIS_Journal_Det.jtype = AMIS_Journal_Hd.Jtype  where " & _
         "(dbo.AMIS_Journal_HD.Jdate >= '" & dtFrom & "' and dbo.AMIS_Journal_HD.Jdate <= '" & dtTo & "') and (AMIS_Journal_Hd.Status = 'P' AND ((Left(AMIS_Journal_Det.Acct_Code,5) = '11-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '11-03' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-07') OR ((AMIS_Journal_HD.JTYPE = 'CCM' AND (left(AMIS_Journal_det.acct_code,5) = '11-02' or left(AMIS_Journal_det.acct_code,5) = '21-01')) OR (AMIS_Journal_HD.JTYPE = 'CSJ' AND (left(AMIS_Journal_det.acct_code,5) = '11-03' or left(AMIS_Journal_det.acct_code,5) = '21-07'))))) AND AMIS_Journal_Hd.CustomerCode = '" & txtCode.Text & "' order by AMIS_Journal_Hd.jdate asc,AMIS_Journal_Hd.id asc", gconDMIS
        'rsJournal_HDDet.Open "SELECT HD.STATUS AS SS,HD.ADJ_AMOUNT AS ADJ_AMOUNT,HD.ADJ_TYPE AS ADJ_TYPE,HD.Debit AS DM,HD.Credit AS CM,HD.AmountToPay,HD.InvoiceAmt,HD.InvoiceNo,HD.ID,HD.JNo,HD.JDate,HD.JType,HD_DET.Debit,HD_DET.Credit, HD.VoucherNo,HD.CheckNo,HD.InvoiceType,HD.VendorCode,HD.JNo FROM AMIS_Journal_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET HD_DET ON HD_DET.VoucherNo  = HD.VoucherNo AND HD_DET.jtype = HD.Jtype " & _
         "where ( HD.Jdate >= '" & dtFrom & "' and HD.Jdate <= '" & dtTo & "') and (((Left(HD_DET.Acct_Code,5) IN ('11-02','11-03','21-02','21-07'))) OR  ((HD.JTYPE = 'GJ' AND right(HD_DET.ENTITY,6) = '" & txtCode.Text & "' AND (left(HD_DET.acct_code,5) IN ('11-02','21-01','11-03','21-07'))))) AND HD.CustomerCode = '" & txtCode.Text & "' and HD.Status = 'P' order by HD.jdate asc,HD.id asc", gconDMIS, adOpenKeyset
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
            'rsJournal_HDDet.Open "select AMIS_Journal_Hd.Debit AS DM,AMIS_Journal_Hd.Credit AS CM,AMIS_Journal_Hd.AmountToPay,AMIS_Journal_Hd.InvoiceAmt,AMIS_Journal_Hd.InvoiceNo,AMIS_Journal_HD.ID,AMIS_Journal_Hd.JNo,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.JType,AMIS_Journal_Det.Debit,AMIS_Journal_Det.Credit,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.CheckNo,AMIS_Journal_Hd.InvoiceType,AMIS_Journal_Hd.VendorCode,AMIS_Journal_Hd.JNo from AMIS_Journal_HD left outer Join AMIS_Journal_det on AMIS_Journal_Det.JNo  = AMIS_Journal_Hd.JNo and AMIS_Journal_Det.jtype = AMIS_Journal_Hd.Jtype  where " & _
             "(dbo.AMIS_Journal_HD.Jdate >= '" & dtFrom & "' and dbo.AMIS_Journal_HD.Jdate <= '" & dtTo & "') and (AMIS_Journal_Hd.Status = 'P' AND AMIS_Journal_Det.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "' AND ((Left(AMIS_Journal_Det.Acct_Code,5) = '11-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '11-03' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-07') OR ((AMIS_Journal_HD.JTYPE = 'CCM' AND (left(AMIS_Journal_det.acct_code,5) = '11-01' or left(AMIS_Journal_det.acct_code,5) = '21-01')) OR (AMIS_Journal_HD.JTYPE = 'CSJ' AND (left(AMIS_Journal_det.acct_code,5) = '11-03' or left(AMIS_Journal_det.acct_code,5) = '21-07'))))) AND AMIS_Journal_Hd.CustomerCode = '" & txtCode.Text & "' order by AMIS_Journal_Hd.jdate asc,AMIS_Journal_Hd.id asc", gconDMIS
            'rsJournal_HDDet.Open "SELECT HD.ADJ_AMOUNT AS ADJ_AMOUNT,HD.ADJ_TYPE AS ADJ_TYPE,HD.Debit AS DM,HD.Credit AS CM,HD.AmountToPay,HD.InvoiceAmt,HD.InvoiceNo,HD.ID,HD.JNo,HD.JDate,HD.JType,HD_DET.Debit,HD_DET.Credit, HD.VoucherNo,HD.CheckNo,HD.InvoiceType,HD.VendorCode,HD.JNo FROM AMIS_Journal_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET HD_DET ON HD_DET.VoucherNo  = HD.VoucherNo AND HD_DET.jtype = HD.Jtype " & _
             "where ( HD.Jdate >= '" & dtFrom & "' and HD.Jdate <= '" & dtTo & "') and (HD_DET.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "' AND ((Left(HD_DET.Acct_Code,5) IN ('11-02','11-03','21-02','21-07'))) OR  ((HD.JTYPE = 'GJ' AND (left(HD_DET.acct_code,5) IN ('11-02','21-01','11-03','21-07'))))) AND HD.CustomerCode = '" & txtCode.Text & "'  and HD.Status = 'P'order by HD.jdate asc,HD.id asc", gconDMIS, adOpenKeyset
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
                        'If CHK_PYMENT_DISPLAY(Null2String(rsJournal_HDDet!VOUCHERNO), Null2String(rsJournal_HDDet!jtype)) = True Then
                        OUTBALANCE = N2Str2Zero(OUTBALANCE) + ((N2Str2Zero(rsJournal_HDDet!DEBIT) - N2Str2Zero(rsJournal_HDDet!CREDIT)))
                        'Else
                        'DON'T COMPUTE
                        'End If
                    Else
                        OUTBALANCE = N2Str2Zero(OUTBALANCE) + ((N2Str2Zero(rsJournal_HDDet!DEBIT) - N2Str2Zero(rsJournal_HDDet!CREDIT)))
                    End If

                End If

                '                If Null2String(rsJournal_HDDet!jtype) = "ADJ" Then
                '                    If Null2String(rsJournal_HDDet!ADJ_TYPE) = "CREDIT" Then
                '                        OUTBALANCE = N2Str2Zero(OUTBALANCE) - N2Str2Zero(rsJournal_HDDet!ADJ_AMOUNT)
                '                    Else
                '                        OUTBALANCE = N2Str2Zero(OUTBALANCE) + N2Str2Zero(rsJournal_HDDet!ADJ_AMOUNT)
                '                    End If
                '                Else
                '                    OUTBALANCE = N2Str2Zero(OUTBALANCE) + ((N2Str2Zero(rsJournal_HDDet!DEBIT) - N2Str2Zero(rsJournal_HDDet!CREDIT)))
                '                End If

            End If

            'If Not rsJournal_HDDet.EOF And Not rsJournal_HDDet.BOF Then
            'Do While Not rsJournal_HDDet.EOF
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
                    'ITEM.SubItems(2) = Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO)
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
                    'ITEM.SubItems(2) = Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO)
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
                    'If CHK_PYMENT_DISPLAY(Null2String(rsJournal_HDDet!VOUCHERNO), Null2String(rsJournal_HDDet!jtype)) = True Then
                    Set Item = lvwLedger.ListItems.Add(, , Null2String(rsJournal_HDDet!JDate))
                    Item.SubItems(1) = Null2String(Reference)
                    Item.SubItems(2) = Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO)
                    Item.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!DEBIT))
                    Item.SubItems(4) = ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!CREDIT))
                    Item.SubItems(5) = ToDoubleNumber(OUTBALANCE)
                    Item.SubItems(6) = rsJournal_HDDet!ID
                    Item.SubItems(7) = Null2String(rsJournal_HDDet!jtype)
                    Item.SubItems(8) = Null2String(rsJournal_HDDet!InvoiceType)
                    'Else
                    '    lvw_COUNT = lvw_COUNT - 1
                    'DONT DISPLAY BECAUSE CUSTOMER CODE FOR PAYMENT IS WRONG EVENTHOUGH THE INVOICETYPE AND INVOICE IS IN SALES JOURNAL
                    'End If
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
                    'If CHK_PYMENT_DISPLAY(Null2String(rsJournal_HDDet!VOUCHERNO), Null2String(rsJournal_HDDet!jtype)) = True Then
                    TUTAL_DEBIT = TUTAL_DEBIT + N2Str2Zero(rsJournal_HDDet!DEBIT)
                    TUTAL_CREDIT = TUTAL_CREDIT + N2Str2Zero(rsJournal_HDDet!CREDIT)
                    'Else
                    'DON'T PAYMENT
                    'End If
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



