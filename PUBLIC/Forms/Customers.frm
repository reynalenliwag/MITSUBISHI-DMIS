VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAllCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers"
   ClientHeight    =   8895
   ClientLeft      =   525
   ClientTop       =   795
   ClientWidth     =   12915
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00F5F5F5&
   Icon            =   "Customers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Customers.frx":08CA
   ScaleHeight     =   8895
   ScaleWidth      =   12915
   Begin VB.PictureBox picCredit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   6000
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   2985
      ScaleWidth      =   3585
      TabIndex        =   120
      Top             =   3600
      Visible         =   0   'False
      Width           =   3615
      Begin VB.ComboBox cboInvoiceType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Customers.frx":0C0C
         Left            =   1320
         List            =   "Customers.frx":0C16
         Style           =   2  'Dropdown List
         TabIndex        =   145
         Top             =   420
         Width           =   2145
      End
      Begin VB.CheckBox chkZeroRated 
         Appearance      =   0  'Flat
         Caption         =   "Zero Rate  Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         TabIndex        =   130
         Top             =   1740
         Width           =   2205
      End
      Begin VB.TextBox txtCreditDays 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   124
         Text            =   "Text1"
         Top             =   1290
         Width           =   2145
      End
      Begin VB.TextBox txtCreditLimit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   122
         Text            =   "Text1"
         Top             =   870
         Width           =   2145
      End
      Begin VB.CommandButton cmdCloseTerm 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Index           =   1
         Left            =   2760
         MouseIcon       =   "Customers.frx":0C32
         MousePointer    =   99  'Custom
         Picture         =   "Customers.frx":0D84
         Style           =   1  'Graphical
         TabIndex        =   127
         ToolTipText     =   "Cancel Entry"
         Top             =   2130
         Width           =   705
      End
      Begin VB.CommandButton cmdSaveTerms 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   2040
         MouseIcon       =   "Customers.frx":10C2
         MousePointer    =   99  'Custom
         Picture         =   "Customers.frx":1214
         Style           =   1  'Graphical
         TabIndex        =   126
         ToolTipText     =   "Save Entry"
         Top             =   2130
         Width           =   705
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   345
         Left            =   0
         TabIndex        =   165
         Top             =   0
         Width           =   3615
         _Version        =   655364
         _ExtentX        =   6376
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "CREDITS AND TERMS"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         GradientColorLight=   16711680
         GradientColorDark=   4194304
         ForeColor       =   16777215
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Type:"
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
         Left            =   120
         TabIndex        =   144
         Top             =   510
         Width           =   1170
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Days:"
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
         Left            =   120
         TabIndex        =   123
         Top             =   1380
         Width           =   1170
      End
      Begin VB.Label labTermID 
         Height          =   555
         Left            =   360
         TabIndex        =   125
         Top             =   2790
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Limit:"
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
         Left            =   120
         TabIndex        =   121
         Top             =   930
         Width           =   1170
      End
   End
   Begin VB.TextBox labOLDCuscde 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00400000&
      Height          =   450
      Left            =   14070
      TabIndex        =   108
      Top             =   2430
      Visible         =   0   'False
      Width           =   1500
   End
   Begin Crystal.CrystalReport rptCustomer 
      Left            =   1230
      Top             =   8430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   9315
      Left            =   0
      ScaleHeight     =   9315
      ScaleWidth      =   14355
      TabIndex        =   0
      Top             =   0
      Width           =   14355
      Begin VB.PictureBox picToolFrame 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   3540
         ScaleHeight     =   945
         ScaleWidth      =   9345
         TabIndex        =   10
         Top             =   0
         Width           =   9375
         Begin VB.CommandButton cmdHistory 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   6420
            MouseIcon       =   "Customers.frx":1564
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":16B6
            Style           =   1  'Graphical
            TabIndex        =   142
            ToolTipText     =   "View Customer Transaction History"
            Top             =   330
            Width           =   585
         End
         Begin VB.CommandButton cmdDuplicate 
            Height          =   555
            Left            =   4770
            MouseIcon       =   "Customers.frx":1D80
            MousePointer    =   99  'Custom
            OLEDropMode     =   1  'Manual
            Picture         =   "Customers.frx":1ED2
            Style           =   1  'Graphical
            TabIndex        =   133
            Tag             =   "1102"
            ToolTipText     =   "View Sales Calculator"
            Top             =   330
            Width           =   585
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   8550
            Top             =   900
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.CommandButton cmdCUSTINFO_Contact 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   3030
            MouseIcon       =   "Customers.frx":2351
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":24A3
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Contact Information"
            Top             =   330
            Width           =   585
         End
         Begin VB.CommandButton cmdCustInfo_Child 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1530
            MouseIcon       =   "Customers.frx":2B95
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":2CE7
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "View Customers Number of Children"
            Top             =   330
            Width           =   585
         End
         Begin VB.CommandButton cmdCustInfo_Credit 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   30
            MouseIcon       =   "Customers.frx":3308
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":345A
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Update Credit and Terms of Customers"
            Top             =   330
            Width           =   585
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "History"
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   7050
            TabIndex        =   143
            Top             =   480
            Width           =   570
         End
         Begin VB.Label lblDuplicate 
            BackStyle       =   0  'Transparent
            Caption         =   "Possible Duplicate Customer"
            ForeColor       =   &H00000000&
            Height          =   645
            Left            =   5400
            TabIndex        =   134
            Top             =   270
            Width           =   870
         End
         Begin VB.Label labCustCode 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "A00001"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   7860
            TabIndex        =   128
            Top             =   390
            Width           =   1425
         End
         Begin VB.Label labCustInfo_Contact 
            Caption         =   "Contact Information"
            Height          =   495
            Left            =   3660
            MouseIcon       =   "Customers.frx":3ABD
            MousePointer    =   99  'Custom
            TabIndex        =   17
            Top             =   390
            Width           =   1215
         End
         Begin XtremeShortcutBar.ShortcutCaption CapInfo 
            Height          =   300
            Index           =   2
            Left            =   0
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   -30
            Width           =   9810
            _Version        =   655364
            _ExtentX        =   17304
            _ExtentY        =   529
            _StockProps     =   14
            Caption         =   "Customers Information"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   64
         End
         Begin VB.Label labCustInfo_Child 
            Caption         =   "Children"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2190
            MouseIcon       =   "Customers.frx":3DC7
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   510
            Width           =   975
         End
         Begin VB.Label labCustInfo_Credit 
            Alignment       =   2  'Center
            Caption         =   "Credit && Terms"
            Height          =   465
            Left            =   660
            MouseIcon       =   "Customers.frx":40D1
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   390
            Width           =   825
         End
         Begin VB.Label labCustCode2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "A00001"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   7890
            TabIndex        =   129
            Top             =   390
            Width           =   1425
         End
      End
      Begin VB.Frame fraSearch 
         Height          =   8865
         Left            =   0
         TabIndex        =   1
         Top             =   -90
         Width           =   3525
         Begin VB.OptionButton Option1 
            Caption         =   "Search By Code"
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
            MouseIcon       =   "Customers.frx":43DB
            MousePointer    =   99  'Custom
            TabIndex        =   131
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox txtSearch 
            Appearance      =   0  'Flat
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
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1470
            Width           =   3405
         End
         Begin VB.OptionButton optSearchKeyLast 
            Caption         =   "Search By Last Name/Company"
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
            MouseIcon       =   "Customers.frx":452D
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   180
            Width           =   3105
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
            Left            =   2700
            MouseIcon       =   "Customers.frx":467F
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   1380
            Visible         =   0   'False
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
            MouseIcon       =   "Customers.frx":47D1
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   450
            Value           =   -1  'True
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
            MouseIcon       =   "Customers.frx":4923
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   705
            Width           =   2295
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
            MouseIcon       =   "Customers.frx":4A75
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   960
            Width           =   2295
         End
         Begin MSComctlLib.ListView lstCustomer 
            Height          =   6945
            Left            =   30
            TabIndex        =   9
            Top             =   1890
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   12250
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
            Appearance      =   0
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
            MouseIcon       =   "Customers.frx":4BC7
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "A/C Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
         End
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
            ItemData        =   "Customers.frx":4D29
            Left            =   60
            List            =   "Customers.frx":4D2B
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1740
            Visible         =   0   'False
            Width           =   2325
         End
      End
      Begin VB.Frame Frame1 
         Height          =   7125
         Left            =   3540
         TabIndex        =   18
         Top             =   900
         Width           =   9315
         Begin VB.TextBox txtAcctName 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00400000&
            Height          =   330
            Left            =   5190
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   22
            Top             =   195
            Width           =   3945
         End
         Begin VB.ComboBox cboCustType 
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
            ItemData        =   "Customers.frx":4D2D
            Left            =   1320
            List            =   "Customers.frx":4D2F
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   180
            Width           =   2835
         End
         Begin VB.Frame Frame3 
            Caption         =   "Notes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   3135
            Left            =   4290
            TabIndex        =   58
            Top             =   3840
            Width           =   4965
            Begin VB.TextBox txtNotes 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   2835
               Left            =   60
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   59
               Top             =   240
               Width           =   4845
            End
         End
         Begin VB.Frame fraEntity 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   1920
            Left            =   30
            TabIndex        =   23
            Top             =   450
            Width           =   9225
            Begin VB.Timer Timer1 
               Interval        =   500
               Left            =   0
               Top             =   -420
            End
            Begin VB.ComboBox cboPersonalCity 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   4200
               TabIndex        =   37
               Text            =   "cboApod"
               Top             =   900
               Width           =   1995
            End
            Begin VB.TextBox txtPersonalStreet 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   375
               Left            =   120
               ScrollBars      =   2  'Vertical
               TabIndex        =   36
               Top             =   900
               Width           =   4035
            End
            Begin VB.TextBox txtPersonalState 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   6240
               MaxLength       =   30
               TabIndex        =   38
               Top             =   900
               Width           =   1695
            End
            Begin VB.TextBox txtPersonalZIP 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   7980
               MaxLength       =   6
               TabIndex        =   39
               Top             =   900
               Width           =   1155
            End
            Begin VB.TextBox txtLastName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1020
               TabIndex        =   29
               ToolTipText     =   "LAST NAME OR COMPANY NAME"
               Top             =   360
               Width           =   2715
            End
            Begin VB.ComboBox cboApod 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   120
               TabIndex        =   28
               Text            =   "cboApod"
               Top             =   360
               Width           =   855
            End
            Begin VB.TextBox txtMiddleName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   6435
               MaxLength       =   50
               TabIndex        =   31
               Top             =   360
               Width           =   2700
            End
            Begin VB.TextBox txtFirstName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   3780
               TabIndex        =   30
               Top             =   360
               Width           =   2625
            End
            Begin VB.ComboBox cboSex 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   2040
               TabIndex        =   44
               Text            =   "cboSex"
               Top             =   1500
               Width           =   855
            End
            Begin VB.TextBox txtBirthDate 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   120
               MaxLength       =   10
               TabIndex        =   43
               Top             =   1500
               Width           =   1875
            End
            Begin VB.TextBox txtSpouse 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   2940
               MaxLength       =   100
               TabIndex        =   45
               Top             =   1500
               Width           =   6195
            End
            Begin VB.Label Label32 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   900
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   180
               Index           =   3
               Left            =   4020
               TabIndex        =   163
               Top             =   600
               Width           =   150
            End
            Begin VB.Label Label32 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   900
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   180
               Index           =   2
               Left            =   -30
               TabIndex        =   162
               Top             =   600
               Width           =   150
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "City"
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
               Height          =   225
               Left            =   4200
               TabIndex        =   34
               Top             =   690
               Width           =   315
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Street"
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
               Height          =   195
               Left            =   120
               TabIndex        =   33
               Top             =   690
               Width           =   525
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Province"
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
               Height          =   225
               Left            =   6240
               TabIndex        =   35
               Top             =   690
               Width           =   750
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Zip Code"
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
               Height          =   225
               Left            =   7980
               TabIndex        =   32
               Top             =   660
               Width           =   735
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Salutation"
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
               Height          =   225
               Left            =   120
               TabIndex        =   24
               Top             =   150
               Width           =   855
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Middle Name"
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
               Height          =   225
               Left            =   6420
               TabIndex        =   27
               Top             =   150
               Width           =   1095
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "First Name"
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
               Height          =   225
               Left            =   3780
               TabIndex        =   26
               Top             =   150
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Last Name"
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
               Height          =   225
               Left            =   1050
               TabIndex        =   25
               Top             =   150
               Width           =   915
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Sex"
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
               Height          =   225
               Left            =   2010
               TabIndex        =   42
               Top             =   1290
               Width           =   330
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Birth Date"
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
               Height          =   225
               Left            =   120
               TabIndex        =   40
               Top             =   1260
               Width           =   840
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Spouse Name"
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
               Height          =   225
               Left            =   2910
               TabIndex        =   41
               Top             =   1290
               Width           =   1185
            End
         End
         Begin VB.Frame fraMiscellenous 
            Height          =   2085
            Left            =   30
            TabIndex        =   46
            Top             =   2340
            Width           =   4185
            Begin VB.CheckBox chkWithholdingTax 
               Caption         =   "Withholding Tax Agent "
               Height          =   315
               Left            =   1230
               TabIndex        =   141
               Top             =   1620
               Visible         =   0   'False
               Width           =   2715
            End
            Begin VB.TextBox txtTin 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   1245
               MaxLength       =   15
               TabIndex        =   48
               Top             =   160
               Width           =   2775
            End
            Begin VB.TextBox txtDepartment 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   1245
               TabIndex        =   54
               Top             =   1260
               Width           =   2775
            End
            Begin VB.TextBox txtTitle 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   52
               Top             =   915
               Width           =   2775
            End
            Begin VB.ComboBox cboLeadSource 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1245
               TabIndex        =   50
               Text            =   "cboLeadSource"
               Top             =   525
               Width           =   2775
            End
            Begin VB.Label Label32 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   900
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   180
               Index           =   0
               Left            =   660
               TabIndex        =   161
               Top             =   120
               Width           =   150
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "TIN"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   840
               TabIndex        =   47
               Top             =   180
               Width           =   285
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Department"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   180
               TabIndex        =   53
               Top             =   1245
               Width           =   975
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Position"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   480
               TabIndex        =   51
               Top             =   915
               Width           =   675
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Lead Source"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   105
               TabIndex        =   49
               Top             =   555
               Width           =   1050
            End
         End
         Begin VB.Frame fraContacts 
            Height          =   2745
            Left            =   30
            TabIndex        =   146
            Top             =   4320
            Width           =   4185
            Begin VB.TextBox txtEmail 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1245
               TabIndex        =   153
               Top             =   150
               Width           =   2775
            End
            Begin VB.TextBox txtAssistant 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   152
               Top             =   2010
               Width           =   2775
            End
            Begin VB.TextBox txtAsstPhone 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1245
               TabIndex        =   151
               Top             =   2370
               Width           =   2775
            End
            Begin VB.TextBox txtCusphon1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   150
               Top             =   540
               Width           =   2775
            End
            Begin VB.TextBox txtMobile 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   149
               Top             =   915
               Width           =   2775
            End
            Begin VB.TextBox txtHomePhone 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   148
               Top             =   1290
               Width           =   2775
            End
            Begin VB.TextBox txtFax 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   147
               Top             =   1665
               Width           =   2775
            End
            Begin VB.Label Label32 
               BackStyle       =   0  'Transparent
               Caption         =   "*"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   900
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   180
               Index           =   4
               Left            =   450
               TabIndex        =   164
               Top             =   840
               Width           =   150
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Email"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   675
               TabIndex        =   160
               Top             =   180
               Width           =   480
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Assistant"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   390
               TabIndex        =   159
               Top             =   2040
               Width           =   765
            End
            Begin VB.Label lblCap 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Asst. Phone"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   165
               TabIndex        =   158
               Top             =   2400
               Width           =   990
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Office Phone"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   105
               TabIndex        =   157
               Top             =   570
               Width           =   1050
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   615
               TabIndex        =   156
               Top             =   930
               Width           =   540
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Home Phone"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   60
               TabIndex        =   155
               Top             =   1290
               Width           =   1095
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Fax"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   840
               TabIndex        =   154
               Top             =   1680
               Width           =   285
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Delivery Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   1335
            Left            =   4290
            TabIndex        =   55
            Top             =   2490
            Width           =   4995
            Begin VB.TextBox txtDeliveryAddress 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   825
               Left            =   60
               MaxLength       =   150
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   57
               Top             =   420
               Width           =   4845
            End
            Begin VB.CommandButton cmdSame 
               Caption         =   "Same As above"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3480
               TabIndex        =   56
               Top             =   150
               Width           =   1395
            End
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Acct. Name"
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
            Height          =   210
            Left            =   4200
            TabIndex        =   21
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label23 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Type"
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
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   11430
         ScaleHeight     =   885
         ScaleWidth      =   1590
         TabIndex        =   60
         Top             =   8070
         Width           =   1590
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
            Height          =   795
            Left            =   690
            MouseIcon       =   "Customers.frx":4D31
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":4E83
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Cancel"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
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
            Left            =   0
            MouseIcon       =   "Customers.frx":51C1
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":5313
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Save this Record"
            Top             =   0
            Width           =   705
         End
      End
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   900
         ScaleHeight     =   960
         ScaleWidth      =   12315
         TabIndex        =   63
         Top             =   8070
         Width           =   12315
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
            Left            =   11220
            MouseIcon       =   "Customers.frx":5663
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":57B5
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Exit Window"
            Top             =   0
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
            Left            =   10530
            MouseIcon       =   "Customers.frx":5B1B
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":5C6D
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Print this Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
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
            Left            =   9845
            MouseIcon       =   "Customers.frx":5FD3
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":6125
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Delete Selected Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
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
            Left            =   9150
            MouseIcon       =   "Customers.frx":6450
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":65A2
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Edit Selected Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
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
            Left            =   8465
            MouseIcon       =   "Customers.frx":68FE
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":6A50
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Add Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton cmdLast 
            Caption         =   "Last"
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
            Left            =   7740
            MouseIcon       =   "Customers.frx":6D63
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":6EB5
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Move to Last Record"
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton cmdFirst 
            Caption         =   "First"
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
            Left            =   7025
            MouseIcon       =   "Customers.frx":7205
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":7357
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Move to First Record"
            Top             =   0
            Width           =   735
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
            Left            =   6330
            MouseIcon       =   "Customers.frx":76B5
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":7807
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Find a Record"
            Top             =   0
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
            Left            =   5645
            MouseIcon       =   "Customers.frx":7B01
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":7C53
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Move to Next Record"
            Top             =   0
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
            Left            =   4950
            MouseIcon       =   "Customers.frx":7FAB
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":80FD
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Move to Previous Record"
            Top             =   0
            Width           =   705
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Select"
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
            Left            =   4260
            MouseIcon       =   "Customers.frx":845C
            MousePointer    =   99  'Custom
            Picture         =   "Customers.frx":85AE
            Style           =   1  'Graphical
            TabIndex        =   132
            ToolTipText     =   "Move to Previous Record"
            Top             =   0
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label lblloyal 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "With Loyalty ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   615
            Left            =   2730
            TabIndex        =   139
            Top             =   60
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "With Loyalty ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   945
            Left            =   2760
            TabIndex        =   140
            Top             =   90
            Width           =   1455
         End
      End
   End
   Begin VB.TextBox txtCuscde 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00400000&
      Height          =   450
      Left            =   14070
      TabIndex        =   83
      Top             =   1920
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picContactAE 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   5820
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   4665
      ScaleWidth      =   4515
      TabIndex        =   89
      Top             =   2880
      Visible         =   0   'False
      Width           =   4545
      Begin VB.TextBox txtContactName 
         Height          =   345
         Left            =   1260
         MaxLength       =   100
         TabIndex        =   91
         Top             =   630
         Width           =   3045
      End
      Begin VB.CommandButton cmdCloseContactsAE 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   1
         Left            =   3600
         MouseIcon       =   "Customers.frx":890D
         MousePointer    =   99  'Custom
         Picture         =   "Customers.frx":8A5F
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Cancel Entry"
         Top             =   3840
         Width           =   705
      End
      Begin VB.CommandButton cmdSaveContact 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2900
         MouseIcon       =   "Customers.frx":8D9D
         MousePointer    =   99  'Custom
         Picture         =   "Customers.frx":8EEF
         Style           =   1  'Graphical
         TabIndex        =   105
         ToolTipText     =   "Save Details"
         Top             =   3840
         Width           =   705
      End
      Begin VB.CommandButton cmdDeleteContact 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2180
         MouseIcon       =   "Customers.frx":923F
         MousePointer    =   99  'Custom
         Picture         =   "Customers.frx":9391
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "Delect Details"
         Top             =   3840
         Width           =   705
      End
      Begin VB.ComboBox cboContactRelation 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00400000&
         Height          =   345
         ItemData        =   "Customers.frx":96BC
         Left            =   1260
         List            =   "Customers.frx":96BE
         TabIndex        =   93
         Top             =   1035
         Width           =   3045
      End
      Begin VB.TextBox txtContactPosition 
         Height          =   345
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   95
         Top             =   1425
         Width           =   3045
      End
      Begin VB.TextBox txtContactDepartment 
         Height          =   345
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   96
         Top             =   1830
         Width           =   3045
      End
      Begin VB.TextBox txtContactPhone 
         Height          =   345
         Left            =   1260
         MaxLength       =   20
         TabIndex        =   98
         Top             =   2235
         Width           =   3045
      End
      Begin VB.TextBox txtContactMobile 
         Height          =   345
         Left            =   1260
         MaxLength       =   20
         TabIndex        =   100
         Top             =   2625
         Width           =   3045
      End
      Begin VB.TextBox txtContactAddress 
         Height          =   645
         Left            =   1260
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   103
         Top             =   3030
         Width           =   3045
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Index           =   1
         Left            =   0
         TabIndex        =   168
         Top             =   0
         Width           =   4545
         _Version        =   655364
         _ExtentX        =   8017
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "CONTACTS INFORMATION"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         GradientColorLight=   16711680
         GradientColorDark=   4194304
         ForeColor       =   16777215
      End
      Begin VB.Label labIDContacts 
         BackColor       =   &H80000013&
         Height          =   555
         Left            =   120
         TabIndex        =   104
         Top             =   3960
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Relation:"
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
         Left            =   180
         TabIndex        =   92
         Top             =   1090
         Width           =   1050
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Left            =   180
         TabIndex        =   90
         Top             =   660
         Width           =   1050
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Position:"
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
         Left            =   180
         TabIndex        =   94
         Top             =   1490
         Width           =   1050
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Department:"
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
         Left            =   180
         TabIndex        =   97
         Top             =   1890
         Width           =   1050
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
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
         Left            =   180
         TabIndex        =   99
         Top             =   2280
         Width           =   1050
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile:"
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
         Index           =   1
         Left            =   180
         TabIndex        =   101
         Top             =   2640
         Width           =   1050
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
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
         Left            =   180
         TabIndex        =   102
         Top             =   2970
         Width           =   1050
      End
   End
   Begin VB.PictureBox picContactList 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4845
      Left            =   5400
      ScaleHeight     =   4815
      ScaleWidth      =   5370
      TabIndex        =   84
      Top             =   2040
      Visible         =   0   'False
      Width           =   5400
      Begin VB.CommandButton cmdCancelContactList 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4535
         MouseIcon       =   "Customers.frx":96C0
         MousePointer    =   99  'Custom
         Picture         =   "Customers.frx":9812
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Cancel"
         Top             =   4030
         Width           =   705
      End
      Begin VB.CommandButton cmdEditContact 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3820
         MouseIcon       =   "Customers.frx":9B50
         MousePointer    =   99  'Custom
         Picture         =   "Customers.frx":9CA2
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Edit Contact"
         Top             =   4030
         Width           =   705
      End
      Begin MSComctlLib.ListView lvContactList 
         Height          =   3495
         Left            =   105
         TabIndex        =   85
         Top             =   450
         Width           =   5150
         _ExtentX        =   9075
         _ExtentY        =   6165
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Customers.frx":9FFE
         NumItems        =   0
      End
      Begin VB.CommandButton cmdAddContactList 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3110
         MouseIcon       =   "Customers.frx":A160
         MousePointer    =   99  'Custom
         Picture         =   "Customers.frx":A2B2
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Add Contact"
         Top             =   4030
         Width           =   705
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   345
         Index           =   2
         Left            =   0
         TabIndex        =   169
         Top             =   0
         Width           =   5685
         _Version        =   655364
         _ExtentX        =   10028
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "LIST OF CONTACTS"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         GradientColorLight=   16711680
         GradientColorDark=   4194304
         ForeColor       =   16777215
      End
   End
   Begin VB.PictureBox PICLoyaltyID 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   6720
      ScaleHeight     =   1875
      ScaleWidth      =   3585
      TabIndex        =   135
      Top             =   3960
      Width           =   3615
      Begin VB.CommandButton CmdloyaltyCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2740
         MouseIcon       =   "Customers.frx":A5C5
         MousePointer    =   99  'Custom
         Picture         =   "Customers.frx":A717
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Cancel"
         Top             =   960
         Width           =   705
      End
      Begin VB.TextBox txtLoyalID 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   120
         TabIndex        =   136
         ToolTipText     =   "Type Loyalty ID here"
         Top             =   470
         Width           =   3345
      End
      Begin VB.CommandButton CmdloyaltySave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2040
         MouseIcon       =   "Customers.frx":AA55
         MousePointer    =   99  'Custom
         Picture         =   "Customers.frx":ABA7
         Style           =   1  'Graphical
         TabIndex        =   138
         ToolTipText     =   "Save this Record"
         Top             =   960
         Width           =   705
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   345
         Left            =   0
         TabIndex        =   170
         Top             =   0
         Width           =   4125
         _Version        =   655364
         _ExtentX        =   7276
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "LOYALTY IDENTIFICATION NO."
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16711680
         GradientColorDark=   4194304
         ForeColor       =   16777215
      End
   End
   Begin VB.PictureBox picChildAE 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      FillColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   2265
      Left            =   5760
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   2235
      ScaleWidth      =   4875
      TabIndex        =   109
      Top             =   3120
      Visible         =   0   'False
      Width           =   4905
      Begin VB.CommandButton cmdCloseChildAE 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   3
         Left            =   3990
         MouseIcon       =   "Customers.frx":AEF7
         MousePointer    =   99  'Custom
         Picture         =   "Customers.frx":B049
         Style           =   1  'Graphical
         TabIndex        =   118
         ToolTipText     =   "Cancel Entry"
         Top             =   1410
         Width           =   705
      End
      Begin VB.CommandButton cmdSaveChild 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3280
         MouseIcon       =   "Customers.frx":B387
         MousePointer    =   99  'Custom
         Picture         =   "Customers.frx":B4D9
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Save Children Information"
         Top             =   1410
         Width           =   705
      End
      Begin VB.CommandButton cmdDeleteChild 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2580
         MouseIcon       =   "Customers.frx":B829
         MousePointer    =   99  'Custom
         Picture         =   "Customers.frx":B97B
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Add Children Information"
         Top             =   1410
         Width           =   705
      End
      Begin VB.TextBox txtChildName 
         Height          =   345
         Left            =   1320
         TabIndex        =   111
         Top             =   510
         Width           =   3375
      End
      Begin VB.ComboBox cboChildSex 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00400000&
         Height          =   345
         ItemData        =   "Customers.frx":BCA6
         Left            =   3840
         List            =   "Customers.frx":BCB3
         TabIndex        =   115
         Top             =   900
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtChildDate 
         Height          =   345
         Left            =   1320
         TabIndex        =   113
         Top             =   900
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "mm/dd/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.Label labIdCHILD 
         BackColor       =   &H80000013&
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   120
         TabIndex        =   119
         Top             =   1560
         Visible         =   0   'False
         Width           =   645
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   345
         Index           =   2
         Left            =   0
         TabIndex        =   166
         Top             =   0
         Width           =   4905
         _Version        =   655364
         _ExtentX        =   8652
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   " CHILDREN INFORMATION"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         GradientColorLight=   16711680
         GradientColorDark=   4194304
         ForeColor       =   16777215
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Left            =   165
         TabIndex        =   110
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Birth:"
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
         Left            =   165
         TabIndex        =   112
         Top             =   990
         Width           =   1125
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SEX:"
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
         Left            =   3405
         TabIndex        =   114
         Top             =   990
         Width           =   375
      End
   End
   Begin VB.PictureBox picChildList 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4905
      Left            =   5400
      ScaleHeight     =   4875
      ScaleWidth      =   5655
      TabIndex        =   77
      Top             =   2280
      Visible         =   0   'False
      Width           =   5685
      Begin VB.CommandButton cmdCancelChildList 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   4800
         MouseIcon       =   "Customers.frx":BCC0
         MousePointer    =   99  'Custom
         Picture         =   "Customers.frx":BE12
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Cancel"
         Top             =   4080
         Width           =   705
      End
      Begin VB.CommandButton cmdSelectChild 
         Caption         =   "&Select"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   4080
         MouseIcon       =   "Customers.frx":C150
         MousePointer    =   99  'Custom
         Picture         =   "Customers.frx":C2A2
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Select"
         Top             =   4080
         Width           =   705
      End
      Begin MSComctlLib.ListView lvChildList 
         Height          =   3495
         Left            =   120
         TabIndex        =   78
         Top             =   450
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   6165
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Customers.frx":C5DE
         NumItems        =   0
      End
      Begin VB.CommandButton cmdAddChildInfo 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   3360
         MouseIcon       =   "Customers.frx":C740
         MousePointer    =   99  'Custom
         Picture         =   "Customers.frx":C892
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Add Children/Dependent"
         Top             =   4080
         Width           =   705
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   345
         Index           =   0
         Left            =   0
         TabIndex        =   167
         Top             =   0
         Width           =   5685
         _Version        =   655364
         _ExtentX        =   10028
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "LIST OF CHILDRENS/DEPENDENTS"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         GradientColorLight=   16711680
         GradientColorDark=   4194304
         ForeColor       =   16777215
      End
   End
   Begin VB.Label labid 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12450
      TabIndex        =   74
      Top             =   210
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label labSEQ 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12450
      TabIndex        =   76
      Top             =   1050
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label labSEQMAX 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12450
      TabIndex        =   82
      Top             =   1500
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label labPrev 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   12450
      TabIndex        =   75
      Top             =   660
      Visible         =   0   'False
      Width           =   1545
   End
End
Attribute VB_Name = "frmAllCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS                                                              As ADODB.Recordset
Dim rsCusCtl                                                        As ADODB.Recordset
Dim AddorEdit                                                       As String
Dim AccountCode                                                     As String
Dim CustType                                                        As String
Dim EntryPoint                                                      As String
Dim GoingWhere                                                      As String
Dim TempProspectID                                                  As Long
Event CustomerSelected(xCUSCODE As String, XaCCOUNTNAME As String)
Event ChangedData(xCUSCODE As String)
Event ProspectConverted(CustomerCode As String, xGoingWhere As String, PROSPECTID As Long)

Dim rsTERMS                                                         As ADODB.Recordset
'****************** REVISED BY DHANG_ERZ 100915 10:30AM
Function GetCustomerCode(lastname As String) As String
    Dim temprs                                                      As ADODB.Recordset
    Dim lAlpha                                                      As String
    Dim xCusNum                                                     As Integer
    Dim xCus2Num                                                    As String
    If Len(lastname) = 0 Then
        Exit Function
    End If
    
    lAlpha = Left(Trim(lastname), 1)
    If IsNumeric(lAlpha) = True Then
        lAlpha = "Z"
    End If
    Set temprs = gconDMIS.Execute("Select Cuscde From ALL_CUSMAS where LEFT(Cuscde,1)='" & lAlpha & "' order by Cuscde desc")
    If Not (temprs.EOF Or temprs.BOF) Then
        If IsNumeric(Left(lastname, 1)) = True Then
            xCusNum = N2Str2IntZero(Right(temprs!CUSCDE, 5)) + 1
            xCus2Num = "" & Format(xCusNum, "00000") & ""
            GetCustomerCode = "Z" + xCus2Num
        Else
            xCusNum = N2Str2IntZero(Right(temprs!CUSCDE, 5)) + 1
            xCus2Num = "" & Format(xCusNum, "00000") & ""
            GetCustomerCode = lAlpha + xCus2Num
        End If
    Else
        If IsNumeric(Left(lastname, 1)) = True Then
            GetCustomerCode = "Z" & "00001"
        Else
            GetCustomerCode = Left(lastname, 1) & "00001"
        End If
    End If
End Function
'****************** REVISED BY DHANG_ERZ 100915 10:30AM

Sub FillSearchGrid(XXX As String)
    Dim rsCustomer2                                                 As ADODB.Recordset

    lstCustomer.Enabled = False
    lstCustomer.Sorted = False
    lstCustomer.ListItems.Clear
    
    Set rsCustomer2 = New ADODB.Recordset
    'XXX = Repleys(LTrim(RTrim(XXX)))
    If optSearchKeyAcctName.Value = True Then
        Set rsCustomer2 = gconDMIS.Execute("SELECT ACCTNAME AS CustomerName,ACCTNAME,id FROM ALL_CUSTOMER WHERE AcctName LIKE '" & XXX & "%' ORDER BY AcctName ASC")
    ElseIf optSearchKeyAddress.Value = True Then
        Set rsCustomer2 = gconDMIS.Execute("SELECT ACCTNAME AS CustomerName,CustomerAdd,id FROM ALL_CUSTOMER WHERE CustomerAdd LIKE '" & XXX & "%' ORDER BY CustomerAdd ASC")
    ElseIf optSearchKeyCompany.Value = True Then
        Set rsCustomer2 = gconDMIS.Execute("SELECT ACCTNAME AS CustomerName,CUSCOMP,id FROM ALL_CUSTOMER WHERE CUSCOMP LIKE '" & XXX & "%' ORDER BY  CUSCOMP ASC")
    ElseIf optSearchKeyLast.Value = True Then
        Set rsCustomer2 = gconDMIS.Execute("SELECT ACCTNAME AS CustomerName,LastName,id FROM ALL_CUSTOMER WHERE LastName LIKE '" & XXX & "%' ORDER BY Lastname ASC")
    ElseIf optSearchKeyEmail.Value = True Then
        Set rsCustomer2 = gconDMIS.Execute("SELECT ACCTNAME AS CustomerName,Email,id FROM ALL_CUSTOMER WHERE Email LIKE '" & XXX & "%' ORDER BY Email ASC")
    ElseIf Option1.Value = True Then
        Set rsCustomer2 = gconDMIS.Execute("SELECT ACCTNAME AS CustomerName,CUSCDE,id FROM ALL_CUSTOMER WHERE CUSCDE LIKE '" & XXX & "%' ORDER BY CUSCDE ASC")
    End If

    If Not (rsCustomer2.EOF Or rsCustomer2.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsCustomer2
        lstCustomer.Enabled = True
        lstCustomer.Refresh
    End If
    '    Dim Key                             As String
    '    Dim LIMITKEY                        As String

    '    Select Case cboSearchCustype.ListIndex
    '        Case 0                                               'Search All
    '            LIMITKEY = "AND CUSTYPE IN ('P','C','F','G', NULL)"
    '        Case 1                                               'Only Personal Customers
    '            LIMITKEY = "AND CUSTYPE IN ('P', NULL)"
    '        Case 2                                               ' Only Company/Agency Customers
    '            LIMITKEY = "'C', NULL"
    '        Case 3                                               'Only Government Customer
    '            LIMITKEY = "'G', NULL"
    '        Case 4                                               'Only Fleet Account Customer
    '            LIMITKEY = "'F', NULL"
    '    End Select

End Sub

Sub InitData()
    Combo_Loadval cboPersonalCity, gconDMIS.Execute("SELECT Distinct CITY FROM ALL_CUSTOMER WHERE CITY IS NOT NULL")
    With cboCustType
        .AddItem ("Personal")
        .AddItem ("Company/Agency")
        .AddItem ("Government")
        .AddItem ("Fleet Account")
        .ListIndex = 0
    End With
    With cboSearchCustype
        .AddItem ("Search All")
        .AddItem ("Individual")
        .AddItem ("Company/Agency")
        .AddItem ("Government")
        .AddItem ("Fleet")
        .ListIndex = 0
    End With
End Sub

Sub initMemvars()
    Dim temprs                                                      As ADODB.Recordset
    
    txtSearch.Text = vbNullString
    labSEQ.Caption = gconDMIS.Execute("SELECT ISNULL(MAX(ID),0) FROM ALL_CUSTOMER").Collect(0)
    labCustCode = ""
    labCustCode2 = ""
    txtCuscde.Text = ""
    txtLastName.Text = ""
    txtFirstName.Text = ""
    txtMiddleName.Text = ""
    txtAcctName.Text = ""
    cboLeadSource.Text = ""
    cboSex.Text = ""
    txtTitle.Text = ""
    txtDepartment.Text = ""
    txtEmail.Text = ""
    txtCusphon1.Text = ""
    txtMobile.Text = ""
    txtHomePhone.Text = ""
    txtFax.Text = ""
    txtAssistant.Text = ""
    txtAsstPhone.Text = ""
    txtPersonalStreet.Text = ""
    cboPersonalCity.Text = ""
    txtPersonalState.Text = ""
    txtPersonalZIP.Text = ""
    txtBirthDate.Text = ""
    txtSpouse.Text = ""
    txtNotes.Text = ""
    cboApod.Clear
    cboCustType.ListIndex = -1
    txtTin = ""
    txtDeliveryAddress = ""
    txtCreditDays = "0"
    txtCreditLimit = "0.00"
    chkWithholdingTax.Value = 0

    Dim rsAPOD                                                      As ADODB.Recordset
    Set rsAPOD = New ADODB.Recordset
    Set rsAPOD = gconDMIS.Execute("SELECT Distinct apod FROM ALL_CustMaster_Smis WHERE APOD IS NOT NULL")

    If Not rsAPOD.EOF And Not rsAPOD.BOF Then
        rsAPOD.MoveFirst
        Do While Not rsAPOD.EOF
            cboApod.AddItem Null2String(rsAPOD!APOD)
            rsAPOD.MoveNext
        Loop
    End If
    Set rsAPOD = Nothing
    
    Set temprs = gconDMIS.Execute("SELECT DataDesc FROM CRIS_vW_MasterPullDown WHERE Masterdesc = 'Lead Source'")
    cboLeadSource.Clear
    While Not temprs.EOF
        cboLeadSource.AddItem temprs.Collect(0)
        temprs.MoveNext
    Wend
    
    cboSex.Clear
    cboSex.AddItem "NA"
    cboSex.AddItem "M"
    cboSex.AddItem "F"
End Sub

Sub rsRefresh()
    Set RS = New ADODB.Recordset
    RS.Open "SELECT * FROM ALL_Customer ORDER BY id DESC", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub SetCustomerAccountName()
'Or AddorEdit = "EDIT"
    If EntryPoint = "PROSPECT" Or AddorEdit = "" Then: Exit Sub
    txtAcctName = UCase(txtLastName & IIf(txtFirstName = "", "", ", ") & txtFirstName & IIf(txtMiddleName = "", "", " ") & Left(txtMiddleName, 1)) & IIf(txtMiddleName = "", "", ".")
End Sub

Sub StoreMemVars()
    If Not RS.EOF And Not RS.BOF Then
        labid.Caption = RS!ID
        cboApod.Text = Null2String(RS!APOD)
        txtCuscde.Text = Null2String(RS!CUSCDE)
        labCustCode = Null2String(RS!CUSCDE)
        labCustCode2 = Null2String(RS!CUSCDE)
        txtLastName.Text = Null2String(RS!lastname)
        txtFirstName.Text = Null2String(RS!FIRSTNAME)
        txtMiddleName.Text = Null2String(RS!MiddleInitial)

        txtTin = Null2String(RS!TIN)

        cboSex.Text = Null2String(RS!Sex)
        txtAcctName = Null2String(RS!AcctName)

        txtPersonalStreet.Text = Null2String(RS!CUSTOMERADD)
        txtPersonalState.Text = Null2String(RS!provincialadd)
        txtPersonalZIP.Text = Null2String(RS!ZIPCODE)
        txtCusphon1.Text = Null2String(RS!TelephoneNo)

        cboLeadSource.Text = Null2String(RS!LeadSource)
        txtTitle.Text = Null2String(RS!TITLE)
        txtDepartment.Text = Null2String(RS!Department)
        txtEmail.Text = Null2String(RS!EMAIL)
        txtMobile.Text = Null2String(RS!Mobile)
        txtHomePhone.Text = Null2String(RS!HomePhone)
        txtFax.Text = Null2String(RS!Fax)
        txtAssistant.Text = Null2String(RS!Assistant)
        txtAsstPhone.Text = Null2String(RS!AsstPhone)
        cboPersonalCity.Text = Null2String(RS!CITY)
        txtBirthDate.Text = Null2String(RS!BirthDate)
        txtSpouse.Text = Null2String(RS!Spouse)
        txtDeliveryAddress = Null2String(RS!DELIVERYADDRESS)
        txtNotes.Text = Null2String(RS!DESCRIPTION)

        txtCreditDays = NumericVal(RS!CREDITDAYS)
        txtCreditLimit = FormatNumber(NumericVal(RS!CREDITLIMIT))
        '
        If Null2Bit(RS!TAX_AGENT) = 1 Then
            chkWithholdingTax.Value = 1
        Else
            chkWithholdingTax.Value = 0
        End If
        '*****************************************
        
        '        If COMPANY_CODE = "HAI" Then
        '            txtLoyalID = Null2String(RS!Loyalty_ID)
        '
        '            If (RS!Loyalty_ID) <> "" Then
        '               lblloyal.Visible = True
        '               Label11.Visible = True
        '            Else
        '               lblloyal.Visible = False
        '               Label11.Visible = False
        '            End If
        '        End If

        If Null2String(RS!CUSTYPE) = "P" Then
            cboCustType.ListIndex = 0
        ElseIf Null2String(RS!CUSTYPE) = "C" Then
            cboCustType.ListIndex = 1
        ElseIf Null2String(RS!CUSTYPE) = "G" Then
            cboCustType.ListIndex = 2
        ElseIf Null2String(RS!CUSTYPE) = "F" Then
            cboCustType.ListIndex = 3
        Else
            cboCustType.ListIndex = -1
        End If

        If Null2String(RS!CUSCAT) = "Z" Then
            chkZeroRated.Value = 1
        Else
            chkZeroRated.Value = 0
        End If
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub ShowPictureBox(cntl As Object, State As Boolean, Optional ByVal MasterObject As Object)
    cntl.Visible = State
    If Not (MasterObject Is Nothing) Then
        MasterObject.Enabled = Not State
    End If
    
    If State = True Then
        cntl.ZOrder 0
    Else
        cntl.ZOrder 1
    End If
End Sub

Friend Sub AddEditCustomer(xAcCode As String)
    AccountCode = xAcCode
End Sub

Private Sub cboApod_KeyPress(KeyAscii As Integer)
    UpperAscii KeyAscii
End Sub

Private Sub cboCustType_Click()
    Select Case cboCustType.Text
    Case "Personal"
        Label1.Caption = "Last Name"
        Label2.Visible = True
        Label3.Visible = True
        txtLastName.Width = 2625
        txtFirstName.Visible = True
        txtMiddleName.Visible = True
        Label7.Caption = "Birth Date"
        Label24.Caption = "Spouse Name"
        CustType = "P"
        
        cmdCustInfo_Child.Enabled = True
        labCustInfo_Child.Enabled = True
        chkWithholdingTax.Visible = False
        Label32(0).Visible = False
    Case "Company/Agency"
        Label7.Caption = "Est Date"
        Label2.Visible = False
        Label3.Visible = False
        txtLastName.Width = 8115
        txtFirstName.Visible = False
        txtMiddleName.Visible = False
        Label1.Caption = "Company Name"
        Label24.Caption = "Contact Person"
        CustType = "C"
        cmdCustInfo_Child.Enabled = False
        labCustInfo_Child.Enabled = False
        chkWithholdingTax.Visible = True
        Label32(0).Visible = True
    Case "Fleet Account"
        CustType = "F"
        Label7.Caption = "Est Date"
        Label2.Visible = False
        Label3.Visible = False
        txtLastName.Width = 8115
        txtFirstName.Visible = False
        txtMiddleName.Visible = False
        Label1.Caption = "Company Name"
        Label24.Caption = "Contact Person"
        cmdCustInfo_Child.Enabled = False
        labCustInfo_Child.Enabled = False
        chkWithholdingTax.Visible = True
        Label32(0).Visible = True
    Case "Government"
        Label7.Caption = "Est Date"
        Label2.Visible = False
        Label3.Visible = False
        txtLastName.Width = 8115
        txtFirstName.Visible = False
        txtMiddleName.Visible = False
        Label1.Caption = "Establishment Name"
        Label24.Caption = "Contact Person"
        CustType = "G"
        cmdCustInfo_Child.Enabled = False
        labCustInfo_Child.Enabled = False
        chkWithholdingTax.Visible = False
        Label32(0).Visible = True
    End Select
End Sub

Private Sub cboCustType_LostFocus()
'    On Error GoTo ADDER:
'
'    Dim i
'    For i = 0 To cboCustType.ListCount - 1
'        If UCase(cboCustType.List(i)) = UCase(cboCustType.Text) Then
'            Exit Sub
'        End If
'    Next
'    cboCustType.ListIndex = -1
'    CustType = ""
'    cboCustType.ListIndex = -1
'    MsgBox "Please select proper Customer Type from the list", vbInformation
'    On Error Resume Next
'    cboCustType.SetFocus
'    Exit Sub
'
'ADDER:
'    ShowVBError
End Sub

Private Sub cboSearchCustype_Click()
    FillSearchGrid txtSearch
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "CUSTOMER") = False Then Exit Sub
    AddorEdit = "ADD"
    Frame1.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    picToolFrame.Enabled = True
    lstCustomer.Enabled = False
    txtSearch.Enabled = False
    initMemvars
    On Error Resume Next
End Sub

Private Sub cmdAddChildInfo_Click()
    cmdDeleteChild.Enabled = False
    txtChildDate = ""
    txtChildName = ""
    cboChildSex = ""
    labIdCHILD = ""
    ShowPictureBox picChildAE, True, picMain
    cmdAddChildInfo.Enabled = False
    cmdSelectChild.Enabled = False
    cmdCancelChildList.Enabled = False
    On Error Resume Next
    txtChildName.SetFocus
End Sub

Private Sub cmdCancel_Click()
    ShowPictureBox picChildList, False
    ShowPictureBox picChildAE, False, picMain
    Frame1.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    picToolFrame.Enabled = True
    lstCustomer.Enabled = True
    fraSearch.Enabled = True
    AddorEdit = ""
    txtSearch.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdCancelChildList_Click()
    ShowPictureBox picChildList, False, picMain
End Sub

Private Sub cmdCancelContactList_Click()
    ShowPictureBox picContactList, False, picMain
End Sub

Private Sub cmdCloseChildAE_Click(Index As Integer)
    ShowPictureBox picChildAE, False, picMain
    cmdCUSTINFO_Child_Click
End Sub

Private Sub cmdCloseContactsAE_Click(Index As Integer)
    ShowPictureBox picContactAE, False, picMain
    cmdAddContactList.Enabled = True
    cmdEditContact.Enabled = True
    cmdCancelContactList.Enabled = True
End Sub

Private Sub cmdCloseTerm_Click(Index As Integer)
    ShowPictureBox picCredit, False, picMain
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "CUSTOMER") = False Then Exit Sub
    On Error GoTo ErrorCode:
    Dim lng                                                         As Integer
    
    Load frmSplash
    Screen.MousePointer = 11
    frmSplash.labCon.Caption = "Checking Customer Record(s)... Please wait..."
    frmSplash.Show
    'IS PROSPECT
    lng = gconDMIS.Execute("SELECT COUNT(CUSCDE) FROM CRIS_PROSPECTS WHERE CUSCDE=" & N2Str2Null(txtCuscde)).Fields(0).Value
    If lng > 0 Then: MessagePop RecLocekd, "Record cannot be deleted", "Customer Information cannot be deleted. Prospect Information Exist": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
    'GOT RECEIPTS
    lng = gconDMIS.Execute("SELECT COUNT(CUSCDE) FROM cmis_off_hd WHERE CUSCDE=" & N2Str2Null(txtCuscde)).Fields(0).Value
    If lng > 0 Then: MessagePop RecLocekd, "Record cannot be deleted", "Customer Information cannot be deleted. Official Receipt exists for this Customer..": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
    'GOT VEHICLES
    lng = gconDMIS.Execute("SELECT COUNT(CUSCDE) FROM csms_cusveh WHERE CUSCDE=" & N2Str2Null(txtCuscde)).Fields(0).Value
    If lng > 0 Then: MessagePop RecLocekd, "Record cannot be deleted", "Customer Information cannot be deleted. Customer has a record for Service..": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
    'GOT Appointment
    lng = gconDMIS.Execute("SELECT COUNT(CUSCDE) FROM csms_appointment WHERE CUSCDE=" & N2Str2Null(txtCuscde)).Fields(0).Value
    If lng > 0 Then: MessagePop RecLocekd, "Record cannot be deleted", "Customer Information cannot be deleted. Customer has an appointment Information..": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
    'GOT PARTS TRANS
    lng = gconDMIS.Execute("SELECT COUNT(CUSTCODE) FROM pmis_ord_hist WHERE CUSTCODE=" & N2Str2Null(txtCuscde)).Fields(0).Value
    If lng > 0 Then: MessagePop RecLocekd, "Record cannot be deleted", "Customer Information cannot be deleted. Customer has a record for Parts Transactions.": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
    'GOT PARTS TRANS
    lng = gconDMIS.Execute("SELECT COUNT(CUSTCODE) FROM pmis_ord_hd WHERE CUSTCODE=" & N2Str2Null(txtCuscde)).Fields(0).Value
    If lng > 0 Then: MessagePop RecLocekd, "Record cannot be deleted", "Customer Information cannot be deleted. Customer is a Parts Customer and has Parts Transactions.": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
    'ACCOUNTING
    lng = gconDMIS.Execute("SELECT COUNT(CUSTCODE) FROM amis_openinvoice WHERE CUSTCODE=" & N2Str2Null(txtCuscde)).Fields(0).Value
    If lng > 0 Then: MessagePop RecLocekd, "Record cannot be deleted", "Customer Information cannot be deleted. Customer has a record Finance and Accounting.": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
    'SALES
    lng = gconDMIS.Execute("SELECT COUNT(CODE) FROM smis_salesorder WHERE CODE=" & N2Str2Null(txtCuscde)).Fields(0).Value
    If lng > 0 Then: MessagePop RecLocekd, "Record cannot be deleted", "Customer Information cannot be deleted. Customer has a Sales Record.": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
    'SERVICE
    lng = gconDMIS.Execute("SELECT COUNT(ACCT_NO) FROM CSMS_repairorder WHERE ACCT_NO=" & N2Str2Null(txtCuscde)).Fields(0).Value
    If lng > 0 Then: MessagePop RecLocekd, "Record cannot be deleted", "Customer Information cannot be deleted. Customer has a Service Record.": Screen.MousePointer = 0: Unload frmSplash: Exit Sub

    Unload frmSplash
    Screen.MousePointer = 0

    If ShowConfirmDelete = True Then
        Screen.MousePointer = 11
        SQL_STATEMENT = "DELETE FROM ALL_CUSTOMER WHERE ID=" & labid
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("X", "CUSTOMER", SQL_STATEMENT, labid, "", "CUST CODE: " & labCustCode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        SQL_STATEMENT = "DELETE FROM ALL_CUSTOMER_CONTACTS WHERE CUSCDE=" & N2Str2Null(txtCuscde)
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("X", "CUSTOMER", SQL_STATEMENT, labid, "", "CUST CODE: " & labCustCode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        SQL_STATEMENT = "DELETE FROM ALL_CUSTOMER_CHILD WHERE CUSCDE=" & N2Str2Null(txtCuscde)
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("X", "CUSTOMER", SQL_STATEMENT, labid, "", "CUST CODE: " & labCustCode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        gconDMIS.Execute "DELETE FROM ALL_CusCtl"

        Dim rsCustomer                                              As ADODB.Recordset
        Dim k                                                       As Integer
        Dim NewCtlCde                                               As String
        
        For k = 65 To 90
            Set rsCustomer = New ADODB.Recordset
            rsCustomer.Open "SELECT Code FROM ALL_CustMaster_Smis WHERE LEFT(Code,1) = '" & Chr(k) & "' ORDER BY Code DESC", gconDMIS
            If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCustomer!Code, 2, 5)) + 1, "00000")
                gconDMIS.Execute "INSERT INTO ALL_CusCtl (ctlcde,ctldsc) VALUES('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
            Else
                gconDMIS.Execute "INSERT INTO ALL_CusCtl (ctlcde,ctldsc) VALUES('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
            End If
        Next
        Screen.MousePointer = 0
        FillSearchGrid ""
        rsRefresh
        StoreMemVars
        MessagePop Delete, "Record Deleted", "Customer Information Deleted. "
        'LogAudit "X", "CUSTOMER MASTER FILE", labCustCode & " ACCOUNT NAME" & txtAcctName
    End If

    rsRefresh
    RS.Bookmark = rsFind(RS.Clone, "ID", labid).Bookmark
    initMemvars
    StoreMemVars
    Exit Sub
    
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdDeleteChild_Click()
    On Error GoTo ErrorCode:
    
    If MsgBox("Are you sure you want to delete this Information", vbQuestion + vbOKCancel, "Delete?") = vbCancel Then: Exit Sub
    
    SQL_STATEMENT = "DELETE FROM ALL_CUSTOMER_CHILD WHERE id = " & labIdCHILD
    gconDMIS.Execute SQL_STATEMENT
    
    Call NEW_LogAudit("XX", "CUSTOMER", SQL_STATEMENT, labid, "", "NAME: " & txtChildName & " - CHILD", "", labIdCHILD)
    ShowDeletedMsg
    
    'Updated by: Joanna J Enriquez 04/10/12 2:45PM
    'To refresh loading of List of Children after delete
    If picChildList.Visible = True Then
        cmdCUSTINFO_Child_Click
    End If
    
    ShowPictureBox picChildAE, False, picMain
    'LogAudit "XX", "CUSTOMER CHILD", labIdCHILD
    Exit Sub
    
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdDeleteContact_Click()
    On Error GoTo ErrorCode:
    
    If MsgBox("Are you sure you want to delete this Information", vbQuestion + vbOKCancel, "Delete?") = vbCancel Then: Exit Sub
    
    SQL_STATEMENT = "DELETE FROM ALL_CUSTOMER_CONTACTS WHERE id=" & labIDContacts
    gconDMIS.Execute SQL_STATEMENT

    Call NEW_LogAudit("XX", "CUSTOMER", SQL_STATEMENT, labid, "", "NAME: " & txtContactName & " - CONTACT", "", labIDContacts)
    ShowDeletedMsg
    
    If picContactList.Visible = True Then
        cmdCUSTINFO_Contact_Click
    End If
    ShowPictureBox picContactAE, False, picMain
    'LogAudit "X", "CUSTOMER CONTACT", labIDContacts
    Exit Sub
    
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdDuplicate_Click()
    On Error Resume Next
    
    If MODULENAME = "AMIS" Then
        If Module_Access(LOGID, "MERGE ACCOUNTS", "SYSTEM") = False Then Exit Sub
        frmCRIS_Inquiry_PossibleDuplicates.Show
    Else
        MessagePop InfoFriend, "Info", "This module is only supported in AMIS module."
    End If
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "CUSTOMER") = False Then Exit Sub
    
    '052716 | SJR
    'DESCRIPTION: EDITING NOT ALLOWED FOR EXISTING TRANSACTIONS. CONCERNS OCCUR BECAUSE OF THE HABITUAL EDITING OF THE CUSTOMER MASTER FILE
    'If CHECK_IF_HAS_TRANSACTION(labCustCode.Caption) = False Then
    '    MsgBox "Customer record can not be edited this time. Transactions have been made from other modules.", vbCritical, "Information"
    '    Exit Sub
    'End If
    
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    picToolFrame.Enabled = False
    lstCustomer.Enabled = False
    fraSearch.Enabled = False
    On Error Resume Next
    txtLastName.SetFocus
End Sub

Function CHECK_IF_HAS_TRANSACTION(xCUSTCODE As String)
    Dim rsHasTrans As ADODB.Recordset
    Set rsHasTrans = New ADODB.Recordset
    rsHasTrans.Open "SELECT * FROM " & _
                    "( " & _
                    "SELECT CUSTCODE FROM PMIS_ORD_HD " & _
                    "Union " & _
                    "(SELECT CUSTCODE FROM PMIS_ORD_HIST) " & _
                    "Union " & _
                    "(SELECT CUSTOMERCODE AS CUSTCODE FROM SMIS_MRRINV_TABLE) " & _
                    "Union " & _
                    "(SELECT ACCT_NO AS CUSTCODE FROM CSMS_REPOR) " & _
                    "Union " & _
                    "(SELECT CUSTOMERCODE AS CUSTCODE FROM AMIS_JOURNAL_HD) " & _
                    ")Y " & _
                    "WHERE CUSTCODE= '" & xCUSTCODE & "'", gconDMIS
    
    If Not rsHasTrans.EOF And rsHasTrans.BOF Then
        CHECK_IF_HAS_TRANSACTION = True
    Else
        CHECK_IF_HAS_TRANSACTION = False
    End If
End Function

Private Sub cmdEditContact_Click()
    lvContactList_KeyPress 13
End Sub

Private Sub cmdExit_Click()
    Unload Me
    'frmCSMSNewAppointment.Show vbModal
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    On Error Resume Next
    
    RS.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdLast_Click()
    On Error Resume Next
    
    RS.MoveLast
    StoreMemVars
End Sub

Private Sub CmdloyaltyCancel_Click()
    Call rsRefresh
    RS.Find "id = " & labid
    Call StoreMemVars

    PICLoyaltyID.ZOrder 1
    PICLoyaltyID.Enabled = False
    picAdds.Enabled = True
    picSaves.Enabled = True
    Frame1.Enabled = True
    picToolFrame.Enabled = True
    fraSearch.Enabled = True
End Sub

Private Sub CmdloyaltySave_Click()
    Dim rsLoyaltyID                                                 As New ADODB.Recordset
    Dim Lid                                                         As String
    Dim CTR                                                         As Long

    On Error GoTo error
    
    If MsgBox("Register this Loyalty ID to this customer, Are you sure?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Lid = N2Str2Null(txtLoyalID.Text)
    
    Set rsLoyaltyID = gconDMIS.Execute("SELECT id,loyalty_id FROM all_customer_table WHERE loyalty_id = " & Lid & "")
    If Not (rsLoyaltyID.BOF And rsLoyaltyID.EOF) Then
        If rsLoyaltyID!ID <> labid Then
            MessagePop RecSaveWarning, "Duplicate Record", "Loyalty ID already registered to another customer"
            Exit Sub
        End If
    End If
    '    If (RS!Loyalty_ID) <> (txtLoyalID.Text) Then
    '        CTR = gconDMIS.Execute("Select count(Loyalty_ID) from all_customer where " & _
             '            " Loyalty_ID = " & LID & "").Fields(0).Value
    '        If CTR >= "1" Then
    '            MessagePop RecSaveWarning, "Duplicate Record", "Loyalty ID Already Exist"
    '            On Error Resume Next
    '            txtLoyalID.SetFocus
    '            Exit Sub
    '        Else
    '            gconDMIS.Execute ("Update ALL_CUSTOMER set " & _
                 '                " Loyalty_ID = " & LID & _
                 '                " where ID = " & labID.Caption & "")
    '        End If
    '    Else
    '        CTR = gconDMIS.Execute("Select count(*) from all_customer where Loyalty_ID = " & LID & "").Fields(0).Value
    '        If CTR >= "1" Then
    '            MessagePop RecSaveWarning, "Duplicate Record", "Loyalty ID Already Exist"
    '            On Error Resume Next
    '            txtLoyalID.SetFocus
    '            Exit Sub
    '        Else
    '
    '        End If
    '
    '    End If
    gconDMIS.Execute ("Update ALL_CUSTOMER Set" & _
                      " Loyalty_ID = " & Lid & _
                      " where id = " & labid.Caption & "")

    MessagePop RecSaveOk, "Info", "Record Saved"

    Call FillSearchGrid(txtSearch)
    Call CmdloyaltyCancel_Click
    Exit Sub

error:
    Call ErrHandler(gconDMIS)
End Sub

Sub ErrHandler(objCon As Object)
    Dim ADOErr                                                      As ADODB.error
    Dim strError                                                    As String

    For Each ADOErr In objCon.Errors
        strError = "Error #: " & ADOErr.Number & vbCrLf & _
                   "Error Description : " & ADOErr.DESCRIPTION
    Next
    MsgBox strError, vbCritical, "Error"
    objCon.Errors.Clear
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    
    RS.MoveNext
    If RS.EOF Then
        RS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    
    RS.MovePrevious
    If RS.BOF Then
        RS.MoveFirst
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "CUSTOMER") = False Then Exit Sub
    CrystalReport1.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    CrystalReport1.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport CrystalReport1, SMIS_REPORT_PATH & "Customers.rpt", "", DMIS_REPORT_Connection, 1
    'PrintSQLReport CrystalReport1, SMIS_REPORT_PATH & "Customer_ENTRYDATE.rpt", "{CRIS_VW_ALLPROFILE.ENTRY_DATE} >= #" & LOGDATE & "# AND {CRIS_VW_ALLPROFILE.ENTRY_DATE} <= #" & LOGDATE & "#", DMIS_REPORT_Connection, 1

    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("V", "CUSTOMER", "", labid, "", "CUST CODE: " & labCustCode, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
    'LogAudit "V", "CUSTOMER MASTER FILE"
    Exit Sub
    
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode:

    Dim vtxtCusCde                                                  As String
    Dim VTXTLASTNAME                                                As String
    Dim VTXTFIRSTNAME                                               As String
    Dim vtxtMiddleInitial                                           As String
    Dim vtxtCUSCOMP                                                 As String

    Dim vcboSex                                                     As String
    Dim vtxtCusadd1                                                 As String
    Dim vtxtCusadd2                                                 As String
    Dim vtxtCuszipc                                                 As String
    Dim vtxtCusphon1                                                As String
    Dim vtxtAcctName                                                As String
    Dim vcboApod                                                    As String
    Dim vcboLeadSource                                              As String
    Dim vtxtTitle                                                   As String
    Dim vtxtDepartment                                              As String
    Dim vtxtEmail                                                   As String
    Dim vtxtMobile                                                  As String
    Dim vtxtHomePhone                                               As String
    Dim VtxtFax                                                     As String
    Dim vtxtAssistant                                               As String
    Dim vtxtAsstPhone                                               As String
    Dim vtxtCity                                                    As String
    Dim vTxtBirthDate                                               As String
    Dim vTxtSpouse                                                  As String
    Dim vtxtDescription                                             As String
    Dim vtxtCustType                                                As String
    Dim vtxtCompanyAdd                                              As String
    Dim TEMPSQL                                                     As String
    Dim vtxtDeliveryAddress                                         As String
    Dim VtxtTIN                                                     As String

    If LTrim(RTrim(txtAcctName)) = "" Then
        ShowIsRequiredMsg "Account Name"
        On Error Resume Next
        txtAcctName.SetFocus
        Exit Sub
    End If
    If CustType = "P" Then
        If LTrim(RTrim(txtFirstName)) = "" Then
            ShowIsRequiredMsg "First Name"
            On Error Resume Next
            txtFirstName.SetFocus
            Exit Sub
        End If
        If LTrim(RTrim(txtLastName)) = "" Then
            ShowIsRequiredMsg "Last Name"
            On Error Resume Next
            txtLastName.SetFocus
            Exit Sub
        End If
    Else
        If LTrim(RTrim(txtLastName)) = "" Then
            ShowIsRequiredMsg "Company Name"
            On Error Resume Next
            txtLastName.SetFocus
            Exit Sub
        End If
        If LTrim(RTrim(txtTin)) = "" Then
            ShowIsRequiredMsg "TIN"
            On Error Resume Next
            txtTin.SetFocus
            Exit Sub
        End If
    End If
    
    If LTrim(RTrim(txtPersonalStreet)) = "" Then
        ShowIsRequiredMsg "Street"
        On Error Resume Next
        txtPersonalStreet.SetFocus
        Exit Sub
    End If
    If LTrim(RTrim(cboPersonalCity)) = "" Then
        ShowIsRequiredMsg "City"
        On Error Resume Next
        cboPersonalCity.SetFocus
        Exit Sub
    End If
    If txtEmail.Text = "" And txtCusphon1.Text = "" And txtMobile.Text = "" And txtHomePhone.Text = "" And txtFax.Text = "" Then
        ShowIsRequiredMsg Chr(13) & "Email, Office Phone, Mobile, Home Phone and Fax" & Chr(13) & "at least one of this field."
        On Error Resume Next
        txtMobile.SetFocus
        Exit Sub
    End If
    
    'JJE MITSUBISHI ONLY WALA SA HYUNDAI
    If LTrim(RTrim(cboCustType)) = "" Then
        ShowIsRequiredMsg "Account Type"
        On Error Resume Next
        cboCustType.SetFocus
        Exit Sub
    End If
    'JJE
    
    'Update By: BTT -4142008
    If Len(txtNotes.Text) > 250 Then
        MsgBox "Number of character exceed! Please simplify your notes.", vbInformation, "Warning!"
        txtNotes.SetFocus
        txtNotes.BackColor = &HFFFFC0
        Exit Sub
    Else
        txtNotes.BackColor = vbWhite
    End If
    
    If AddorEdit = "ADD" Then
        Dim rsfindDup                                               As ADODB.Recordset
        Set rsfindDup = New ADODB.Recordset
        rsfindDup.Open "SELECT * FROM ALL_CUSTOMER_TABLE WHERE CUSCDE = '" & txtCuscde & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsfindDup.EOF And Not rsfindDup.BOF Then
            MsgSpeechBox "Code already exist!"
            Exit Sub
        End If
        txtCuscde = GetCustomerCode(LTrim(RTrim(txtLastName)))
    End If

    Dim tmpcount                                                    As Integer
    Dim rsCheck                                                     As ADODB.Recordset

    If AddorEdit = "ADD" Then
        tmpcount = 0
        Set rsCheck = gconDMIS.Execute("SELECT COUNT(*) TCOUNT FROM all_customer WHERE REPLACE(LTRIM(RTRIM(ISNULL(REPLACE(ACCTNAME,' ',''),''))),',','') = " & N2Str2Null(LTrim(RTrim(Replace(Replace(txtAcctName, " ", ""), ",", "")))))
        tmpcount = rsCheck!TCOUNT
        If tmpcount > 0 Then
            If MsgBox("Customer with similar Account Name exists!" & vbCrLf & "Do you want to continue saving this Information?", vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        tmpcount = 0
          
        Set rsCheck = gconDMIS.Execute("SELECT COUNT(*) TCOUNT FROM all_customer WHERE REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(LASTNAME,''))),' ',''),',','') + REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(FIRSTNAME,''))),' ',''),',','')= " & N2Str2Null(Replace(Replace(LTrim(RTrim(txtLastName)) & LTrim(RTrim(txtFirstName)), ",", ""), " ", "")))
        tmpcount = rsCheck!TCOUNT
        If tmpcount > 0 Then
            If MsgBox("Customer with similar Last Name and First Name exists!" & vbCrLf & "Do you want to continue saving this Information?", vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
    End If

    VtxtTIN = N2Str2Null(txtTin)
    vtxtCompanyAdd = N2Str2Null(UCase(txtPersonalStreet))
    vtxtCustType = N2Str2Null(CustType)
    vcboApod = N2Str2Null(UCase(cboApod))
    vtxtCusCde = N2Str2Null(RTrim(LTrim(txtCuscde)))
    VTXTLASTNAME = N2Str2Null((txtLastName))
    VTXTFIRSTNAME = N2Str2Null((txtFirstName))
    vtxtMiddleInitial = N2Str2Null(txtMiddleName)
    vtxtAcctName = N2Str2Null(txtAcctName)
    vtxtCUSCOMP = N2Str2Null((txtLastName))
    vcboSex = N2Str2Null(cboSex)
    vtxtCusadd1 = N2Str2Null(Trim(UCase(txtPersonalStreet)))
    vtxtCusadd2 = N2Str2Null(UCase(txtPersonalState))
    vtxtCity = N2Str2Null(UCase(cboPersonalCity))
    vtxtCuszipc = N2Str2Null(txtPersonalZIP)
    vtxtCusphon1 = N2Str2Null(txtCusphon1)

    vcboLeadSource = N2Str2Null(cboLeadSource)
    vtxtTitle = N2Str2Null(UCase(txtTitle))
    vtxtDepartment = N2Str2Null(UCase(txtDepartment))
    vtxtEmail = N2Str2Null(txtEmail)
    vtxtMobile = N2Str2Null(txtMobile)
    vtxtHomePhone = N2Str2Null(txtHomePhone)
    VtxtFax = N2Str2Null(txtFax)
    vtxtAssistant = N2Str2Null(txtAssistant)
    vtxtAsstPhone = N2Str2Null(txtAsstPhone)
    
    vTxtBirthDate = N2Str2Null(txtBirthDate)
    vTxtSpouse = N2Str2Null(txtSpouse)
    vtxtDescription = N2Str2Null(txtNotes)
    vtxtDeliveryAddress = N2Str2Null(txtDeliveryAddress)

    If AddorEdit = "ADD" Then
        TEMPSQL = "INSERT INTO ALL_CUSTOMER(" & vbCrLf
        TEMPSQL = TEMPSQL & " TIN,CUSCOMP,APOD,CUSCDE,LASTNAME,FIRSTNAME,MIDDLEINITIAL,ACCTNAME,SEX,CUSTOMERADD,PROVINCIALADD,ZIPCODE,TELEPHONENO,LEADSOURCE,TITLE,DEPARTMENT,EMAIL,MOBILE,HOMEPHONE,FAX,ASSISTANT,ASSTPHONE,CITY,BIRTHDATE,SPOUSE,DESCRIPTION,CUSTYPE,COMPANYADD,DELIVERYADDRESS,Tax_Agent,USERCODE " & vbCrLf
        TEMPSQL = TEMPSQL & " ) VALUES ( " & vbCrLf
        TEMPSQL = TEMPSQL & VtxtTIN & ", "
        TEMPSQL = TEMPSQL & vtxtCUSCOMP & ", "
        TEMPSQL = TEMPSQL & vcboApod & ","
        TEMPSQL = TEMPSQL & vtxtCusCde & ", "
        TEMPSQL = TEMPSQL & VTXTLASTNAME & ", "
        TEMPSQL = TEMPSQL & VTXTFIRSTNAME & ", "
        TEMPSQL = TEMPSQL & vtxtMiddleInitial & ", "
        TEMPSQL = TEMPSQL & vtxtAcctName & ","
        TEMPSQL = TEMPSQL & vcboSex & "," & vbCrLf
        TEMPSQL = TEMPSQL & vtxtCusadd1 & ", "
        TEMPSQL = TEMPSQL & vtxtCusadd2 & ", "
        TEMPSQL = TEMPSQL & vtxtCuszipc & ", "
        TEMPSQL = TEMPSQL & vtxtCusphon1 & ","
        TEMPSQL = TEMPSQL & vcboLeadSource & ","
        TEMPSQL = TEMPSQL & vtxtTitle & ","
        TEMPSQL = TEMPSQL & vtxtDepartment & ","
        TEMPSQL = TEMPSQL & vtxtEmail & ","
        TEMPSQL = TEMPSQL & vtxtMobile & ","
        TEMPSQL = TEMPSQL & vtxtHomePhone & ","
        TEMPSQL = TEMPSQL & VtxtFax & ","
        TEMPSQL = TEMPSQL & vtxtAssistant & ","
        TEMPSQL = TEMPSQL & vtxtAsstPhone & ","
        TEMPSQL = TEMPSQL & vtxtCity & ","
        TEMPSQL = TEMPSQL & vTxtBirthDate & ","
        TEMPSQL = TEMPSQL & vTxtSpouse & ","
        TEMPSQL = TEMPSQL & vtxtDescription & ","
        TEMPSQL = TEMPSQL & vtxtCustType & ","
        TEMPSQL = TEMPSQL & vtxtCompanyAdd & ","
        TEMPSQL = TEMPSQL & vtxtDeliveryAddress & ","
        TEMPSQL = TEMPSQL & chkWithholdingTax.Value & ","
        TEMPSQL = TEMPSQL & N2Str2Null(LOGCODE)
        TEMPSQL = TEMPSQL & ")"

        gconDMIS.Execute TEMPSQL
        SQL_STATEMENT = TEMPSQL
        
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("A", "CUSTOMER", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtCuscde), "CUSCDE", "ALL_CUSTOMER"), "", "CUST CODE: " & Null2String(vtxtCusCde), "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        'LogAudit "A", "CUSTOMER MASTER FILE", labCustCode & " ACCOUNT NAME" & txtAcctName
    Else
        TEMPSQL = "UPDATE ALL_CUSTOMER SET" & vbCrLf
        TEMPSQL = TEMPSQL & " CUSCOMP = " & vtxtCUSCOMP & "," & vbCrLf
        TEMPSQL = TEMPSQL & " TIN = " & VtxtTIN & "," & vbCrLf
        TEMPSQL = TEMPSQL & " COMPANYADD = " & vtxtCompanyAdd & "," & vbCrLf
        TEMPSQL = TEMPSQL & " APOD = " & vcboApod & "," & vbCrLf
        TEMPSQL = TEMPSQL & " LASTNAME = " & VTXTLASTNAME & "," & vbCrLf
        TEMPSQL = TEMPSQL & " FIRSTNAME = " & VTXTFIRSTNAME & "," & vbCrLf
        TEMPSQL = TEMPSQL & " MIDDLEINITIAL = " & vtxtMiddleInitial & "," & vbCrLf
        TEMPSQL = TEMPSQL & " ACCTNAME = " & vtxtAcctName & "," & vbCrLf
        TEMPSQL = TEMPSQL & " SEX = " & vcboSex & "," & vbCrLf
        TEMPSQL = TEMPSQL & " CUSTOMERADD = " & vtxtCusadd1 & "," & vbCrLf
        TEMPSQL = TEMPSQL & " PROVINCIALADD = " & vtxtCusadd2 & "," & vbCrLf
        TEMPSQL = TEMPSQL & " ZIPCODE = " & vtxtCuszipc & "," & vbCrLf
        TEMPSQL = TEMPSQL & " LEADSOURCE = " & vcboLeadSource & "," & vbCrLf
        TEMPSQL = TEMPSQL & " TITLE = " & vtxtTitle & "," & vbCrLf
        TEMPSQL = TEMPSQL & " DEPARTMENT = " & vtxtDepartment & "," & vbCrLf
        TEMPSQL = TEMPSQL & " EMAIL = " & vtxtEmail & "," & vbCrLf
        TEMPSQL = TEMPSQL & " MOBILE = " & vtxtMobile & "," & vbCrLf
        TEMPSQL = TEMPSQL & " TELEPHONENO  = " & vtxtCusphon1 & "," & vbCrLf
        TEMPSQL = TEMPSQL & " HOMEPHONE = " & vtxtHomePhone & "," & vbCrLf
        TEMPSQL = TEMPSQL & " FAX = " & VtxtFax & "," & vbCrLf
        TEMPSQL = TEMPSQL & " ASSISTANT = " & vtxtAssistant & "," & vbCrLf
        TEMPSQL = TEMPSQL & " ASSTPHONE = " & vtxtAsstPhone & "," & vbCrLf
        TEMPSQL = TEMPSQL & " CITY = " & vtxtCity & "," & vbCrLf
        TEMPSQL = TEMPSQL & " BIRTHDATE = " & vTxtBirthDate & "," & vbCrLf
        TEMPSQL = TEMPSQL & " SPOUSE = " & vTxtSpouse & "," & vbCrLf
        TEMPSQL = TEMPSQL & " CUSTYPE = " & vtxtCustType & "," & vbCrLf
        TEMPSQL = TEMPSQL & " DESCRIPTION = " & vtxtDescription & "," & vbCrLf
        TEMPSQL = TEMPSQL & " LASTUPDATE = '" & LOGDATE & "'," & vbCrLf
        TEMPSQL = TEMPSQL & " TIMEUPDATE = '" & LOGTIME & "'," & vbCrLf
        TEMPSQL = TEMPSQL & " DeliveryAddress = " & vtxtDeliveryAddress & "," & vbCrLf
        TEMPSQL = TEMPSQL & " Tax_Agent = '" & chkWithholdingTax.Value & "'," & vbCrLf
        TEMPSQL = TEMPSQL & " USERCODE = '" & LOGCODE & "'" & vbCrLf
        TEMPSQL = TEMPSQL & " WHERE CUSCDE = '" & txtCuscde & "'" & vbCrLf
        
        gconDMIS.Execute TEMPSQL
        SQL_STATEMENT = TEMPSQL
        
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "CUSTOMER", SQL_STATEMENT, labid, "", "CUST CODE: " & labCustCode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        'LogAudit "E", "CUSTOMER MASTER FILE", labCustCode & " ACCOUNT NAME" & txtAcctName
    End If

    Screen.MousePointer = 11
    gconDMIS.Execute "Delete from ALL_CusCtl"
    
    '072216
    'DESCRIPTION: UPDATE AMIS_AR CUSTOMERNAME
    If Add_o_Edit = "ADD" Then
    Else
        Dim RSGETFULLNAME As New ADODB.Recordset
        Set RSGETFULLNAME = New ADODB.Recordset
        Set RSGETFULLNAME = gconDMIS.Execute("SELECT ACCTNAME FROM ALL_CUSTOMER_TABLE WHERE CUSCDE='" & labCustCode.Caption & "'")
        gconDMIS.Execute ("UPDATE AMIS_AR SET CUSTOMERNAME= '" & RSGETFULLNAME!AcctName & "' WHERE CUSTOMERCODE = '" & labCustCode.Caption & "'")
        Set RSGETFULLNAME = Nothing
    End If
    
    Dim rsCustomer                                                  As ADODB.Recordset
    Dim k                                                           As Integer
    Dim NewCtlCde                                                   As String
    
    For k = 65 To 90
        Set rsCustomer = New ADODB.Recordset
        rsCustomer.Open "SELECT Code FROM ALL_CustMaster_Smis WHERE LEFT(Code,1) = '" & Chr(k) & "' ORDER BY Code DESC", gconDMIS
        If Not rsCustomer.EOF And Not rsCustomer.BOF Then
            NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCustomer!Code, 2, 5)) + 1, "00000")
            gconDMIS.Execute "INSERT INTO ALL_CusCtl (ctlcde,ctldsc) VALUES('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
        Else
            gconDMIS.Execute "INSERT INTO ALL_CusCtl (ctlcde,ctldsc) VALUES('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
        End If
    Next
    Screen.MousePointer = 0

    If EntryPoint = "PROSPECT" Then
        SQL_STATEMENT = "UPDATE CRIS_PROSPECTS SET PROSPECTTYPE=" & vtxtCustType & " WHERE PROSPECTID= " & TempProspectID
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "PROSPECT", SQL_STATEMENT, N2Str2Null(TempProspectID), "", "CUST CODE: " & labCustCode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        RaiseEvent ProspectConverted(Replace(vtxtCusCde, "'", ""), GoingWhere, TempProspectID)
        Screen.MousePointer = 0
    Else
        RaiseEvent ChangedData(Replace(vtxtCusCde, "'", ""))
        Screen.MousePointer = 0
        MessagePop RecSave, "Record Saved", "Customer Information Saved"

        RS.Requery
        If AddorEdit = "EDIT" Then
            RS.Find "id =" & labid
        End If
        FillSearchGrid txtSearch
    End If
    cmdCancel.Value = True
    '
    '    If MsgBox("Do you want to add Loyalty ID?", vbInformation + vbYesNo) = vbYes Then
    '
    '            PICLoyaltyID.ZOrder 0
    '            PICLoyaltyID.Enabled = True
    '            picAdds.Enabled = False
    '            picSaves.Enabled = False
    '            Frame1.Enabled = False
    '            picToolFrame.Enabled = False
    '            fraSearch.Enabled = False
    '            Exit Sub
    '    Else
    '            Exit Sub
    '    End If
    Exit Sub
    
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSaveContact_Click()
    On Error GoTo ErrorCode:
    
    If RTrim(LTrim(txtContactName)) = "" Then
        ShowIsRequiredMsg "CONTACT NAME"
        On Error Resume Next
        txtContactName.SetFocus
        Exit Sub
    End If
    
    Dim vtxtCusCde, ContactName, Relation, ContactPosition, Department, Phone, Mobile, Address, SQL

    vtxtCusCde = N2Str2Null(txtCuscde)
    ContactName = N2Str2Null(txtContactName)
    Relation = N2Str2Null(cboContactRelation)
    ContactPosition = N2Str2Null(txtContactPosition)
    Department = N2Str2Null(txtContactDepartment)
    Phone = N2Str2Null(txtContactPhone)
    Mobile = N2Str2Null(txtContactMobile)
    Address = N2Str2Null(txtContactAddress)

    If NumericVal(labIDContacts) = 0 Then
        SQL = "INSERT INTO ALL_CUSTOMER_CONTACTS "
        SQL = SQL & "(ContactName,CUSCDE,Relation,ContactPosition,Department,Phone,Mobile,Address) VALUES ("
        SQL = SQL & ContactName & " ,"
        SQL = SQL & vtxtCusCde & " ,"
        SQL = SQL & Relation & " ,"
        SQL = SQL & ContactPosition & " ,"
        SQL = SQL & Department & " ,"
        SQL = SQL & Phone & " ,"
        SQL = SQL & Mobile & " ,"
        SQL = SQL & Address & " )"

        gconDMIS.Execute SQL
        SQL_STATEMENT = SQL

        Call NEW_LogAudit("AA", "CUSTOMER", SQL_STATEMENT, labid, "", "NAME: " & txtContactName & " - CONTACTS", "", FindTransactionID(N2Str2Null(labCustCode), "CUSCDE", "ALL_CUSTOMER_CONTACTS", "DETAILS", N2Str2Null(txtContactName), "CONTACTNAME"))
        ShowSuccessFullyAdded
        'LogAudit "A", "CONTACTS INFORMAION"
    Else
        SQL = "UPDATE ALL_CUSTOMER_CONTACTS SET "
        SQL = SQL & " ContactName =" & ContactName & ", "
        SQL = SQL & " Relation =" & Relation & ", "
        SQL = SQL & " ContactPosition =" & ContactPosition & ", "
        SQL = SQL & " Department =" & Department & ", "
        SQL = SQL & " Phone =" & Phone & ", "
        SQL = SQL & " Address =" & Address & ", "
        SQL = SQL & " Mobile =" & Mobile
        SQL = SQL & " WHERE id=" & labIDContacts
        'LogAudit "E", "CONTACTS INFORMAION"

        gconDMIS.Execute SQL
        SQL_STATEMENT = SQL

        Call NEW_LogAudit("EE", "CUSTOMER", SQL_STATEMENT, labid, "", "NAME: " & txtContactName & " - CONTACTS", "", labIDContacts)
        ShowSuccessFullyUpdated
    End If

    'gconDMIS.Execute SQL

    If picContactList.Visible = True Then
        cmdCUSTINFO_Contact_Click
    End If
    ShowPictureBox picContactAE, False, picMain
    Exit Sub
    
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSelectChild_Click()
    lvChildList_KeyPress 13
End Sub

Private Sub cmdCUSTINFO_Child_Click()
'Updated By Joanna Enriquez 04/13/2012 11:04am
'Child Info Data Entry will be added in listview
'Child's Age will be computed
    Dim temprs                                                      As ADODB.Recordset
    Dim Item                                                        As ListItem
       
    Set temprs = gconDMIS.Execute("SELECT CHILDNAME,SEX,DOB,DATEDIFF(YEAR,DOB,GETDATE()) AS AGE,ID FROM ALL_CUSTOMER_CHILD WHERE CUSCDE=" & N2Str2Null(txtCuscde))
    lvChildList.ListItems.Clear
    Listview_Loadval lvChildList.ListItems, temprs
    
'    If Not (temprs.EOF And temprs.BOF) Then
'        Do While Not temprs.EOF
'            Set Item = lvChildList.ListItems.Add(, , Null2String(temprs!CHILDNAME))
'            Item.SubItems(1) = Null2String(temprs!Sex)
'            Item.SubItems(2) = Null2String(temprs!DOB)
'            Item.SubItems(3) = Null2String(temprs!AGE)
'            Item.SubItems(4) = temprs!Id
'            temprs.MoveNext
'        Loop
'    End If
    
'    Set temprs = Nothing
    lvChildList.Refresh
    
        '''Listview_Loadval lvChildList.ListItems, temprs
        '''lvChildList.Refresh

    ShowPictureBox picChildList, True, picMain
    cmdAddChildInfo.Enabled = True
    cmdSelectChild.Enabled = True
    cmdCancelChildList.Enabled = True
   
    If lvChildList.ListItems.Count = 0 Then
        cmdSelectChild.Enabled = False
    Else
        cmdSelectChild.Enabled = True
    End If
End Sub

Private Sub cmdServiceInternal_Click()
    If COMPANY_CODE <> "HLP" Or COMPANY_CODE <> "HAM" Or COMPANY_CODE <> "HSP" Then
        MessagePop InfoFriend, "Module Info.", "This module is not supported by your dealer. For more information Kindly contact Netspeed Software Inc. about this Module"
        Exit Sub
    End If

    If Module_Access(LOGID, "INPUT LOYALTY NO", "SYSTEM") = False Then Exit Sub

    If AddorEdit = "ADD" Then Exit Sub
    PICLoyaltyID.ZOrder 0
    PICLoyaltyID.Enabled = True
    picAdds.Enabled = False
    picSaves.Enabled = False
    Frame1.Enabled = False
    picToolFrame.Enabled = False
    fraSearch.Enabled = False
End Sub

Private Sub Command1_Click()
    If Len(LTrim(RTrim(txtCuscde))) > 0 Then
        RaiseEvent CustomerSelected(txtCuscde, txtAcctName)
    End If
End Sub

Private Sub cmdSaveTerms_Click()
    Dim cLimit, cDays, zrated
    
    'JJE MITSUBISHI ONLY WALA SA HYUNDAI
    If cboInvoiceType.Text = "" Then
        MsgBox "Please select Invoice Type", vbInformation, "System Message"
        cboInvoiceType.SetFocus
        Exit Sub
    End If
    
    cLimit = NumericVal(txtCreditLimit)
    cDays = NumericVal(txtCreditDays)
    If chkZeroRated.Value = 1 Then
        zrated = "'Z'"
    Else
        zrated = "NULL"
    End If

    SQL_STATEMENT = "UPDATE all_Customer SET CUSCAT=" & zrated & " , CreditLimit=" & cLimit & ", CREDITTERM='C', CREDITDAYS=" & cDays & " WHERE ID=" & labid
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("EE", "CUSTOMER", SQL_STATEMENT, labid, "", "CUST CODE: " & labCustCode & " - CREDIT LIMIT", "", "")
    'NEW LOG AUDIT-----------------------------------------------------
    
    'JJE WALA SA HYUNDAI
    Set rsTERMS = New ADODB.Recordset
    rsTERMS.Open "SELECT * FROM ALL_CUSTOMER_TERMS WHERE CUSTOMERCODE = '" & labCustCode & "' AND INVOICETYPE= '" & cboInvoiceType.Text & "'", gconDMIS, adOpenForwardOnly
    If Not rsTERMS.EOF And Not rsTERMS.BOF Then
        SQL_STATEMENT = "UPDATE ALL_CUSTOMER_TERMS SET CREDITLIMIT=" & cLimit & ", CREDITTERM=" & cDays & " WHERE CUSTOMERCODE = '" & labCustCode & "' AND INVOICETYPE= '" & cboInvoiceType.Text & "'"
        gconDMIS.Execute SQL_STATEMENT
    Else
        SQL_STATEMENT = "INSERT INTO ALL_CUSTOMER_TERMS(CUSTOMERCODE,INVOICETYPE,CREDITLIMIT,CREDITTERM) VALUES("
        SQL_STATEMENT = SQL_STATEMENT & N2Str2Null(labCustCode.Caption) & ","
        SQL_STATEMENT = SQL_STATEMENT & N2Str2Null(cboInvoiceType.Text) & ","
        SQL_STATEMENT = SQL_STATEMENT & cLimit & ","
        SQL_STATEMENT = SQL_STATEMENT & cDays & ")"
        gconDMIS.Execute SQL_STATEMENT
    End If
    
    Set rsTERMS = Nothing
    On Error Resume Next
    RS.Requery
    RS.Find ("ID=" & labid)
    StoreMemVars
    MessagePop RecSaveOk, "Credit Info", "Credit Information Updated"
    ShowPictureBox picCredit, False, picMain
End Sub

Private Sub cmdHistory_Click()
    Dim frmTraHist                                                  As frmCRIS_Inquiry_CustomerTransHistory
    Set frmTraHist = New frmCRIS_Inquiry_CustomerTransHistory
    frmTraHist.SHOWTRANSACTION txtCuscde
    frmTraHist.Show
End Sub

Private Sub cmdSame_Click()
    txtDeliveryAddress = txtPersonalStreet & ", " & cboPersonalCity & ", " & txtPersonalState & ", " & txtPersonalZIP
End Sub

Private Sub cmdAddContactList_Click()
    labCUSTINFO_Contact_Click
    ShowPictureBox picContactAE, True
    txtContactName.SetFocus
    cmdAddContactList.Enabled = False
    cmdEditContact.Enabled = False
    cmdCancelContactList.Enabled = False
End Sub

Private Sub cmdSaveChild_Click()
    On Error GoTo ErrorCode:

    If RTrim(LTrim(txtChildName)) = "" Then
        ShowIsRequiredMsg "Children Name "
        On Error Resume Next
        txtChildName.SetFocus
        Exit Sub
    End If

    Dim vtxtCHILDNAME                                               As String
    Dim vtxtSEX                                                     As String
    Dim vtxtDOB                                                     As String
    Dim SQL                                                         As String
    Dim vtxtCusCde                                                  As String
    
    vtxtCusCde = N2Str2Null(txtCuscde)
    vtxtCHILDNAME = N2Str2Null(txtChildName)
    If cboChildSex = "M" Then
        vtxtSEX = "'M'"
    ElseIf cboChildSex = "F" Then
        vtxtSEX = "'F'"
    Else
        vtxtSEX = "'U'"
    End If
    vtxtDOB = N2Date2Null(txtChildDate)

    If NumericVal(labIdCHILD) = 0 Then
        SQL = "INSERT INTO ALL_CUSTOMER_CHILD (CUSCDE,CHILDNAME,SEX,DOB) VALUES("
        SQL = SQL & vtxtCusCde
        SQL = SQL & "," & vtxtCHILDNAME & " ,"
        SQL = SQL & vtxtSEX & " ,"
        SQL = SQL & vtxtDOB & " )"

        gconDMIS.Execute SQL
        SQL_STATEMENT = SQL
        
        Call NEW_LogAudit("AA", "CUSTOMER", SQL_STATEMENT, labid, "", "NAME: " & txtChildName & " - CHILD", "", FindTransactionID(N2Str2Null(labCustCode), "cuscde", "ALL_CUSTOMER_CHILD", "DETAILS", N2Str2Null(txtChildName), "CHILDNAME"))
        'LogAudit "A", "CUSTOMER CHILD"
        ShowSuccessFullyAdded
    Else
        SQL = "UPDATE ALL_CUSTOMER_CHILD SET "
        SQL = SQL & " CHILDNAME =" & vtxtCHILDNAME & " , "
        SQL = SQL & " SEX=" & vtxtSEX & " , "
        SQL = SQL & " DOB=" & vtxtDOB
        SQL = SQL & " WHERE id=" & labIdCHILD

        'LogAudit "E", "CUSTOMER CHILD"

        gconDMIS.Execute SQL
        SQL_STATEMENT = SQL

        Call NEW_LogAudit("EE", "CUSTOMER", SQL_STATEMENT, labid, "", "NAME: " & txtChildName & " - CHILD", "", labIdCHILD)
        ShowSuccessFullyUpdated
    End If

    'gconDMIS.Execute SQL

    If picChildList.Visible = True Then
        cmdCUSTINFO_Child_Click
    End If
    
    ShowPictureBox picChildAE, False, picMain
    Exit Sub
    
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCUSTINFO_Contact_Click()
    Dim temprs                                                      As ADODB.Recordset
    Set temprs = gconDMIS.Execute("SELECT CONTACTNAME,RELATION,PHONE,MOBILE,CONTACTPOSITION,DEPARTMENT,ADDRESS,ID FROM ALL_CUSTOMER_CONTACTS WHERE CUSCDE=" & N2Str2Null(txtCuscde))
    
    'Updated By: Joanna J. Enriquez 04/23/2012  1:58PM
    'Load new records added into ContactListview
    lvContactList.ListItems.Clear
    Listview_Loadval lvContactList.ListItems, temprs

    ShowPictureBox picContactList, True, picMain
    cmdAddContactList.Enabled = True
    cmdCancelContactList.Enabled = True
    
    If lvContactList.ListItems.Count = 0 Then
        cmdEditContact.Enabled = False
    Else
        cmdEditContact.Enabled = True
    End If
End Sub

Private Sub cmdCUSTINFO_CREDIT_Click()
    If Module_Access(LOGID, "CUSTOMER CREDIT LIMIT", "DATA ENTRY") = False Then Exit Sub
    ShowPictureBox picCredit, True, picMain
    On Error Resume Next
    cboInvoiceType.SetFocus
    
    'JJE WALA SA HYUNDAI
    Dim rsTERMS2 As ADODB.Recordset
    Set rsTERMS2 = New ADODB.Recordset
    rsTERMS2.Open "SELECT TOP 1 * FROM ALL_CUSTOMER_TERMS WHERE CUSTOMERCODE = '" & labCustCode.Caption & "'", gconDMIS
    If Not rsTERMS2.EOF And Not rsTERMS2.BOF Then
        cboInvoiceType.Text = Null2String(rsTERMS2!INVOICETYPE)
        txtCreditLimit.Text = ToDoubleNumber(NumericVal(rsTERMS2!CREDITLIMIT))
        txtCreditDays.Text = NumericVal(rsTERMS2!CREDITTERM)
    End If
    Set rsTERMS2 = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If picChildAE.Visible = True And KeyCode = 27 Then
        cmdCloseChildAE_Click 0
    ElseIf picChildList.Visible = True And KeyCode = 27 Then
        cmdCancelChildList_Click

    ElseIf picContactAE.Visible = True And KeyCode = 27 Then
        cmdCloseContactsAE_Click 0
    ElseIf picContactList.Visible = True And KeyCode = 27 Then
        cmdCancelContactList_Click

    ElseIf picCredit.Visible = True And KeyCode = 27 Then
        cmdCloseTerm_Click 0
    Else
        MoveKeyPress KeyCode
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF1 And Shift = 1:
        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        If picAdds.Visible = False Then Exit Sub
        Unload frmALL_AuditInquiry

        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        frmALL_AuditInquiry.Caption = "Audit Inquiry (CUSTOMER MASTER FILE)"
        Call frmALL_AuditInquiry.DisplayHistory(labid, "CUSTOMER", "")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Frame1.Enabled = False
    picAdds.Visible = True
    picToolFrame.Enabled = True
    picSaves.Visible = False
    initMemvars
    InitData

    With lvChildList.ColumnHeaders
        .Add 1, , "CHILDNAME", 0.4 * lvChildList.Width
        .Add 2, , "SEX", 0.1 * lvChildList.Width
        .Add 3, , "DATE OF BIRTH", 0.15 * lvChildList.Width
        .Add 4, , "AGE", 0.2 * lvChildList.Width
        .Add 5, , "ID", 0.1 * lvChildList.Width
    End With
    With lvContactList.ColumnHeaders
        .Add 1, , "CONTACTNAME", 0.4 * lvChildList.Width
        .Add 2, , "RELATION", 0.2 * lvChildList.Width
        .Add 3, , "PHONE", 0.17 * lvChildList.Width
        .Add 4, , "MOBILE", 0.17 * lvChildList.Width
        .Add 5, , "POSITION", 0.4 * lvChildList.Width
        .Add 6, , "DEPARTMENT", 0.2 * lvChildList.Width
        .Add 7, , "ADDRESS", 0.17 * lvChildList.Width
        .Add 8, , "ID", 0.17 * lvChildList.Width
    End With
    rsRefresh

    If AccountCode <> "" Then
        RS.Find ("CUSCDE=" & N2Str2Null(AccountCode))
        StoreMemVars
        'cmdEdit.Value = True
    End If
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsCusCtl = Nothing
    AddorEdit = vbNullString
    AccountCode = vbNullString
    CustType = vbNullString
    EntryPoint = vbNullString
    TempProspectID = 0
End Sub

Private Sub labCustInfo_Child_Click()
    cmdAddChildInfo_Click
End Sub

Private Sub labCustInfo_Child_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labCustInfo_Child.ForeColor = &H400000
    labCustInfo_Child.FontBold = True
End Sub

Private Sub labCUSTINFO_CREDIT_Click()
    cmdCUSTINFO_CREDIT_Click
End Sub

Private Sub labCUSTINFO_CREDIT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labCustInfo_Credit.ForeColor = &H400000
    labCustInfo_Credit.FontBold = True
End Sub

Private Sub labCUSTINFO_Contact_Click()
    labIDContacts = 0
    txtContactName = ""
    cboContactRelation = ""
    txtContactPosition = ""
    txtContactDepartment = ""
    txtContactPhone = ""
    txtContactMobile = ""
    txtContactAddress = ""
    cmdDeleteContact.Enabled = False
End Sub

Private Sub labCUSTINFO_Contact_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labCustInfo_Contact.ForeColor = &H400000
    labCustInfo_Contact.FontBold = True
End Sub

Private Sub lblDuplicate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblDuplicate.ForeColor = &H400000
    lblDuplicate.FontBold = True
End Sub

Private Sub lstCustomer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstCustomer
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstCustomer_DblClick()
    If lstCustomer.Enabled = True Then
        cmdEdit.Value = True
    End If
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RS.MoveFirst
    RS.Find ("ID=" & Item.ListSubItems(2).Text)
    StoreMemVars
End Sub

Private Sub lvChildList_DblClick()
    lvChildList_KeyPress 13
End Sub

Private Sub lvChildList_KeyPress(KeyAscii As Integer)
    If lvChildList.SelectedItem Is Nothing Then Exit Sub
    On Error GoTo ADDER:

    'Updated By: Joanna J. Enriquez 04/23/2012  4:15PM
    'Load new records added into Child Information Listview
    If KeyAscii <> 13 Then Exit Sub
    txtChildName = lvChildList.SelectedItem
    cboChildSex = lvChildList.SelectedItem.ListSubItems(1).Text
    txtChildDate = lvChildList.SelectedItem.ListSubItems(2).Text
    labIdCHILD = lvChildList.SelectedItem.ListSubItems(4).Text
    cmdDeleteChild.Enabled = True

    ShowPictureBox picChildAE, True, picMain
    cmdAddChildInfo.Enabled = False
    cmdSelectChild.Enabled = False
    cmdCancelChildList.Enabled = False
    On Error Resume Next
    txtChildName.SetFocus
    Exit Sub
    
ADDER:
    ShowVBError
End Sub

Private Sub lvContactList_DblClick()
    lvContactList_KeyPress 13
End Sub

Private Sub lvContactList_KeyPress(KeyAscii As Integer)
    If lvContactList.SelectedItem Is Nothing Then Exit Sub
    If KeyAscii <> 13 Then Exit Sub
    
    ShowPictureBox picContactAE, True, picMain
    With lvContactList.SelectedItem
        txtContactName = .Text
        cboContactRelation = .ListSubItems(1).Text
        txtContactPhone = .ListSubItems(2).Text
        txtContactMobile = .ListSubItems(3).Text
        txtContactPosition = .ListSubItems(4).Text
        txtContactDepartment = .ListSubItems(5).Text
        txtContactAddress = .ListSubItems(6).Text
        labIDContacts = .ListSubItems(7).Text
    End With

    ShowPictureBox picContactAE, True, picMain
    cmdAddContactList.Enabled = False
    cmdEditContact.Enabled = False
    cmdCancelContactList.Enabled = False
    cmdDeleteContact.Enabled = True
    On Error Resume Next

    txtContactName.SetFocus
End Sub

Private Sub Option1_Click()
'    lstCustomer.ColumnHeaders(1).Text = "CODE"
    FillSearchGrid (txtSearch.Text)
End Sub

Private Sub optSearchKeyAcctName_Click()
'    lstCustomer.ColumnHeaders(1).Text = "A/C Name"
    FillSearchGrid (txtSearch.Text)
End Sub

Private Sub optSearchKeyAddress_Click()
'    lstCustomer.ColumnHeaders(1).Text = "Address"
    FillSearchGrid (txtSearch.Text)
End Sub

Private Sub optSearchKeyCompany_Click()
'    lstCustomer.ColumnHeaders(1).Text = "Company"
    FillSearchGrid (txtSearch.Text)
End Sub

Private Sub optSearchKeyEmail_Click()
'    lstCustomer.ColumnHeaders(1).Text = "Email"
    FillSearchGrid (txtSearch.Text)
End Sub

Private Sub optSearchKeyLast_Click()
'    lstCustomer.ColumnHeaders(1).Text = "LastName"
    FillSearchGrid (txtSearch.Text)
End Sub

Private Sub picToolFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labCustInfo_Child.ForeColor = vbBlack
    labCustInfo_Child.FontBold = False
    labCustInfo_Credit.ForeColor = vbBlack
    labCustInfo_Credit.FontBold = False
    labCustInfo_Contact.ForeColor = vbBlack
    labCustInfo_Contact.FontBold = False
    lblDuplicate.ForeColor = vbBlack
    lblDuplicate.FontBold = False
End Sub
'JJE WALA SA HYUNDAI
Private Sub picCredit_Click()
Limit:
End Sub

Private Sub Timer1_Timer()
    If lblloyal.Visible = True Then
        Label11.Visible = True
        lblloyal.Visible = False
    ElseIf Label11.Visible = True Then
        Label11.Visible = False
        'lblloyal.Visible = True
        lblloyal.Visible = False
    End If
End Sub

Private Sub txtBirthDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtCreditDays_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCreditDays_GotFocus()
    If NumericVal(txtCreditDays.Text) <= 0 Then txtCreditDays = ""
End Sub

Private Sub txtCreditDays_LostFocus()
    If NumericVal(txtCreditDays) <= 0 Then txtCreditDays = "0"
    txtCreditDays = NumericVal(txtCreditDays)
End Sub

Private Sub txtCreditLimit_GotFocus()
    If NumericVal(txtCreditLimit.Text) <= 0 Then txtCreditLimit = ""
End Sub

Private Sub txtCreditLimit_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCreditLimit_LostFocus()
    If NumericVal(txtCreditLimit) <= 0 Then txtCreditLimit = "0.00"
    txtCreditLimit = FormatNumber(NumericVal(txtCreditLimit))
End Sub

Private Sub txtFirstName_Change()
    SetCustomerAccountName
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtLastName_Change()
    If AddorEdit = "ADD" And LTrim(RTrim(txtLastName)) <> "" Then
        txtCuscde = GetCustomerCode(txtLastName)
    End If
    SetCustomerAccountName
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMiddleName_Change()
    SetCustomerAccountName
End Sub

Private Sub txtMiddleName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNotes_Change()
    If Len(txtNotes.Text) > 250 Then
        MsgBox "Number of character exceed..please simplify your notes.", vbInformation, "Warning!"
        txtNotes.SetFocus
        txtNotes.BackColor = &HFFFFC0
    Else
        txtNotes.BackColor = vbWhite
        Exit Sub
    End If
End Sub

Private Sub txtPersonalStreet_KeyPress(KeyAscii As Integer)
    UpperAscii KeyAscii
End Sub

Private Sub txtPersonalZIP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSearch_Change()
    FillSearchGrid (txtSearch.Text)
End Sub

Public Sub AddCustomerFromProspect(oRs As Recordset, xGoingWhere As String)
    Dim ar                                                          As Variant
    
    GoingWhere = xGoingWhere
    If Not (oRs.EOF Or oRs.BOF) Then
        EntryPoint = "PROSPECT"
        AddorEdit = "ADD"
        picAdds.Visible = False
        picSaves.Visible = True
        Frame1.Enabled = True
        fraSearch.Enabled = False
        txtAcctName.Text = Null2String(oRs!AcctName)
        TempProspectID = oRs!PROSPECTID
        CustType = Null2String(oRs!ProspectType)
        If CustType = "P" Then
            ar = Split(Null2String(oRs!AcctName))
            If UBound(ar) = 0 Then
                txtLastName.Text = ar(0)
            ElseIf UBound(ar) = 1 Then
                txtFirstName.Text = ar(0)
                txtLastName.Text = ar(1)
            ElseIf UBound(ar) >= 2 Then
                txtFirstName.Text = ar(0)
                txtLastName.Text = ar(2)
                txtMiddleName.Text = ar(1)
            End If
        Else
            ar = Split(Null2String(oRs!ContactPerson))
            If UBound(ar) = 0 Then
                txtLastName.Text = ar(0)
            ElseIf UBound(ar) = 1 Then
                txtFirstName.Text = ar(0)
                txtLastName.Text = ar(1)
            ElseIf UBound(ar) >= 2 Then
                txtFirstName.Text = ar(0)
                txtLastName.Text = ar(1)
                txtMiddleName.Text = ar(2)
            End If
        End If
        txtCusphon1 = Null2String(oRs!Telephone)
        txtMobile = Null2String(oRs!Mobile)
        txtEmail = Null2String(oRs!EMAIL)
        txtPersonalStreet = Null2String(oRs!Address)
        txtNotes = Null2String(oRs!Notes)

        If CustType = "P" Then
            cboCustType.ListIndex = 0
        ElseIf CustType = "C" Then
            cboCustType.ListIndex = 1
        ElseIf CustType = "F" Then
            cboCustType.ListIndex = 2
        ElseIf CustType = "G" Then
            cboCustType.ListIndex = 3
        End If
        cboLeadSource.Text = Null2String(oRs!LeadSource)
    End If
End Sub

Public Sub FillCustomerListView(RS As Recordset, grd As ListView, Optional WithSN As Boolean = False, Optional WITHCOLUMNHEADER As Boolean = False)
    Dim fld                                                         As Field
    Dim j                                                           As Long
    Dim ijx                                                         As Integer
    Dim LST                                                         As ListItem
    Dim i                                                           As Integer

    grd.Enabled = False
    grd.ListItems.Clear

    If WithSN = True And WITHCOLUMNHEADER = True Then
        grd.ColumnHeaders.Clear
        Call grd.ColumnHeaders.Add(, , "Item")
        For i = 0 To RS.Fields.Count - 1
            Call grd.ColumnHeaders.Add(, , RS.Fields(i).Name)
        Next
        While Not RS.EOF
            j = j + 1
            Set LST = grd.ListItems.Add(, , j)
            For Each fld In RS.Fields
                If IsNull(fld.Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , fld.Value
                End If
            Next
            RS.MoveNext
        Wend
    ElseIf WithSN = True And WITHCOLUMNHEADER = False Then
        While Not RS.EOF
            j = j + 1
            Set LST = grd.ListItems.Add(, , j)
            For Each fld In RS.Fields
                If IsNull(fld.Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , fld.Value
                End If
            Next
            RS.MoveNext
        Wend
    ElseIf WithSN = False And WITHCOLUMNHEADER = True Then
        grd.ColumnHeaders.Clear
        For i = 0 To RS.Fields.Count - 1
            Call grd.ColumnHeaders.Add(, , RS.Fields(i).Name)
        Next
        j = RS.Fields.Count
        While Not RS.EOF
            Set LST = grd.ListItems.Add(, , RS.Fields(0).Value)
            For ijx = 1 To j - 1
                If IsNull(RS.Fields(ijx).Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , RS.Fields(ijx).Value
                End If
            Next
            RS.MoveNext
        Wend
    Else
        j = RS.Fields.Count
        While Not RS.EOF
            Set LST = grd.ListItems.Add(, , Null2String(RS.Fields(0).Value))
            For ijx = 1 To j - 1
                If IsNull(RS.Fields(ijx).Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , RS.Fields(ijx).Value
                End If
            Next
            RS.MoveNext
        Wend
    End If
    grd.Enabled = True
    Set LST = Nothing
End Sub
'JJE WALA HYUNDAI
Private Sub cboInvoiceType_Click()
    txtCreditDays.Text = "0"
    txtCreditLimit.Text = "0.00"
    Set rsTERMS = New ADODB.Recordset
    rsTERMS.Open "SELECT * FROM ALL_CUSTOMER_TERMS WHERE CUSTOMERCODE = '" & labCustCode.Caption & "' AND INVOICETYPE= '" & cboInvoiceType.Text & "'", gconDMIS
    If Not rsTERMS.EOF And Not rsTERMS.BOF Then
        txtCreditLimit.Text = ToDoubleNumber(NumericVal(rsTERMS!CREDITLIMIT))
        txtCreditDays.Text = NumericVal(rsTERMS!CREDITTERM)
    End If
    Set rsTERMS = Nothing
End Sub
