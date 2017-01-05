VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmAMIS_UNAPPLIED_PAYMENT 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Un-applied Payment"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   11715
   Begin VB.PictureBox pic_Dropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2955
      Left            =   4800
      ScaleHeight     =   2925
      ScaleWidth      =   5475
      TabIndex        =   35
      Top             =   3150
      Visible         =   0   'False
      Width           =   5505
      Begin MSComctlLib.ListView lvwInvoice 
         Height          =   2565
         Left            =   30
         TabIndex        =   38
         Top             =   330
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   4524
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
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Invoice No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Invioce Type"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Invoice Amt"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5070
         TabIndex        =   37
         Top             =   30
         Width           =   375
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   330
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   5505
         _Version        =   655364
         _ExtentX        =   9710
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "Select Invoice"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
      End
   End
   Begin VB.PictureBox pic_Tagging 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6525
      Left            =   420
      ScaleHeight     =   6495
      ScaleWidth      =   10965
      TabIndex        =   18
      Top             =   510
      Width           =   10995
      Begin VB.TextBox txtTotalApplieAmount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6030
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   5730
         Width           =   2685
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Height          =   495
         Left            =   10020
         TabIndex        =   40
         Top             =   2670
         Width           =   825
      End
      Begin VB.CommandButton cmdPVCancel 
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
         Left            =   10080
         MouseIcon       =   "frmAMIS_UNAPPLIED_PAYMENT.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmAMIS_UNAPPLIED_PAYMENT.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   5640
         Width           =   795
      End
      Begin VB.ComboBox cboTagAR 
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
         Left            =   4350
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2760
         Width           =   4305
      End
      Begin VB.TextBox txtInvoiceAmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8730
         TabIndex        =   30
         Top             =   2190
         Width           =   2115
      End
      Begin VB.TextBox txtInvoiceDate 
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
         Left            =   6510
         TabIndex        =   29
         Top             =   2205
         Width           =   2145
      End
      Begin VB.TextBox txtInvoiceType 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   27
         Top             =   2220
         Width           =   2085
      End
      Begin VB.TextBox txtVoucherNo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   2220
         Width           =   2085
      End
      Begin VB.TextBox txtInvoiceNo 
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
         Left            =   4380
         TabIndex        =   28
         Top             =   2205
         Width           =   2085
      End
      Begin MSComctlLib.ListView lstDetails 
         Height          =   1785
         Left            =   120
         TabIndex        =   20
         Top             =   90
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   3149
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
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmAMIS_UNAPPLIED_PAYMENT.frx":0490
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ITEM #"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ACCOUNT CODE"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ACCOUNT DESCRIPTION"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "DEBIT"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "CREDIT"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin MSComctlLib.ListView lvwUplydList 
         Height          =   2385
         Left            =   120
         TabIndex        =   39
         Top             =   3210
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   4207
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
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ITEM #  "
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Invioce Type"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Invoice No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Invoice Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Voucher No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "AR Account"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.CommandButton cmdPVSave 
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
         Left            =   9300
         MouseIcon       =   "frmAMIS_UNAPPLIED_PAYMENT.frx":05F2
         MousePointer    =   99  'Custom
         Picture         =   "frmAMIS_UNAPPLIED_PAYMENT.frx":0744
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   5640
         Width           =   795
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Applied Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   41
         Top             =   5850
         Width           =   2625
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   10830
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Apply to"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   32
         Top             =   2790
         Width           =   1785
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8760
         TabIndex        =   21
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6510
         TabIndex        =   25
         Top             =   1950
         Width           =   1785
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4380
         TabIndex        =   24
         Top             =   1950
         Width           =   1785
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         TabIndex        =   23
         Top             =   1950
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   1950
         Width           =   1785
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   6585
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   11835
         _Version        =   655364
         _ExtentX        =   20876
         _ExtentY        =   11615
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7830
      Left            =   0
      ScaleHeight     =   7830
      ScaleWidth      =   11715
      TabIndex        =   0
      Top             =   0
      Width           =   11715
      Begin VB.CheckBox Check1 
         Caption         =   "Display Un-Applied Payment With Invoice"
         Height          =   435
         Left            =   2130
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6690
         Width           =   3705
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   375
         Left            =   1350
         TabIndex        =   12
         Top             =   6720
         Width           =   435
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   7170
         Width           =   11625
      End
      Begin VB.PictureBox picLoading 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   3300
         ScaleHeight     =   1425
         ScaleWidth      =   5445
         TabIndex        =   6
         Top             =   2700
         Visible         =   0   'False
         Width           =   5475
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   555
            Left            =   -30
            TabIndex        =   9
            Top             =   300
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   979
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label labvoucherno 
            BackStyle       =   0  'Transparent
            Caption         =   "labvoucherno"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   30
            TabIndex        =   7
            Top             =   0
            Width           =   2745
         End
         Begin VB.Label labpercent 
            BackStyle       =   0  'Transparent
            Caption         =   "labpercent"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   30
            TabIndex        =   10
            Top             =   900
            Width           =   2745
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption5 
            Height          =   1545
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   5445
            _Version        =   655364
            _ExtentX        =   9604
            _ExtentY        =   2725
            _StockProps     =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            VisualTheme     =   3
         End
      End
      Begin VB.TextBox txtTotalAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9210
         TabIndex        =   16
         Top             =   6720
         Width           =   2415
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   60
         Width           =   6915
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmAMIS_UNAPPLIED_PAYMENT.frx":0A94
         Left            =   8940
         List            =   "frmAMIS_UNAPPLIED_PAYMENT.frx":0A9E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   30
         Width           =   2505
      End
      Begin FlexCell.Grid grdUnppliedPayment 
         Height          =   6105
         Left            =   30
         TabIndex        =   4
         Top             =   510
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   10769
         Appearance      =   0
         BackColorBkg    =   -2147483645
         Cols            =   5
         DefaultFontName =   "Verdana"
         DefaultFontSize =   9
         Rows            =   30
         SelectionMode   =   1
         EnterKeyMoveTo  =   1
      End
      Begin VB.PictureBox DataGrid1 
         Height          =   6105
         Left            =   60
         ScaleHeight     =   6045
         ScaleWidth      =   11625
         TabIndex        =   5
         Top             =   510
         Width           =   11685
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1830
         TabIndex        =   14
         Top             =   6795
         Width           =   285
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Show Top"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   6795
         Width           =   1275
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   1245
         Left            =   0
         TabIndex        =   11
         Top             =   6630
         Width           =   11745
         _Version        =   655364
         _ExtentX        =   20717
         _ExtentY        =   2196
         _StockProps     =   14
         Caption         =   "                                                                                            Total Amount:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   555
         Left            =   -60
         TabIndex        =   1
         Top             =   -30
         Width           =   11865
         _Version        =   655364
         _ExtentX        =   20929
         _ExtentY        =   979
         _StockProps     =   14
         Caption         =   $"frmAMIS_UNAPPLIED_PAYMENT.frx":0ABE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
      End
   End
End
Attribute VB_Name = "frmAMIS_UNAPPLIED_PAYMENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUNAPLYD_PAYMENT                             As ADODB.Recordset
Dim rsGetInvoice                                  As ADODB.Recordset
Dim rsCHECK_INVOICE                               As ADODB.Recordset
Dim xVOUCHERNO                                    As String
Dim XCustomerCode                                 As String
Dim BILANG                                        As Integer
Dim TOTAL_AR_AMOUNT                               As Double
Dim TotalARAmountToPay                            As Double
Dim TOTAL_APPLIED                                 As Double
Dim TOTDEBIT                                      As Double
Dim TOTCREDIT                                     As Double
Dim TOTTAX                                        As Double
'Dim xCHK                                               As Integer
Private Sub Check1_Click()
    rsRefresh
    StoreMemVars
End Sub

Private Sub cmdOK_Click()
    Dim Item                                      As ListItem
    Dim xCount                                    As Integer
    Dim lvwCount                                  As Integer
    Dim invCount                                  As Integer

    TOTAL_APPLIED = 0

    If Not (txtInvoiceType = "AI" Or txtInvoiceType = "VI" Or txtInvoiceType = "SI" Or txtInvoiceType = "PI" Or txtInvoiceType = "MI" Or txtInvoiceType = "CI") Then
        MessagePop RecSaveError, "WARNING!", "Invalid Invoice Type!"
        txtInvoiceType.BackColor = &HFFFF80
        Exit Sub
    End If

    If lvwUplydList.ListItems.Count = 0 Then
        If ToDoubleNumber(txtInvoiceAmt.Text) <= TOTCREDIT Then
            'proceed
        Else
            MessagePop InfoFriend, "WARNING!", "Total Details amount must be less than or equal to the total AR amount"
            Exit Sub
        End If
    End If

    For invCount = 1 To lvwUplydList.ListItems.Count
        TOTAL_APPLIED = TOTAL_APPLIED + ToDoubleNumber(lvwUplydList.ListItems(invCount).ListSubItems.Item(4))
    Next invCount

    TOTAL_APPLIED = TOTAL_APPLIED + ToDoubleNumber(txtInvoiceAmt.Text)

    If TOTAL_APPLIED <= TOTCREDIT Then
        'ok
    Else
        MessagePop InfoFriend, "WARNING!", "Total Details amount must be less than or equal to the total AR amount"
        Exit Sub
    End If

    For lvwCount = 1 To lvwUplydList.ListItems.Count
        If lvwUplydList.ListItems(lvwCount).ListSubItems.Item(1) = txtInvoiceType.Text And lvwUplydList.ListItems(lvwCount).ListSubItems.Item(2) = txtInvoiceNo.Text Then
            MessagePop InfoFriend, "WARNING!", "Invoice No. and Invoice Type is already in use."
            Exit Sub
        Else
            'pass
        End If
    Next lvwCount

    If cboTagAR.Text = "" Then
        MessagePop RecSaveError, "WARNING!", "Tagging of AR is required.."
        cboTagAR.BackColor = &HFFFF80
        Exit Sub
    End If

    If lvwUplydList.ListItems.Count = 0 Then
        BILANG = 0
    End If

    If BILANG <> 1 Then
        Set Item = lvwUplydList.ListItems.Add(, , GetItemNO(txtVoucherNo, "CRJ"))
    Else
        Set Item = lvwUplydList.ListItems.Add(, , Format(NumericVal(lvwUplydList.SelectedItem.Text) + 1, "0000"))
    End If

    Item.SubItems(1) = txtInvoiceType.Text
    Item.SubItems(2) = txtInvoiceNo.Text
    Item.SubItems(3) = txtInvoiceDate.Text
    Item.SubItems(4) = txtInvoiceAmt.Text
    Item.SubItems(5) = txtVoucherNo.Text
    Item.SubItems(6) = cboTagAR.Text

    txtInvoiceType.Text = ""
    txtInvoiceNo.Text = ""
    txtInvoiceAmt.Text = ""
    cboTagAR.Clear
    InitCbo_AR

    txtTotalApplieAmount.Text = ToDoubleNumber(TOTAL_APPLIED)
    pic_Dropdown.Visible = True

End Sub
Function GetItemNO(xVOUCHERNO As String, xJType As String) As String
    Dim getItem                                   As ADODB.Recordset
    Set getItem = gconDMIS.Execute("Select cast(ITEMNO as integer) as ItemNo from Amis_Crj_detail where voucherno = '" & xVOUCHERNO & "' and CR_TYPE = '" & xJType & "' order by  ItemNo desc")
    If Not getItem.EOF And Not getItem.BOF Then
        GetItemNO = Format(NumericVal(getItem!ItemNo) + 1, "0000")
        BILANG = 1
    Else
        GetItemNO = "0001"
        BILANG = 1
    End If
    Set getItem = Nothing
End Function

Private Sub cmdPVCancel_Click()
    pic_Tagging.Visible = False
    pic_Dropdown.Visible = False
    grdUnppliedPayment.Enabled = True
    grdUnppliedPayment.SetFocus
End Sub

Private Sub cmdPVSave_Click()
    Dim rsCheck_ItemNo                            As ADODB.Recordset
    Dim Kawnt                                     As Integer
    Dim xItemNo                                   As String
    Dim xJ_Class                                  As String
    Dim xVouch                                    As String
    Dim xINVOICETYPE                              As String
    Dim xINVOICENO                                As String
    Dim xInvoicedate                              As String
    Dim XINVOICEAMT                               As String

    For Kawnt = 1 To lvwUplydList.ListItems.Count
        xJ_Class = Null2String(lvwUplydList.ListItems(Kawnt).ListSubItems.Item(6))
        xVouch = Null2String(lvwUplydList.ListItems(Kawnt).ListSubItems.Item(5))
        xItemNo = Null2String(lvwUplydList.ListItems(Kawnt).Text)
        xINVOICETYPE = Null2String(lvwUplydList.ListItems(Kawnt).ListSubItems.Item(1))
        xINVOICENO = Null2String(lvwUplydList.ListItems(Kawnt).ListSubItems.Item(2))
        xInvoicedate = N2Date2Null(lvwUplydList.ListItems(Kawnt).ListSubItems.Item(3))
        XINVOICEAMT = NumericVal(lvwUplydList.ListItems(Kawnt).ListSubItems.Item(4))

        gconDMIS.Execute ("INSERT INTO AMIS_CRJ_DETAIL(J_CLASS,CR_TYPE,VoucherNo,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,STATUS)" & _
                          "VALUES('" & Setacctcode(xJ_Class) & "','CRJ','" & xVouch & "','" & xItemNo & "','" & xINVOICETYPE & "','" & xINVOICENO & "'," & xInvoicedate & "," & XINVOICEAMT & ",'" & SetStatus(xVouch) & "')")
        MessagePop RecSaveInfo, "INFORMATION", "Payment Succesfully applied!"
    Next Kawnt

    pic_Tagging.Visible = False
    grdUnppliedPayment.Enabled = True
    grdUnppliedPayment.RemoveItem (grdUnppliedPayment.ActiveCell.Row)
    txtRemarks = ""
    grdUnppliedPayment.SetFocus
End Sub
Function Setacctcode(JJJ As Variant) As String
    Dim rsChartAccount2                           As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where Description = " & N2Str2Null(JJJ), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctcode = UCase(Null2String(rsChartAccount2!ACCTCODE))
    Else
        Setacctcode = ""
    End If
End Function
Function SetStatus(XXX As String) As String
    Dim rsSTATUS                                  As ADODB.Recordset
    Set rsSTATUS = gconDMIS.Execute("SELECT STATUS FROM AMIS_JOURNAL_HD WHERE JTYPE = 'CRJ' AND VOUCHERNO = '" & XXX & "'")
    If Not rsSTATUS.EOF And Not rsSTATUS.BOF Then
        SetStatus = UCase(Null2String(rsSTATUS!Status))
    Else
        'THE VOUCHER NO IS NO HEADER
    End If
    Set rsSTATUS = Nothing
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If pic_Tagging.Visible = True And KeyCode = vbKeyEscape Then
        cmdPVCancel_Click
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    InitGrid
    rsRefresh
    pic_Tagging.Visible = False
    StoreMemVars
    Screen.MousePointer = 0
End Sub
Sub InitGrid()
    With grdUnppliedPayment
        .Cols = 9: .Rows = 2
        .DisplayFocusRect = True: .AllowUserResizing = True

        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Cell(0, 1).Text = "CODE"
        .Cell(0, 2).Text = "CUSTOMER NAME"
        .Cell(0, 3).Text = "VOUCHER#"
        .Cell(0, 4).Text = "OR#"
        .Cell(0, 5).Text = "OR AMOUNT"
        .Cell(0, 6).Text = "INV TYPE"
        .Cell(0, 6).WrapText = True

        .Cell(0, 7).Text = "OR DATE"
        .Cell(0, 8).Text = "REMARKS"


        .Column(1).CellType = cellTextBox: .Column(1).Alignment = cellCenterGeneral
        .Column(2).CellType = cellTextBox:                 '.Column(2).MaxLength = 50
        .Column(3).CellType = cellTextBox: .Column(3).Alignment = cellCenterGeneral
        .Column(4).CellType = cellTextBox: .Column(4).Alignment = cellCenterGeneral
        .Column(5).CellType = cellTextBox: .Column(5).Alignment = cellRightGeneral
        .Column(6).CellType = cellTextBox
        .Column(7).CellType = cellTextBox
        .Column(8).CellType = cellTextBox

        .Column(1).Width = 60: .Column(1).Locked = True
        .Column(2).Width = 245: .Column(2).Locked = True
        .Column(3).Width = 80: .Column(3).Locked = True
        .Column(4).Width = 70: .Column(4).Locked = True
        .Column(5).Width = 110: .Column(5).Locked = True
        .Column(6).Width = 65: .Column(6).Locked = True

        .Column(7).Width = 85: .Column(7).Locked = True
        .Column(8).Width = 300: .Column(8).Locked = True
        .AllowUserSort = False
        .RowHeight(0) = 25
        .Range(1, 3, .Rows - 1, 3).ForeColor = RGB(0, 0, 128)
    End With
End Sub
Sub rsRefresh()
    Set rsUNAPLYD_PAYMENT = New ADODB.Recordset
    If IsNumeric(Text1) = True Then
        rsUNAPLYD_PAYMENT.Open "SELECT top " & Text1 & " PERCENT  HD.CUSTOMERCODE as Code,CUST.ACCTNAME as AcctName,HD.VOUCHERNO as VoucherNo,HD.INVOICENO as InvoiceNo,HD.INVOICETYPE as InvoiceType,HD.REMARKS as Remarks,HD.INVOICEAMT as InvoiceAmt,HD.INVOICEDATE AS InvDate From AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_CUSTOMER_TABLE CUST ON HD.CustomerCode = CUST.CUSCDE WHERE JTYPE = 'CRJ' AND VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_CRJ_DETAIL WHERE CR_TYPE = 'CRJ') order by AcctName asc", gconDMIS
    Else
        rsUNAPLYD_PAYMENT.Open "SELECT HD.CUSTOMERCODE as Code,CUST.ACCTNAME as AcctName,HD.VOUCHERNO as VoucherNo,HD.INVOICENO as InvoiceNo,HD.INVOICETYPE as InvoiceType,HD.REMARKS as Remarks,HD.INVOICEAMT as InvoiceAmt,HD.INVOICEDATE AS InvDate From AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_CUSTOMER_TABLE CUST ON HD.CustomerCode = CUST.CUSCDE WHERE JTYPE = 'CRJ'   AND VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_CRJ_DETAIL WHERE CR_TYPE = 'CRJ') order by AcctName asc", gconDMIS
    End If
End Sub

Sub StoreMemVars()
    grdUnppliedPayment.Rows = 1
    grdUnppliedPayment.AutoRedraw = False

    Dim xTOTAL                                    As Double

    xTOTAL = 0
    If Check1.Value = 1 Then
        picLoading.Visible = True
        ProgressBar1.Value = 0
        ProgressBar1.Max = rsUNAPLYD_PAYMENT.RecordCount
        If Not rsUNAPLYD_PAYMENT.EOF And Not rsUNAPLYD_PAYMENT.BOF Then
            Do While Not rsUNAPLYD_PAYMENT.EOF
                DoEvents
                'CHECK_IS_SCHEDULE_ACCOUNT (Null2String(rsUNAPLYD_PAYMENT!VOUCHERNO))
                'If xCHK > 0 Then
                Set rsCHECK_INVOICE = gconDMIS.Execute("SELECT * FROM AMIS_JOURNAL_HD WHERE CUSTOMERCODE = '" & Null2String(rsUNAPLYD_PAYMENT!code) & "' AND JTYPE = 'SJ' AND INVOICENO IS NOT NULL AND INVOICETYPE IS NOT NULL")
                If Not rsCHECK_INVOICE.EOF And Not rsCHECK_INVOICE.BOF Then
                    grdUnppliedPayment.AddItem rsUNAPLYD_PAYMENT!code & Chr(9) & rsUNAPLYD_PAYMENT!AcctName & Chr(9) & rsUNAPLYD_PAYMENT!VOUCHERNO & Chr(9) & rsUNAPLYD_PAYMENT!INVOICENO & Chr(9) & ToDoubleNumber(rsUNAPLYD_PAYMENT!InvoiceAmt) & Chr(9) & rsUNAPLYD_PAYMENT!InvoiceType & Chr(9) & rsUNAPLYD_PAYMENT!INVDATE & Chr(9) & rsUNAPLYD_PAYMENT!remarks
                    xTOTAL = xTOTAL + NumericVal(rsUNAPLYD_PAYMENT!InvoiceAmt)
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    labPercent.Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "% Completed"
                    labvoucherno.Caption = Null2String(rsUNAPLYD_PAYMENT!VOUCHERNO)
                    rsUNAPLYD_PAYMENT.MoveNext
                    grdUnppliedPayment.TopRow = grdUnppliedPayment.Rows - 1
                    grdUnppliedPayment.Refresh

                Else
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    labPercent.Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "% Completed"
                    labvoucherno.Caption = Null2String(rsUNAPLYD_PAYMENT!VOUCHERNO)
                    rsUNAPLYD_PAYMENT.MoveNext
                End If
            Loop

            grdUnppliedPayment.AutoRedraw = True
            grdUnppliedPayment.Refresh
            txtTotalAmount = ToDoubleNumber(xTOTAL)
            Set rsCHECK_INVOICE = Nothing
            picLoading.Visible = False
        Else
            MessagePop InfoFriend, "INFORMATION", "There is no record found"
            Exit Sub
        End If
    Else
        If Not rsUNAPLYD_PAYMENT.EOF And Not rsUNAPLYD_PAYMENT.BOF Then
            Do While Not rsUNAPLYD_PAYMENT.EOF
                grdUnppliedPayment.AddItem rsUNAPLYD_PAYMENT!code & Chr(9) & rsUNAPLYD_PAYMENT!AcctName & Chr(9) & rsUNAPLYD_PAYMENT!VOUCHERNO & Chr(9) & rsUNAPLYD_PAYMENT!INVOICENO & Chr(9) & ToDoubleNumber(rsUNAPLYD_PAYMENT!InvoiceAmt) & Chr(9) & rsUNAPLYD_PAYMENT!InvoiceType & Chr(9) & rsUNAPLYD_PAYMENT!INVDATE & Chr(9) & rsUNAPLYD_PAYMENT!remarks
                xTOTAL = xTOTAL + NumericVal(rsUNAPLYD_PAYMENT!InvoiceAmt)
                rsUNAPLYD_PAYMENT.MoveNext
                DoEvents
            Loop
            txtTotalAmount = ToDoubleNumber(xTOTAL)
            grdUnppliedPayment.AutoRedraw = True
            grdUnppliedPayment.Refresh
        Else
            MessagePop InfoFriend, "INFORMATION", "There is no record found"
            Exit Sub
        End If
    End If
    Set rsUNAPLYD_PAYMENT = Nothing
End Sub

Function CHECK_IS_SCHEDULE_ACCOUNT(xVOUCHERNO As String)
    Dim RSDET                                     As ADODB.Recordset
    Dim rsAccount                                 As ADODB.Recordset
    'xCHK = 0
    Set RSDET = gconDMIS.Execute("Select ACCT_CODE from AMis_journal_det where voucherno = '" & xVOUCHERNO & "' and Jtype = 'CRJ'")
    If Not RSDET.EOF And Not RSDET.BOF Then
        Do While Not RSDET.EOF
            Set rsAccount = gconDMIS.Execute("Select *  FROM AMIS_CHARTACCOUNT WHERE ACCTCODE = '" & Null2String(RSDET!Acct_code) & "' and IS_SCHEDULE_ACCNT = 1 ")
            If Not rsAccount.EOF And Not rsAccount.BOF Then
                'xCHK = xCHK + 1
            End If
            RSDET.MoveNext
        Loop
    End If
    Set RSDET = Nothing
    Set rsAccount = Nothing
End Function

Private Sub grdUnppliedPayment_Click()
    xVOUCHERNO = grdUnppliedPayment.Cell(grdUnppliedPayment.ActiveCell.Row, 3).Text
    XCustomerCode = LTrim(RTrim(grdUnppliedPayment.Cell(grdUnppliedPayment.ActiveCell.Row, 1).Text))
    TOTDEBIT = 0
    TOTCREDIT = 0
    TOTTAX = 0
End Sub

Private Sub grdUnppliedPayment_DblClick()
    initMemvars

    txtVoucherNo.Text = grdUnppliedPayment.Cell(grdUnppliedPayment.ActiveCell.Row, 3).Text
    txtInvoiceDate.Text = grdUnppliedPayment.Cell(grdUnppliedPayment.ActiveCell.Row, 7).Text
    txtInvoiceAmt.Text = grdUnppliedPayment.Cell(grdUnppliedPayment.ActiveCell.Row, 5).Text

    InitCbo_AR
    FillDetails
    grdUnppliedPayment.Enabled = False
    pic_Tagging.Visible = True
    lvwUplydList.ListItems.Clear
    On Error Resume Next
    txtInvoiceNo.SetFocus
End Sub

Private Sub grdUnppliedPayment_RowColChange(ByVal Row As Long, ByVal Col As Long)
    grdUnppliedPayment.Range(Row, 0, Row, 8).Selected
    txtRemarks.Text = grdUnppliedPayment.Cell(Row, 8).Text
End Sub

Private Sub Label7_Click()
    pic_Dropdown.Visible = False
End Sub

Private Sub lvwInvoice_DblClick()
    If lvwInvoice.ListItems.Count = 0 Then Exit Sub
    txtInvoiceNo.Text = lvwInvoice.SelectedItem.Text
    txtInvoiceType.Text = lvwInvoice.SelectedItem.SubItems(1)
    txtInvoiceAmt.Text = lvwInvoice.SelectedItem.SubItems(2)
    pic_Dropdown.Visible = False
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyInteger(KeyAscii)
End Sub

Private Sub txtInvoiceNo_Change()
    Dim Item                                      As ListItem
    Dim rsGetInvoice                              As ADODB.Recordset
    Dim I                                         As Integer
    Dim item_exists                               As Boolean
    If txtInvoiceNo.Text <> "" Then
        Set rsGetInvoice = gconDMIS.Execute("SELECT INVOICETYPE,INVOICENO,INVOICEAMT FROM AMIS_JOURNAL_HD WHERE JTYPE = 'SJ' AND CUSTOMERCODE = '" & XCustomerCode & "' AND INVOICENO LIKE '" & txtInvoiceNo & "%'")
    Else
        Set rsGetInvoice = gconDMIS.Execute("SELECT INVOICETYPE,INVOICENO,INVOICEAMT FROM AMIS_JOURNAL_HD WHERE JTYPE = 'SJ' AND CUSTOMERCODE = '" & XCustomerCode & "'")
    End If
    Me.lvwInvoice.Sorted = True
    Me.lvwInvoice.ListItems.Clear
    Me.lvwInvoice.Enabled = False

    If Not rsGetInvoice.EOF And Not rsGetInvoice.BOF Then
        Do While Not rsGetInvoice.EOF
            item_exists = False
            If lvwUplydList.ListItems.Count = 0 Then
                Set Item = lvwInvoice.ListItems.Add(, , Null2String(rsGetInvoice!INVOICENO))
                Item.SubItems(1) = Null2String(rsGetInvoice!InvoiceType)
                Item.SubItems(2) = Round(ToDoubleNumber(rsGetInvoice!InvoiceAmt), 2)
            Else
                For I = 1 To lvwUplydList.ListItems.Count
                    If lvwUplydList.ListItems(I).ListSubItems(2).Text = Null2String(rsGetInvoice!INVOICENO) And _
                       lvwUplydList.ListItems(I).ListSubItems(1).Text = Null2String(rsGetInvoice!InvoiceType) Then
                        item_exists = True
                        Exit For
                    End If
                Next

                If item_exists = False Then
                    Set Item = lvwInvoice.ListItems.Add(, , Null2String(rsGetInvoice!INVOICENO))
                    Item.SubItems(1) = Null2String(rsGetInvoice!InvoiceType)
                    Item.SubItems(2) = Round(ToDoubleNumber(rsGetInvoice!InvoiceAmt), 2)
                End If
            End If
            rsGetInvoice.MoveNext
        Loop
    End If
    Me.lvwInvoice.Enabled = True: Me.lvwInvoice.Sorted = False: Me.lvwInvoice.Refresh
    pic_Dropdown.Visible = True
End Sub

Private Sub txtInvoiceNo_GotFocus()
    If txtInvoiceNo = "" Then
        pic_Dropdown.Visible = True
        txtInvoiceNo_Change
    End If
End Sub

Private Sub txtInvoiceNo_KeyPress(KeyAscii As Integer)
'    Dim Item                                           As ListItem
'    Dim i                                              As Integer
'    Dim item_exists                                    As Boolean
'
'    If KeyAscii = 13 Then
'        Set rsGetInvoice = gconDMIS.Execute("SELECT INVOICETYPE,INVOICENO, INVOICEAMT FROM AMIS_JOURNAL_HD WHERE JTYPE = 'SJ' AND CUSTOMERCODE = '" & XCustomerCode & "'")
'        Me.lvwInvoice.Sorted = True: Me.lvwInvoice.ListItems.Clear: Me.lvwInvoice.Enabled = False
'        Do While Not rsGetInvoice.EOF
'            If lvwUplydList.ListItems.Count = 0 Then
'                Set Item = lvwInvoice.ListItems.Add(, , Null2String(rsGetInvoice!INVOICENO))
'                Item.SubItems(1) = Null2String(rsGetInvoice!InvoiceType)
'                Item.SubItems(2) = Round(ToDoubleNumber(rsGetInvoice!INVOICEAMT), 2)
'            Else
'                For i = 1 To lvwUplydList.ListItems.Count
'                    If lvwUplydList.ListItems(i).ListSubItems(2).Text = Null2String(rsGetInvoice!INVOICENO) And _
                     '                       lvwUplydList.ListItems(i).ListSubItems(1).Text = Null2String(rsGetInvoice!InvoiceType) Then
'                        item_exists = True
'                        Exit For
'                    End If
'                Next
'                If item_exists = False Then
'                    Set Item = lvwInvoice.ListItems.Add(, , Null2String(rsGetInvoice!INVOICENO))
'                    Item.SubItems(1) = Null2String(rsGetInvoice!InvoiceType)
'                    Item.SubItems(2) = Round(ToDoubleNumber(rsGetInvoice!INVOICEAMT), 2)
'                End If
'            End If
'            rsGetInvoice.MoveNext
'        Loop
'        Me.lvwInvoice.Enabled = True: Me.lvwInvoice.Sorted = False: Me.lvwInvoice.Refresh
'
'    End If
End Sub

Private Sub txtInvoiceNo_LostFocus()
'pic_Dropdown.Visible = False
End Sub

Private Sub txtSEARCH_Change()
    Dim rssearch                                  As ADODB.Recordset
    Dim xTOTAL                                    As Double
    grdUnppliedPayment.Rows = 1
    grdUnppliedPayment.AutoRedraw = False

    If Combo1.Text = "" Then
        MessagePop InfoFriend, "INFORMATION", "Please select search option"
        Combo1.SetFocus
        Exit Sub
    End If
    If Combo1.Text = "Customer Name" Then
        If txtSearch = "" Then
            rsRefresh
            StoreMemVars
            Exit Sub
        Else
            Set rssearch = gconDMIS.Execute("SELECT HD.CUSTOMERCODE as Code,CUST.ACCTNAME as AcctName,HD.VOUCHERNO as VoucherNo,HD.INVOICENO as InvoiceNo,HD.INVOICETYPE as InvoiceType,HD.REMARKS as Remarks,HD.INVOICEAMT as InvoiceAmt,HD.INVOICEDATE AS INVDATE From AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_CUSTOMER_TABLE CUST ON HD.CustomerCode = CUST.CUSCDE WHERE JTYPE = 'CRJ'  and CUST.ACCTNAME LIKE '" & RTrim(LTrim(txtSearch.Text)) & "%' AND VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_CRJ_DETAIL WHERE CR_TYPE = 'CRJ') order by AcctName asc")
        End If
    ElseIf Combo1.Text = "Voucher No." Then
        Set rssearch = gconDMIS.Execute("SELECT HD.CUSTOMERCODE as Code,CUST.ACCTNAME as AcctName,HD.VOUCHERNO as VoucherNo,HD.INVOICENO as InvoiceNo,HD.INVOICETYPE as InvoiceType,HD.REMARKS as Remarks,HD.INVOICEAMT as InvoiceAmt, HD.INVOICEDATE AS INVDATE From AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_CUSTOMER_TABLE CUST ON HD.CustomerCode = CUST.CUSCDE WHERE JTYPE = 'CRJ'  and HD.VOUCHERNO LIKE '" & RTrim(LTrim(txtSearch.Text)) & "%' AND VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_CRJ_DETAIL WHERE CR_TYPE = 'CRJ') order by AcctName asc")
    End If

    xTOTAL = 0
    If Not rssearch.EOF And Not rssearch.BOF Then
        Do While Not rssearch.EOF
            grdUnppliedPayment.AddItem rssearch!code & Chr(9) & rssearch!AcctName & Chr(9) & rssearch!VOUCHERNO & Chr(9) & rssearch!INVOICENO & Chr(9) & ToDoubleNumber(rssearch!InvoiceAmt) & Chr(9) & rssearch!InvoiceType & Chr(9) & rssearch!INVDATE & Chr(9) & rssearch!remarks
            xTOTAL = xTOTAL + NumericVal(rssearch!InvoiceAmt)
            rssearch.MoveNext
        Loop
        txtTotalAmount = ToDoubleNumber(xTOTAL)
        grdUnppliedPayment.AutoRedraw = True
        grdUnppliedPayment.Refresh
    Else
        MessagePop InfoFriend, "INFORMATION", "There is no record found"
        Exit Sub
    End If
    Set rssearch = Nothing
End Sub

Sub InitCbo_AR()
    Dim rsAR                                      As ADODB.Recordset
    Set rsAR = gconDMIS.Execute("SELECT DESCRIPTION FROM AMIS_CHARTACCOUNT WHERE LEFT(ACCTCODE,5) = '11-02'")
    If Not rsAR.EOF And Not rsAR.BOF Then
        Do While Not rsAR.EOF
            cboTagAR.AddItem (Null2String(rsAR!Description))
            rsAR.MoveNext
        Loop
    End If
    Set rsAR = Nothing
End Sub

Sub initMemvars()
    txtVoucherNo.Text = ""
    txtInvoiceNo.Text = ""
    txtInvoiceAmt.Text = ""
    txtInvoiceType.Text = ""
    txtInvoiceDate.Text = ""
    cboTagAR.Clear
End Sub

Sub FillDetails()
    Dim rsJournal_Det                             As ADODB.Recordset
    Dim kcnt                                      As Integer
    Dim J_ITemNo                                  As Integer
    lstDetails.Sorted = False: lstDetails.ListItems.Clear
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select id,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax from AMIS_Journal_Det where VoucherNo = " & N2Str2Null(txtVoucherNo.Text) & " and jtype = 'CRJ' order by jitemno asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        Screen.MousePointer = 11
        rsJournal_Det.MoveFirst:                           'TOTAL_AR_AMOUNT = 0
        Do While Not rsJournal_Det.EOF
            kcnt = kcnt + 1
            If Null2String(rsJournal_Det!jitemno) = "" Then J_ITemNo = kcnt Else J_ITemNo = Null2String(rsJournal_Det!jitemno)
            lstDetails.ListItems.Add kcnt, , Format(J_ITemNo, "0000")
            lstDetails.ListItems(kcnt).ListSubItems.Add 1, , Null2String(rsJournal_Det!Acct_code)
            lstDetails.ListItems(kcnt).ListSubItems.Add 2, , Null2String(rsJournal_Det!acct_Name)
            lstDetails.ListItems(kcnt).ListSubItems.Add 3, , ToDoubleNumber(N2Str2Zero(rsJournal_Det!DEBIT))
            If Left(Null2String(rsJournal_Det!Acct_code), 5) = "11-02" Or Left(Null2String(rsJournal_Det!Acct_code), 5) = "11-03" Then
                TOTAL_AR_AMOUNT = TOTAL_AR_AMOUNT + N2Str2Zero(rsJournal_Det!CREDIT)
                TotalARAmountToPay = TotalARAmountToPay + N2Str2Zero(rsJournal_Det!DEBIT)
            End If
            '                If Left(Null2String(rsJournal_Det!acct_code), 5) = "21-01" Or Left(Null2String(rsJournal_Det!acct_code), 5) = "21-02" Then
            '                    TOTAL_AP_AMOUNT = TOTAL_AP_AMOUNT + N2Str2Zero(rsJournal_Det!CREDIT)
            '                    TotalAPAmountToPay = TotalAPAmountToPay + N2Str2Zero(rsJournal_Det!DEBIT)
            '                End If
            lstDetails.ListItems(kcnt).ListSubItems.Add 4, , ToDoubleNumber(N2Str2Zero(rsJournal_Det!CREDIT))
            lstDetails.ListItems(kcnt).ListSubItems.Add 5, , rsJournal_Det!ID
            'If xJOURNALTYPE = "SJ" Or xJOURNALTYPE = "CSJ" Then COMP_SJ_OUTPUT_TAX = 0
            TOTDEBIT = TOTDEBIT + Round(NumericVal(N2Str2Zero(rsJournal_Det!DEBIT)), 2)
            TOTCREDIT = TOTCREDIT + Round(NumericVal(N2Str2Zero(rsJournal_Det!CREDIT)), 2)
            TOTTAX = TOTTAX + NumericVal(N2Str2Zero(rsJournal_Det!tax))
            rsJournal_Det.MoveNext
        Loop
        lstDetails.Sorted = True: lstDetails.Refresh
        '            txtTotDebit.Text = ToDoubleNumber(TOTDEBIT)
        '            txtTotCredit.Text = ToDoubleNumber(TOTCREDIT)
        '            OUTBALANCE = Round(TOTDEBIT - TOTCREDIT, 2)
    End If
    Screen.MousePointer = 0
    Set rsJournal_Det = Nothing
End Sub


Sub SearchVoucherNo(xVOUCHERNO As String)
    Dim rsSearchVoucherNo                         As ADODB.Recordset
    Dim xxTOTAL                                   As Double
    Set rsSearchVoucherNo = gconDMIS.Execute("SELECT HD.CUSTOMERCODE as Code,CUST.ACCTNAME as AcctName,HD.VOUCHERNO as VoucherNo,HD.INVOICENO as InvoiceNo,HD.INVOICETYPE as InvoiceType,HD.REMARKS as Remarks,HD.INVOICEAMT as InvoiceAmt, HD.INVOICEDATE AS INVDATE From AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_CUSTOMER_TABLE CUST ON HD.CustomerCode = CUST.CUSCDE WHERE JTYPE = 'CRJ'  and HD.VOUCHERNO = '" & xVOUCHERNO & "' AND VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_CRJ_DETAIL WHERE CR_TYPE = 'CRJ') order by AcctName asc")
    xxTOTAL = 0
    If Not rsSearchVoucherNo.EOF And Not rsSearchVoucherNo.BOF Then
        Do While Not rsSearchVoucherNo.EOF
            grdUnppliedPayment.AddItem rsSearchVoucherNo!code & Chr(9) & rsSearchVoucherNo!AcctName & Chr(9) & rsSearchVoucherNo!VOUCHERNO & Chr(9) & rsSearchVoucherNo!INVOICENO & Chr(9) & ToDoubleNumber(rsSearchVoucherNo!InvoiceAmt) & Chr(9) & rsSearchVoucherNo!InvoiceType & Chr(9) & rsSearchVoucherNo!INVDATE & Chr(9) & rsSearchVoucherNo!remarks
            xxTOTAL = xxTOTAL + NumericVal(rsSearchVoucherNo!InvoiceAmt)
            rsSearchVoucherNo.MoveNext
        Loop
        txtTotalAmount = ToDoubleNumber(xxTOTAL)
        grdUnppliedPayment.AutoRedraw = True
        grdUnppliedPayment.Refresh
    Else
        MessagePop InfoFriend, "INFORMATION", "There is no record found"
        Exit Sub
    End If
    Set rsSearchVoucherNo = Nothing
End Sub
