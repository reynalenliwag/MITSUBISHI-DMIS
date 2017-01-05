VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
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
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
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
      Format          =   54067201
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
      Left            =   1860
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
         Left            =   2520
         TabIndex        =   8
         Top             =   180
         Width           =   225
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
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   2595
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
         Top             =   180
         Width           =   2415
      End
      Begin MSComctlLib.ListView lstCustomer 
         Height          =   7095
         Left            =   60
         TabIndex        =   2
         Top             =   600
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   12515
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
         MouseIcon       =   "CustomersLedger.frx":030A
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
      Left            =   10965
      MouseIcon       =   "CustomersLedger.frx":046C
      MousePointer    =   99  'Custom
      Picture         =   "CustomersLedger.frx":05BE
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
      MouseIcon       =   "CustomersLedger.frx":0924
      MousePointer    =   99  'Custom
      Picture         =   "CustomersLedger.frx":0A76
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
      MouseIcon       =   "CustomersLedger.frx":0DDC
      MousePointer    =   99  'Custom
      Picture         =   "CustomersLedger.frx":0F2E
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
      MouseIcon       =   "CustomersLedger.frx":1228
      MousePointer    =   99  'Custom
      Picture         =   "CustomersLedger.frx":137A
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
      MouseIcon       =   "CustomersLedger.frx":16D2
      MousePointer    =   99  'Custom
      Picture         =   "CustomersLedger.frx":1824
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Move to Previous Record"
      Top             =   7515
      Width           =   705
   End
   Begin VB.Frame fraDetails 
      Height          =   5985
      Left            =   2670
      TabIndex        =   15
      Top             =   1425
      Width           =   9135
      Begin MSComctlLib.ListView lvwLedger 
         Height          =   5175
         Left            =   30
         TabIndex        =   31
         Top             =   150
         Width           =   9045
         _ExtentX        =   15954
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
         Height          =   5085
         Left            =   60
         TabIndex        =   16
         Top             =   180
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   8969
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
         MouseIcon       =   "CustomersLedger.frx":1B83
      End
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   345
      Left            =   9900
      TabIndex        =   30
      Top             =   60
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
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
      Format          =   54067201
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
Dim rsCUSTOMER                                                        As ADODB.Recordset
Dim rsJournal_hd                                                      As ADODB.Recordset
Dim rsJournal_HDDet                                                   As ADODB.Recordset
Dim AddorEdit, ORDER_BY                                               As String
Attribute ORDER_BY.VB_VarUserMemId = 1073938435
Dim TUTAL_DEBIT, TUTAL_CREDIT, TUTAL_BALANCE                          As Double
Attribute TUTAL_DEBIT.VB_VarUserMemId = 1073938437
Attribute TUTAL_CREDIT.VB_VarUserMemId = 1073938437
Attribute TUTAL_BALANCE.VB_VarUserMemId = 1073938437
Dim LocalAcess                                                        As String

Dim rsCUSTOMER_OPENING                                            As ADODB.Recordset
'Dim xBALANCE        As Double

Function SetCustomerName(VVV As Variant)
    Dim rsCustomerDup                                                 As ADODB.Recordset
    Set rsCustomerDup = New ADODB.Recordset
    rsCustomerDup.Open "Select CustCode,Custname from ALL_CUSTMASTER_AMIS where CustCode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCustomerDup.EOF And Not rsCustomerDup.BOF Then SetCustomerName = Null2String(rsCustomerDup!CUSTNAME) Else SetCustomerName = ""
End Function

Sub rsRefresh()
    Set rsCUSTOMER = New ADODB.Recordset
    If cboAccountName.Text = "ALL ACCOUNTS" Then
        rsCUSTOMER.Open "SELECT DISTINCT dbo.ALL_Customer.ACCTNAME as CUSTNAME,dbo.ALL_Customer.ID,dbo.ALL_Customer.CUSCDE AS CUSTCODE FROM dbo.AMIS_Journal_HD left outer JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType INNER JOIN dbo.ALL_Customer ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer.CUSCDE WHERE ((LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') or (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03')) ORDER BY dbo.ALL_Customer.ACCTNAME", gconDMIS, adOpenKeyset
    Else
        rsCUSTOMER.Open "SELECT DISTINCT dbo.ALL_Customer.ACCTNAME as CUSTNAME,dbo.ALL_Customer.ID,dbo.ALL_Customer.CUSCDE AS CUSTCODE FROM dbo.AMIS_Journal_HD left outer JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType INNER JOIN dbo.ALL_Customer ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer.CUSCDE WHERE ((LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') or (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03')) AND dbo.AMIS_Journal_Det.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "' ORDER BY dbo.ALL_Customer.ACCTNAME", gconDMIS, adOpenKeyset
    End If
End Sub

Sub InitMemVars()
    Frame1.Enabled = True
    txtCode.Text = "": txtCode1.Text = "": txtCode2.Text = "": txtCode3.Text = ""
    txtCustName.Text = "":
    txtTotalDebit.Text = ZERO: txtTotalCredit.Text = ZERO
    txtTotalBalance.Text = ZERO:
End Sub

Sub StoreMemvars()
    If Not rsCUSTOMER.EOF And Not rsCUSTOMER.BOF Then
        Frame1.Enabled = False
        labID.Caption = rsCUSTOMER!ID
        txtCode.Text = Null2String(rsCUSTOMER!CUSTCODE)
        txtCustName.Text = Null2String(rsCUSTOMER!CUSTNAME)
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
    'DESCRIPTION: CUSTOMER OPENING BALANCE
    
    
    If cboAccountName.Text = "ALL ACCOUNTS" Then
        rsCUSTOMER_OPENING.Open "SELECT ROUND(SUM(DET.DEBIT) - SUM(DET.CREDIT),2) AS CUST_BALANCE FROM AMIS_JOURNAL_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET DET ON DET.JNO  = HD.JNO AND DET.JTYPE = HD.JTYPE WHERE (HD.JDATE < '" & dtFrom & "'  AND HD.STATUS = 'P' AND HD.CUSTOMERCODE = '" & txtCode.Text & "') " & _
                                "AND ((LEFT(DET.ACCT_CODE,5) in ('11-02' , '11-03' ,'21-02' ,'21-07'))OR ((HD.JTYPE = 'CCM' AND (LEFT(DET.ACCT_CODE,5) in ('11-02','21-01'))) OR (HD.JTYPE = 'CSJ' AND (LEFT(DET.ACCT_CODE,5) in ( '11-03' , '21-07'))) ))", gconDMIS
        Dim xCOB As Double
        Dim rsCHECK_COB As ADODB.Recordset
        Set rsCHECK_COB = New ADODB.Recordset
        
        xCOB = 0
            
        rsCHECK_COB.Open "SELECT HD.JTYPE as xJTYPE,HD.INVOICEAMT as xINVOICE_AMT FROM AMIS_JOURNAL_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET DET ON DET.JNO  = HD.JNO AND DET.JTYPE = HD.JTYPE WHERE (HD.JDATE < '" & dtFrom & "'  AND HD.STATUS = 'P' AND HD.CUSTOMERCODE = '" & txtCode.Text & "') " & _
                                "AND ((LEFT(DET.ACCT_CODE,5) in ('11-02' , '11-03' ,'21-02' ,'21-07'))OR ((HD.JTYPE = 'CCM' AND (LEFT(DET.ACCT_CODE,5) in ('11-02','21-01'))) OR (HD.JTYPE = 'CSJ' AND (LEFT(DET.ACCT_CODE,5) in ( '11-03' , '21-07'))) ))", gconDMIS
        If Not rsCHECK_COB.EOF And Not rsCHECK_COB.BOF Then
            Do While Not rsCHECK_COB.EOF
                If Null2String(rsCHECK_COB!xJtype) = "COB" Then
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
    Else
        Dim rsCUSTOMER_OPENING_ACCT As ADODB.Recordset
        Dim rsCHECK_COB_ACCT As ADODB.Recordset
        Dim xCOB_ACCT As Double
        
        Set rsCUSTOMER_OPENING_ACCT = New ADODB.Recordset
        rsCUSTOMER_OPENING_ACCT.Open "SELECT ROUND(SUM(DET.DEBIT) - SUM(DET.CREDIT),2) AS CUST_BALANCE FROM AMIS_JOURNAL_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET DET ON DET.JNO  = HD.JNO AND DET.JTYPE = HD.JTYPE WHERE (HD.JDATE < '" & dtFrom & "'  AND HD.STATUS = 'P' AND DET.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "'  AND HD.CUSTOMERCODE = '" & txtCode.Text & "') " & _
                                "AND ((LEFT(DET.ACCT_CODE,5) in ('11-02' , '11-03' ,'21-02' ,'21-07'))OR ((HD.JTYPE = 'CCM' AND (LEFT(DET.ACCT_CODE,5) in ('11-02','21-01'))) OR (HD.JTYPE = 'CSJ' AND (LEFT(DET.ACCT_CODE,5) in ( '11-03' , '21-07'))) ))", gconDMIS

        xCOB_ACCT = 0
        Set rsCHECK_COB_ACCT = New ADODB.Recordset
        rsCHECK_COB_ACCT.Open "SELECT HD.JTYPE as xJTYPE,HD.INVOICEAMT as xINVOICE_AMT FROM AMIS_JOURNAL_HD HD LEFT OUTER JOIN AMIS_JOURNAL_DET DET ON DET.JNO  = HD.JNO AND DET.JTYPE = HD.JTYPE WHERE (HD.JDATE < '" & dtFrom & "'  AND HD.STATUS = 'P'AND DET.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "' AND HD.CUSTOMERCODE = '" & txtCode.Text & "') " & _
                                "AND ((LEFT(DET.ACCT_CODE,5) in ('11-02' , '11-03' ,'21-02' ,'21-07'))OR ((HD.JTYPE = 'CCM' AND (LEFT(DET.ACCT_CODE,5) in ('11-02','21-01'))) OR (HD.JTYPE = 'CSJ' AND (LEFT(DET.ACCT_CODE,5) in ( '11-03' , '21-07'))) ))", gconDMIS
        
        If Not rsCHECK_COB_ACCT.EOF And Not rsCHECK_COB_ACCT.BOF Then
            Do While Not rsCHECK_COB_ACCT.EOF
                If Null2String(rsCHECK_COB_ACCT!xJtype) = "COB" Then
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

'        If Not rsCUSTOMER_OPENING.BOF And Not rsCUSTOMER_OPENING.EOF Then
'            If Null2String(rsCUSTOMER_OPENING!CUST_BALANCE) = "" Then
'                xBALANCE = ToDoubleNumber(0)
'            Else
'                xBALANCE = ToDoubleNumber(rsCUSTOMER_OPENING!CUST_BALANCE)
'            End If
'        Else
'            xBALANCE = ToDoubleNumber(0)
'        End If
    End If
    
    Set rsCUSTOMER_OPENING = Nothing
    Set rsCHECK_COB = Nothing
End Sub


Sub FillGrids()
    Dim OUTBALANCE                                                    As Double
    Dim Reference                                                     As String
    Dim cnt                                                           As Integer
    Dim CREDIT                                                        As Double
    Dim DEBIT                                                         As Double
    Dim cnt_adjusment As Integer
    Dim tmp_voucher As String
    
    
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
                    grdAccountsLedger.AddItem Null2String(rsJournal_HDDet!jdate) & Chr(9) & _
                                          Reference & Chr(9) & _
                                        " " & Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO) & Chr(9) & _
                                          "0.00" & Chr(9) & _
                                          ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!CM)) & Chr(9) & _
                                          ToDoubleNumber(OUTBALANCE) & Chr(9) & rsJournal_HDDet!ID & Chr(9) & Null2String(rsJournal_HDDet!jtype)
                 
            ElseIf Null2String(rsJournal_HDDet!jtype) = "CSJ" Then
                    grdAccountsLedger.AddItem Null2String(rsJournal_HDDet!jdate) & Chr(9) & _
                                          Reference & Chr(9) & _
                                        " " & Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO) & Chr(9) & _
                                          ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!DM)) & Chr(9) & _
                                          "0.00" & Chr(9) & _
                                          ToDoubleNumber(OUTBALANCE) & Chr(9) & rsJournal_HDDet!ID & Chr(9) & Null2String(rsJournal_HDDet!jtype)
            ElseIf Null2String(rsJournal_HDDet!jtype) = "COB" Then
                grdAccountsLedger.AddItem Null2String(rsJournal_HDDet!jdate) & Chr(9) & _
                                          Reference & Chr(9) & _
                                        " " & Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO) & Chr(9) & _
                                          ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!amounttopay)) & Chr(9) & _
                                          "0.00" & Chr(9) & _
                                          ToDoubleNumber(OUTBALANCE) & Chr(9) & rsJournal_HDDet!ID & Chr(9) & Null2String(rsJournal_HDDet!jtype)
            Else
                
                grdAccountsLedger.AddItem Null2String(rsJournal_HDDet!jdate) & Chr(9) & _
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
    Dim rsCustomers                                                   As ADODB.Recordset
    lstCustomer.Enabled = False
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set rsCustomers = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))

    'Set rsCustomers = gconDMIS.Execute("SELECT DISTINCT dbo.ALL_Customer.ACCTNAME as CUSTNAME,dbo.ALL_Customer.ID,dbo.ALL_Customer.CUSCDE AS CUSTCODE FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType INNER JOIN dbo.ALL_Customer ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer.CUSCDE WHERE ((LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') OR (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03')) AND dbo.ALL_Customer.ACCTNAME like '" & ReplaceQuote(XXX) & "%' ORDER BY dbo.ALL_Customer.ACCTNAME")
    Set rsCustomers = gconDMIS.Execute("SELECT DISTINCT dbo.ALL_Customer.ACCTNAME as CUSTNAME,dbo.ALL_Customer.ID,dbo.ALL_Customer.CUSCDE AS CUSTCODE FROM dbo.AMIS_Journal_HD left outer JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType INNER JOIN dbo.ALL_Customer ON dbo.AMIS_Journal_HD.CustomerCode = dbo.ALL_Customer.CUSCDE WHERE ((LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') OR (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03')) AND dbo.ALL_Customer.ACCTNAME like '" & ReplaceQuote(XXX) & "%' ORDER BY dbo.ALL_Customer.ACCTNAME")
    If Not (rsCustomers.EOF And rsCustomers.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsCustomers
        lstCustomer.Refresh
        lstCustomer.Enabled = True
        lstCustomer.Enabled = True
    Else
        lstCustomer.Enabled = False
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

    rsCUSTOMER.MoveNext
    If rsCUSTOMER.EOF Then
        rsCUSTOMER.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo Errorcode:

    rsCUSTOMER.MovePrevious
    If rsCUSTOMER.BOF Then
        rsCUSTOMER.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
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
    rsCUSTOMER.MoveFirst
    InitTotal
    rsCUSTOMER.Bookmark = rsFind(rsCUSTOMER.Clone, "ID", lstCustomer.SelectedItem.SubItems(1)).Bookmark
    StoreMemvars
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
    Dim rsSetMinMaxDate                                               As ADODB.Recordset
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
    rsRefresh
    TextSearch.Text = "": Frame2.ZOrder 1
    InitMemVars
    StoreMemvars
    Screen.MousePointer = 0
End Sub

Function Setacctcode(XXX As String) As String
    Dim rsCOA                                                         As ADODB.Recordset
    Set rsCOA = New ADODB.Recordset
    Set rsCOA = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where Description = '" & XXX & "'")
    If Not rsCOA.EOF And Not rsCOA.BOF Then
        Setacctcode = Null2String(rsCOA!acctcode)
    End If
    Set rsCOA = Nothing
End Function

Sub InitCbo()
    Dim rsCOA                                                         As ADODB.Recordset
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
    Dim VARVOUCHERNO                                                  As String
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
    Dim RETURNVOUCHERNO                                               As ADODB.Recordset
    Set RETURNVOUCHERNO = gconDMIS.Execute("Select VoucherNo from AMIS_Journal_HD Where ID = " & NumericVal(grdAccountsLedger.Text))
    If Not RETURNVOUCHERNO.EOF And Not RETURNVOUCHERNO.BOF Then
        VARVOUCHERNO = Null2String(RETURNVOUCHERNO!VOUCHERNO)   'Right(grdAccountsLedger.Text, 6)
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
    rsCUSTOMER.Bookmark = rsFind(rsCUSTOMER.Clone, "ID", lstCustomer.SelectedItem.SubItems(1)).Bookmark
    StoreMemvars
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    rsCUSTOMER.Bookmark = rsFind(rsCUSTOMER.Clone, "ID", lstCustomer.SelectedItem.SubItems(1)).Bookmark
    StoreMemvars
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
    Dim VARVOUCHERNO                                                  As String
    Dim RETURNVOUCHERNO                                               As ADODB.Recordset
    
    JOURNALTYPE = lvwLedger.SelectedItem.SubItems(7)
    Set RETURNVOUCHERNO = gconDMIS.Execute("Select VoucherNo from AMIS_Journal_HD Where ID = " & NumericVal(lvwLedger.SelectedItem.SubItems(6)))
    If Not RETURNVOUCHERNO.EOF And Not RETURNVOUCHERNO.BOF Then
        VARVOUCHERNO = Null2String(RETURNVOUCHERNO!VOUCHERNO)   'Right(grdAccountsLedger.Text, 6)
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

Private Sub textSearch_Change()
    If Trim(TextSearch.Text) = "" Then FillGrid Else FillSearchGrid (TextSearch.Text)
End Sub

Private Sub FillGrid()
    Dim rsCustomers                                                   As ADODB.Recordset
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

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Frame2.ZOrder 0
    If KeyCode = vbKeyDown Then
        If lstCustomer.ListItems.Count > 0 And lstCustomer.Enabled = True Then
            lstCustomer.SetFocus
        End If
    End If
End Sub
Function checkaccount(XXX As String, EVAN_PARAKITO As String)
    Dim rs As New ADODB.Recordset
    Dim cnt As Integer
    Dim Account_code As String
    Set rs = gconDMIS.Execute("select acct_code,voucherno from amis_journal_det where VoucherNo='" & XXX & "' and jtype='" & EVAN_PARAKITO & "'")
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        cnt = 0
        Do While Not rs.EOF
            Account_code = Null2String(rs!ACCT_CODE)
            If Left(Account_code, 5) = "11-02" Or Left(Account_code, 5) = "11-03" Then
                cnt = cnt + 1
                checkaccount = Left(Account_code, 5)
            End If
            checkaccount = cnt
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
End Function


Sub FILL_LEDGER()
    Dim OUTBALANCE                                                    As Double
    Dim Reference                                                     As String
    Dim cnt                                                           As Integer
    Dim CREDIT                                                        As Double
    Dim DEBIT                                                         As Double
    Dim cnt_adjusment As Integer
    Dim tmp_voucher As String
    Dim lvw_COUNT                                                     As Integer
    Dim Item As ListItem
    
    lvw_COUNT = 1
    
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
        If cboAccountName.Text = "A/REC CREDIT CARD" Then
            If txtCode = "C00067" Then
                rsJournal_HDDet.Open "select AMIS_Journal_Hd.Debit AS DM,AMIS_Journal_Hd.Credit AS CM,AMIS_Journal_Hd.AmountToPay,AMIS_Journal_Hd.InvoiceAmt,AMIS_Journal_Hd.InvoiceNo,AMIS_Journal_HD.ID,AMIS_Journal_Hd.JNo,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.JType,AMIS_Journal_Det.Debit,AMIS_Journal_Det.Credit,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.CheckNo,AMIS_Journal_Hd.InvoiceType,AMIS_Journal_Hd.VendorCode,AMIS_Journal_Hd.JNo,AMIS_Journal_Hd.Bank from AMIS_Journal_HD left outer Join AMIS_Journal_det on AMIS_Journal_Det.VoucherNo  = AMIS_Journal_Hd.VoucherNo and AMIS_Journal_Det.jtype = AMIS_Journal_Hd.Jtype  where " & _
                                         "(dbo.AMIS_Journal_HD.Jdate >= '" & dtFrom & "' and dbo.AMIS_Journal_HD.Jdate <= '" & dtTo & "') and (AMIS_Journal_Hd.Status = 'P' AND AMIS_Journal_Det.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "' AND ((Left(AMIS_Journal_Det.Acct_Code,5) = '11-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '11-03' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-07') OR ((AMIS_Journal_HD.JTYPE = 'CCM' AND (left(AMIS_Journal_det.acct_code,5) = '11-01' or left(AMIS_Journal_det.acct_code,5) = '21-01')) OR (AMIS_Journal_HD.JTYPE = 'CSJ' AND (left(AMIS_Journal_det.acct_code,5) = '11-03' or left(AMIS_Journal_det.acct_code,5) = '21-07'))))) AND AMIS_Journal_Hd.Bank = '" & txtCode.Text & "' order by AMIS_Journal_Hd.jdate asc,AMIS_Journal_Hd.id asc", gconDMIS
            Else
                rsJournal_HDDet.Open "SELECT AMIS_vw_JOURNAL_HD.DEBIT AS DM,AMIS_vw_JOURNAL_HD.CREDIT AS CM,AMIS_vw_JOURNAL_HD.AMOUNTTOPAY,AMIS_vw_JOURNAL_HD.INVOICEAMT,AMIS_vw_JOURNAL_HD.INVOICENO,AMIS_vw_JOURNAL_HD.ID,AMIS_vw_JOURNAL_HD.JNO,AMIS_vw_JOURNAL_HD.JDATE,AMIS_vw_JOURNAL_HD.JTYPE,AMIS_JOURNAL_DET.DEBIT,AMIS_JOURNAL_DET.CREDIT,AMIS_vw_JOURNAL_HD.VOUCHERNO,AMIS_vw_JOURNAL_HD.CHECKNO,AMIS_vw_JOURNAL_HD.INVOICETYPE,AMIS_vw_JOURNAL_HD.VENDORCODE,AMIS_vw_JOURNAL_HD.JNO,AMIS_vw_JOURNAL_HD.REFERENCENO FROM AMIS_vw_JOURNAL_HD LEFT OUTER JOIN AMIS_JOURNAL_DET ON AMIS_JOURNAL_DET.VOUCHERNO  = AMIS_vw_JOURNAL_HD.VOUCHERNO AND AMIS_JOURNAL_DET.JTYPE = AMIS_vw_JOURNAL_HD.JTYPE AND AMIS_vw_JOURNAL_HD.REFERENCENO=AMIS_JOURNAL_DET.REFERENCENO WHERE " & _
                                    "(dbo.AMIS_vw_Journal_HD.Jdate >= '" & dtFrom & "' and dbo.AMIS_vw_Journal_HD.Jdate <= '" & dtTo & "') and (AMIS_vw_Journal_Hd.Status = 'P' AND AMIS_Journal_Det.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "' AND ((Left(AMIS_Journal_Det.Acct_Code,5) = '11-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '11-03' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-07') OR ((AMIS_vw_Journal_HD.JTYPE = 'CCM' AND (left(AMIS_Journal_det.acct_code,5) in ('11-02','21-07'))) OR (AMIS_vw_Journal_HD.JTYPE = 'CSJ' AND (left(AMIS_Journal_det.acct_code,5) in ('11-03','21-07')))))) AND AMIS_vw_JOURNAL_HD.REFERENCENO = '" & txtCode.Text & "' order by AMIS_vw_Journal_Hd.jdate asc,AMIS_vw_Journal_Hd.id asc", gconDMIS
            End If
        Else
            rsJournal_HDDet.Open "select AMIS_Journal_Hd.Debit AS DM,AMIS_Journal_Hd.Credit AS CM,AMIS_Journal_Hd.AmountToPay,AMIS_Journal_Hd.InvoiceAmt,AMIS_Journal_Hd.InvoiceNo,AMIS_Journal_HD.ID,AMIS_Journal_Hd.JNo,AMIS_Journal_Hd.JDate,AMIS_Journal_Hd.JType,AMIS_Journal_Det.Debit,AMIS_Journal_Det.Credit,AMIS_Journal_Hd.VoucherNo,AMIS_Journal_Hd.CheckNo,AMIS_Journal_Hd.InvoiceType,AMIS_Journal_Hd.VendorCode,AMIS_Journal_Hd.JNo from AMIS_Journal_HD left outer Join AMIS_Journal_det on AMIS_Journal_Det.JNo  = AMIS_Journal_Hd.JNo and AMIS_Journal_Det.jtype = AMIS_Journal_Hd.Jtype  where " & _
                                 "(dbo.AMIS_Journal_HD.Jdate >= '" & dtFrom & "' and dbo.AMIS_Journal_HD.Jdate <= '" & dtTo & "') and (AMIS_Journal_Hd.Status = 'P' AND AMIS_Journal_Det.Acct_Code = '" & Setacctcode(cboAccountName.Text) & "' AND ((Left(AMIS_Journal_Det.Acct_Code,5) = '11-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '11-03' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-02' or Left(AMIS_Journal_Det.Acct_Code,5) = '21-07') OR ((AMIS_Journal_HD.JTYPE = 'CCM' AND (left(AMIS_Journal_det.acct_code,5) = '11-01' or left(AMIS_Journal_det.acct_code,5) = '21-01')) OR (AMIS_Journal_HD.JTYPE = 'CSJ' AND (left(AMIS_Journal_det.acct_code,5) = '11-03' or left(AMIS_Journal_det.acct_code,5) = '21-07'))))) AND AMIS_Journal_Hd.CustomerCode = '" & txtCode.Text & "' order by AMIS_Journal_Hd.jdate asc,AMIS_Journal_Hd.id asc", gconDMIS
        End If
    End If
    
    Me.lvwLedger.ListItems.Clear: Me.lvwLedger.Enabled = False
    If Not rsJournal_HDDet.EOF And Not rsJournal_HDDet.BOF Then
        'Me.lvwLedger.Sorted = True:
        rsJournal_HDDet.MoveFirst
        Do While Not rsJournal_HDDet.EOF
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
                If Null2String(rsJournal_HDDet!jtype) = "CCM" Then
                        OUTBALANCE = N2Str2Zero(OUTBALANCE) - N2Str2Zero(rsJournal_HDDet!CM)
                ElseIf Null2String(rsJournal_HDDet!jtype) = "CSJ" Then
                        OUTBALANCE = N2Str2Zero(OUTBALANCE) + N2Str2Zero(rsJournal_HDDet!DM)
                Else
                    OUTBALANCE = N2Str2Zero(OUTBALANCE) + ((N2Str2Zero(rsJournal_HDDet!DEBIT) - N2Str2Zero(rsJournal_HDDet!CREDIT)))
                End If
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
         
            If Null2String(rsJournal_HDDet!jtype) = "CCM" Then
               Set Item = lvwLedger.ListItems.Add(, , Null2String(rsJournal_HDDet!jdate))
               Item.SubItems(1) = Null2String(Reference)
               Item.SubItems(2) = Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO)
               Item.SubItems(3) = "0.00"
               Item.SubItems(4) = ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!CM))
               Item.SubItems(5) = ToDoubleNumber(OUTBALANCE)
               Item.SubItems(6) = rsJournal_HDDet!ID
               Item.SubItems(7) = Null2String(rsJournal_HDDet!jtype)
            ElseIf Null2String(rsJournal_HDDet!jtype) = "CSJ" Then
               Set Item = lvwLedger.ListItems.Add(, , Null2String(rsJournal_HDDet!jdate))
               Item.SubItems(1) = Null2String(Reference)
               Item.SubItems(2) = Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO)
               Item.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!DM))
               Item.SubItems(4) = "0.00"
               Item.SubItems(5) = ToDoubleNumber(OUTBALANCE)
               Item.SubItems(6) = rsJournal_HDDet!ID
               Item.SubItems(7) = Null2String(rsJournal_HDDet!jtype)
            ElseIf Null2String(rsJournal_HDDet!jtype) = "COB" Then
               Set Item = lvwLedger.ListItems.Add(, , Null2String(rsJournal_HDDet!jdate))
               Item.SubItems(1) = Null2String(Reference)
               Item.SubItems(2) = Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO)
               Item.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!amounttopay))
               Item.SubItems(4) = "0.00"
               Item.SubItems(5) = ToDoubleNumber(OUTBALANCE)
               Item.SubItems(6) = rsJournal_HDDet!ID
               Item.SubItems(7) = Null2String(rsJournal_HDDet!jtype)
            Else
               Set Item = lvwLedger.ListItems.Add(, , Null2String(rsJournal_HDDet!jdate))
               Item.SubItems(1) = Null2String(Reference)
               Item.SubItems(2) = Null2String(rsJournal_HDDet!CheckNo) & Null2String(rsJournal_HDDet!INVOICENO)
               Item.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!DEBIT))
               Item.SubItems(4) = ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!CREDIT))
               Item.SubItems(5) = ToDoubleNumber(OUTBALANCE)
               Item.SubItems(6) = rsJournal_HDDet!ID
               Item.SubItems(7) = Null2String(rsJournal_HDDet!jtype)
               Item.SubItems(8) = Null2String(rsJournal_HDDet!InvoiceType)
               
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
            Else
               TUTAL_DEBIT = TUTAL_DEBIT + N2Str2Zero(rsJournal_HDDet!DEBIT)
               TUTAL_CREDIT = TUTAL_CREDIT + N2Str2Zero(rsJournal_HDDet!CREDIT)
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
                OUTBALANCE = xBALANCE
    End If
    
    Me.lvwLedger.Enabled = True: Me.lvwLedger.Sorted = False: Me.lvwLedger.Refresh
    
    txtTotalDebit.Text = ToDoubleNumber(TUTAL_DEBIT)
    txtTotalCredit.Text = ToDoubleNumber(TUTAL_CREDIT)
    txtTotalBalance.Text = ToDoubleNumber(TUTAL_BALANCE + N2Str2Zero(OUTBALANCE))
End Sub

Function CHECK_PAYMENT(xVoucherNo As String, xJtype As String) As Boolean
    Dim rsINV_INVTYPE As ADODB.Recordset
    Dim rsCHECK_PAYMENT  As ADODB.Recordset
    Set rsINV_INVTYPE = gconDMIS.Execute("SELECT INVOICENO,INVOICETYPE FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & xVoucherNo & "' AND JTYPE = '" & xJtype & "'")
        If Not rsINV_INVTYPE.EOF And Not rsINV_INVTYPE.BOF Then
            Set rsCHECK_PAYMENT = gconDMIS.Execute("SELECT * FROM AMIS_CRJ_DETAIL WHERE INVOICENO = '" & Null2String(rsINV_INVTYPE!INVOICENO) & "' AND INVOICETYPE = '" & Null2String(rsINV_INVTYPE!InvoiceType) & "'")
            If Not rsCHECK_PAYMENT.BOF And Not rsCHECK_PAYMENT.EOF Then
                CHECK_PAYMENT = True
            Else
                CHECK_PAYMENT = False
            End If
        Else
            'vocher number has no header
        End If
    Set rsINV_INVTYPE = Nothing
End Function

