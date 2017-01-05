VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISCustomerAdjustment 
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Adjustments"
   ClientHeight    =   6360
   ClientLeft      =   585
   ClientTop       =   330
   ClientWidth     =   9510
   ForeColor       =   &H8000000F&
   Icon            =   "CustomerAdjustment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   9510
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
      Left            =   6660
      MouseIcon       =   "CustomerAdjustment.frx":08CA
      MousePointer    =   99  'Custom
      Picture         =   "CustomerAdjustment.frx":0A1C
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   "Move to Previous Record"
      Top             =   5550
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
      Left            =   7350
      MouseIcon       =   "CustomerAdjustment.frx":0D7B
      MousePointer    =   99  'Custom
      Picture         =   "CustomerAdjustment.frx":0ECD
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "Move to Next Record"
      Top             =   5550
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
      Left            =   8040
      MouseIcon       =   "CustomerAdjustment.frx":1225
      MousePointer    =   99  'Custom
      Picture         =   "CustomerAdjustment.frx":1377
      Style           =   1  'Graphical
      TabIndex        =   68
      ToolTipText     =   "Find a Record"
      Top             =   5550
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
      Left            =   5730
      MouseIcon       =   "CustomerAdjustment.frx":1671
      MousePointer    =   99  'Custom
      Picture         =   "CustomerAdjustment.frx":17C3
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "Add Record"
      Top             =   6870
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
      Left            =   6480
      MouseIcon       =   "CustomerAdjustment.frx":1AD6
      MousePointer    =   99  'Custom
      Picture         =   "CustomerAdjustment.frx":1C28
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Edit Selected Record"
      Top             =   6840
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
      Left            =   7290
      MouseIcon       =   "CustomerAdjustment.frx":1F84
      MousePointer    =   99  'Custom
      Picture         =   "CustomerAdjustment.frx":20D6
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Delete Selected Record"
      Top             =   6840
      Width           =   705
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
      Left            =   8730
      MouseIcon       =   "CustomerAdjustment.frx":2401
      MousePointer    =   99  'Custom
      Picture         =   "CustomerAdjustment.frx":2553
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Exit Window"
      Top             =   5550
      Width           =   705
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   3285
      Left            =   2700
      TabIndex        =   27
      Top             =   -30
      Width           =   6705
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   2070
         ScaleHeight     =   825
         ScaleWidth      =   4425
         TabIndex        =   39
         Top             =   1410
         Width           =   4455
         Begin VB.TextBox txtCredit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   390
            Left            =   2250
            MaxLength       =   17
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   330
            Width           =   2055
         End
         Begin VB.TextBox txtDebit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   390
            Left            =   60
            MaxLength       =   17
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   330
            Width           =   2055
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "CREDIT"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2310
            TabIndex        =   41
            Top             =   30
            Width           =   1935
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DEBIT"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   150
            TabIndex        =   40
            Top             =   30
            Width           =   1935
         End
      End
      Begin VB.TextBox txtCustCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   180
         Width           =   855
      End
      Begin VB.TextBox txtParticular 
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
         Height          =   720
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "CustomerAdjustment.frx":28B9
         Top             =   2460
         Width           =   6585
      End
      Begin VB.TextBox txtAccountNo 
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
         Left            =   4350
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   180
         Width           =   2175
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
         Left            =   1710
         MaxLength       =   150
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   630
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Balances"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   60
         TabIndex        =   38
         Top             =   1110
         Width           =   1995
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   6750
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label7 
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
         Left            =   0
         TabIndex        =   34
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Particulars"
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
         Left            =   90
         TabIndex        =   32
         Top             =   2190
         Width           =   1785
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   6720
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account No"
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
         Left            =   2760
         TabIndex        =   31
         Top             =   210
         Width           =   1455
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
         Left            =   4710
         TabIndex        =   30
         Top             =   720
         Width           =   225
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
         Left            =   3780
         TabIndex        =   29
         Top             =   690
         Width           =   465
      End
      Begin VB.Label Label3 
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
         Left            =   60
         TabIndex        =   28
         Top             =   720
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2205
      Left            =   2700
      TabIndex        =   42
      Top             =   3300
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   3889
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "[F3 - Add Debit Memo]"
      TabPicture(0)   =   "CustomerAdjustment.frx":28BF
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picOpenInv"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOpenInv"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstOpenInvoice"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "[F4 - Add Credit Memo]"
      TabPicture(1)   =   "CustomerAdjustment.frx":28DB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstCreditMemo"
      Tab(1).Control(1)=   "cmdCreditMemo"
      Tab(1).Control(2)=   "picCreditMemo"
      Tab(1).ControlCount=   3
      Begin MSComctlLib.ListView lstOpenInvoice 
         Height          =   1755
         Left            =   60
         TabIndex        =   6
         Top             =   90
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   3096
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "CustomerAdjustment.frx":28F7
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Terms"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Invoice #"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Dealer"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Invoice Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin MSComctlLib.ListView lstCreditMemo 
         Height          =   1755
         Left            =   -74940
         TabIndex        =   7
         Top             =   60
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   3096
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "CustomerAdjustment.frx":2A59
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Reference"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ref Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "OR No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "OR Date"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.CommandButton cmdOpenInv 
         Caption         =   "Command1"
         Height          =   1725
         Left            =   390
         TabIndex        =   43
         Top             =   90
         Width           =   5925
      End
      Begin VB.CommandButton cmdCreditMemo 
         Caption         =   "Command1"
         Height          =   1725
         Left            =   -74610
         TabIndex        =   50
         Top             =   90
         Width           =   5925
      End
      Begin VB.PictureBox picOpenInv 
         Height          =   1605
         Left            =   450
         ScaleHeight     =   1545
         ScaleWidth      =   5745
         TabIndex        =   44
         Top             =   150
         Width           =   5805
         Begin VB.TextBox txtInvoiceType 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   9
            Text            =   "Text2"
            Top             =   330
            Width           =   1755
         End
         Begin wizButton.cmd cmdOISave 
            Height          =   315
            Left            =   3900
            TabIndex        =   15
            Top             =   1200
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            TX              =   "&Save"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "CustomerAdjustment.frx":2BBB
         End
         Begin VB.Frame Frame2 
            Caption         =   "Frame2"
            Height          =   1875
            Left            =   2850
            TabIndex        =   56
            Top             =   -150
            Width           =   30
         End
         Begin VB.TextBox txtOIParticular 
            Appearance      =   0  'Flat
            Height          =   555
            Left            =   2970
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Text            =   "CustomerAdjustment.frx":2BD7
            Top             =   630
            Width           =   2745
         End
         Begin VB.TextBox txtOIAmount 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3930
            TabIndex        =   13
            Text            =   "Text2"
            Top             =   30
            Width           =   1755
         End
         Begin VB.TextBox txtInvoiceDate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   12
            Text            =   "Text2"
            Top             =   1230
            Width           =   1755
         End
         Begin VB.TextBox txtDealer 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   11
            Text            =   "Text2"
            Top             =   930
            Width           =   1755
         End
         Begin VB.TextBox txtInvoiceNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   10
            Text            =   "Text2"
            Top             =   630
            Width           =   1755
         End
         Begin VB.TextBox txtTerms 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   8
            Text            =   "Text2"
            Top             =   30
            Width           =   675
         End
         Begin wizButton.cmd cmdOICancel 
            Height          =   315
            Left            =   4800
            TabIndex        =   16
            Top             =   1200
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            TX              =   "&Cancel"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "CustomerAdjustment.frx":2BDD
         End
         Begin wizButton.cmd cmdOIDelete 
            Height          =   315
            Left            =   2970
            TabIndex        =   17
            Top             =   1200
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            TX              =   "&Delete"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "CustomerAdjustment.frx":2BF9
         End
         Begin VB.Label Label9 
            Caption         =   "Invoice Type"
            Height          =   255
            Left            =   60
            TabIndex        =   61
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label labCMID 
            Caption         =   "Label20"
            Height          =   255
            Left            =   1950
            TabIndex        =   60
            Top             =   600
            Width           =   495
         End
         Begin VB.Label labOIID 
            Caption         =   "Label20"
            Height          =   165
            Left            =   1950
            TabIndex        =   59
            Top             =   600
            Width           =   285
         End
         Begin VB.Label Label16 
            Caption         =   "Particular"
            Height          =   255
            Left            =   2970
            TabIndex        =   55
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label Label13 
            Caption         =   "Amount"
            Height          =   255
            Left            =   2970
            TabIndex        =   49
            Top             =   60
            Width           =   1155
         End
         Begin VB.Label Label12 
            Caption         =   "Invoice Date"
            Height          =   255
            Left            =   60
            TabIndex        =   48
            Top             =   1260
            Width           =   1155
         End
         Begin VB.Label Label11 
            Caption         =   "Dealer"
            Height          =   255
            Left            =   60
            TabIndex        =   47
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label Label10 
            Caption         =   "Invoice #"
            Height          =   255
            Left            =   60
            TabIndex        =   46
            Top             =   660
            Width           =   1155
         End
         Begin VB.Label Label5 
            Caption         =   "Terms"
            Height          =   255
            Left            =   60
            TabIndex        =   45
            Top             =   60
            Width           =   555
         End
      End
      Begin VB.PictureBox picCreditMemo 
         BackColor       =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         ForeColor       =   &H8000000A&
         Height          =   1635
         Left            =   -74490
         ScaleHeight     =   1575
         ScaleWidth      =   5745
         TabIndex        =   51
         Top             =   150
         Width           =   5805
         Begin VB.TextBox txtCMORDate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   21
            Text            =   "Text2"
            Top             =   930
            Width           =   1755
         End
         Begin VB.TextBox txtCMORNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   20
            Text            =   "Text2"
            Top             =   630
            Width           =   1755
         End
         Begin VB.Frame Frame3 
            Caption         =   "Frame3"
            Height          =   1725
            Left            =   2850
            TabIndex        =   58
            Top             =   -120
            Width           =   30
         End
         Begin VB.TextBox txtCMParticular 
            Appearance      =   0  'Flat
            Height          =   825
            Left            =   2940
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Text            =   "CustomerAdjustment.frx":2C15
            Top             =   300
            Width           =   2745
         End
         Begin VB.TextBox txtCMReference 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   18
            Text            =   "Text2"
            Top             =   30
            Width           =   1755
         End
         Begin VB.TextBox txtCMRefDate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   19
            Text            =   "Text2"
            Top             =   330
            Width           =   1755
         End
         Begin VB.TextBox txtCMAmount 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   22
            Text            =   "Text2"
            Top             =   1230
            Width           =   1755
         End
         Begin wizButton.cmd cmdCMSave 
            Height          =   315
            Left            =   3900
            TabIndex        =   24
            Top             =   1200
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            TX              =   "Save"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "CustomerAdjustment.frx":2C1B
         End
         Begin wizButton.cmd cmdCMCancel 
            Height          =   315
            Left            =   4800
            TabIndex        =   25
            Top             =   1200
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            TX              =   "Cancel"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "CustomerAdjustment.frx":2C37
         End
         Begin wizButton.cmd cmdCMDelete 
            Height          =   315
            Left            =   2970
            TabIndex        =   26
            Top             =   1200
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            TX              =   "Delete"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "CustomerAdjustment.frx":2C53
         End
         Begin VB.Label Label20 
            Caption         =   "OR Date"
            Height          =   255
            Left            =   60
            TabIndex        =   63
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label Label19 
            Caption         =   "OR No."
            Height          =   255
            Left            =   60
            TabIndex        =   62
            Top             =   660
            Width           =   1155
         End
         Begin VB.Label Label18 
            Caption         =   "Particular"
            Height          =   255
            Left            =   2970
            TabIndex        =   57
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label17 
            Caption         =   "Reference"
            Height          =   255
            Left            =   60
            TabIndex        =   54
            Top             =   60
            Width           =   1155
         End
         Begin VB.Label Label15 
            Caption         =   "Ref. Date"
            Height          =   255
            Left            =   60
            TabIndex        =   53
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label14 
            Caption         =   "Amount"
            Height          =   255
            Left            =   60
            TabIndex        =   52
            Top             =   1260
            Width           =   1155
         End
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   5340
      Left            =   2700
      ScaleHeight     =   5280
      ScaleWidth      =   2535
      TabIndex        =   33
      Top             =   60
      Width           =   2595
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   11640
         Left            =   2520
         Picture         =   "CustomerAdjustment.frx":2C6F
         Top             =   -180
         Width           =   2535
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   5565
      Left            =   60
      TabIndex        =   35
      Top             =   -30
      Width           =   2595
      Begin VB.TextBox TextSearch 
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
         Left            =   60
         MaxLength       =   19
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   180
         Width           =   2475
      End
      Begin MSComctlLib.ListView lstCustomer 
         Height          =   4905
         Left            =   30
         TabIndex        =   37
         Top             =   570
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   8652
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
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
         MouseIcon       =   "CustomerAdjustment.frx":12B67
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CUSTOMER NAME"
            Object.Width           =   8290
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAMISCustomerAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCustomer                                    As ADODB.Recordset
Dim AddOrEditOI                                   As String
Dim AddOrEditCM                                   As String

Function NewOpenInvoiceVoucherNo() As String
    Dim rsAMIS_Journal_HDNew                      As ADODB.Recordset
    Set rsAMIS_Journal_HDNew = New ADODB.Recordset
    Set rsAMIS_Journal_HDNew = gconDMIS.Execute("Select * from AMIS_Journal_HD Where Jtype = 'CSJ' order by VoucherNo Desc")
    If Not rsAMIS_Journal_HDNew.EOF And Not rsAMIS_Journal_HDNew.BOF Then
        NewOpenInvoiceVoucherNo = Format(N2Str2Zero(rsAMIS_Journal_HDNew!VOUCHERNO) + 1, "000000")
    Else
        NewOpenInvoiceVoucherNo = "000001"
    End If
End Function

Function NewCreditMemoVoucherNo() As String
    Dim rsAMIS_Journal_HDNew                      As ADODB.Recordset
    Set rsAMIS_Journal_HDNew = New ADODB.Recordset
    Set rsAMIS_Journal_HDNew = gconDMIS.Execute("Select * from AMIS_Journal_HD Where Jtype = 'CCM' order by VoucherNo Desc")
    If Not rsAMIS_Journal_HDNew.EOF And Not rsAMIS_Journal_HDNew.BOF Then
        NewCreditMemoVoucherNo = Format(N2Str2Zero(rsAMIS_Journal_HDNew!VOUCHERNO) + 1, "000000")
    Else
        NewCreditMemoVoucherNo = "000001"
    End If
End Function

Function NewOpenJNo() As String
    Dim rsAMIS_Journal_HDNew                      As ADODB.Recordset
    Set rsAMIS_Journal_HDNew = New ADODB.Recordset
    Set rsAMIS_Journal_HDNew = gconDMIS.Execute("Select * from AMIS_Journal_HD order by JNo Desc")
    If Not rsAMIS_Journal_HDNew.EOF And Not rsAMIS_Journal_HDNew.BOF Then
        NewOpenJNo = Format(N2Str2Zero(rsAMIS_Journal_HDNew!JNo) + 1, "000000")
    Else
        NewOpenJNo = "000001"
    End If
End Function

Function SetCustomerName(CCC As Variant)
    Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "Select custcode,custname from All_Customer where custcode = " & N2Str2Null(CCC), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerName = Null2String(rsCustomer!CUSTNAME)
    Else
        SetCustomerName = ""
    End If
End Function

Sub rsRefresh()
    Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "select * from all_Customer Order by CUSCDE asc", gconDMIS, adOpenKeyset
End Sub

Sub StoreMemVars()
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        txtCustCode.Text = Null2String(rsCustomer!CUSCDE)
        txtAccountNo.Text = Null2String(rsCustomer!ACCOUNTNO)
        txtCustName.Text = Null2String(rsCustomer!AcctName)
        txtDebit.Text = 0
        txtCredit.Text = 0
        txtParticular.Text = 0
        FillOIGrid
        FillCMGrid
    End If
End Sub

Sub initMemvars()
    txtCustCode.Text = ""
    txtAccountNo.Text = ""
    txtCustName.Text = ""
    txtDebit.Text = ""
    txtCredit.Text = ""
    txtParticular.Text = ""
End Sub

Sub FillOIGrid()
    Dim rsOpenInvoice                             As ADODB.Recordset
    lstOpenInvoice.Sorted = False: lstOpenInvoice.ListItems.Clear
    lstOpenInvoice.Enabled = False
    Set rsOpenInvoice = New ADODB.Recordset
    Set rsOpenInvoice = gconDMIS.Execute("select Terms,InvoiceNo,Dealer,InvoiceDate,InvoiceAmt,ID from AMIS_Journal_HD where jtype = 'CSJ' and customercode = " & N2Str2Null(txtCustCode.Text) & " order by InvoiceNo asc")
    If Not (rsOpenInvoice.EOF And rsOpenInvoice.BOF) Then
        lstOpenInvoice.Enabled = True
        Listview_Loadval Me.lstOpenInvoice.ListItems, rsOpenInvoice
        lstOpenInvoice.Refresh
        lstOpenInvoice.Enabled = True
    Else
        lstOpenInvoice.Enabled = False
    End If
    'Update By Btt: 07302008
    Set rsOpenInvoice = New ADODB.Recordset
    Set rsOpenInvoice = gconDMIS.Execute("select SUM(InvoiceAmt) AS TOTALDEBIT from AMIS_Journal_HD where jtype = 'CSJ' and customercode = " & N2Str2Null(txtCustCode.Text))
    If Not (rsOpenInvoice.EOF And rsOpenInvoice.BOF) Then
        txtDebit.Text = ToDoubleNumber(N2Str2Zero(rsOpenInvoice!TotalDebit))
    End If

End Sub

Sub StoreOIMemvars(XXX As Variant)
    Dim rsOpenInvoice                             As ADODB.Recordset
    Set rsOpenInvoice = New ADODB.Recordset
    Set rsOpenInvoice = gconDMIS.Execute("Select * from AMIS_Journal_HD Where ID = " & XXX)
    If Not rsOpenInvoice.EOF And Not rsOpenInvoice.BOF Then
        AddOrEditOI = "EDIT"
        SSTab1.Tab = 0
        cmdOpenInv.ZOrder 0: picOpenInv.ZOrder 0
        labOIID.Caption = rsOpenInvoice!ID
        txtTerms.Text = Null2String(rsOpenInvoice!TERMS)
        txtInvoiceType.Text = Null2String(rsOpenInvoice!InvoiceType)
        txtInvoiceNo.Text = Null2String(rsOpenInvoice!INVOICENO)
        txtDealer.Text = Null2String(rsOpenInvoice!Dealer)
        txtInvoiceDate.Text = Null2String(rsOpenInvoice!invoicedate)
        txtOIAmount.Text = Null2String(rsOpenInvoice!InvoiceAmt)
        txtOIParticular.Text = Null2String(rsOpenInvoice!remarks)
        On Error Resume Next
        txtTerms.SetFocus
    End If
End Sub

Sub InitOIMemvars()
    txtTerms.Text = ""
    txtInvoiceType.Text = ""
    txtInvoiceNo.Text = ""
    txtDealer.Text = ""
    txtInvoiceDate.Text = ""
    txtOIAmount.Text = ""
    txtOIParticular.Text = ""
End Sub

Sub FillCMGrid()
    Dim rsCreditMemo                              As ADODB.Recordset
    lstCreditMemo.Enabled = False
    lstCreditMemo.Sorted = False: lstCreditMemo.ListItems.Clear
    lstCreditMemo.Enabled = False
    Set rsCreditMemo = New ADODB.Recordset
    Set rsCreditMemo = gconDMIS.Execute("select RefNo,RefDate,InvoiceNo,InvoiceDate,InvoiceAmt,ID from AMIS_Journal_HD where Jtype = 'CCM' and customercode = " & N2Str2Null(txtCustCode.Text) & " order by RefDate asc")
    If Not (rsCreditMemo.EOF And rsCreditMemo.BOF) Then
        lstCreditMemo.Enabled = True
        Listview_Loadval Me.lstCreditMemo.ListItems, rsCreditMemo
        lstCreditMemo.Refresh
        lstCreditMemo.Enabled = True
    Else
        lstCreditMemo.Enabled = False
    End If

    Set rsCreditMemo = New ADODB.Recordset
    Set rsCreditMemo = gconDMIS.Execute("select SUM(invoiceAMT) AS TOTALCredit from AMIS_Journal_HD where jtype = 'CCM' and customercode = " & N2Str2Null(txtCustCode.Text))
    If Not (rsCreditMemo.EOF And rsCreditMemo.BOF) Then
        txtCredit.Text = ToDoubleNumber(N2Str2Zero(rsCreditMemo!Totalcredit))
    End If
End Sub

Sub StoreCMMemvars(XXX As Variant)
    Dim rsCreditMemo                              As ADODB.Recordset
    Set rsCreditMemo = New ADODB.Recordset
    Set rsCreditMemo = gconDMIS.Execute("Select * from AMIS_Journal_HD Where ID = " & XXX)
    If Not rsCreditMemo.EOF And Not rsCreditMemo.BOF Then
        AddOrEditCM = "EDIT"
        SSTab1.Tab = 1
        cmdCreditMemo.ZOrder 0: picCreditMemo.ZOrder 0
        labCMID.Caption = rsCreditMemo!ID
        txtCMReference.Text = Null2String(rsCreditMemo!refno)
        txtCMRefDate.Text = Null2String(rsCreditMemo!RefDate)
        txtCMORNo.Text = Null2String(rsCreditMemo!INVOICENO)
        txtCMORDate.Text = Null2String(rsCreditMemo!invoicedate)
        txtCMAmount.Text = Null2String(rsCreditMemo!InvoiceAmt)
        txtCMParticular.Text = Null2String(rsCreditMemo!remarks)
        On Error Resume Next
        txtCMReference.SetFocus
    End If
End Sub

Sub InitCMMemvars()
    txtCMReference.Text = ""
    txtCMRefDate.Text = ""
    txtCMORNo.Text = ""
    txtCMORDate.Text = ""
    txtCMAmount.Text = "0.00"
    txtCMParticular.Text = ""
End Sub

Sub FillGrid()
    Dim rsCustomers                               As ADODB.Recordset
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set rsCustomers = New ADODB.Recordset
    Set rsCustomers = gconDMIS.Execute("select AcctName,ID from All_Customer where ACCTNAME <> '' order by AcctName asc")
    If Not (rsCustomers.EOF And rsCustomers.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsCustomers
        lstCustomer.Refresh
        lstCustomer.Enabled = True
    Else
        lstCustomer.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    XXX = Repleys(LTrim(RTrim(XXX)))
    Dim rsCustomers                               As ADODB.Recordset
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set rsCustomers = New ADODB.Recordset
    Set rsCustomers = gconDMIS.Execute("select ACCTNAME,ID from All_Customer where ACCTNAME like'" & XXX & "%' order by ACCTNAME asc")
    If Not (rsCustomers.EOF And rsCustomers.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsCustomers
        lstCustomer.Refresh
        lstCustomer.Enabled = True
    Else
        lstCustomer.Enabled = False
    End If
End Sub

Private Sub cmdCMCancel_Click()
    cmdCreditMemo.ZOrder 1
    picCreditMemo.ZOrder 1
End Sub

Private Sub cmdCMDelete_Click()
    gconDMIS.Execute ("delete from AMIS_Journal_HD Where ID = " & labCMID.Caption)
    LogAudit "X", "CUSTOMER ADJUSTMENTS", labCMID
    FillCMGrid
End Sub

Private Sub cmdCMSave_Click()
    Dim NewVoucherNo                              As String
    Dim NewJNo                                    As String
    NewVoucherNo = N2Str2Null(NewCreditMemoVoucherNo)
    NewJNo = N2Str2Null(NewOpenJNo)
    On Error GoTo error:
    If AddOrEditCM = "ADD" Then
        SQL_STATEMENT = "Insert Into AMIS_Journal_HD " & _
                        "(JType,JDate,DueDate,JNo,VoucherNo,CustomerCode,VendorCode,RefNo,RefDate,InvoiceNo,InvoiceType,InvoiceAmt,Remarks,Status,PaidStatus,ReceiveStatus)" & _
                        " values ('CCM'," & N2Date2Null(txtCMRefDate.Text) & "," & N2Date2Null(LOGDATE) & "," & NewJNo & "," & NewVoucherNo & "," & N2Str2Null(txtCustCode.Text) & ",'999999'," & N2Str2Null(txtCMReference.Text) & "," & N2Str2Null(txtCMRefDate.Text) & "," & N2Str2Null(txtCMORNo.Text) & "," & N2Str2Null(txtCMORDate.Text) & "," & NumericVal(txtCMAmount.Text) & "," & N2Str2Null(txtCMParticular.Text) & ",'P','N','N')"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "VENDOR ADJUSTMENT", SQL_STATEMENT, labCMID.Caption, "CREDIT MEMO", N2Str2Null(NewVoucherNo), "", ""
    Else
        SQL_STATEMENT = "Update AMIS_Journal_HD Set " & _
                        " CustomerCode = " & N2Str2Null(txtCustCode.Text) & "," & _
                        " RefNo = " & N2Str2Null(txtCMReference.Text) & "," & _
                        " RefDate = " & N2Str2Null(txtCMRefDate.Text) & "," & _
                        " JDate = " & N2Str2Null(txtCMRefDate.Text) & "," & _
                        " InvoiceNo = " & N2Str2Null(txtCMORNo.Text) & "," & _
                        " InvoiceDate = " & N2Str2Null(txtCMORDate.Text) & "," & _
                        " InvoiceAmt = " & NumericVal(txtCMAmount.Text) & "," & _
                        " Remarks = " & N2Str2Null(txtCMParticular.Text) & _
                        " Where ID = " & labCMID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "CUSTOMER ADJUSTMENT", SQL_STATEMENT, labCMID.Caption, "CREDIT MEMO", N2Str2Null(NewVoucherNo), "", ""
    End If
    If AddOrEditCM = "ADD" Then
        AddOrEditCM = "ADD"
        SSTab1.Tab = 0: InitOIMemvars
        cmdCreditMemo.ZOrder 0: picCreditMemo.ZOrder 0
        On Error Resume Next
        txtCMReference.SetFocus
    Else
        cmdCMCancel_Click
    End If
    FillCMGrid
    Exit Sub
error:
    MsgBox "Please check/Validate the data..", vbInformation, "Information"
End Sub

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
End Sub

Private Sub cmdFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraDetails.ZOrder 0
    On Error Resume Next
    TextSearch.SetFocus
End Sub

Private Sub cmdNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Errorcode:
    rsCustomer.MoveNext
    If rsCustomer.EOF Then
        rsCustomer.MoveLast
        MsgBox "Last Record!"
    End If
    StoreMemVars
    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub cmdOICancel_Click()
    cmdOpenInv.ZOrder 1
    picOpenInv.ZOrder 1
End Sub

Private Sub cmdOIDelete_Click()
    If ShowConfirmDelete = True Then
        SQL_STATEMENT = ("delete from AMIS_Journal_HD Where ID = " & labOIID.Caption)
        gconDMIS.Execute (SQL_STATEMENT)
        ShowDeletedMsg
        FillOIGrid
        LogAudit "X", "Open Inv", labOIID
        NEW_LogAudit "X", "VENDOR ADJUSTMENT", SQL_STATEMENT, labOIID.Caption, "OPEN INVOICE", txtCustCode, "", ""
    End If
End Sub

Private Sub cmdOISave_Click()
    Dim NewVoucherNo                              As String
    Dim NewJNo                                    As String
    Dim jtype                                     As String
    NewVoucherNo = N2Str2Null(NewOpenInvoiceVoucherNo)
    NewJNo = N2Str2Null(NewOpenJNo)
    jtype = "CSJ"
    If AddOrEditOI = "ADD" Then
        SQL_STATEMENT = "Insert Into AMIS_Journal_HD " & _
                        "(JType,JDate,DueDate,JNo,VoucherNo,CustomerCode,VendorCode,Terms,InvoiceNo,InvoiceType,Dealer,InvoiceDate,InvoiceAmt,AmountPaid,Balance,Remarks,Status,PaidStatus,ReceiveStatus)" & _
                        " values ('CSJ'," & N2Date2Null(LOGDATE) & "," & N2Date2Null(LOGDATE) & "," & NewJNo & "," & NewVoucherNo & "," & N2Str2Null(txtCustCode.Text) & ",'999999'," & N2Str2Null(txtTerms.Text) & "," & N2Str2Null(txtInvoiceNo.Text) & "," & N2Str2Null(txtInvoiceType.Text) & "," & N2Str2Null(txtDealer.Text) & "," & N2Str2Null(txtInvoiceDate.Text) & "," & NumericVal(txtOIAmount.Text) & ",0," & NumericVal(txtOIAmount.Text) & "," & N2Str2Null(txtOIParticular.Text) & ",'P','N','N')"
        gconDMIS.Execute SQL_STATEMENT
        TransactionID = (FindTransactionID(N2Str2Null(NewVoucherNo), "voucherno", "AMIS_Journal_HD", "X", N2Str2Null(jtype), "Jtype"))
        NEW_LogAudit "A", "VENDOR ADJUSTMENT", SQL_STATEMENT, TransactionID, "OPEN INVOICE", N2Str2Null(NewVoucherNo), "", ""
    Else
        SQL_STATEMENT = "Update AMIS_Journal_HD Set " & _
                        " CustomerCode = " & N2Str2Null(txtCustCode.Text) & "," & _
                        " Terms = " & N2Str2Null(txtTerms.Text) & "," & _
                        " InvoiceType = " & N2Str2Null(txtInvoiceType.Text) & "," & _
                        " InvoiceNo = " & N2Str2Null(txtInvoiceNo.Text) & "," & _
                        " Dealer = " & N2Str2Null(txtDealer.Text) & "," & _
                        " InvoiceDate = " & N2Str2Null(txtInvoiceDate.Text) & "," & _
                        " InvoiceAmt = " & NumericVal(txtOIAmount.Text) & "," & _
                        " Remarks = " & N2Str2Null(txtOIParticular.Text) & _
                        " Where ID = " & labOIID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "VENDOR ADJUSTMENT", SQL_STATEMENT, labOIID.Caption, "OPEN INVOICE", N2Str2Null(NewVoucherNo), "", ""
    End If
    If AddOrEditOI = "ADD" Then
        AddOrEditOI = "ADD"
        SSTab1.Tab = 0: InitOIMemvars
        cmdOpenInv.ZOrder 0: picOpenInv.ZOrder 0
        On Error Resume Next
        txtTerms.SetFocus
    Else
        cmdOICancel_Click
    End If
    FillOIGrid
End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo Errorcode:
    rsCustomer.MovePrevious
    If rsCustomer.BOF Then
        rsCustomer.MoveFirst
        MsgBox "First Record!"
    End If
    StoreMemVars
    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        AddOrEditOI = "ADD"
        SSTab1.Tab = 0: InitOIMemvars
        cmdOpenInv.ZOrder 0: picOpenInv.ZOrder 0
        On Error Resume Next
        txtTerms.SetFocus
    Case vbKeyF4
        AddOrEditCM = "ADD"
        SSTab1.Tab = 1: InitCMMemvars
        cmdCreditMemo.ZOrder 0: picCreditMemo.ZOrder 0
        On Error Resume Next
        txtCMReference.SetFocus
    Case vbKeyEscape
        cmdOpenInv.ZOrder 1: picOpenInv.ZOrder 1
        cmdCreditMemo.ZOrder 1: picCreditMemo.ZOrder 1
    Case Else
        MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF1 And Shift = 1:
        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        Unload frmALL_AuditInquiry
        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        frmALL_AuditInquiry.Caption = "VENDOR ADJUSTMENT"
        Call frmALL_AuditInquiry.DisplayHistory(labID, "VENDOR ADJUSTMENT")
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Frame1.Enabled = False: rsRefresh: initMemvars: StoreMemVars
    TextSearch.Text = ""
End Sub

Private Sub lstCustomer_GotFocus()
    On Error Resume Next
    rsCustomer.Bookmark = rsFind(rsCustomer.Clone, "ID", lstCustomer.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    rsCustomer.Bookmark = rsFind(rsCustomer.Clone, "ID", lstCustomer.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstCreditmemo_DblClick()
    If Trim(lstCreditMemo.SelectedItem.SubItems(5)) <> "" Then
        StoreCMMemvars (lstCreditMemo.SelectedItem.SubItems(5))
    End If
End Sub

Private Sub lstOpenInvoice_DblClick()
    If Trim(lstOpenInvoice.SelectedItem.SubItems(5)) <> "" Then
        StoreOIMemvars (lstOpenInvoice.SelectedItem.SubItems(5))
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstCustomer.ListItems.Count > 0 And lstCustomer.Enabled = True Then
            lstCustomer.SetFocus
        End If
    End If
End Sub

Private Sub txtCMAmount_GotFocus()
    If NumericVal(txtCMAmount.Text) > 0 Then txtCMAmount.Text = NumericVal(txtCMAmount.Text) Else txtCMAmount.Text = ""
End Sub

Private Sub txtCMAmount_LostFocus()
    If NumericVal(txtCMAmount.Text) > 0 Then txtCMAmount.Text = ToDoubleNumber(txtCMAmount.Text) Else txtCMAmount.Text = "0.00"
End Sub

Private Sub txtDealer_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtInvoiceDate_GotFocus()
    txtInvoiceDate.Text = Format(txtInvoiceDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtInvoiceDate_LostFocus()
    If txtInvoiceDate.Text <> "" Then
        If IsDate(txtInvoiceDate.Text) = True Then
            txtInvoiceDate.Text = Format(txtInvoiceDate.Text, "DD-MMM-YY")
        Else
            MsgBoxXP "Invalid Invoice Date!", "Error", XP_OKOnly, msg_Exclamation
            On Error Resume Next
            txtInvoiceDate.SetFocus
        End If
    End If
End Sub

Private Sub txtInvoiceNo_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtInvoiceType_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtOIAmount_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtCMRefDate_GotFocus()
    txtCMRefDate.Text = Format(txtCMRefDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtCMRefDate_LostFocus()
    If txtCMRefDate.Text <> "" Then
        If IsDate(txtCMRefDate.Text) = True Then
            txtCMRefDate.Text = Format(txtCMRefDate.Text, "DD-MMM-YY")
        Else
            MsgBoxXP "Invalid Reference Date!", "Error", XP_OKOnly, msg_Exclamation
            On Error Resume Next
            txtCMRefDate.SetFocus
        End If
    End If
End Sub

Private Sub txtCMORDate_GotFocus()
    txtCMORDate.Text = Format(txtCMORDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtCMORDate_LostFocus()
    If txtCMORDate.Text <> "" Then
        If IsDate(txtCMORDate.Text) = True Then
            txtCMORDate.Text = Format(txtCMORDate.Text, "DD-MMM-YY")
        Else
            MsgBoxXP "Invalid OR Date!", "Error", XP_OKOnly, msg_Exclamation
            On Error Resume Next
            txtCMORDate.SetFocus
        End If
    End If
End Sub

Private Sub lstCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        TextSearch.SetFocus
    End If
End Sub

Private Sub textSearch_Change()
    If Trim(TextSearch.Text) = "" Then FillGrid Else FillSearchGrid (TextSearch.Text)
End Sub

Private Sub txtTerms_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

