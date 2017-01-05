VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISVendorAdjustment 
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendor Adjustment"
   ClientHeight    =   6390
   ClientLeft      =   585
   ClientTop       =   330
   ClientWidth     =   9405
   ForeColor       =   &H8000000F&
   Icon            =   "VendorAdjustment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   9405
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
      Left            =   6630
      MouseIcon       =   "VendorAdjustment.frx":08CA
      MousePointer    =   99  'Custom
      Picture         =   "VendorAdjustment.frx":0A1C
      Style           =   1  'Graphical
      TabIndex        =   84
      ToolTipText     =   "Move to Previous Record"
      Top             =   5580
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
      Left            =   7320
      MouseIcon       =   "VendorAdjustment.frx":0D7B
      MousePointer    =   99  'Custom
      Picture         =   "VendorAdjustment.frx":0ECD
      Style           =   1  'Graphical
      TabIndex        =   83
      ToolTipText     =   "Move to Next Record"
      Top             =   5580
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
      Left            =   8010
      MouseIcon       =   "VendorAdjustment.frx":1225
      MousePointer    =   99  'Custom
      Picture         =   "VendorAdjustment.frx":1377
      Style           =   1  'Graphical
      TabIndex        =   82
      ToolTipText     =   "Find a Record"
      Top             =   5580
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
      Left            =   8700
      MouseIcon       =   "VendorAdjustment.frx":1671
      MousePointer    =   99  'Custom
      Picture         =   "VendorAdjustment.frx":17C3
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "Exit Window"
      Top             =   5580
      Width           =   705
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   3285
      Left            =   2700
      TabIndex        =   35
      Top             =   0
      Width           =   6705
      Begin VB.PictureBox Picture4 
         Enabled         =   0   'False
         Height          =   945
         Left            =   2070
         ScaleHeight     =   885
         ScaleWidth      =   4395
         TabIndex        =   46
         Top             =   1470
         Width           =   4455
         Begin VB.TextBox txtCredit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   330
            Width           =   2055
         End
         Begin VB.TextBox txtDebit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            TabIndex        =   2
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
            TabIndex        =   48
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
            TabIndex        =   47
            Top             =   30
            Width           =   1935
         End
      End
      Begin VB.TextBox txtVendorCode 
         Appearance      =   0  'Flat
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
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "VendorAdjustment.frx":1B29
         Top             =   2520
         Width           =   6585
      End
      Begin VB.TextBox txtVendorName 
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
         Left            =   1710
         MaxLength       =   150
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   660
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Balances"
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
         Left            =   90
         TabIndex        =   45
         Top             =   1170
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
         Caption         =   "Vendor Code"
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
         TabIndex        =   41
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
         TabIndex        =   39
         Top             =   2190
         Width           =   1785
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   6720
         Y1              =   1080
         Y2              =   1080
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
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   690
         Width           =   465
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Name"
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
         TabIndex        =   36
         Top             =   720
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2205
      Left            =   2700
      TabIndex        =   49
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
      TabPicture(0)   =   "VendorAdjustment.frx":1B2F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picDebitMemo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdDebitMemo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstDebitMemo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "[F4 - Add Credit Memo]"
      TabPicture(1)   =   "VendorAdjustment.frx":1B4B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picCustDeposit"
      Tab(1).Control(1)=   "cmdCreditMemo"
      Tab(1).Control(2)=   "picCreditMemo"
      Tab(1).Control(3)=   "lstCreditMemo"
      Tab(1).ControlCount=   4
      Begin MSComctlLib.ListView lstCreditMemo 
         Height          =   1755
         Left            =   -74940
         TabIndex        =   6
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
         MouseIcon       =   "VendorAdjustment.frx":1B67
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Bank Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Check No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Check Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "PV No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amt. Paid"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin MSComctlLib.ListView lstDebitMemo 
         Height          =   1755
         Left            =   60
         TabIndex        =   5
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
         MouseIcon       =   "VendorAdjustment.frx":1CC9
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
            Text            =   "Invoice Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Due Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amt. To Pay"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.PictureBox picCreditMemo 
         Height          =   1605
         Left            =   -74550
         ScaleHeight     =   1545
         ScaleWidth      =   5745
         TabIndex        =   71
         Top             =   150
         Width           =   5805
         Begin VB.TextBox txtCheckNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   19
            Text            =   "Text2"
            Top             =   630
            Width           =   1755
         End
         Begin VB.ComboBox cboBankName 
            Height          =   315
            Left            =   60
            TabIndex        =   18
            Text            =   "cboBankName"
            Top             =   330
            Width           =   2745
         End
         Begin VB.TextBox txtBankCode 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1020
            TabIndex        =   17
            Text            =   "Text2"
            Top             =   30
            Width           =   945
         End
         Begin VB.TextBox txtCheckDate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   20
            Text            =   "Text2"
            Top             =   930
            Width           =   1755
         End
         Begin VB.TextBox txtPVNumber 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   21
            Text            =   "Text2"
            Top             =   1230
            Width           =   1755
         End
         Begin VB.TextBox txtAmountPaid 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4110
            TabIndex        =   22
            Text            =   "Text2"
            Top             =   30
            Width           =   1575
         End
         Begin VB.TextBox txtCreditParticular 
            Appearance      =   0  'Flat
            Height          =   555
            Left            =   2970
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Text            =   "VendorAdjustment.frx":1E2B
            Top             =   630
            Width           =   2745
         End
         Begin VB.Frame Frame5 
            Caption         =   "Frame2"
            Height          =   1875
            Left            =   2850
            TabIndex        =   72
            Top             =   -150
            Width           =   30
         End
         Begin wizButton.cmd cmdCreditSave 
            Height          =   315
            Left            =   3900
            TabIndex        =   24
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
            MICON           =   "VendorAdjustment.frx":1E31
         End
         Begin wizButton.cmd cmdCreditCancel 
            Height          =   315
            Left            =   4800
            TabIndex        =   25
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
            MICON           =   "VendorAdjustment.frx":1E4D
         End
         Begin wizButton.cmd cmdCreditDelete 
            Height          =   315
            Left            =   2970
            TabIndex        =   26
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
            MICON           =   "VendorAdjustment.frx":1E69
         End
         Begin VB.Label Label28 
            Caption         =   "Bank Code"
            Height          =   255
            Left            =   60
            TabIndex        =   80
            Top             =   60
            Width           =   945
         End
         Begin VB.Label Label27 
            Caption         =   "Check No."
            Height          =   255
            Left            =   60
            TabIndex        =   79
            Top             =   660
            Width           =   1155
         End
         Begin VB.Label Label26 
            Caption         =   "Check Date"
            Height          =   255
            Left            =   60
            TabIndex        =   78
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label Label25 
            Caption         =   "PV Number"
            Height          =   255
            Left            =   60
            TabIndex        =   77
            Top             =   1260
            Width           =   1155
         End
         Begin VB.Label Label24 
            Caption         =   "Amount Paid"
            Height          =   255
            Left            =   2970
            TabIndex        =   76
            Top             =   60
            Width           =   1155
         End
         Begin VB.Label Label23 
            Caption         =   "Particular"
            Height          =   255
            Left            =   2970
            TabIndex        =   75
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label Label22 
            Caption         =   "Label20"
            Height          =   165
            Left            =   1950
            TabIndex        =   74
            Top             =   600
            Width           =   285
         End
         Begin VB.Label Label21 
            Caption         =   "Label20"
            Height          =   255
            Left            =   1950
            TabIndex        =   73
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdCreditMemo 
         Caption         =   "Command1"
         Height          =   1725
         Left            =   -74610
         TabIndex        =   57
         Top             =   90
         Width           =   5925
      End
      Begin VB.PictureBox picCustDeposit 
         Height          =   1605
         Left            =   -74550
         ScaleHeight     =   1545
         ScaleWidth      =   5745
         TabIndex        =   58
         Top             =   150
         Width           =   5805
         Begin VB.TextBox txtRefInvNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   30
            Text            =   "Text2"
            Top             =   1230
            Width           =   1755
         End
         Begin VB.Frame Frame3 
            Caption         =   "Frame3"
            Height          =   1725
            Left            =   2850
            TabIndex        =   65
            Top             =   -120
            Width           =   30
         End
         Begin VB.TextBox txtCDParticular 
            Appearance      =   0  'Flat
            Height          =   825
            Left            =   2970
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Text            =   "VendorAdjustment.frx":1E85
            Top             =   330
            Width           =   2745
         End
         Begin VB.TextBox txtOrNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   27
            Text            =   "Text2"
            Top             =   30
            Width           =   1755
         End
         Begin VB.TextBox txtOrDate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   28
            Text            =   "Text2"
            Top             =   330
            Width           =   1755
         End
         Begin VB.TextBox txtCDAmount 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   29
            Text            =   "Text2"
            Top             =   630
            Width           =   1755
         End
         Begin VB.Frame Frame4 
            Height          =   30
            Left            =   -90
            TabIndex        =   66
            Top             =   960
            Width           =   2985
         End
         Begin wizButton.cmd cmdCDSave 
            Height          =   315
            Left            =   3900
            TabIndex        =   33
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
            MICON           =   "VendorAdjustment.frx":1E8B
         End
         Begin wizButton.cmd cmdCDCancel 
            Height          =   315
            Left            =   4800
            TabIndex        =   34
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
            MICON           =   "VendorAdjustment.frx":1EA7
         End
         Begin wizButton.cmd cmdCDDelete 
            Height          =   315
            Left            =   2970
            TabIndex        =   32
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
            MICON           =   "VendorAdjustment.frx":1EC3
         End
         Begin VB.Label Label19 
            Caption         =   "Reference Invoice No."
            Height          =   255
            Left            =   60
            TabIndex        =   67
            Top             =   990
            Width           =   2835
         End
         Begin VB.Label Label18 
            Caption         =   "Particular"
            Height          =   255
            Left            =   2970
            TabIndex        =   64
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label17 
            Caption         =   "OR #"
            Height          =   255
            Left            =   60
            TabIndex        =   61
            Top             =   60
            Width           =   1155
         End
         Begin VB.Label Label15 
            Caption         =   "OR Date"
            Height          =   255
            Left            =   60
            TabIndex        =   60
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label14 
            Caption         =   "Amount"
            Height          =   255
            Left            =   60
            TabIndex        =   59
            Top             =   660
            Width           =   1155
         End
      End
      Begin VB.CommandButton cmdDebitMemo 
         Caption         =   "Command1"
         Height          =   1725
         Left            =   390
         TabIndex        =   50
         Top             =   90
         Width           =   5925
      End
      Begin VB.PictureBox picDebitMemo 
         Height          =   1605
         Left            =   450
         ScaleHeight     =   1545
         ScaleWidth      =   5745
         TabIndex        =   51
         Top             =   150
         Width           =   5805
         Begin VB.TextBox txtPayType 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1020
            TabIndex        =   8
            Text            =   "Text2"
            Top             =   330
            Width           =   1755
         End
         Begin wizButton.cmd cmdDebitSave 
            Height          =   315
            Left            =   3900
            TabIndex        =   14
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
            MICON           =   "VendorAdjustment.frx":1EDF
         End
         Begin VB.Frame Frame2 
            Caption         =   "Frame2"
            Height          =   1875
            Left            =   2850
            TabIndex        =   63
            Top             =   -150
            Width           =   30
         End
         Begin VB.TextBox txtDebitParticular 
            Appearance      =   0  'Flat
            Height          =   555
            Left            =   2970
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Text            =   "VendorAdjustment.frx":1EFB
            Top             =   630
            Width           =   2745
         End
         Begin VB.TextBox txtAmountToPay 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4110
            TabIndex        =   12
            Text            =   "Text2"
            Top             =   30
            Width           =   1575
         End
         Begin VB.TextBox txtInvoiceDate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   11
            Text            =   "Text2"
            Top             =   1230
            Width           =   1755
         End
         Begin VB.TextBox txtInvoiceNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   10
            Text            =   "Text2"
            Top             =   930
            Width           =   1755
         End
         Begin VB.TextBox txtPONo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   9
            Text            =   "Text2"
            Top             =   630
            Width           =   1755
         End
         Begin VB.TextBox txtTerms 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   7
            Text            =   "Text2"
            Top             =   30
            Width           =   675
         End
         Begin wizButton.cmd cmdDebitCancel 
            Height          =   315
            Left            =   4800
            TabIndex        =   15
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
            MICON           =   "VendorAdjustment.frx":1F01
         End
         Begin wizButton.cmd cmdDebitDelete 
            Height          =   315
            Left            =   2970
            TabIndex        =   16
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
            MICON           =   "VendorAdjustment.frx":1F1D
         End
         Begin VB.Label Label9 
            Caption         =   "Pay Type"
            Height          =   255
            Left            =   60
            TabIndex        =   70
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label labCDID 
            Caption         =   "Label20"
            Height          =   255
            Left            =   1950
            TabIndex        =   69
            Top             =   600
            Width           =   495
         End
         Begin VB.Label labOIID 
            Caption         =   "Label20"
            Height          =   165
            Left            =   1950
            TabIndex        =   68
            Top             =   600
            Width           =   285
         End
         Begin VB.Label Label16 
            Caption         =   "Particular"
            Height          =   255
            Left            =   2970
            TabIndex        =   62
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label Label13 
            Caption         =   "Amount to Pay"
            Height          =   255
            Left            =   2970
            TabIndex        =   56
            Top             =   60
            Width           =   1155
         End
         Begin VB.Label Label12 
            Caption         =   "Invoice Date"
            Height          =   255
            Left            =   60
            TabIndex        =   55
            Top             =   1260
            Width           =   1155
         End
         Begin VB.Label Label11 
            Caption         =   "Invoice No."
            Height          =   255
            Left            =   60
            TabIndex        =   54
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label Label10 
            Caption         =   "PO Number"
            Height          =   255
            Left            =   60
            TabIndex        =   53
            Top             =   660
            Width           =   1155
         End
         Begin VB.Label Label5 
            Caption         =   "Terms"
            Height          =   255
            Left            =   60
            TabIndex        =   52
            Top             =   60
            Width           =   555
         End
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   5370
      Left            =   2730
      ScaleHeight     =   5310
      ScaleWidth      =   2535
      TabIndex        =   40
      Top             =   90
      Width           =   2595
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   11640
         Left            =   2520
         Picture         =   "VendorAdjustment.frx":1F39
         Top             =   -150
         Width           =   2535
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   5625
      Left            =   60
      TabIndex        =   42
      Top             =   0
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
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   180
         Width           =   2475
      End
      Begin MSComctlLib.ListView lstVendor 
         Height          =   4995
         Left            =   30
         TabIndex        =   44
         Top             =   570
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   8811
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
         MouseIcon       =   "VendorAdjustment.frx":11E31
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
Attribute VB_Name = "frmAMISVendorAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsVENDOR                                                          As ADODB.Recordset
Dim AddOrEditDebit, AddOrEditCredit                                   As String
Attribute AddOrEditCredit.VB_VarUserMemId = 1073938433

Function NewOpenJNo() As String
    Dim rsJournal_HDNew                                               As ADODB.Recordset
    Set rsJournal_HDNew = New ADODB.Recordset
    Set rsJournal_HDNew = gconDMIS.Execute("Select * from AMIS_Journal_HD order by JNo Desc")
    If Not rsJournal_HDNew.EOF And Not rsJournal_HDNew.BOF Then
        NewOpenJNo = Format(N2Str2Zero(rsJournal_HDNew!JNo) + 1, "000000")
    Else
        NewOpenJNo = "000001"
    End If
End Function

Function SetVendorName(CCC As Variant)
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "Select custcode,custname from ALL_Vendor where Vendorcode = " & N2Str2Null(CCC), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorName = Null2String(rsVENDOR!VendorName)
    Else
        SetVendorName = ""
    End If
End Function

Function NewVDJVoucherNo() As String
    Dim rsJournal_HDNew                                               As ADODB.Recordset
    Set rsJournal_HDNew = New ADODB.Recordset
    Set rsJournal_HDNew = gconDMIS.Execute("Select * from AMIS_Journal_HD Where Jtype = 'VDJ' order by VoucherNo Desc")
    If Not rsJournal_HDNew.EOF And Not rsJournal_HDNew.BOF Then
        NewVDJVoucherNo = Format(N2Str2Zero(rsJournal_HDNew!VOUCHERNO) + 1, "000000")
    Else
        NewVDJVoucherNo = "000001"
    End If
End Function

Function SetBankCode(XXX As String) As String
    Dim rsBanks                                                       As ADODB.Recordset
    Set rsBanks = New ADODB.Recordset
    Set rsBanks = gconDMIS.Execute("Select * from ALL_BANKS where bankNAME = '" & XXX & "'")
    If Not rsBanks.EOF And Not rsBanks.BOF Then
        SetBankCode = Null2String(rsBanks!BankCode)
    End If
End Function

Function SetBankName(XXX As String) As String
    Dim rsBanks                                                       As ADODB.Recordset
    Set rsBanks = New ADODB.Recordset
    Set rsBanks = gconDMIS.Execute("Select * from ALL_BANKS where bankcode = '" & XXX & "'")
    If Not rsBanks.EOF And Not rsBanks.BOF Then
        SetBankName = Null2String(rsBanks!bankname)
    End If
End Function

Function NewVCJVoucherNo() As String
    Dim rsJournal_HDNew                                               As ADODB.Recordset
    Set rsJournal_HDNew = New ADODB.Recordset
    Set rsJournal_HDNew = gconDMIS.Execute("Select * from AMIS_Journal_HD Where Jtype = 'VCJ' order by VoucherNo Desc")
    If Not rsJournal_HDNew.EOF And Not rsJournal_HDNew.BOF Then
        NewVCJVoucherNo = Format(N2Str2Zero(rsJournal_HDNew!VOUCHERNO) + 1, "000000")
    Else
        NewVCJVoucherNo = "000001"
    End If
End Function

Sub rsRefresh()
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "select * from ALL_Vendor Order by NameOFVendor asc", gconDMIS, adOpenKeyset
End Sub

Sub StoreMemvars()
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        txtVendorCode.Text = Null2String(rsVENDOR!code)
        txtVendorName.Text = Null2String(rsVENDOR!nameofvendor)
        FillDebitGrid
        FillCreditGrid
    End If
End Sub

Sub InitMemVars()
    txtVendorCode.Text = ""
    txtVendorName.Text = ""
    txtDebit.Text = ""
    txtCredit.Text = ""
    txtParticular.Text = ""
End Sub

Sub FillGrid()
    Dim rsVendors                                                     As ADODB.Recordset
    lstVendor.Sorted = False: lstVendor.ListItems.Clear
    Set rsVendors = New ADODB.Recordset
    Set rsVendors = gconDMIS.Execute("select NAMEOFVENDOR,Code from ALL_Vendor where NAMEOFVENDOR <> '' order by NAMEOFVENDOR asc")
    If Not (rsVendors.EOF And rsVendors.BOF) Then
        Listview_Loadval Me.lstVendor.ListItems, rsVendors
        lstVendor.Refresh
        lstVendor.Enabled = True
    Else
        lstVendor.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsVendors                                                     As ADODB.Recordset
    lstVendor.Sorted = False: lstVendor.ListItems.Clear
    Set rsVendors = New ADODB.Recordset
    Set rsVendors = gconDMIS.Execute("select NameOfVendor,Code from ALL_Vendor where NameOfVendor like'" & XXX & "%' order by NameOfVendor asc")
    If Not (rsVendors.EOF And rsVendors.BOF) Then
        Listview_Loadval Me.lstVendor.ListItems, rsVendors
        lstVendor.Refresh
        lstVendor.Enabled = True
    Else
        lstVendor.Enabled = False
    End If
End Sub

Sub FillDebitGrid()
    Dim rsDebitMemo                                                   As ADODB.Recordset
    lstDebitMemo.Sorted = False: lstDebitMemo.ListItems.Clear
    Set rsDebitMemo = New ADODB.Recordset
    Set rsDebitMemo = gconDMIS.Execute("select Terms,InvoiceNo,InvoiceDate,DueDate,AmountToPay,ID from AMIS_Journal_HD where jtype = 'VDJ' and Vendorcode = " & N2Str2Null(txtVendorCode.Text) & " order by VoucherNo asc")
    If Not (rsDebitMemo.EOF And rsDebitMemo.BOF) Then
        lstDebitMemo.Enabled = True
        Listview_Loadval Me.lstDebitMemo.ListItems, rsDebitMemo
        lstDebitMemo.Refresh
    Else
        lstDebitMemo.Enabled = False
    End If
    Set rsDebitMemo = New ADODB.Recordset
    Set rsDebitMemo = gconDMIS.Execute("select SUM(Balance) AS TOTALDEBIT from AMIS_Journal_HD where jtype = 'VDJ' and Vendorcode = " & N2Str2Null(txtVendorCode.Text))
    If Not (rsDebitMemo.EOF And rsDebitMemo.BOF) Then
        txtDebit.Text = ToDoubleNumber(N2Str2Zero(rsDebitMemo!TotalDebit))
    End If
End Sub

Sub StoreDMMemvars(XXX As Variant)
    Dim rsDebitMemo                                                   As ADODB.Recordset
    Set rsDebitMemo = New ADODB.Recordset
    Set rsDebitMemo = gconDMIS.Execute("Select * from AMIS_Journal_HD Where ID = " & XXX)
    If Not rsDebitMemo.EOF And Not rsDebitMemo.BOF Then
        AddOrEditDebit = "EDIT"
        SSTab1.Tab = 0
        cmdDebitMemo.ZOrder 0: picDebitMemo.ZOrder 0
        labOIID.Caption = rsDebitMemo!ID
        txtTerms.Text = Null2String(rsDebitMemo!TERMS)
        txtPayType.Text = Null2String(rsDebitMemo!paytype)
        txtPONo.Text = Null2String(rsDebitMemo!RefNo)
        txtInvoiceNo.Text = Null2String(rsDebitMemo!InvoiceNo)
        txtInvoiceDate.Text = Null2String(rsDebitMemo!InvoiceDate)
        txtAmountToPay.Text = Null2String(rsDebitMemo!AmountToPay)
        txtDebitParticular.Text = Null2String(rsDebitMemo!Remarks)
        On Error Resume Next
        txtTerms.SetFocus
    End If
End Sub

Sub InitDebitMemvars()
    txtTerms.Text = ""
    txtPayType.Text = ""
    txtPONo.Text = ""
    txtInvoiceNo.Text = ""
    txtInvoiceDate.Text = ""
    txtAmountToPay.Text = ""
    txtDebitParticular.Text = ""
End Sub

'credit memo

Sub FillBankName()
    Dim rsBank                                                        As ADODB.Recordset
    Set rsBank = New ADODB.Recordset
    Set rsBank = gconDMIS.Execute("Select * from ALL_BANKS order by bankname asc")
    If Not rsBank.EOF And Not rsBank.BOF Then
        rsBank.MoveFirst: cboBankName.Clear
        Do While Not rsBank.EOF
            cboBankName.AddItem Null2String(rsBank!bankname)
            rsBank.MoveNext
        Loop
    End If
End Sub

Sub FillCreditGrid()
    Dim rsCreditMemo                                                  As ADODB.Recordset
    lstCreditMemo.Sorted = False: lstCreditMemo.ListItems.Clear
    Set rsCreditMemo = New ADODB.Recordset
    Set rsCreditMemo = gconDMIS.Execute("select BankCode,CheckNo,CheckDate,RefNo,Credit,ID from AMIS_Journal_HD where jtype = 'VCJ' and Vendorcode = " & N2Str2Null(txtVendorCode.Text) & " order by VoucherNo asc")
    If Not (rsCreditMemo.EOF And rsCreditMemo.BOF) Then
        lstCreditMemo.Enabled = True
        Listview_Loadval Me.lstCreditMemo.ListItems, rsCreditMemo
        lstCreditMemo.Refresh
    Else
        lstCreditMemo.Enabled = False
    End If
    Set rsCreditMemo = New ADODB.Recordset
    Set rsCreditMemo = gconDMIS.Execute("select SUM(Balance) AS TOTALCredit from AMIS_Journal_HD where jtype = 'VCJ' and Vendorcode = " & N2Str2Null(txtVendorCode.Text))
    If Not (rsCreditMemo.EOF And rsCreditMemo.BOF) Then
        txtCredit.Text = ToDoubleNumber(N2Str2Zero(rsCreditMemo!Totalcredit))
    End If
End Sub

Sub StoreCMMemvars(XXX As Variant)
    Dim rsCreditMemo                                                  As ADODB.Recordset
    Set rsCreditMemo = New ADODB.Recordset
    Set rsCreditMemo = gconDMIS.Execute("Select * from AMIS_Journal_HD Where ID = " & XXX)
    If Not rsCreditMemo.EOF And Not rsCreditMemo.BOF Then
        AddOrEditCredit = "EDIT"
        SSTab1.Tab = 1
        cmdCreditMemo.ZOrder 0: picCreditMemo.ZOrder 0
        labCDID.Caption = rsCreditMemo!ID
        txtBankCode.Text = Null2String(rsCreditMemo!BankCode)
        cboBankName.Text = SetBankName(Null2String(rsCreditMemo!BankCode))
        txtCheckNo.Text = Null2String(rsCreditMemo!CheckNo)
        txtCheckDate.Text = Null2String(rsCreditMemo!CheckDate)
        txtPVNumber.Text = Null2String(rsCreditMemo!RefNo)
        txtAmountPaid.Text = Null2String(rsCreditMemo!CREDIT)
        txtCreditParticular.Text = Null2String(rsCreditMemo!Remarks)
        On Error Resume Next
        txtCheckNo.SetFocus
    End If
End Sub

Sub InitCreditMemvars()
    txtBankCode.Text = ""
    FillBankName
    txtCheckNo.Text = ""
    txtCheckDate.Text = ""
    txtPVNumber.Text = ""
    txtAmountPaid.Text = ""
    txtCreditParticular.Text = ""
End Sub

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Unload Me
End Sub

Private Sub cmdFind_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraDetails.ZOrder 0
    On Error Resume Next
    TextSearch.SetFocus
End Sub

Private Sub cmdNext_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    rsVENDOR.MoveNext
    If rsVENDOR.EOF Then
        rsVENDOR.MoveLast
        MsgBox "Last Record!"
    End If
    StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
    rsVENDOR.MovePrevious
    If rsVENDOR.BOF Then
        rsVENDOR.MoveFirst
        MsgBox "First Record!"
    End If
    StoreMemvars

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            AddOrEditDebit = "ADD"
            SSTab1.Tab = 0: InitDebitMemvars
            cmdDebitMemo.ZOrder 0: picDebitMemo.ZOrder 0
            On Error Resume Next
            txtTerms.SetFocus
        Case vbKeyF4
            AddOrEditCredit = "ADD"
            SSTab1.Tab = 1: InitCreditMemvars
            cmdCreditMemo.ZOrder 0: picCreditMemo.ZOrder 0
            On Error Resume Next
            txtOrNo.SetFocus
        Case vbKeyEscape
            cmdDebitMemo.ZOrder 1: picDebitMemo.ZOrder 1
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
    Frame1.Enabled = False: rsRefresh: InitMemVars: StoreMemvars
    TextSearch.Text = ""
End Sub

Private Sub lstVendor_GotFocus()
    On Error Resume Next
    rsVENDOR.Bookmark = rsFind(rsVENDOR.Clone, "Code", lstVendor.SelectedItem.SubItems(1)).Bookmark
    StoreMemvars
End Sub

Private Sub lstVendor_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    rsVENDOR.Bookmark = rsFind(rsVENDOR.Clone, "Code", lstVendor.SelectedItem.SubItems(1)).Bookmark
    StoreMemvars
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then lstVendor.SetFocus
End Sub

Private Sub lstVendor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then TextSearch.SetFocus
End Sub

Private Sub textSearch_Change()
    If Trim(TextSearch.Text) = "" Then FillGrid Else FillSearchGrid (TextSearch.Text)
End Sub

'debit memo
Private Sub cmdDebitCancel_Click()
    cmdDebitMemo.ZOrder 1
    picDebitMemo.ZOrder 1
End Sub

Private Sub cmdDebitDelete_Click()
    If ShowConfirmDelete = True Then
        SQL_STATEMENT = ("delete from AMIS_Journal_HD Where ID = " & labOIID.Caption)
        gconDMIS.Execute SQL_STATEMENT
        ShowDeletedMsg
        FillDebitGrid
        NEW_LogAudit "X", "VENDOR ADJUSTMENT", SQL_STATEMENT, labOIID.Caption, "DEDIT MEMO", txtVendorCode, "", ""
    End If
End Sub

Private Sub cmdDebitSave_Click()
    Dim NewVoucherNo                                                  As String
    Dim NewJNo                                                        As String
    NewVoucherNo = N2Str2Null(NewVDJVoucherNo)
    NewJNo = N2Str2Null(NewOpenJNo)
    Dim jtype As String

    Dim VtxtTerms                                                     As String
    Dim VtxtPayType                                                   As String
    Dim VtxtPONo                                                      As String
    Dim VtxtInvoiceNo                                                 As String
    Dim VtxtInvoiceDate                                               As String
    Dim VDueDate                                                      As String
    Dim VtxtAmountToPay                                               As Double
    Dim VtxtDebitParticular                                           As String

    VDueDate = N2Date2Null(Format(DateAdd("d", NumericVal(txtTerms.Text), txtInvoiceDate.Text), "DD-MMM-YY"))
    
    VtxtTerms = N2Str2Null(txtTerms.Text)
    VtxtPayType = N2Str2Null(txtPayType.Text)
    VtxtPONo = N2Str2Null(txtPONo.Text)
    VtxtInvoiceNo = N2Str2Null(Format(txtInvoiceNo.Text, "000000"))
    VtxtInvoiceDate = N2Date2Null(txtInvoiceDate.Text)
    VtxtAmountToPay = N2Str2Zero(txtAmountToPay.Text)
    VtxtDebitParticular = N2Str2Null(txtDebitParticular.Text)
    jtype = "VDJ"
    If AddOrEditDebit = "ADD" Then
        SQL_STATEMENT = "Insert Into AMIS_Journal_HD " & _
                         "(JType,JDate,DueDate,JNo,VoucherNo,VendorCode,CustomerCode,Terms,InvoiceNo,RefNo,PayType,InvoiceDate,Debit,AmountPaid,Remarks,Status,PaidStatus,ReceiveStatus)" & _
                       " values ('VDJ'," & N2Date2Null(LOGDATE) & "," & VDueDate & "," & NewJNo & "," & NewVoucherNo & "," & N2Str2Null(txtVendorCode.Text) & ",'999999'," & VtxtTerms & "," & VtxtInvoiceNo & "," & VtxtPONo & "," & VtxtPayType & "," & VtxtInvoiceDate & "," & VtxtAmountToPay & "," & VtxtAmountToPay & "," & VtxtDebitParticular & ",'P','N','N')"
        gconDMIS.Execute SQL_STATEMENT
        
        TransactionID = (FindTransactionID(N2Str2Null(NewVoucherNo), "voucherno", "AMIS_Journal_HD", "X", N2Str2Null(jtype), "Jtype"))
        NEW_LogAudit "A", "VENDOR ADJUSTMENT", SQL_STATEMENT, TransactionID, "DEBIT MEMO", N2Str2Null(NewVoucherNo), "", ""
        
    Else
        SQL_STATEMENT = "Update AMIS_Journal_HD Set " & _
                       " JType = 'VDJ'," & _
                       " JDate = " & N2Date2Null(LOGDATE) & "," & _
                       " DueDate = " & VDueDate & "," & _
                       " JNo = " & NewJNo & "," & _
                       " VoucherNo = " & NewVoucherNo & "," & _
                       " VendorCode = " & N2Str2Null(txtVendorCode.Text) & "," & _
                       " CustomerCode = '999999'," & _
                       " Terms = " & VtxtTerms & "," & _
                       " InvoiceNo = " & VtxtInvoiceNo & "," & _
                       " RefNo = " & VtxtPONo & "," & _
                       " PayType = " & VtxtPayType & "," & _
                       " InvoiceDate = " & VtxtInvoiceDate & "," & _
                       " AmountToPay = " & VtxtAmountToPay & "," & _
                       " AmountPaid = 0," & _
                       " Balance = " & VtxtAmountToPay & "," & _
                       " Remarks = " & VtxtDebitParticular & "," & _
                       " Status = 'P', PaidStatus = 'N', ReceiveStatus = 'N'" & _
                       " Where ID = " & labOIID.Caption
        gconDMIS.Execute SQL_STATEMENT
        TransactionID = (FindTransactionID(N2Str2Null(NewVoucherNo), "voucherno", "AMIS_Journal_HD", "X", N2Str2Null(jtype), "Jtype"))
        NEW_LogAudit "E", "VENDOR ADJUSTMENT", SQL_STATEMENT, TransactionID, "DEBIT MEMO", N2Str2Null(NewVoucherNo), "", ""
    End If
    If AddOrEditDebit = "ADD" Then
        AddOrEditDebit = "ADD"
        SSTab1.Tab = 0: InitDebitMemvars
        cmdDebitMemo.ZOrder 0: picDebitMemo.ZOrder 0
        On Error Resume Next
        txtTerms.SetFocus
    Else
        cmdDebitCancel_Click
    End If
    FillDebitGrid
End Sub

Private Sub lstdebitmemo_DblClick()
    If Trim(lstDebitMemo.SelectedItem.SubItems(5)) <> "" Then
        StoreDMMemvars (lstDebitMemo.SelectedItem.SubItems(5))
    End If
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

Private Sub txtAmountToPay_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtTerms_Change()
    If NumericVal(txtTerms.Text) = 0 Then
        txtPayType.Text = "CSH"
    Else
        txtPayType.Text = txtTerms.Text & "D"
    End If
End Sub

Private Sub txtTerms_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub cboBankName_Change()
    txtBankCode.Text = SetBankCode(cboBankName.Text)
End Sub

Private Sub cboBankName_Click()
    txtBankCode.Text = SetBankCode(cboBankName.Text)
End Sub

Private Sub cmdCreditCancel_Click()
    cmdCreditMemo.ZOrder 1
    picCreditMemo.ZOrder 1
End Sub

Private Sub cmdCreditDelete_Click()
    If ShowConfirmDelete = True Then
        SQL_STATEMENT = ("delete from AMIS_Journal_HD Where ID = " & labCDID.Caption)
        gconDMIS.Execute (SQL_STATEMENT)
        ShowDeletedMsg
        FillCreditGrid
        NEW_LogAudit "X", "VENDOR ADJUSTMENT", SQL_STATEMENT, labCDID.Caption, "CREDIT MEMO", txtVendorCode, "", ""
    End If
End Sub

Private Sub cmdCreditSave_Click()
    Dim NewVoucherNo                                                  As String
    Dim NewJNo                                                        As String
    NewVoucherNo = N2Str2Null(NewVCJVoucherNo)
    NewJNo = N2Str2Null(NewOpenJNo)

    Dim VtxtBankCode                                                  As String
    Dim VcboBankName                                                  As String
    Dim VtxtCheckNo                                                   As String
    Dim VtxtCheckDate                                                 As String
    Dim VtxtPVNumber                                                  As String
    Dim VtxtAmountPaid                                                As Double
    Dim VtxtCreditParticular                                          As String
    Dim jtype As String
    VtxtBankCode = N2Str2Null(txtBankCode.Text)
    VcboBankName = N2Str2Null(cboBankName.Text)
    VtxtCheckNo = N2Str2Null(txtCheckNo.Text)
    VtxtCheckDate = N2Date2Null(txtCheckDate.Text)
    VtxtPVNumber = N2Str2Null(txtPVNumber.Text)
    VtxtAmountPaid = N2Str2Zero(txtAmountPaid.Text)
    VtxtCreditParticular = N2Str2Null(txtCreditParticular.Text)
    jtype = "VCJ"
    If AddOrEditCredit = "ADD" Then
        SQL_STATEMENT = "Insert Into AMIS_Journal_HD " & _
                         "(JType,JDate,JNo,VoucherNo,VendorCode,CustomerCode,BankCode,CheckNo,CheckDate,RefNo,Credit,Remarks,Status,PaidStatus,ReceiveStatus)" & _
                       " values ('VCJ'," & N2Date2Null(LOGDATE) & "," & NewJNo & "," & NewVoucherNo & "," & N2Str2Null(txtVendorCode.Text) & ",'999999'," & VtxtBankCode & "," & VtxtCheckNo & "," & VtxtCheckDate & "," & VtxtPVNumber & "," & VtxtAmountPaid & "," & VtxtCreditParticular & ",'P','N','N')"
        gconDMIS.Execute SQL_STATEMENT
        TransactionID = (FindTransactionID(N2Str2Null(NewVoucherNo), "voucherno", "AMIS_Journal_HD", "X", N2Str2Null(jtype), "Jtype"))
        NEW_LogAudit "A", "VENDOR ADJUSTMENT", SQL_STATEMENT, labCDID.Caption, "CREDIT MEMO", N2Str2Null(NewVoucherNo), "", ""
    Else
        SQL_STATEMENT = "Update AMIS_Journal_HD Set " & _
                       " JType = 'VCJ'," & _
                       " JDate = " & N2Date2Null(LOGDATE) & "," & _
                       " JNo = " & NewJNo & "," & _
                       " VoucherNo = " & NewVoucherNo & "," & _
                       " VendorCode = " & N2Str2Null(txtVendorCode.Text) & "," & _
                       " CustomerCode = '999999'," & _
                       " BankCode = " & VtxtBankCode & "," & _
                       " CheckNo = " & VtxtCheckNo & "," & _
                       " CheckDate = " & VtxtCheckDate & "," & _
                       " RefNo = " & VtxtPVNumber & "," & _
                       " Credit = " & VtxtAmountPaid & "," & _
                       " Remarks = " & VtxtCreditParticular & "," & _
                       " Status = 'P', PaidStatus = 'N', ReceiveStatus = 'N'" & _
                       " Where ID = " & labCDID.Caption
        gconDMIS.Execute SQL_STATEMENT
         NEW_LogAudit "E", "VENDOR ADJUSTMENT", SQL_STATEMENT, labCDID.Caption, "CREDIT MEMO", N2Str2Null(NewVoucherNo), "", ""
    End If
    If AddOrEditCredit = "ADD" Then
        AddOrEditCredit = "ADD"
        SSTab1.Tab = 0: InitCreditMemvars
        cmdCreditMemo.ZOrder 0: picCreditMemo.ZOrder 0
        On Error Resume Next
        txtCheckNo.SetFocus
    Else
        cmdCreditCancel_Click
    End If
    FillCreditGrid
End Sub

Private Sub lstCreditmemo_DblClick()
    If Trim(lstCreditMemo.SelectedItem.SubItems(5)) <> "" Then
        StoreCMMemvars (lstCreditMemo.SelectedItem.SubItems(5))
    End If
End Sub

Private Sub txtAmountPaid_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtCheckDate_GotFocus()
    txtCheckDate.Text = Format(txtCheckDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtCheckDate_LostFocus()
    If txtCheckDate.Text <> "" Then
        If IsDate(txtCheckDate.Text) = True Then
            txtCheckDate.Text = Format(txtCheckDate.Text, "DD-MMM-YY")
        Else
            MsgBoxXP "Invalid Check Date!", "Error", XP_OKOnly, msg_Exclamation
            On Error Resume Next
            txtCheckDate.SetFocus
        End If
    End If
End Sub

