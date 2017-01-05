VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A06473E6-73D7-426E-82F2-6CD4F1FA4DBE}#1.0#0"; "WIZMACBUT.OCX"
Begin VB.Form frmAMISCustomerOpeningBalance 
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Opening Balances"
   ClientHeight    =   6015
   ClientLeft      =   585
   ClientTop       =   330
   ClientWidth     =   9495
   ForeColor       =   &H8000000F&
   Icon            =   "CustomerOpeningBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   9495
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   3285
      Left            =   2700
      TabIndex        =   26
      Top             =   0
      Width           =   6705
      Begin VB.PictureBox Picture4 
         Enabled         =   0   'False
         Height          =   855
         Left            =   2070
         ScaleHeight     =   795
         ScaleWidth      =   4395
         TabIndex        =   38
         Top             =   1560
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
            TabIndex        =   4
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
            TabIndex        =   40
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
            TabIndex        =   39
            Top             =   30
            Width           =   1935
         End
      End
      Begin VB.TextBox txtCustCode 
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
         TabIndex        =   5
         Text            =   "CustomerOpeningBalance.frx":08CA
         Top             =   2520
         Width           =   6585
      End
      Begin VB.TextBox txtAccountNo 
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
         Left            =   4350
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   180
         Width           =   2175
      End
      Begin VB.TextBox txtCustName 
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
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   90
         TabIndex        =   37
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
         TabIndex        =   33
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   720
         Width           =   1575
      End
   End
   Begin wizMacBut.MacBut cmdExit 
      Height          =   345
      Left            =   7740
      TabIndex        =   62
      Top             =   5610
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   609
      Caption         =   "Exit"
      Caption_Xpos    =   600
   End
   Begin wizMacBut.MacBut cmdFind 
      Height          =   345
      Left            =   6060
      TabIndex        =   63
      Top             =   5610
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   609
      Caption         =   "Search"
      Caption_Xpos    =   400
   End
   Begin wizMacBut.MacBut cmdNext 
      Height          =   345
      Left            =   4380
      TabIndex        =   64
      Top             =   5610
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   609
      Caption         =   "Next >>"
      Caption_Xpos    =   400
   End
   Begin wizMacBut.MacBut cmdPrev 
      Height          =   345
      Left            =   2700
      TabIndex        =   65
      Top             =   5610
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   609
      Caption         =   "<< Previous"
      Caption_Xpos    =   100
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2205
      Left            =   2700
      TabIndex        =   41
      Top             =   3330
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   3889
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "[F3 - Add Invoices]"
      TabPicture(0)   =   "CustomerOpeningBalance.frx":08D0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdOpenInv"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "picOpenInv"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstOpenInvoice"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "[F4 - Add Deposits]"
      TabPicture(1)   =   "CustomerOpeningBalance.frx":08EC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstCustDeposit"
      Tab(1).Control(1)=   "cmdCustDeposit"
      Tab(1).Control(2)=   "picCustDeposit"
      Tab(1).ControlCount=   3
      Begin MSComctlLib.ListView lstOpenInvoice 
         Height          =   1755
         Left            =   60
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
         MouseIcon       =   "CustomerOpeningBalance.frx":0908
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
      Begin MSComctlLib.ListView lstCustDeposit 
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
         MouseIcon       =   "CustomerOpeningBalance.frx":0A6A
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "OR Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "OR Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Ref. Inv. No."
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.CommandButton cmdCustDeposit 
         Caption         =   "Command1"
         Height          =   1725
         Left            =   -74610
         TabIndex        =   49
         Top             =   90
         Width           =   5925
      End
      Begin VB.PictureBox picCustDeposit 
         Height          =   1605
         Left            =   -74550
         ScaleHeight     =   1545
         ScaleWidth      =   5745
         TabIndex        =   50
         Top             =   150
         Width           =   5805
         Begin VB.TextBox txtRefInvNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   21
            Text            =   "Text2"
            Top             =   1230
            Width           =   1755
         End
         Begin VB.Frame Frame3 
            Caption         =   "Frame3"
            Height          =   1725
            Left            =   2850
            TabIndex        =   57
            Top             =   -120
            Width           =   30
         End
         Begin VB.TextBox txtCDParticular 
            Appearance      =   0  'Flat
            Height          =   825
            Left            =   2970
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Text            =   "CustomerOpeningBalance.frx":0BCC
            Top             =   330
            Width           =   2745
         End
         Begin VB.TextBox txtOrNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   18
            Text            =   "Text2"
            Top             =   30
            Width           =   1755
         End
         Begin VB.TextBox txtOrDate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   19
            Text            =   "Text2"
            Top             =   330
            Width           =   1755
         End
         Begin VB.TextBox txtCDAmount 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1020
            TabIndex        =   20
            Text            =   "Text2"
            Top             =   630
            Width           =   1755
         End
         Begin VB.Frame Frame4 
            Height          =   30
            Left            =   -90
            TabIndex        =   58
            Top             =   960
            Width           =   2985
         End
         Begin wizButton.cmd cmdCDSave 
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
            MICON           =   "CustomerOpeningBalance.frx":0BD2
         End
         Begin wizButton.cmd cmdCDCancel 
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
            MICON           =   "CustomerOpeningBalance.frx":0BEE
         End
         Begin wizButton.cmd cmdCDDelete 
            Height          =   315
            Left            =   2970
            TabIndex        =   23
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
            MICON           =   "CustomerOpeningBalance.frx":0C0A
         End
         Begin VB.Label Label19 
            Caption         =   "Reference Invoice No."
            Height          =   255
            Left            =   60
            TabIndex        =   59
            Top             =   990
            Width           =   2835
         End
         Begin VB.Label Label18 
            Caption         =   "Particular"
            Height          =   255
            Left            =   2970
            TabIndex        =   56
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label17 
            Caption         =   "OR #"
            Height          =   255
            Left            =   60
            TabIndex        =   53
            Top             =   60
            Width           =   1155
         End
         Begin VB.Label Label15 
            Caption         =   "OR Date"
            Height          =   255
            Left            =   60
            TabIndex        =   52
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label14 
            Caption         =   "Amount"
            Height          =   255
            Left            =   60
            TabIndex        =   51
            Top             =   660
            Width           =   1155
         End
      End
      Begin VB.PictureBox picOpenInv 
         Height          =   1605
         Left            =   450
         ScaleHeight     =   1545
         ScaleWidth      =   5745
         TabIndex        =   43
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
            MICON           =   "CustomerOpeningBalance.frx":0C26
         End
         Begin VB.Frame Frame2 
            Caption         =   "Frame2"
            Height          =   1875
            Left            =   2850
            TabIndex        =   55
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
            Text            =   "CustomerOpeningBalance.frx":0C42
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
            MICON           =   "CustomerOpeningBalance.frx":0C48
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
            MICON           =   "CustomerOpeningBalance.frx":0C64
         End
         Begin VB.Label Label9 
            Caption         =   "Invoice Type"
            Height          =   255
            Left            =   60
            TabIndex        =   66
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label labCDID 
            Caption         =   "Label20"
            Height          =   255
            Left            =   1950
            TabIndex        =   61
            Top             =   600
            Width           =   495
         End
         Begin VB.Label labOIID 
            Caption         =   "Label20"
            Height          =   165
            Left            =   1950
            TabIndex        =   60
            Top             =   600
            Width           =   285
         End
         Begin VB.Label Label16 
            Caption         =   "Particular"
            Height          =   255
            Left            =   2970
            TabIndex        =   54
            Top             =   360
            Width           =   1395
         End
         Begin VB.Label Label13 
            Caption         =   "Amount"
            Height          =   255
            Left            =   2970
            TabIndex        =   48
            Top             =   60
            Width           =   1155
         End
         Begin VB.Label Label12 
            Caption         =   "Invoice Date"
            Height          =   255
            Left            =   60
            TabIndex        =   47
            Top             =   1260
            Width           =   1155
         End
         Begin VB.Label Label11 
            Caption         =   "Dealer"
            Height          =   255
            Left            =   60
            TabIndex        =   46
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label Label10 
            Caption         =   "Invoice #"
            Height          =   255
            Left            =   60
            TabIndex        =   45
            Top             =   660
            Width           =   1155
         End
         Begin VB.Label Label5 
            Caption         =   "Terms"
            Height          =   255
            Left            =   60
            TabIndex        =   44
            Top             =   60
            Width           =   555
         End
      End
      Begin VB.CommandButton cmdOpenInv 
         Caption         =   "Command1"
         Height          =   1725
         Left            =   390
         TabIndex        =   42
         Top             =   90
         Width           =   5925
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   5880
      Left            =   60
      ScaleHeight     =   5820
      ScaleWidth      =   2535
      TabIndex        =   32
      Top             =   90
      Width           =   2595
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   11640
         Left            =   0
         Picture         =   "CustomerOpeningBalance.frx":0C80
         Top             =   0
         Width           =   2550
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   5955
      Left            =   60
      TabIndex        =   34
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
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   180
         Width           =   2475
      End
      Begin MSComctlLib.ListView lstCustomer 
         Height          =   5355
         Left            =   30
         TabIndex        =   36
         Top             =   570
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   9446
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
         MouseIcon       =   "CustomerOpeningBalance.frx":155F8
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
Attribute VB_Name = "frmAMISCustomerOpeningBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCustomer As ADODB.Recordset
Dim rsJOURNAL_HD As ADODB.Recordset
Dim AddOrEditOI, AddOrEditCD As String

Private Sub cmdCDCancel_Click()
cmdCustDeposit.ZOrder 1
picCustDeposit.ZOrder 1
End Sub

'Private Sub cmdCDDelete_Click()
'gconAMIS.Execute ("delete from CustDeposit Where ID = " & labCDID.Caption)
'FillCDGrid
'End Sub

'Private Sub cmdCDSave_Click()
'If AddOrEditCD = "ADD" Then
'   gconAMIS.Execute "Insert Into CustDeposit " & _
'                    "(CustCode,OrNo,ORDate,CDAmount,RefInvNo,CDParticular)" & _
'                    " values (" & N2Str2Null(txtCustCode.Text) & "," & N2Str2Null(txtOrNo.Text) & "," & N2Str2Null(txtOrDate.Text) & "," & N2Str2Zero(txtCDAmount.Text) & "," & N2Str2Null(txtRefInvNo.Text) & "," & N2Str2Null(txtCDParticular.Text) & ")"
'Else
'   gconAMIS.Execute "Update CustDeposit Set " & _
'                    " CustCode = " & N2Str2Null(txtCustCode.Text) & "," & _
'                    " ORNo = " & N2Str2Null(txtOrNo.Text) & "," & _
'                    " ORDate = " & N2Str2Null(txtOrDate.Text) & "," & _
'                    " CDAmount = " & N2Str2Zero(txtCDAmount.Text) & "," & _
'                    " RefInvNo = " & N2Str2Null(txtRefInvNo.Text) & "," & _
'                    " CDParticular = " & N2Str2Null(txtCDParticular.Text) & _
'                    " Where ID = " & labCDID.Caption
'End If
'cmdCDCancel_Click
'FillCDGrid
'End Sub

Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Unload Me
End Sub

Private Sub cmdFind_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
fraDetails.ZOrder 0
On Error Resume Next
TextSearch.SetFocus
End Sub

Private Sub cmdNext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
rsCustomer.MoveNext
If rsCustomer.EOF Then
   rsCustomer.MoveLast
   MsgBox "Last Record!"
End If
StoreMemvars
End Sub

Private Sub cmdOICancel_Click()
cmdOpenInv.ZOrder 1
picOpenInv.ZOrder 1
End Sub

Private Sub cmdOIDelete_Click()
If ShowConfirmDelete = True Then
   gconAmis.Execute ("delete from Journal_HD Where ID = " & labOIID.Caption)
   ShowDeletedMsg
   FillOIGrid
End If
End Sub

Private Sub cmdOISave_Click()
Dim NewVoucherNo As String
Dim NewJNo As String
NewVoucherNo = N2Str2Null(NewOpenInvoiceVoucherNo)
NewJNo = N2Str2Null(NewOpenJNo)
If AddOrEditOI = "ADD" Then
   gconAmis.Execute "Insert Into Journal_HD " & _
                    "(JType,JDate,DueDate,JNo,VoucherNo,CustomerCode,VendorCode,Terms,InvoiceNo,InvoiceType,Dealer,InvoiceDate,InvoiceAmt,AmountPaid,Balance,Remarks,Status,PaidStatus,ReceiveStatus)" & _
                    " values ('CSJ'," & N2Date2Null(LOGDATE) & "," & N2Date2Null(LOGDATE) & "," & NewJNo & "," & NewVoucherNo & "," & N2Str2Null(txtCustCode.Text) & ",'999999'," & N2Str2Null(txtTerms.Text) & "," & N2Str2Null(Format(txtInvoiceNo.Text, "000000")) & "," & N2Str2Null(txtInvoiceType.Text) & "," & N2Str2Null(txtDealer.Text) & "," & N2Str2Null(txtInvoiceDate.Text) & "," & N2Str2Zero(txtOIAmount.Text) & ",0," & N2Str2Zero(txtOIAmount.Text) & "," & N2Str2Null(txtOIParticular.Text) & ",'P','N','N')"
Else
   gconAmis.Execute "Update Journal_HD Set " & _
                    " CustomerCode = " & N2Str2Null(txtCustCode.Text) & "," & _
                    " Terms = " & N2Str2Null(txtTerms.Text) & "," & _
                    " InvoiceType = " & N2Str2Null(txtInvoiceType.Text) & "," & _
                    " InvoiceNo = " & N2Str2Null(Format(txtInvoiceNo.Text, "000000")) & "," & _
                    " Dealer = " & N2Str2Null(txtDealer.Text) & "," & _
                    " InvoiceDate = " & N2Str2Null(txtInvoiceDate.Text) & "," & _
                    " InvoiceAmt = " & N2Str2Zero(txtOIAmount.Text) & "," & _
                    " Remarks = " & N2Str2Null(txtOIParticular.Text) & _
                    " Where ID = " & labOIID.Caption
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

Function NewOpenInvoiceVoucherNo() As String
Dim rsJournal_HDNew As ADODB.Recordset
Set rsJournal_HDNew = New ADODB.Recordset
Set rsJournal_HDNew = gconAmis.Execute("Select * from Journal_HD Where Jtype = 'CSJ' order by VoucherNo Desc")
If Not rsJournal_HDNew.EOF And Not rsJournal_HDNew.BOF Then
   NewOpenInvoiceVoucherNo = Format(N2Str2Zero(rsJournal_HDNew!VoucherNo) + 1, "000000")
Else
   NewOpenInvoiceVoucherNo = "000001"
End If
End Function

Function NewOpenJNo() As String
Dim rsJournal_HDNew As ADODB.Recordset
Set rsJournal_HDNew = New ADODB.Recordset
Set rsJournal_HDNew = gconAmis.Execute("Select * from Journal_HD order by JNo Desc")
If Not rsJournal_HDNew.EOF And Not rsJournal_HDNew.BOF Then
   NewOpenJNo = Format(N2Str2Zero(rsJournal_HDNew!JNo) + 1, "000000")
Else
   NewOpenJNo = "000001"
End If
End Function

Private Sub cmdPrev_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
rsCustomer.MovePrevious
If rsCustomer.BOF Then
   rsCustomer.MoveFirst
   MsgBox "First Record!"
End If
StoreMemvars
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
            AddOrEditCD = "ADD"
            SSTab1.Tab = 1: InitCDMemvars
            cmdCustDeposit.ZOrder 0: picCustDeposit.ZOrder 0
            On Error Resume Next
            txtOrNo.SetFocus
       Case vbKeyEscape
            cmdOpenInv.ZOrder 1: picOpenInv.ZOrder 1
            cmdCustDeposit.ZOrder 1: picCustDeposit.ZOrder 1
       Case Else
            MoveKeyPress KeyCode
End Select
End Sub

Private Sub Form_Load()
CenterMe frmMain, Me, 1
Frame1.Enabled = False: rsRefresh: initMemvars: StoreMemvars
TextSearch.Text = ""
End Sub

Sub rsRefresh()
Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "select * from Customer Order by CustName asc", gconAmis, adOpenKeyset
End Sub

Sub StoreMemvars()
If Not rsCustomer.EOF And Not rsCustomer.BOF Then
   txtCustCode.Text = Null2String(rsCustomer!custcode)
   txtAccountNo.Text = Null2String(rsCustomer!accountno)
   txtCustName.Text = Null2String(rsCustomer!CustName)
   txtDebit.Text = ToDoubleNumber(N2Str2Zero(rsCustomer!Debit))
   txtCredit.Text = ToDoubleNumber(N2Str2Zero(rsCustomer!Credit))
   txtParticular.Text = Null2String(rsCustomer!particular)
   FillOIGrid
   FillCDGrid
End If
End Sub

Function SetCustomerName(CCC As Variant)
Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "Select custcode,custname from Customer where custcode = " & N2Str2Null(CCC), gconAmis, adOpenForwardOnly, adLockReadOnly
If Not rsCustomer.EOF And Not rsCustomer.BOF Then
   SetCustomerName = Null2String(rsCustomer!CustName)
Else
   SetCustomerName = ""
End If
End Function

Sub initMemvars()
txtCustCode.Text = ""
txtAccountNo.Text = ""
txtCustName.Text = ""
txtDebit.Text = ""
txtCredit.Text = ""
txtParticular.Text = ""
End Sub

Sub FillOIGrid()
Dim rsOpenInvoice As ADODB.Recordset
lstOpenInvoice.Sorted = False: lstOpenInvoice.ListItems.Clear
Set rsOpenInvoice = New ADODB.Recordset
Set rsOpenInvoice = gconAmis.Execute("select Terms,InvoiceNo,Dealer,InvoiceDate,InvoiceAmt,ID from Journal_HD where jtype = 'CSJ' and customercode = " & N2Str2Null(txtCustCode.Text) & " order by InvoiceNo asc")
If Not (rsOpenInvoice.EOF And rsOpenInvoice.BOF) Then
   lstOpenInvoice.Enabled = True
   Listview_Loadval Me.lstOpenInvoice.ListItems, rsOpenInvoice
   lstOpenInvoice.Refresh
Else
   lstOpenInvoice.Enabled = False
End If
Set rsOpenInvoice = New ADODB.Recordset
Set rsOpenInvoice = gconAmis.Execute("select SUM(InvoiceAmt) AS TOTALDEBIT from Journal_HD where jtype = 'CSJ' and customercode = " & N2Str2Null(txtCustCode.Text))
If Not (rsOpenInvoice.EOF And rsOpenInvoice.BOF) Then
   txtDebit.Text = ToDoubleNumber(N2Str2Zero(rsOpenInvoice!TotalDebit))
End If
End Sub

Sub StoreOIMemvars(XXX As Variant)
Dim rsOpenInvoice As ADODB.Recordset
Set rsOpenInvoice = New ADODB.Recordset
Set rsOpenInvoice = gconAmis.Execute("Select * from Journal_HD Where ID = " & XXX)
If Not rsOpenInvoice.EOF And Not rsOpenInvoice.BOF Then
   AddOrEditOI = "EDIT"
   SSTab1.Tab = 0
   cmdOpenInv.ZOrder 0: picOpenInv.ZOrder 0
   labOIID.Caption = rsOpenInvoice!ID
   txtTerms.Text = Null2String(rsOpenInvoice!Terms)
   txtInvoiceType.Text = Null2String(rsOpenInvoice!InvoiceType)
   txtInvoiceNo.Text = Null2String(rsOpenInvoice!InvoiceNo)
   txtDealer.Text = Null2String(rsOpenInvoice!Dealer)
   txtInvoiceDate.Text = Null2String(rsOpenInvoice!InvoiceDate)
   txtOIAmount.Text = Null2String(rsOpenInvoice!InvoiceAmt)
   txtOIParticular.Text = Null2String(rsOpenInvoice!Remarks)
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

Sub FillCDGrid()
Dim rsCustDeposit As ADODB.Recordset
lstCustDeposit.Sorted = False: lstCustDeposit.ListItems.Clear
Set rsCustDeposit = New ADODB.Recordset
Set rsCustDeposit = gconAmis.Execute("select OrNo,OrDate,CDAmount,RefInvNo,ID from CustDeposit where custcode = " & N2Str2Null(txtCustCode.Text) & " order by ORNo asc")
If Not (rsCustDeposit.EOF And rsCustDeposit.BOF) Then
   lstCustDeposit.Enabled = True
   Listview_Loadval Me.lstCustDeposit.ListItems, rsCustDeposit
   lstCustDeposit.Refresh
Else
   lstCustDeposit.Enabled = False
End If
End Sub

Sub StoreCDMemvars(XXX As Variant)
Dim rsCustDeposit As ADODB.Recordset
Set rsCustDeposit = New ADODB.Recordset
Set rsCustDeposit = gconAmis.Execute("Select * from CustDeposit Where ID = " & XXX)
If Not rsCustDeposit.EOF And Not rsCustDeposit.BOF Then
   AddOrEditCD = "EDIT"
   SSTab1.Tab = 1
   cmdCustDeposit.ZOrder 0: picCustDeposit.ZOrder 0
   labCDID.Caption = rsCustDeposit!ID
   txtOrNo.Text = Null2String(rsCustDeposit!OrNo)
   txtOrDate.Text = Null2String(rsCustDeposit!OrDate)
   txtCDAmount.Text = Null2String(rsCustDeposit!CDAmount)
   txtRefInvNo.Text = Null2String(rsCustDeposit!RefInvNo)
   txtCDParticular.Text = Null2String(rsCustDeposit!CDParticular)
   On Error Resume Next
   txtOrNo.SetFocus
End If
End Sub

Sub InitCDMemvars()
txtOrNo.Text = ""
txtOrDate.Text = ""
txtCDAmount.Text = ""
txtRefInvNo.Text = ""
txtCDParticular.Text = ""
End Sub

Private Sub lstCustomer_GotFocus()
On Error Resume Next
rsCustomer.Bookmark = rsFind(rsCustomer.Clone, "ID", lstCustomer.SelectedItem.SubItems(1)).Bookmark
StoreMemvars
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
rsCustomer.Bookmark = rsFind(rsCustomer.Clone, "ID", lstCustomer.SelectedItem.SubItems(1)).Bookmark
StoreMemvars
End Sub

Private Sub lstCustDeposit_DblClick()
If Trim(lstCustDeposit.SelectedItem.SubItems(4)) <> "" Then
   StoreCDMemvars (lstCustDeposit.SelectedItem.SubItems(4))
End If
End Sub

Private Sub lstOpenInvoice_DblClick()
If Trim(lstOpenInvoice.SelectedItem.SubItems(5)) <> "" Then
   StoreOIMemvars (lstOpenInvoice.SelectedItem.SubItems(5))
End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then lstCustomer.SetFocus
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

Private Sub txtORDate_GotFocus()
txtOrDate.Text = Format(txtOrDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtORDate_LostFocus()
If txtOrDate.Text <> "" Then
   If IsDate(txtOrDate.Text) = True Then
      txtOrDate.Text = Format(txtOrDate.Text, "DD-MMM-YY")
   Else
      MsgBoxXP "Invalid OR Date!", "Error", XP_OKOnly, msg_Exclamation
      On Error Resume Next
      txtOrDate.SetFocus
   End If
End If
End Sub

Private Sub lstCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then TextSearch.SetFocus
End Sub

Private Sub textSearch_Change()
If Trim(TextSearch.Text) = "" Then FillGrid Else FillSearchGrid (TextSearch.Text)
End Sub

Sub FillGrid()
Dim rsCustomers As ADODB.Recordset
lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
Set rsCustomers = New ADODB.Recordset
Set rsCustomers = gconAmis.Execute("select CustName,ID from Customer where CUSTNAME <> '' order by CustName asc")
If Not (rsCustomers.EOF And rsCustomers.BOF) Then
   Listview_Loadval Me.lstCustomer.ListItems, rsCustomers
   lstCustomer.Refresh
   lstCustomer.Enabled = True
Else
   lstCustomer.Enabled = False
End If
End Sub

Sub FillSearchGrid(XXX As String)
Dim rsCustomers As ADODB.Recordset
lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
Set rsCustomers = New ADODB.Recordset
Set rsCustomers = gconAmis.Execute("select CustName,ID from Customer where CustName like'" & XXX & "%' order by CustName asc")
If Not (rsCustomers.EOF And rsCustomers.BOF) Then
   Listview_Loadval Me.lstCustomer.ListItems, rsCustomers
   lstCustomer.Refresh
   lstCustomer.Enabled = True
Else
   lstCustomer.Enabled = False
End If
End Sub

Private Sub txtTerms_KeyUp(KeyCode As Integer, Shift As Integer)
KeyCode = OnlyNumeric(KeyCode)
End Sub
